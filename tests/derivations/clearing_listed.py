"""Expected values of the illustrative client-clearing cases (#359).

A separate implementation of the SA-CCR formulas these cases use, written
from the cited rules and sharing no code with the VBA engine: listed equity
index future and call option in one margined netting set, for a clearing
member facing its client and for a client facing its clearing member, under
CRR and BCBS. It prints the expected outputs; run it from the repository root
with ``python tests/derivations/clearing_listed.py``.

Rules: margined RC CRE52.18, CRR Art. 275(2); maturity factor CRE52.52,
CRR Art. 279c(1)(b); MPOR floors CRE54.12 and CRE52.50, CRR Art. 279c(1) and
285(2); option delta CRE52.40, CRR Art. 279a(1)(a); equity add-on CRE52.68,
CRE52.72, CRR Art. 280d; multiplier CRE52.23, CRR Art. 278(3); exposure value
CRE52.1, CRR Art. 274(2); cap CRE52.2, CRR Art. 274(3).
"""

from datetime import date
from math import erf, exp, log, sqrt

VALUATION = date(2026, 9, 30)
DAYS_PER_YEAR = 365.0
BUSINESS_DAYS_PER_YEAR = 250.0
ALPHA = 1.4
FLOOR = 0.05
SF_EQ_INDEX = 0.20
VOL_EQ_INDEX = 0.75
RHO_EQ_INDEX = 0.80

# NS1: margined, no threshold, MTA, VM or NICA. The clearing member's case
# remargins daily, the client's weekly (N = 5 business days).
THRESHOLD = MTA = NICA = VM = 0.0
REMARGIN = {"clearing_member": 1, "client": 5}
TRADES = [
    # id, instrument, notional (current value of the underlying), market value, maturity, expiry, price, strike
    ("T1", "future", 5_000_000.0, 20_000.0, date(2026, 12, 18), None, None, None),
    ("T2", "call", 2_500_000.0, 60_000.0, date(2026, 12, 18), date(2026, 12, 18), 5_000.0, 5_200.0),
]
# MPOR floor F by role and regime: CRE54.12 and CRR Art. 279c(1) give the
# clearing member 5 days in both regimes and the client 5 under CRR; a BCBS
# client keeps 10 (CRE54.8(2), CRE54.14-.17, CRE52.50). MPOR = F + N - 1
# (CRR Art. 285(5); CRE52.50(2)).
FLOOR_BD = {("clearing_member", "CRR"): 5, ("clearing_member", "BCBS"): 5,
            ("client", "CRR"): 5, ("client", "BCBS"): 10}


def phi(x: float) -> float:
    return 0.5 * (1.0 + erf(x / sqrt(2.0)))


def years(d: date) -> float:
    return (d - VALUATION).days / DAYS_PER_YEAR


def delta(trade) -> float:
    _, kind, _, _, _, expiry, price, strike = trade
    if kind == "future":
        return 1.0
    t = years(expiry)
    d1 = (log(price / strike) + 0.5 * VOL_EQ_INDEX ** 2 * t) / (VOL_EQ_INDEX * sqrt(t))
    return phi(d1)


def add_on(maturity_factor) -> float:
    effective = sum(delta(t) * t[2] * maturity_factor(t) for t in TRADES)
    entity = SF_EQ_INDEX * effective
    # One reference entity: sqrt((rho A)^2 + (1 - rho^2) A^2) = |A|.
    return sqrt((RHO_EQ_INDEX * entity) ** 2 + (1 - RHO_EQ_INDEX ** 2) * entity ** 2)


def outputs(role: str, regime: str) -> dict:
    mpor = FLOOR_BD[(role, regime)] + REMARGIN[role] - 1
    value = sum(t[3] for t in TRADES)
    collateral = VM + NICA
    margined_add_on = add_on(lambda t: 1.5 * sqrt(mpor / BUSINESS_DAYS_PER_YEAR))
    unmargined_add_on = add_on(
        lambda t: sqrt(min(max(years(t[4]), 10 / BUSINESS_DAYS_PER_YEAR), 1.0)))
    def pfe(addon: float) -> tuple[float, float]:
        multiplier = min(1.0, FLOOR + (1 - FLOOR) * exp((value - collateral) / (2 * (1 - FLOOR) * addon)))
        return multiplier, multiplier * addon
    rc_margined = max(value - collateral, THRESHOLD + MTA - NICA, 0.0)
    multiplier, pfe_margined = pfe(margined_add_on)
    ead_margined = ALPHA * (rc_margined + pfe_margined)
    _, pfe_unmargined = pfe(unmargined_add_on)
    # The cap's collateral is NICA (CRE52.10; EBA-QA-2023_6962).
    ead_cap = ALPHA * (max(value - NICA, 0.0) + pfe_unmargined)
    return {
        "margin_period_of_risk": mpor,
        "replacement_cost": rc_margined,
        "add_on.equity": margined_add_on,
        "multiplier": multiplier,
        "potential_future_exposure": pfe_margined,
        "exposure_value": min(ead_margined, ead_cap),
        "cap_applied": "Y" if ead_cap < ead_margined else "N",
        "supervisory_delta T2": delta(TRADES[1]),
    }


if __name__ == "__main__":
    for role in ("clearing_member", "client"):
        for regime in ("CRR", "BCBS"):
            print(role, regime)
            for name, result in outputs(role, regime).items():
                print(f"  {name} = {result!r}")
