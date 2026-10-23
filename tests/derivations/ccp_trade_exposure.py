"""Expected values of the illustrative CCP trade-exposure cases (#112).

A separate implementation of the SA-CCR formulas these cases use, written
from the cited rules and sharing no code with the VBA engine. Each case is
one margined netting set in which the bank, as clearing member, faces a CCP,
remargined daily, without collateral, threshold or MTA; the trades are on
one equity index. The rules are the same under CRR and BCBS for these
trades, and QCCP status changes the risk weight only, not the exposure
value. It prints the expected outputs; run it from the repository root with
``python tests/derivations/ccp_trade_exposure.py``.

Rules: MPOR floor of a trade exposure to a CCP, 10 business days (CRE54.8(2);
CRR Art. 304(3)(b)), plus N - 1 for remargining every N days (CRE52.50; CRR
Art. 285(5)); with VM retained by the CCP and collateral not protected, each
trade's horizon is the lower of one year and its remaining maturity, floored
at 10 business days (CRE54.8(3); CRR Art. 304(3)(d)), applied where it
exceeds the MPOR; margined maturity factor 1.5 sqrt(horizon / 250) (CRE52.52;
CRR Art. 279c(1)(b)); unmargined maturity factor CRE52.48, CRR Art.
279c(1)(a); option delta CRE52.40, CRR Art. 279a(1)(a); equity add-on
CRE52.68, CRE52.72, CRR Art. 280d; margined RC CRE52.18, CRR Art. 275(2);
multiplier CRE52.23, CRR Art. 278(3); exposure value and cap CRE52.1-.2, CRR
Art. 274(2)-(3), the cap's collateral being NICA.
"""

from datetime import date
from math import erf, exp, log, sqrt

VALUATION = date(2026, 9, 30)
DAYS_PER_YEAR = 365.0
BUSINESS_DAYS_PER_YEAR = 250.0
FLOOR_BD = 10.0
CCP_FLOOR_BD = 10.0
REMARGIN_BD = 1.0
ALPHA = 1.4
FLOOR = 0.05
SF_EQ_INDEX = 0.20
VOL_EQ_INDEX = 0.75
RHO_EQ_INDEX = 0.80
THRESHOLD = MTA = NICA = VM = 0.0

CASES = {
    # A listed equity index future and call, as in the clearing-member case
    # of #359, now facing the CCP: MPOR 10 instead of 5.
    "illustrative-ccp-trade-exposure": {
        "retained_vm": False,
        "trades": [
            # id, kind, direction, notional, market value, maturity, expiry, price, strike
            ("T1", "linear", +1, 5_000_000.0, 20_000.0, date(2026, 12, 18), None, None, None),
            ("T2", "call", +1, 2_500_000.0, 60_000.0, date(2026, 12, 18), date(2026, 12, 18), 5_000.0, 5_200.0),
        ],
    },
    # Equity forwards against which the CCP retains VM, with the bank's
    # collateral not protected: maturities of 7 days (horizon at the floor,
    # not above the MPOR), 182 days (between the MPOR and a year) and two
    # years (capped at a year).
    "illustrative-ccp-retained-vm": {
        "retained_vm": True,
        "trades": [
            ("T1", "linear", +1, 4_000_000.0, 10_000.0, date(2026, 10, 7), None, None, None),
            ("T2", "linear", -1, 3_000_000.0, -15_000.0, date(2027, 3, 31), None, None, None),
            ("T3", "linear", +1, 2_000_000.0, 25_000.0, date(2028, 9, 29), None, None, None),
        ],
    },
}
# The trades of illustrative-ccp-trade-exposure facing a non-qualifying CCP:
# QCCP status changes the risk weight only, so every output is the same.
CASES["illustrative-ccp-non-qualifying"] = CASES["illustrative-ccp-trade-exposure"]


def phi(x: float) -> float:
    return 0.5 * (1.0 + erf(x / sqrt(2.0)))


def years(d: date) -> float:
    return (d - VALUATION).days / DAYS_PER_YEAR


def delta(trade) -> float:
    kind, direction, expiry, price, strike = trade[1], trade[2], trade[6], trade[7], trade[8]
    if kind == "linear":
        return float(direction)
    t = years(expiry)
    d1 = (log(price / strike) + 0.5 * VOL_EQ_INDEX ** 2 * t) / (VOL_EQ_INDEX * sqrt(t))
    return direction * phi(d1)


def mpor() -> float:
    return CCP_FLOOR_BD + REMARGIN_BD - 1


def horizon(trade, retained_vm: bool) -> float:
    """Business days of the trade's margined maturity factor."""
    if not retained_vm:
        return mpor()
    own = max(FLOOR_BD, min(BUSINESS_DAYS_PER_YEAR, years(trade[5]) * BUSINESS_DAYS_PER_YEAR))
    return max(mpor(), own)


def mf_margined(trade, retained_vm: bool) -> float:
    return 1.5 * sqrt(horizon(trade, retained_vm) / BUSINESS_DAYS_PER_YEAR)


def mf_unmargined(trade) -> float:
    return sqrt(min(max(years(trade[5]), FLOOR_BD / BUSINESS_DAYS_PER_YEAR), 1.0))


def add_on(trades, maturity_factor) -> float:
    effective = sum(delta(t) * t[3] * maturity_factor(t) for t in trades)
    entity = SF_EQ_INDEX * effective
    # One reference entity: sqrt((rho A)^2 + (1 - rho^2) A^2) = |A|.
    return sqrt((RHO_EQ_INDEX * entity) ** 2 + (1 - RHO_EQ_INDEX ** 2) * entity ** 2)


def outputs(case: str) -> dict:
    spec = CASES[case]
    trades, retained = spec["trades"], spec["retained_vm"]
    value = sum(t[4] for t in trades)
    collateral = VM + NICA

    def pfe(addon: float) -> tuple[float, float]:
        multiplier = min(1.0, FLOOR + (1 - FLOOR) * exp((value - collateral) / (2 * (1 - FLOOR) * addon)))
        return multiplier, multiplier * addon

    margined_add_on = add_on(trades, lambda t: mf_margined(t, retained))
    unmargined_add_on = add_on(trades, mf_unmargined)
    rc = max(value - collateral, THRESHOLD + MTA - NICA, 0.0)
    multiplier, pfe_margined = pfe(margined_add_on)
    ead_margined = ALPHA * (rc + pfe_margined)
    _, pfe_unmargined = pfe(unmargined_add_on)
    ead_cap = ALPHA * (max(value - NICA, 0.0) + pfe_unmargined)
    result = {
        "margin_period_of_risk": mpor(),
        "replacement_cost": rc,
        "add_on.equity": margined_add_on,
        "multiplier": multiplier,
        "potential_future_exposure": pfe_margined,
        "exposure_value": min(ead_margined, ead_cap),
        "cap_applied": "Y" if ead_cap < ead_margined else "N",
    }
    for t in trades:
        result[f"maturity_factor_margined {t[0]}"] = mf_margined(t, retained)
        if t[1] == "call":
            result[f"supervisory_delta {t[0]}"] = delta(t)
    return result


if __name__ == "__main__":
    for case in CASES:
        print(case)
        for name, result in outputs(case).items():
            print(f"  {name} = {result!r}")
