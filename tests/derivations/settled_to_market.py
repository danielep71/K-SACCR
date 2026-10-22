"""Expected values of the illustrative settled-to-market cases (#360).

A separate implementation of the SA-CCR formulas these cases use, written
from the cited rules and sharing no code with the VBA engine. Both netting
sets are unmargined, and every trade is settled-to-market (STM): its value
after the day's settlement is zero, and its remaining maturity M is the time
to the next reset date (CRE52.37(5); CRR Art. 279c(2)), floored at 10
business days in the maturity factor (CRE52.48; CRR Art. 279c(1)(a)). S, E
and the option expiry T keep their meaning. The rules are the same under CRR
and BCBS for these trades. It prints the expected outputs; run it from the
repository root with ``python tests/derivations/settled_to_market.py``.

Rules: unmargined RC CRE52.10, CRR Art. 275(1); supervisory duration
CRE52.34, CRR Art. 279b(1)(a); option delta CRE52.40, CRR Art. 279a(1)(a);
interest-rate add-on CRE52.57, CRR Art. 280a; equity add-on CRE52.68 and
CRE52.72, CRR Art. 280d; commodity add-on CRE52.69 and CRE52.73, CRR Art.
280e; multiplier CRE52.23, CRR Art. 278(3); exposure value CRE52.1, CRR Art.
274(2).
"""

from datetime import date
from math import erf, exp, log, sqrt

VALUATION = date(2026, 9, 30)
DAYS_PER_YEAR = 365.0
BUSINESS_DAYS_PER_YEAR = 250.0
FLOOR_BD = 10.0
ALPHA = 1.4
FLOOR = 0.05
SF = {"IR": 0.005, "EQ_INDEX": 0.20, "CO_OIL_GAS": 0.18}
VOL_EQ_INDEX = 0.75
RHO_EQ_INDEX = 0.80
RHO_CO = 0.40

# Daily: an equity index future and call, both reset the next business day.
# Periodic: a five-year EUR swap, an equity index call and a Brent forward,
# reset in 30 days; the forward's reset falls on its maturity date.
CASES = {
    "illustrative-settled-to-market-daily": [
        # id, class, kind, notional, start, end, maturity, expiry, price, strike, reset
        ("T1", "EQ", "linear", 5_000_000.0, None, date(2026, 12, 18), date(2026, 12, 18), None, None, None,
         date(2026, 10, 1)),
        ("T2", "EQ", "call", 2_500_000.0, None, date(2026, 12, 18), date(2026, 12, 18), date(2026, 12, 18),
         5_000.0, 5_200.0, date(2026, 10, 1)),
    ],
    "illustrative-settled-to-market-periodic": [
        ("T1", "IR", "linear", 10_000_000.0, None, date(2031, 9, 30), date(2031, 9, 30), None, None, None,
         date(2026, 10, 30)),
        ("T2", "EQ", "call", 2_500_000.0, None, date(2026, 12, 18), date(2026, 12, 18), date(2026, 12, 18),
         5_000.0, 5_200.0, date(2026, 10, 30)),
        ("T3", "CO", "linear", 1_000_000.0, None, date(2026, 10, 30), date(2026, 10, 30), None, None, None,
         date(2026, 10, 30)),
    ],
}


def phi(x: float) -> float:
    return 0.5 * (1.0 + erf(x / sqrt(2.0)))


def years(d: date) -> float:
    return (d - VALUATION).days / DAYS_PER_YEAR


def remaining_maturity(trade) -> float:
    # STM: the time to the next reset date, which is after the valuation
    # date and not after the contractual maturity.
    reset, maturity = trade[10], trade[6]
    assert VALUATION < reset <= maturity
    return years(reset)


def maturity_factor(trade) -> float:
    m = max(remaining_maturity(trade), FLOOR_BD / BUSINESS_DAYS_PER_YEAR)
    return sqrt(min(m, 1.0))


def delta(trade) -> float:
    kind, expiry, price, strike = trade[2], trade[7], trade[8], trade[9]
    if kind == "linear":
        return 1.0
    t = years(expiry)
    d1 = (log(price / strike) + 0.5 * VOL_EQ_INDEX ** 2 * t) / (VOL_EQ_INDEX * sqrt(t))
    return phi(d1)


def adjusted_notional(trade) -> float:
    if trade[1] != "IR":
        return trade[3]
    s, e = max(years(trade[4]), 0.0) if trade[4] else 0.0, years(trade[5])
    duration = max((exp(-0.05 * s) - exp(-0.05 * e)) / 0.05, FLOOR_BD / BUSINESS_DAYS_PER_YEAR)
    return trade[3] * duration


def bucket(trade) -> str:
    e = years(trade[5])
    return "<1Y" if e < 1 else ("1-5Y" if e <= 5 else ">5Y")


def single_entity(addon: float, rho: float) -> float:
    # One reference entity or commodity type: sqrt((rho A)^2 + (1 - rho^2) A^2) = |A|.
    return sqrt((rho * addon) ** 2 + (1 - rho ** 2) * addon ** 2)


def outputs(case: str) -> dict:
    trades = CASES[case]
    effective = {"IR": 0.0, "EQ": 0.0, "CO": 0.0}
    for t in trades:
        effective[t[1]] += delta(t) * adjusted_notional(t) * maturity_factor(t)
    add_ons = {
        "interest_rate": SF["IR"] * abs(effective["IR"]),
        "equity": single_entity(SF["EQ_INDEX"] * effective["EQ"], RHO_EQ_INDEX),
        "commodity": single_entity(SF["CO_OIL_GAS"] * effective["CO"], RHO_CO),
    }
    aggregate = sum(add_ons.values())
    value = collateral = 0.0
    multiplier = min(1.0, FLOOR + (1 - FLOOR) * exp((value - collateral) / (2 * (1 - FLOOR) * aggregate)))
    rc = max(value - collateral, 0.0)
    result = {"replacement_cost": rc, "aggregate_add_on": aggregate, "multiplier": multiplier,
              "potential_future_exposure": multiplier * aggregate,
              "exposure_value": ALPHA * (rc + multiplier * aggregate)}
    result.update({f"add_on.{k}": v for k, v in add_ons.items() if any(
        t[1] == {"interest_rate": "IR", "equity": "EQ", "commodity": "CO"}[k] for t in trades)})
    for t in trades:
        result[f"remaining_maturity {t[0]}"] = remaining_maturity(t)
        result[f"maturity_factor {t[0]}"] = maturity_factor(t)
        if t[2] == "call":
            result[f"supervisory_delta {t[0]}"] = delta(t)
        if t[1] == "IR":
            result[f"adjusted_notional {t[0]}"] = adjusted_notional(t)
            result[f"bucket {t[0]}"] = bucket(t)
    return result


if __name__ == "__main__":
    for case in CASES:
        print(case)
        for name, result in outputs(case).items():
            print(f"  {name} = {result!r}")
