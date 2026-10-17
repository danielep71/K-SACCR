#!/usr/bin/env python3
"""Cross-check the engine's expected values with the CRAN package SACCR (#352).

The fixtures in tests/fixtures are translated into the package's trades, CSA
and collateral, one counterparty per netting set; tools/rsaccr_run.R runs the
package; each numeric value in tests/expected is then compared with the
package's figure for the same quantity. A fixture the package cannot represent
is listed with the reason and not run.

    python tools/rsaccr_crosscheck.py [--rscript Rscript] [--out DIR] [--json]

Needs R with SACCR 3.4 and Trading 3.2 installed. It is not part of
tools/check.py, which needs only Python. Its output compares two
implementations; it is evidence about agreement, not about correctness.
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
R_DRIVER = Path(__file__).resolve().parent / "rsaccr_run.R"

# Credit quality steps (CRR Art. 280c) mapped to the package's ratings with
# the same supervisory factor, correlation and volatility.
CQS_RATING = {"CQS1": "AA", "CQS2": "A", "CQS3": "BBB", "CQS4": "BB", "CQS5": "B", "CQS6": "CCC"}
INDEX_GRADE = {"IG_INDEX": "IG", "SG_INDEX": "SG"}
# Commodity sub-classes mapped to the package's hedging sets (its SubClass).
COMMODITY_SET = {"ELECTRICITY": "Energy", "OIL_GAS": "Energy", "METALS": "Metals",
                 "AGRICULTURAL": "Agricultural", "OTHER": "Other", "CLIMATIC": "Climatic"}
TRADE_GROUP = {"interest_rate": "IRD", "foreign_exchange": "FX", "credit": "Credit", "equity": "EQ",
               "commodity": "Commodity", "other": "OtherExposure"}
# Expected quantities with a package counterpart, by where the figure is found.
NETTING_SET_QUANTITY = {"exposure_value": "EAD", "replacement_cost": "RC", "aggregate_add_on": "add_on",
                        "potential_future_exposure": "PFE",
                        "multiplier": "multiplier", "margin_period_of_risk": "mpor_business_days"}
CLASS_QUANTITY = {"add_on.interest_rate": "IRD", "add_on.foreign_exchange": "FX", "add_on.credit": "Credit",
                  "add_on.equity": "EQ", "add_on.commodity": "Commodity", "add_on.other": "OtherExposure"}
TRADE_QUANTITY = {"adjusted_notional": "adjusted_notional", "supervisory_delta": "supervisory_delta",
                  "supervisory_factor": "supervisory_factor"}


class Unsupported(Exception):
    """The package cannot represent a trade or netting set as the engine does."""


def years(date: str | None, valuation: dt.date, days_per_year: float) -> float | None:
    """Years from the valuation date, on the DaysPerYear basis, as the engine computes them."""
    if date is None:
        return None
    return (dt.date.fromisoformat(date) - valuation).days / days_per_year


def trade_times(trade: dict[str, Any], valuation: dt.date, days_per_year: float) -> tuple[float, float, float]:
    """S, E and M as the engine fills them: maturity from the end date, then from the expiry."""
    end = years(trade["end_date"], valuation, days_per_year)
    maturity = years(trade["maturity_date"], valuation, days_per_year)
    expiry = years(trade["option_expiry_date"], valuation, days_per_year)
    if maturity is None:
        maturity = end if end is not None else expiry
    if maturity is None:
        raise Unsupported("no maturity, end or expiry date")
    if end is None:
        end = maturity
    start = years(trade["start_date"], valuation, days_per_year)
    return max(0.0, start or 0.0), end, maturity


def option_args(trade: dict[str, Any]) -> dict[str, Any]:
    return {"OptionType": str(trade["option_type"]).capitalize(),
            "UnderlyingPrice": trade["underlying_price"], "StrikePrice": trade["strike_price"]}


def interest_rate(trade: dict[str, Any], start: float, end: float, maturity: float,
                  expiry: float | None) -> tuple[str, dict[str, Any]]:
    # The package has no maturity apart from E, and takes a swaption's expiry as S.
    if abs(maturity - end) > 1e-12:
        raise Unsupported("maturity differs from the end date; the package uses E as the maturity")
    args = {"Currency": trade["risk_factor"], "Si": start, "Ei": end}
    if trade["instrument"] == "linear":
        return "IRDSwap", args
    if expiry is None or abs(expiry - start) > 1e-12:
        raise Unsupported("option expiry differs from S; the package takes a swaption's expiry as S")
    return "IRDSwaption", {**args, **option_args(trade)}


def credit(trade: dict[str, Any], start: float, end: float, maturity: float) -> tuple[str, dict[str, Any]]:
    if abs(maturity - end) > 1e-12:
        raise Unsupported("maturity differs from the end date; the package uses E as the maturity")
    if trade["instrument"] != "linear":
        raise Unsupported("the package has no credit option")
    sub_class = trade["sub_class"]
    args = {"RefEntity": trade["risk_factor"], "Si": start, "Ei": end}
    if sub_class in INDEX_GRADE:
        return "CDX", {**args, "SubClass": INDEX_GRADE[sub_class]}
    return "CDS", {**args, "SubClass": CQS_RATING.get(sub_class, sub_class)}


def commodity(trade: dict[str, Any], maturity: float) -> tuple[str, dict[str, Any]]:
    if trade["instrument"] != "linear":
        raise Unsupported("the package has no commodity option")
    sub_class = trade["sub_class"]
    if sub_class not in COMMODITY_SET:
        raise Unsupported(f"commodity sub-class {sub_class} has no package hedging set")
    # The package gives electricity its factor by commodity type, so an
    # electricity risk factor is entered under that name.
    kind = "Electricity" if sub_class == "ELECTRICITY" else trade["risk_factor"]
    return "Commodity", {"SubClass": COMMODITY_SET[sub_class], "commodity_type": kind, "Si": 0, "Ei": maturity}


def equity(trade: dict[str, Any], maturity: float, expiry: float | None) -> tuple[str, dict[str, Any]]:
    # The package uses E as both an equity option's expiry and its maturity.
    if trade["instrument"] != "linear" and expiry is not None and abs(expiry - maturity) > 1e-12:
        raise Unsupported("option expiry differs from the maturity; the package uses E for both")
    index = trade["sub_class"] == "INDEX"
    args: dict[str, Any] = {"Si": 0, "Ei": maturity}
    args.update({"Underlying_Instrument": trade["risk_factor"], "SubClass": "Index"} if index
                else {"Issuer": trade["risk_factor"]})
    if trade["instrument"] == "linear":
        return "Equity", args
    return ("EquityOptionIndex" if index else "EquityOptionSingle"), {**args, **option_args(trade)}


def translate_trade(trade: dict[str, Any], valuation: dt.date, days_per_year: float,
                    currency: str) -> tuple[str, dict[str, Any]]:
    """The package's trade class and constructor arguments for one fixture trade."""
    if trade["nature"] != "standard":
        raise Unsupported(f"{trade['nature']} transactions are not translated yet")
    start, end, maturity = trade_times(trade, valuation, days_per_year)
    expiry = years(trade["option_expiry_date"], valuation, days_per_year)
    asset_class = trade["asset_class"]
    if asset_class == "interest_rate":
        name, args = interest_rate(trade, start, end, maturity, expiry)
    elif asset_class == "credit":
        name, args = credit(trade, start, end, maturity)
    elif asset_class == "commodity":
        name, args = commodity(trade, maturity)
    elif asset_class == "equity":
        name, args = equity(trade, maturity, expiry)
    elif asset_class == "foreign_exchange" and trade["instrument"] == "linear":
        name, args = "FxForward", {"ccyPair": trade["risk_factor"], "Si": 0, "Ei": maturity}
    elif asset_class == "other" and trade["instrument"] == "linear":
        name, args = "OtherExposure", {"SubClass": trade["risk_factor"], "Si": 0, "Ei": maturity}
    else:
        raise Unsupported(f"{asset_class} {trade['instrument']} has no package class")
    args.setdefault("Currency", currency)
    return name, {**args, "external_id": trade["id"], "Notional": trade["notional_amount"],
                  "MtM": trade["market_value_amount"],
                  "BuySell": "Buy" if trade["direction"] == "long" else "Sell"}


def margin(fixture: dict[str, Any], trades: list[dict[str, Any]]) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    """The CSA and collateral of a margined netting set."""
    ns = fixture["netting_set"]
    if ns["clearing_role"] != "none" or ns["large_or_illiquid"] or ns["margin_disputes"] \
            or ns["mpor_override_business_days"] is not None:
        raise Unsupported("the package's MPOR is 10 + remargining period - 1 business days only")
    # The engine reads a blank remargining period as 1 business day and
    # raises 0 to 1.
    remargin = max(1, ns["remargin_period_business_days"] or 1)
    csa_id = "csa_" + ns["id"]
    csa = {"id": csa_id, "counterparty": ns["id"], "threshold": ns["threshold_amount"] or 0,
           "mta": ns["minimum_transfer_amount"] or 0, "nica": ns["independent_collateral_net_amount"] or 0,
           "remargin_freq": remargin, "mpor_days": 0,
           "currencies": sorted({t["args"]["Currency"] for t in trades}),
           "trade_groups": sorted({TRADE_GROUP[a] for a in fixture_classes(fixture)})}
    # The package applies the threshold and MTA floor only when the CSA has
    # collateral, so the variation margin is always entered, even when 0.
    collateral = [{"id": csa_id + "_vm", "csa_id": csa_id, "type": "VariationMargin",
                   "amount": ns["variation_margin_net_amount"] or 0}]
    if ns["independent_collateral_net_amount"]:
        collateral.append({"id": csa_id + "_ica", "csa_id": csa_id, "type": "ICA",
                           "amount": ns["independent_collateral_net_amount"]})
    return csa, collateral


def fixture_classes(fixture: dict[str, Any]) -> set[str]:
    return {t["asset_class"] for t in fixture["trades"]}


def translate_fixture(fixture: dict[str, Any]) -> dict[str, Any]:
    """The package's netting set for one fixture."""
    valuation = dt.date.fromisoformat(fixture["valuation_date"])
    days_per_year = float(fixture["parameters"].get("DaysPerYear", 365))
    for name, value in fixture["parameters"].items():
        if name != "DaysPerYear":
            raise Unsupported(f"parameter {name}={value} has no package counterpart")
    if fixture["netting_set"]["alpha_factor"] not in (None, 1.4):
        raise Unsupported("the package's alpha is always 1.4")
    ns_id = fixture["netting_set"]["id"]
    trades = []
    for trade in fixture["trades"]:
        _, _, maturity = trade_times(trade, valuation, days_per_year)
        # The engine excludes a matured trade and withholds the netting set's
        # EAD as INCOMPLETE; running the other trades would compare a
        # different netting set.
        if maturity <= 0:
            raise Unsupported(f"{trade['id']} has matured: the engine withholds the netting set's EAD")
        name, args = translate_trade(trade, valuation, days_per_year, fixture["calculation_currency"])
        trades.append({"id": trade["id"], "class": name, "args": {**args, "Counterparty": ns_id}})
    netting_set: dict[str, Any] = {"fixture": fixture["id"], "id": ns_id, "trades": trades,
                                   "csa": None, "collateral": []}
    if fixture["netting_set"]["margined"]:
        netting_set["csa"], netting_set["collateral"] = margin(fixture, trades)
    elif fixture["netting_set"]["variation_margin_net_amount"] or \
            fixture["netting_set"]["independent_collateral_net_amount"]:
        raise Unsupported("collateral on an unmargined netting set has no package counterpart")
    return netting_set


def build_inputs(root: Path) -> dict[str, Any]:
    netting_sets, skipped = [], []
    for path in sorted((root / "tests" / "fixtures").glob("*.json")):
        fixture = json.loads(path.read_text(encoding="utf-8"))
        try:
            netting_set = translate_fixture(fixture)
        except Unsupported as reason:
            skipped.append({"fixture": fixture["id"], "reason": str(reason)})
            continue
        netting_sets.append(netting_set)
    return {"netting_sets": netting_sets, "skipped": skipped}


def package_value(result: dict[str, Any], output: dict[str, Any]) -> float | None:
    quantity = output["quantity"]
    if quantity in NETTING_SET_QUANTITY:
        value = result.get(NETTING_SET_QUANTITY[quantity])
    elif quantity in CLASS_QUANTITY:
        value = result["add_on_by_class"].get(CLASS_QUANTITY[quantity])
    elif quantity in TRADE_QUANTITY and output.get("trade") in result["trades"]:
        value = result["trades"][output["trade"]][TRADE_QUANTITY[quantity]]
    else:
        return None
    return None if value is None else float(value)


def regime_difference(fixture: dict[str, Any]) -> str | None:
    """Why a BCBS expectation cannot be compared with the package, which follows CRR2.

    These are the recorded regime differences (docs/methodology/README.md
    #regime-differences) that a fixture the translator runs can still meet.
    """
    valuation = dt.date.fromisoformat(fixture["valuation_date"])
    days_per_year = float(fixture["parameters"].get("DaysPerYear", 365))
    for trade in fixture["trades"]:
        asset_class, sub_class = trade["asset_class"], trade["sub_class"] or ""
        if asset_class == "other":
            return "other risks exist only under CRR"
        if asset_class == "commodity" and sub_class == "CLIMATIC":
            return "climatic commodities are a hedging set under CRR only"
        if asset_class == "credit" and sub_class.startswith("CQS"):
            return "credit quality steps are the CRR scale"
        if trade["instrument"] == "option" and asset_class in ("interest_rate", "commodity") \
                and min(trade["underlying_price"], trade["strike_price"]) < 0.001:
            return "the CRR lambda shift changes the delta of an option priced near or below zero"
        _, end, _ = trade_times(trade, valuation, days_per_year)
        if asset_class == "interest_rate" and abs(end - 1) < 1e-12:
            return "an end date of exactly 1 year is in bucket 1 under CRR and bucket 2 under BCBS"
    return None


def compare(root: Path, results: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[dict[str, str]]]:
    """One row per numeric expected value that the package also computes, and the
    BCBS expected files left out because their rules differ from CRR2."""
    by_fixture = {r["fixture"]: r for r in results}
    rows, left_out = [], []
    for path in sorted((root / "tests" / "expected").glob("*.json")):
        expected = json.loads(path.read_text(encoding="utf-8"))
        result = by_fixture.get(expected["fixture"])
        if result is None:
            continue
        if expected.get("regime") == "BCBS":
            fixture = json.loads((root / "tests" / "fixtures" / f"{expected['fixture']}.json")
                                 .read_text(encoding="utf-8"))
            reason = regime_difference(fixture)
            if reason is not None:
                left_out.append({"case": path.stem, "reason": reason})
                continue
        for output in expected["outputs"]:
            if not isinstance(output.get("value"), (int, float)) or isinstance(output["value"], bool):
                continue
            value = package_value(result, output)
            if value is None:
                continue
            tolerance = output["tolerance"]
            difference = value - output["value"]
            allowed = max(tolerance["absolute"], tolerance["relative"] * abs(output["value"]))
            rows.append({"case": path.stem, "quantity": output["quantity"], "trade": output.get("trade", ""),
                         "expected": output["value"], "package": value, "difference": difference,
                         "agrees": abs(difference) <= allowed, "reference": output["reference"]["class"]})
    return rows, left_out


def run_package(rscript: str, inputs: Path, results: Path) -> dict[str, Any]:
    subprocess.run([rscript, str(R_DRIVER), str(inputs), str(results)], check=True)
    loaded: dict[str, Any] = json.loads(results.read_text(encoding="utf-8"))
    return loaded


def report(inputs: dict[str, Any], package: dict[str, Any], rows: list[dict[str, Any]],
           left_out: list[dict[str, str]]) -> str:
    lines = ["Versions: " + ", ".join(f"{k} {v}" for k, v in package["versions"].items()), ""]
    lines += [f"{'case':42} {'quantity':22} {'trade':9} {'expected':>18} {'package':>18} {'difference':>14}  result"]
    for row in rows:
        lines.append(f"{row['case']:42} {row['quantity']:22} {row['trade']:9} {row['expected']:18.10g} "
                     f"{row['package']:18.10g} {row['difference']:14.6g}  "
                     f"{'agrees' if row['agrees'] else 'DIFFERS'} ({row['reference']})")
    compared = {row["case"].rsplit(".", 1)[0] for row in rows}
    lines += ["", "Run, with no expected figure to compare (package EAD, RC and add-on):"]
    lines += [f"  {r['fixture']}: EAD={r['EAD']:.10g}; RC={r['RC']:.10g}; add_on={r['add_on']:.10g}"
              for r in package["results"] if r["fixture"] not in compared]
    lines += ["", "BCBS expectations not compared (the package follows CRR2):"]
    lines += [f"  {c['case']}: {c['reason']}" for c in left_out]
    lines += ["", "Not run:"] + [f"  {s['fixture']}: {s['reason']}" for s in inputs["skipped"]]
    differs = sum(1 for row in rows if not row["agrees"])
    lines += ["", f"RESULT compared={len(rows)}; differ={differs}; fixtures_run={len(inputs['netting_sets'])}; "
              f"fixtures_not_run={len(inputs['skipped'])}; bcbs_not_compared={len(left_out)}"]
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--rscript", default="Rscript")
    parser.add_argument("--out", type=Path, help="keep the inputs and results here")
    parser.add_argument("--json", action="store_true", help="print the comparison rows as JSON")
    options = parser.parse_args(argv)
    inputs = build_inputs(options.root)
    with tempfile.TemporaryDirectory() as temp:
        out = options.out or Path(temp)
        out.mkdir(parents=True, exist_ok=True)
        (out / "rsaccr-inputs.json").write_text(json.dumps(inputs, indent=2), encoding="utf-8")
        package = run_package(options.rscript, out / "rsaccr-inputs.json", out / "rsaccr-results.json")
    rows, left_out = compare(options.root, package["results"])
    print(json.dumps(rows, indent=2) if options.json else report(inputs, package, rows, left_out))
    return 0


if __name__ == "__main__":
    sys.exit(main())
