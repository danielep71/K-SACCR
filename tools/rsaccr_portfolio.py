#!/usr/bin/env python3
"""Compare a workbook run of a synthetic portfolio with the CRAN package SACCR (#352).

Three steps, each a command:

    python tools/rsaccr_portfolio.py generate --out DIR [--seed 352] [--netting-sets 40]
    python tools/rsaccr_portfolio.py run --out DIR [--rscript Rscript]
    python tools/rsaccr_portfolio.py compare --out DIR --evidence EVIDENCE.csv

``generate`` writes a deterministic synthetic portfolio: ``portfolio.json`` in
the fixture format of tests/fixtures, one entry per netting set, and the two
CSV files of the input contract, ``K-SACCR_NettingSets_v1.csv`` and
``K-SACCR_Trades_v1.csv``, which the workbook imports. Every trade is one the
package represents as the engine does (tools/rsaccr_crosscheck.py), and every
amount is in EUR, so no FX conversion enters. ``run`` translates the portfolio
with rsaccr_crosscheck.translate_fixture and runs tools/rsaccr_run.R.
``compare`` reads the evidence CSV that Export evidence writes after the
workbook's run of the same files, and compares each netting set's Results with
the package: RC, the add-on of each asset class, the aggregate add-on, the
multiplier, PFE and EAD, with the full-precision tolerance of
docs/methodology/TEST_CASES.md. It writes ``comparison.csv`` with an empty
classification column for each difference.

Agreement between two implementations is evidence about agreement, not about
correctness; this is not a gate of tools/check.py and needs R.
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import io
import json
import math
import random
import sys
from decimal import Decimal
from pathlib import Path
from typing import Any

import csv_contract
import rsaccr_crosscheck as crosscheck

ROOT = Path(__file__).resolve().parent.parent
VALUATION = dt.date(2026, 9, 30)
CURRENCY = "EUR"
REGIME = "CRR"
NS_FILE = "K-SACCR_NettingSets_v1.csv"
TR_FILE = "K-SACCR_Trades_v1.csv"
# Full precision, docs/methodology/TEST_CASES.md: currency amounts and factors.
AMOUNT_TOLERANCE = (1e-6, 1e-9)
FACTOR_TOLERANCE = (1e-9, 1e-9)

ASSET_CODE = {"interest_rate": "IR", "foreign_exchange": "FX", "credit": "CR", "equity": "EQ",
              "commodity": "CO", "other": "OT"}
WORDS = {"linear": "Linear", "option": "Option", "long": "Long", "short": "Short", "call": "Call", "put": "Put",
         "standard": "Standard"}
IR_CURRENCIES = ("EUR", "USD", "GBP", "JPY")
FX_PAIRS = ("EUR/USD", "USD/JPY", "GBP/USD", "EUR/GBP")
# One sub-class per reference across the portfolio, as the engine requires per netting set.
CREDIT_NAMES = {"CORP-A": "CQS1", "CORP-B": "CQS2", "CORP-C": "CQS3", "CORP-D": "CQS4", "CORP-E": "CQS5",
                "CORP-F": "CQS6", "ITRAXX-MAIN": "IG_INDEX", "CDX-HY": "SG_INDEX"}
EQUITY_NAMES = {"STOCK-1": "SINGLE", "STOCK-2": "SINGLE", "STOCK-3": "SINGLE", "STOCK-4": "SINGLE",
                "EUROSTOXX50": "INDEX", "SP500": "INDEX"}
COMMODITY_NAMES = {"BRENT": "OIL_GAS", "WTI": "OIL_GAS", "GOLD": "METALS", "COPPER": "METALS",
                   "WHEAT": "AGRICULTURAL", "CORN": "AGRICULTURAL", "TIMBER": "OTHER"}
OTHER_NAMES = ("LONGEVITY", "INFLATION-VOL")
CLASS_WEIGHTS = (("interest_rate", 40), ("foreign_exchange", 15), ("credit", 15), ("equity", 15),
                 ("commodity", 10), ("other", 5))
# Results columns compared, with the package figure and the tolerance.
RESULT_FIGURES = (("RC", "RC", AMOUNT_TOLERANCE), ("AddOn aggregate", "add_on", AMOUNT_TOLERANCE),
                  ("Multiplier", "multiplier", FACTOR_TOLERANCE), ("PFE", "PFE", AMOUNT_TOLERANCE),
                  ("EAD", "EAD", AMOUNT_TOLERANCE))
CLASS_FIGURES = (("AddOn IR", "IRD"), ("AddOn FX", "FX"), ("AddOn CR", "Credit"), ("AddOn EQ", "EQ"),
                 ("AddOn CO", "Commodity"), ("AddOn OT", "OtherExposure"))


class EvidenceError(Exception):
    """The evidence file is not the report of a run of this portfolio."""


# ---------------------------------------------------------------------------
# Generation
# ---------------------------------------------------------------------------

def iso(days: int) -> str:
    return (VALUATION + dt.timedelta(days=days)).isoformat()


def amount(value: float) -> float:
    """An amount rounded to the cent, as a user would enter it."""
    return float(round(value, 2))


def tenor_days(rng: random.Random, low_years: float, high_years: float) -> int:
    """Days to a date between the bounds, never exactly one year and never under 30 days, where the
    package's unmargined maturity-factor floor differs from the engine's."""
    days = max(30, int(rng.uniform(low_years, high_years) * 365))
    return days + 1 if days == 365 else days


def blank_trade(trade_id: str, asset_class: str, rng: random.Random) -> dict[str, Any]:
    notional = float(rng.randrange(1_000, 50_000) * 1_000)
    return {"id": trade_id, "asset_class": asset_class, "sub_class": None, "risk_factor": None,
            "instrument": "linear", "direction": rng.choice(("long", "short")), "option_type": None,
            "nature": "standard", "hedging_set_label": None, "notional_amount": notional,
            "market_value_amount": amount(notional * rng.uniform(-0.03, 0.03)),
            "start_date": None, "end_date": None, "maturity_date": None, "option_expiry_date": None,
            "underlying_price": None, "strike_price": None, "lambda_price": None, "attachment_rate": None,
            "detachment_rate": None, "settlement_type": "ordinary", "next_reset_date": None,
            "description": None}


def make_option(trade: dict[str, Any], rng: random.Random, price: float, strike: float) -> None:
    trade.update(instrument="option", option_type=rng.choice(("call", "put")),
                 underlying_price=price, strike_price=strike)


def interest_rate_trade(trade: dict[str, Any], rng: random.Random, linear: bool) -> None:
    trade["risk_factor"] = rng.choice(IR_CURRENCIES)
    start = 0 if linear and rng.random() < 0.7 else tenor_days(rng, 0.1, 2.0)
    end = start + tenor_days(rng, 1.5, 15.0)
    trade.update(start_date=iso(start) if start else None, end_date=iso(end), maturity_date=iso(end))
    if not linear:
        # The package takes a swaption's expiry as its start, S.
        trade["option_expiry_date"] = iso(start)
        make_option(trade, rng, round(rng.uniform(0.01, 0.05), 4), round(rng.uniform(0.01, 0.05), 4))


def foreign_exchange_trade(trade: dict[str, Any], rng: random.Random) -> None:
    trade.update(risk_factor=rng.choice(FX_PAIRS), maturity_date=iso(tenor_days(rng, 0.1, 5.0)))


def credit_trade(trade: dict[str, Any], rng: random.Random) -> None:
    name = rng.choice(sorted(CREDIT_NAMES))
    end = tenor_days(rng, 1.0, 10.0)
    trade.update(risk_factor=name, sub_class=CREDIT_NAMES[name], end_date=iso(end), maturity_date=iso(end))


def equity_trade(trade: dict[str, Any], rng: random.Random, linear: bool) -> None:
    name = rng.choice(sorted(EQUITY_NAMES))
    maturity = tenor_days(rng, 0.2, 3.0)
    trade.update(risk_factor=name, sub_class=EQUITY_NAMES[name], maturity_date=iso(maturity))
    if not linear:
        # The package uses E for both an equity option's expiry and its maturity.
        trade["option_expiry_date"] = iso(maturity)
        make_option(trade, rng, 100.0, float(rng.randrange(80, 121)))


def commodity_trade(trade: dict[str, Any], rng: random.Random, used: dict[str, int]) -> None:
    # At most two trades on one commodity: with three or more, the package's
    # add-on can depend on their order (tools/README.md).
    name = rng.choice([n for n in sorted(COMMODITY_NAMES) if used.get(n, 0) < 2])
    used[name] = used.get(name, 0) + 1
    trade.update(risk_factor=name, sub_class=COMMODITY_NAMES[name], maturity_date=iso(tenor_days(rng, 0.2, 5.0)))


def make_trade(trade_id: str, rng: random.Random, linear_only: bool, used: dict[str, int]) -> dict[str, Any]:
    asset_class = rng.choices([c for c, _ in CLASS_WEIGHTS], [w for _, w in CLASS_WEIGHTS])[0]
    trade = blank_trade(trade_id, asset_class, rng)
    linear = linear_only or rng.random() < 0.75
    if asset_class == "interest_rate":
        interest_rate_trade(trade, rng, linear)
    elif asset_class == "foreign_exchange":
        foreign_exchange_trade(trade, rng)
    elif asset_class == "credit":
        credit_trade(trade, rng)
    elif asset_class == "equity":
        equity_trade(trade, rng, linear)
    elif asset_class == "commodity":
        commodity_trade(trade, rng, used)
    else:
        trade.update(risk_factor=rng.choice(OTHER_NAMES), maturity_date=iso(tenor_days(rng, 0.5, 5.0)))
    return trade


def make_netting_set(index: int, rng: random.Random) -> dict[str, Any]:
    ns_id = f"RS-{index:03d}"
    margined = rng.random() < 0.3
    netting_set = {"id": ns_id, "margined": margined, "clearing_role": "none",
                   "remargin_period_business_days": rng.choice((1, 5, 10)) if margined else None,
                   "large_or_illiquid": False, "margin_disputes": False, "mpor_override_business_days": None,
                   "variation_margin_net_amount": amount(rng.uniform(-2e5, 2e5)) if margined else 0,
                   "independent_collateral_net_amount": amount(rng.uniform(0, 5e5)) if margined else 0,
                   "threshold_amount": float(rng.choice((0, 0, 100_000, 250_000))) if margined else 0,
                   "minimum_transfer_amount": float(rng.choice((0, 10_000, 50_000))) if margined else 0,
                   "alpha_factor": None, "qccp": None, "ccp_retains_variation_margin": None}
    # The first trade is linear, so a netting set is never only sold options,
    # which the package sets to zero and the engine does not.
    used: dict[str, int] = {}
    trades = [make_trade(f"{ns_id}-T{t:02d}", rng, t == 1, used) for t in range(1, rng.randint(3, 15) + 1)]
    return {"schema_version": 1, "kind": "saccr-fixture", "id": ns_id, "synthetic": True,
            "description": f"Synthetic netting set {index} of the #352 portfolio", "category": "nominal",
            "calculation_currency": CURRENCY, "valuation_date": VALUATION.isoformat(), "parameters": {},
            "netting_set": netting_set, "trades": trades}


def generate(seed: int, count: int) -> dict[str, Any]:
    """The portfolio: ``count`` netting sets drawn from ``seed``, each one the package can run."""
    rng = random.Random(seed)
    netting_sets = [make_netting_set(i, rng) for i in range(1, count + 1)]
    for fixture in netting_sets:
        crosscheck.translate_fixture(fixture)  # Raises Unsupported if the generator strays.
    return {"seed": seed, "valuation_date": VALUATION.isoformat(), "reporting_currency": CURRENCY,
            "regime": REGIME, "netting_sets": netting_sets}


# ---------------------------------------------------------------------------
# CSV files of the input contract
# ---------------------------------------------------------------------------

def field(value: Any) -> str:
    """A CSV field as the contract writes it: numbers in plain decimal notation, flags as Y or N."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "Y" if value else "N"
    if isinstance(value, (int, float)):
        text = format(Decimal(repr(value)), "f")
        return text.rstrip("0").rstrip(".") if "." in text else text
    return str(value)


def netting_set_row(fixture: dict[str, Any]) -> dict[str, Any]:
    ns = fixture["netting_set"]
    return {"netting_set_id": ns["id"], "counterparty": f"Synthetic counterparty {ns['id']}",
            "margined": ns["margined"], "clearing_role": "N",
            "remargin_period_business_days": ns["remargin_period_business_days"],
            "large_or_illiquid": False, "margin_disputes": False,
            "variation_margin_net_amount": ns["variation_margin_net_amount"],
            "independent_collateral_net_amount": ns["independent_collateral_net_amount"],
            "threshold_amount": ns["threshold_amount"], "minimum_transfer_amount": ns["minimum_transfer_amount"],
            "regime_override": REGIME}


def trade_row(ns_id: str, trade: dict[str, Any]) -> dict[str, Any]:
    return {"trade_id": trade["id"], "netting_set_id": ns_id, "asset_class": ASSET_CODE[trade["asset_class"]],
            "sub_class": trade["sub_class"], "risk_factor": trade["risk_factor"],
            "instrument": WORDS[trade["instrument"]], "direction": WORDS[trade["direction"]],
            "option_type": WORDS[trade["option_type"]] if trade["option_type"] else None,
            "nature": WORDS[trade["nature"]], "notional_amount": trade["notional_amount"],
            "notional_currency": CURRENCY, "market_value_amount": trade["market_value_amount"],
            "market_value_currency": CURRENCY, "start_date": trade["start_date"], "end_date": trade["end_date"],
            "maturity_date": trade["maturity_date"], "option_expiry_date": trade["option_expiry_date"],
            "underlying_price": trade["underlying_price"], "strike_price": trade["strike_price"]}


def csv_text(kind: str, rows: list[dict[str, Any]], contract: dict[str, Any]) -> str:
    names = [column["name"] for column in contract["files"][kind]["columns"]]
    lines = [";".join(names)] + [";".join(field(row.get(name)) for name in names) for row in rows]
    return "\n".join(lines) + "\n"


def csv_files(portfolio: dict[str, Any], contract: dict[str, Any]) -> dict[str, str]:
    """The two input files, checked with the reference validator of the contract."""
    fixtures = portfolio["netting_sets"]
    files = {NS_FILE: csv_text("NettingSets", [netting_set_row(f) for f in fixtures], contract),
             TR_FILE: csv_text("Trades", [trade_row(f["id"], t) for f in fixtures for t in f["trades"]], contract)}
    findings = []
    records: dict[str, list[dict[str, str]]] = {}
    for name, kind in ((NS_FILE, "NettingSets"), (TR_FILE, "Trades")):
        records[kind], found = csv_contract.validate(files[name].encode("utf-8"), kind, contract)
        findings += [f"{name}: {f}" for f in found]
    trades = [(line, row) for line, row in enumerate(records["Trades"], start=2)]
    findings += [f"{TR_FILE}: {f}" for f in csv_contract.check_portfolio(
        trades, [row["netting_set_id"] for row in records["NettingSets"]], contract)]
    if findings:
        raise ValueError("generated files break the input contract: " + "; ".join(findings))
    return files


# ---------------------------------------------------------------------------
# Comparison with the workbook's evidence
# ---------------------------------------------------------------------------

def read_evidence(text: str) -> tuple[dict[str, str], dict[str, dict[str, str]]]:
    """The RUN entries and the RESULTS rows by netting set of an evidence CSV."""
    run: dict[str, str] = {}
    header: list[str] | None = None
    results: dict[str, dict[str, str]] = {}
    for record in csv.reader(io.StringIO(text), delimiter=";"):
        if not record:
            continue
        if record[0] == "RUN" and len(record) >= 3:
            run[record[1]] = record[2]
        elif record[0] == "RESULTS" and header is None:
            header = record[1:]
        elif record[0] == "RESULTS" and header is not None:
            row = dict(zip(header, record[1:]))
            results[row.get("NettingSetID", "")] = row
    if header is None:
        raise EvidenceError("the evidence has no RESULTS table")
    return run, results


def number(text: str) -> float | None:
    try:
        value = float(text)
    except ValueError:
        return None
    return value if math.isfinite(value) else None


def check_run(run: dict[str, str], portfolio: dict[str, Any]) -> None:
    expected = {"reporting_date": portfolio["valuation_date"], "reporting_ccy": portfolio["reporting_currency"],
                "results_status": "CURRENT"}
    wrong = [f"{key} is {run.get(key)!r}, expected {value!r}" for key, value in expected.items()
             if run.get(key) != value]
    if wrong:
        raise EvidenceError("the evidence is not a current run of this portfolio: " + "; ".join(wrong))


def compare_figure(ns_id: str, label: str, workbook: str, package: float,
                   tolerance: tuple[float, float]) -> dict[str, Any]:
    value = number(workbook)
    difference = None if value is None else value - package
    allowed = max(tolerance[0], tolerance[1] * abs(package))
    agrees = difference is not None and abs(difference) <= allowed
    return {"netting_set": ns_id, "quantity": label, "workbook": workbook, "package": package,
            "difference": difference, "agrees": agrees, "classification": ""}


def compare_netting_set(fixture: dict[str, Any], row: dict[str, str] | None,
                        result: dict[str, Any]) -> list[dict[str, Any]]:
    ns_id = fixture["id"]
    if row is None:
        return [{"netting_set": ns_id, "quantity": "row", "workbook": "missing", "package": "",
                 "difference": None, "agrees": False, "classification": ""}]
    rows = []
    for label, expected in (("Regime", REGIME), ("Status", "VALID"),
                            ("Trades used", str(len(fixture["trades"])))):
        found = row.get(label, "")
        same = found == expected or (number(found) is not None and number(found) == number(expected))
        rows.append({"netting_set": ns_id, "quantity": label, "workbook": found, "package": expected,
                     "difference": None, "agrees": same, "classification": ""})
    for label, key, tolerance in RESULT_FIGURES:
        rows.append(compare_figure(ns_id, label, row.get(label, ""), float(result[key]), tolerance))
    by_class = result.get("add_on_by_class") or {}
    for label, key in CLASS_FIGURES:
        rows.append(compare_figure(ns_id, label, row.get(label, "") or "0", float(by_class.get(key, 0.0)),
                                   AMOUNT_TOLERANCE))
    return rows


def compare(portfolio: dict[str, Any], package: dict[str, Any], evidence: str) -> list[dict[str, Any]]:
    """One row per compared figure of every netting set, and one per unexpected workbook row."""
    run, results = read_evidence(evidence)
    check_run(run, portfolio)
    by_id = {r["netting_set"]: r for r in package["results"]}
    rows: list[dict[str, Any]] = []
    for fixture in portfolio["netting_sets"]:
        rows += compare_netting_set(fixture, results.get(fixture["id"]), by_id[fixture["id"]])
    known = {f["id"] for f in portfolio["netting_sets"]}
    rows += [{"netting_set": ns_id, "quantity": "row", "workbook": "not in the portfolio", "package": "",
              "difference": None, "agrees": False, "classification": ""}
             for ns_id in sorted(results) if ns_id not in known and not ns_id.upper().startswith("TOTAL")]
    return rows


def comparison_csv(rows: list[dict[str, Any]]) -> str:
    out = io.StringIO()
    writer = csv.writer(out, delimiter=";", lineterminator="\n")
    writer.writerow(["netting_set", "quantity", "workbook", "package", "difference", "agrees", "classification"])
    for row in rows:
        writer.writerow([row["netting_set"], row["quantity"], row["workbook"], row["package"],
                         "" if row["difference"] is None else repr(row["difference"]),
                         "Y" if row["agrees"] else "N", row["classification"]])
    return out.getvalue()


def summary(portfolio: dict[str, Any], package: dict[str, Any], rows: list[dict[str, Any]]) -> str:
    differ = [row for row in rows if not row["agrees"]]
    largest = max((abs(r["difference"]) for r in rows if r["difference"] is not None), default=0.0)
    trades = sum(len(f["trades"]) for f in portfolio["netting_sets"])
    lines = ["Versions: " + ", ".join(f"{k} {v}" for k, v in package["versions"].items()),
             f"Portfolio: seed {portfolio['seed']}, {len(portfolio['netting_sets'])} netting sets, {trades} trades",
             ""]
    lines += [f"DIFFERS {r['netting_set']} {r['quantity']}: workbook {r['workbook']}, package {r['package']}"
              for r in differ]
    lines.append(f"RESULT compared={len(rows)}; differ={len(differ)}; largest_difference={largest:.3g}")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Commands
# ---------------------------------------------------------------------------

def command_generate(options: argparse.Namespace) -> int:
    portfolio = generate(options.seed, options.netting_sets)
    files = csv_files(portfolio, csv_contract.load_contract(options.root))
    options.out.mkdir(parents=True, exist_ok=True)
    (options.out / "portfolio.json").write_text(json.dumps(portfolio, indent=2) + "\n", encoding="utf-8")
    for name, text in files.items():
        (options.out / name).write_bytes(text.encode("utf-8"))
    trades = sum(len(f["trades"]) for f in portfolio["netting_sets"])
    print(f"RESULT=GENERATED; seed={options.seed}; netting_sets={len(portfolio['netting_sets'])}; "
          f"trades={trades}; reporting_date={portfolio['valuation_date']}; reporting_ccy={CURRENCY}; "
          f"regime={REGIME}")
    return 0


def command_run(options: argparse.Namespace) -> int:
    portfolio = json.loads((options.out / "portfolio.json").read_text(encoding="utf-8"))
    inputs = {"netting_sets": [crosscheck.translate_fixture(f) for f in portfolio["netting_sets"]], "skipped": []}
    (options.out / "rsaccr-inputs.json").write_text(json.dumps(inputs, indent=2), encoding="utf-8")
    package = crosscheck.run_package(options.rscript, options.out / "rsaccr-inputs.json",
                                     options.out / "rsaccr-results.json")
    print(f"RESULT=RUN; netting_sets={len(package['results'])}; versions=" +
          ", ".join(f"{k} {v}" for k, v in package["versions"].items()))
    return 0


def command_compare(options: argparse.Namespace) -> int:
    portfolio = json.loads((options.out / "portfolio.json").read_text(encoding="utf-8"))
    package = json.loads((options.out / "rsaccr-results.json").read_text(encoding="utf-8"))
    evidence = options.evidence.read_bytes().decode("utf-8-sig")
    try:
        rows = compare(portfolio, package, evidence)
    except EvidenceError as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2
    (options.out / "comparison.csv").write_text(comparison_csv(rows), encoding="utf-8")
    print(summary(portfolio, package, rows))
    return 1 if any(not row["agrees"] for row in rows) else 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--root", type=Path, default=ROOT)
    commands = parser.add_subparsers(dest="command", required=True)
    gen = commands.add_parser("generate", help="write the portfolio and its CSV files")
    gen.add_argument("--out", type=Path, required=True)
    gen.add_argument("--seed", type=int, default=352)
    gen.add_argument("--netting-sets", type=int, default=40)
    run = commands.add_parser("run", help="run the package on the portfolio")
    run.add_argument("--out", type=Path, required=True)
    run.add_argument("--rscript", default="Rscript")
    cmp = commands.add_parser("compare", help="compare the workbook's evidence with the package")
    cmp.add_argument("--out", type=Path, required=True)
    cmp.add_argument("--evidence", type=Path, required=True)
    options = parser.parse_args(argv)
    handlers = {"generate": command_generate, "run": command_run, "compare": command_compare}
    return handlers[options.command](options)


if __name__ == "__main__":
    sys.exit(main())
