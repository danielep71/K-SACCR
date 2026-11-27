#!/usr/bin/env python3
"""Reference of the Dashboard sheet's figures (#675).

``docs/workbook/dashboard.json`` lists the blocks of the Dashboard's data
table: the figures a run writes there, which the sheet's charts and cards
read. This module computes every block from the run's own records (the rows
of Results, TradeCalc and Checks, the inputs of NettingSets and the results of
the two reconciliation sheets), so that the definitions in
``docs/workbook/DASHBOARD.md`` are executable and the engine's figures can be
derived independently of it.

Usage::

    python tools/dashboard.py --check

checks the specification: block and row names, kinds and the one scale step, 1 (#810).
"""

from __future__ import annotations

import argparse
import json
import sys
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
SPEC = ROOT / "docs" / "workbook" / "dashboard.json"
REGIMES = ("CRR", "BCBS")
ASSET_CLASSES = ("IR", "FX", "CR", "EQ", "CO", "OT")
SECTORS = ("CB", "GG", "CI", "OFC", "NFC", "HH")
KINDS = {"date", "datetime", "text", "scale", "exposure", "notional", "count", "share", "flag", "amount"}
VALID = "VALID"
RECON_COUNTED = ("MATCH", "DIFFERENCE", "MISSING IN CALCULATOR")
CLEAN_CODE = "RUN-CLEAN"

Cell = Any
Block = list[list[Cell]]


@dataclass
class NettingSetResult:
    """One netting-set row of Results (TOTAL rows are not netting sets)."""

    id: str
    counterparty: str
    regime: str
    margined: bool
    v: float
    rc: float
    add_ons: dict[str, float]
    pfe: float
    ead: float
    cap_applied: bool
    status: str


@dataclass
class NettingSetInput:
    """The inputs of NettingSets the Dashboard reads for a netting set."""

    sector: str = ""
    clearing_role: str = ""
    qccp: str = ""


@dataclass
class TradeLeg:
    """One row of TradeCalc."""

    id: str
    netting_set: str
    asset_class: str
    status: str
    notional: float = 0.0
    mtm: float = 0.0
    maturity: float = 0.0
    regime: str = ""


@dataclass
class CheckRow:
    """One row of Checks."""

    severity: str
    code: str


@dataclass
class RunInfo:
    reporting_date: str
    run_time: str
    reporting_currency: str
    default_regime: str
    parameter_set: str
    inputs: str


@dataclass
class Run:
    """Everything a run gives the Dashboard."""

    info: RunInfo
    results: list[NettingSetResult]
    inputs: dict[str, NettingSetInput]
    legs: list[TradeLeg]
    checks: list[CheckRow]
    recon_netting_sets: list[str] | None = None
    recon_trades: list[str] | None = None


# ---------------------------------------------------------------------------
# Specification
# ---------------------------------------------------------------------------

def load_spec(path: Path = SPEC) -> dict[str, Any]:
    data: dict[str, Any] = json.loads(path.read_text(encoding="utf-8"))
    return data


def check_spec(spec: dict[str, Any]) -> list[str]:
    """Findings in the specification; empty when it is sound."""
    findings: list[str] = []
    if spec.get("schema") != "k-saccr-dashboard" or spec.get("version") != 1:
        findings.append("schema must be k-saccr-dashboard, version 1")
    if spec.get("regimes") != list(REGIMES):
        findings.append(f"regimes must be {list(REGIMES)}")
    factors = [s.get("factor") for s in spec.get("scales", {}).get("steps", [])]
    if factors != [1]:
        findings.append("the one scale step must be 1: amounts in units (#810)")
    names: set[str] = set()
    for block in spec.get("blocks", []):
        findings += check_block(block, names)
    return findings


def check_block(block: dict[str, Any], names: set[str]) -> list[str]:
    """Findings in one block of the specification; adds its name to ``names``."""
    findings: list[str] = []
    name = block.get("name", "")
    where = f"block {name or '?'}"
    if not name.startswith("Dashboard_") or not name.replace("_", "").isalnum():
        findings.append(f"{where}: a name is Dashboard_ and letters")
    if name in names:
        findings.append(f"{where}: duplicate name")
    names.add(name)
    columns = block.get("columns", [])
    if not columns:
        findings.append(f"{where}: no columns")
    kinds = block.get("kinds")
    if kinds is not None and (len(kinds) != len(columns) or set(kinds) - KINDS):
        findings.append(f"{where}: kinds must name a known kind per column")
    if "rank_rows" in block:
        if "rows" in block or not isinstance(block["rank_rows"], int) or block["rank_rows"] < 1:
            findings.append(f"{where}: a ranked block has rank_rows and no rows")
        return findings
    keys = [r.get("key") for r in block.get("rows", [])]
    if not keys or len(set(keys)) != len(keys):
        findings.append(f"{where}: rows need unique keys")
    for row in block.get("rows", []):
        kind = row.get("kind")
        if kind is None and kinds is None:
            findings.append(f"{where}.{row.get('key')}: a row needs a kind when the columns have none")
        if kind is not None and kind not in KINDS:
            findings.append(f"{where}.{row.get('key')}: unknown kind {kind}")
        if not row.get("label"):
            findings.append(f"{where}.{row.get('key')}: no label")
    return findings


# ---------------------------------------------------------------------------
# Rules
# ---------------------------------------------------------------------------

def percent_text(fraction: float) -> str:
    """A fraction as a percentage with one decimal, halves away from zero: 0.3846 -> 38.5%."""
    # The fraction's shortest decimal form first, so that 0.0045 is the half 0.45%, not 0.44999...%.
    value = (Decimal(repr(fraction)) * 100).quantize(Decimal("0.1"), rounding=ROUND_HALF_UP)
    return f"{value}%"


def reading(gross: float, netted: float, rc: float) -> str:
    """The one-line reading of view 11, from gross positive value, value after netting and RC."""
    if gross <= 0:
        return "No trade has a positive market value."
    text = f"Netting removes {percent_text((gross - netted) / gross)} of gross positive value"
    collateral = (netted - rc) / gross
    if collateral >= 0:
        return f"{text}; collateral a further {percent_text(collateral)}."
    return f"{text}; collateral and margin terms add {percent_text(-collateral)}."


def share(part: float, whole: float) -> float | None:
    return part / whole if whole else None


def clearing_role(netting_input: NettingSetInput) -> str:
    role = netting_input.clearing_role.strip().upper()
    if role in ("", "N"):
        return "N"
    if role == "CCP":
        return "CCP_Q" if netting_input.qccp.strip().upper() == "Y" else "CCP_NQ"
    return role


def sector(netting_input: NettingSetInput) -> str:
    code = netting_input.sector.strip().upper()
    return code if code in SECTORS else "NC"


def ranked(valid: list[NettingSetResult]) -> list[NettingSetResult]:
    """Valid netting sets by EAD, largest first; ties by netting-set ID."""
    return sorted(valid, key=lambda r: (-r.ead, r.id))


def top_block(valid: list[NettingSetResult], factor: float) -> Block:
    """View 5 for one regime: ten rows, blank below the last netting set."""
    whole = float(sum(r.ead for r in valid))
    rows: Block = [[r.id, r.counterparty, r.ead / factor, share(r.ead, whole)] for r in ranked(valid)[:10]]
    return rows + [[None] * 4 for _ in range(10 - len(rows))]


def concentration(valid: list[NettingSetResult]) -> tuple[float | None, float | None]:
    """View 6 for one regime: share of the ten largest and the Herfindahl index; blank without EAD."""
    whole = float(sum(r.ead for r in valid))
    if not whole:
        return None, None
    top = sum(r.ead for r in ranked(valid)[:10])
    return top / whole, sum((r.ead / whole) ** 2 for r in valid)


def codes_block(checks: list[CheckRow]) -> Block:
    """View 9: the five codes of most ERROR and WARNING rows; ties by severity, then code."""
    counted: dict[str, list[str]] = {}
    for check in checks:
        if check.severity in ("ERROR", "WARNING"):
            counted.setdefault(check.code or "(none)", []).append(check.severity)
    codes = sorted(counted.items(), key=lambda c: (-len(c[1]), "ERROR" not in c[1], c[0]))[:5]
    rows: Block = [[code, "ERROR" if "ERROR" in severities else "WARNING", len(severities)]
                   for code, severities in codes]
    return rows + [[None] * 3 for _ in range(5 - len(rows))]


def recon(results: list[str] | None) -> list[Cell]:
    """View 10 for one reconciliation sheet: loaded, then the rows of each result."""
    if not results:
        return ["N", 0, 0, 0, 0]
    return ["Y", *(results.count(k) for k in RECON_COUNTED), sum(r not in RECON_COUNTED for r in results)]


# ---------------------------------------------------------------------------
# Blocks
# ---------------------------------------------------------------------------

def compute(run: Run, spec: dict[str, Any] | None = None) -> dict[str, Block]:
    """Every block of the data table, values as written: amounts in units of the reporting currency (#810)."""
    spec = spec or load_spec()
    valid = {g: [r for r in run.results if r.regime == g and r.status == VALID] for g in REGIMES}
    every = {g: [r for r in run.results if r.regime == g] for g in REGIMES}
    valid_ids = {g: {r.id for r in valid[g]} for g in REGIMES}
    used = {g: [t for t in run.legs if t.status == "OK" and t.regime == g and t.netting_set in valid_ids[g]]
            for g in REGIMES}

    def total(g: str, attribute: str) -> float:
        return float(sum(getattr(r, attribute) for r in valid[g]))

    def ead_where(g: str, keep: Any) -> float:
        return float(sum(r.ead for r in valid[g] if keep(r)))

    def inputs(r: NettingSetResult) -> NettingSetInput:
        return run.inputs.get(r.id, NettingSetInput())

    exposure: dict[str, list[float]] = {}
    notional: dict[str, list[float]] = {}
    exposure["key"] = [x for g in REGIMES for x in (total(g, "ead"), total(g, "rc"), total(g, "pfe"))]
    exposure["sector"] = [ead_where(g, lambda r, k=k: sector(inputs(r)) == k)
                          for k in (*SECTORS, "NC") for g in REGIMES]
    roles = ("N", "CM", "CLIENT", "CCP_Q", "CCP_NQ")
    exposure["role"] = [ead_where(g, lambda r, k=k: clearing_role(inputs(r)) == k) for k in roles for g in REGIMES]
    exposure["addon"] = [float(sum(r.add_ons.get(a, 0.0) for r in valid[g])) for a in ASSET_CLASSES for g in REGIMES]
    composition = {g: (total(g, "rc"), total(g, "pfe"), ead_where(g, lambda r: r.margined),
                       ead_where(g, lambda r: not r.margined)) for g in REGIMES}
    exposure["composition"] = [x for g in REGIMES for x in composition[g]]

    def maturity_bucket(t: TradeLeg) -> str:
        return "under_1y" if t.maturity < 1 else ("1y_5y" if t.maturity <= 5 else "over_5y")

    maturity = {g: {k: float(sum(t.notional for t in used[g] if maturity_bucket(t) == k))
                    for k in ("under_1y", "1y_5y", "over_5y")} for g in REGIMES}
    notional["maturity"] = [v for g in REGIMES for v in maturity[g].values()]
    by_class = {g: {a: [t for t in used[g] if t.asset_class == a] for a in ASSET_CLASSES} for g in REGIMES}
    notional["class"] = [float(sum(t.notional for t in by_class[g][a])) for a in ASSET_CLASSES for g in REGIMES]
    path = {g: {"notional": float(sum(t.notional for t in used[g])),
                "gross_positive": float(sum(max(t.mtm, 0.0) for t in used[g])),
                "netted": float(sum(max(r.v, 0.0) for r in valid[g])),
                "rc": total(g, "rc"), "pfe": total(g, "pfe"), "ead": total(g, "ead")} for g in REGIMES}
    notional["path"] = [path[g]["notional"] for g in REGIMES]
    exposure["path"] = [path[g][k] for g in REGIMES for k in ("gross_positive", "netted", "rc", "pfe", "ead")]

    # Amounts in units of the reporting currency (#810): the one step of the specification.
    fx = fn = spec["scales"]["steps"][0]["factor"]
    unit_text = run.info.reporting_currency.strip()

    blocks: dict[str, Block] = {}
    i = run.info
    blocks["Dashboard_Run"] = [[i.reporting_date], [i.run_time], [i.reporting_currency], [i.default_regime],
                               [i.parameter_set], [i.inputs]]
    blocks["Dashboard_Scale"] = [[fx, unit_text], [fn, unit_text]]

    def counterparties(rows: list[NettingSetResult]) -> int:
        return len({r.counterparty.strip().upper() for r in rows if r.counterparty.strip()})

    def key_column(g: str) -> list[Cell]:
        rows = every[g]
        return [total(g, "ead") / fx, total(g, "rc") / fx, total(g, "pfe") / fx, len(rows), len(valid[g]),
                sum(r.status.startswith("INVALID") for r in rows), sum(r.status.startswith("INCOMPLETE") for r in rows),
                sum(r.status == "NO TRADES" for r in rows), sum(r.margined for r in rows),
                sum(not r.margined for r in rows), sum(r.cap_applied for r in valid[g]), counterparties(rows)]

    columns = [key_column(g) for g in REGIMES]
    blocks["Dashboard_KeyFigures"] = [list(row) for row in zip(*columns, strict=True)]
    severity = [c.severity for c in run.checks if c.code != CLEAN_CODE]
    blocks["Dashboard_Counts"] = [
        [len(run.legs)], [sum(t.status == "OK" for t in run.legs)], [sum(t.status == "EXCLUDED" for t in run.legs)],
        [sum(r.regime not in REGIMES for r in run.results)], [counterparties(run.results)],
        [severity.count("ERROR")], [severity.count("WARNING")], [severity.count("INFO")]]

    def per_regime(keys: tuple[str, ...], value: Any) -> Block:
        return [[value(g, k) for g in REGIMES] for k in keys]

    blocks["Dashboard_Sector"] = per_regime((*SECTORS, "NC"),
                                            lambda g, k: ead_where(g, lambda r: sector(inputs(r)) == k) / fx)
    blocks["Dashboard_ClearingRole"] = per_regime(
        roles, lambda g, k: ead_where(g, lambda r: clearing_role(inputs(r)) == k) / fx)
    blocks["Dashboard_AddOn"] = per_regime(ASSET_CLASSES,
                                           lambda g, a: float(sum(r.add_ons.get(a, 0.0) for r in valid[g])) / fx)
    blocks["Dashboard_Composition"] = [[composition[g][n] / fx for g in REGIMES] for n in range(4)]

    for g in REGIMES:
        blocks[f"Dashboard_Top{g}"] = top_block(valid[g], fx)
    measures = {g: concentration(valid[g]) for g in REGIMES}
    blocks["Dashboard_Concentration"] = [[measures[g][n] for g in REGIMES] for n in range(2)]
    blocks["Dashboard_Maturity"] = per_regime(("under_1y", "1y_5y", "over_5y"), lambda g, k: maturity[g][k] / fn)
    blocks["Dashboard_AssetClass"] = [
        [x for g in REGIMES for x in (len(by_class[g][a]), float(sum(t.notional for t in by_class[g][a])) / fn)]
        for a in ASSET_CLASSES]

    blocks["Dashboard_DataQuality"] = [[sum(t.status == "EXCLUDED" for t in run.legs)],
                                       [sum(r.status != VALID for r in run.results)]]
    blocks["Dashboard_Codes"] = codes_block(run.checks)
    blocks["Dashboard_Reconciliation"] = [list(row) for row in
                                          zip(recon(run.recon_netting_sets), recon(run.recon_trades), strict=True)]

    def path_row(key: str) -> list[Cell]:
        factor = fn if key == "notional" else fx
        return [x for g in REGIMES
                for x in (path[g][key] / factor, share(path[g][key], path[g]["notional"]))]

    blocks["Dashboard_NotionalToEAD"] = [path_row(k) for k in
                                         ("notional", "gross_positive", "netted", "rc", "pfe", "ead")]
    blocks["Dashboard_Reading"] = [[reading(path[g]["gross_positive"], path[g]["netted"], path[g]["rc"])
                                    if valid[g] else ""] for g in REGIMES]
    return blocks


def shape_problems(blocks: dict[str, Block], spec: dict[str, Any]) -> list[str]:
    """Blocks whose size differs from the specification's."""
    findings = []
    for block in spec["blocks"]:
        values = blocks.get(block["name"])
        rows = block.get("rank_rows") or len(block["rows"])
        if values is None or len(values) != rows or any(len(r) != len(block["columns"]) for r in values):
            findings.append(f"{block['name']}: expected {rows} x {len(block['columns'])}")
    return findings


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n", 1)[0])
    parser.add_argument("--check", action="store_true", help="check docs/workbook/dashboard.json")
    options = parser.parse_args(argv)
    if not options.check:
        parser.print_help()
        return 2
    findings = check_spec(load_spec())
    for finding in findings:
        print(f"- {finding}")
    print(f"Dashboard specification: {'FAIL' if findings else 'PASS'}")
    return 1 if findings else 0


if __name__ == "__main__":
    sys.exit(main())
