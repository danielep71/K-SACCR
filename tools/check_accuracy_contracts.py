#!/usr/bin/env python3
"""Check the accuracy contracts in tests/accuracy-contracts.json against the
expected files' tolerances (#537).

A contract states, for one quantity family, the tolerance K-SACCR claims at
full precision and the cases that test it, split into calibration and
held-out cases. The gate fails when:

- the file does not follow its format, or two families claim one quantity;
- a case is listed that has no expected file, is listed twice, or has no
  output of the family; or a case with such an output is not listed;
- an output's tolerance is looser than the contract's: the claimed tolerance,
  or for a case listed under ``rounded`` the rounding of its expected value;
- a ``rounded`` entry is not needed, because every output of the case meets
  the claimed tolerance;
- no held-out case tests the claimed tolerance at full precision.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

from _gatelib import parse_report_args, run_gate

CONTRACTS = "tests/accuracy-contracts.json"
EXPECTED = "tests/expected"
CONTRACT_KEYS = {"family", "description", "quantities", "calibration", "held_out", "rounded"}


def matches(quantity: str, pattern: str) -> bool:
    return quantity == pattern or (pattern.endswith("*") and quantity.startswith(pattern[:-1]))


def tolerance_errors(where: str, tolerance: Any) -> list[str]:
    if not isinstance(tolerance, dict) or set(tolerance) != {"absolute", "relative"} or any(
            isinstance(v, bool) or not isinstance(v, (int, float)) or v < 0 for v in tolerance.values()):
        return [f"{where}: tolerance must be {{absolute, relative}}, both non-negative numbers"]
    return []


def within(tolerance: dict[str, Any], bound: dict[str, Any]) -> bool:
    return tolerance["absolute"] <= bound["absolute"] and tolerance["relative"] <= bound["relative"]


def check(document: dict[str, Any], expected: dict[str, dict[str, Any]]) -> list[str]:
    """Check the contracts against expected files keyed by case (file stem)."""
    if document.get("schema_version") != 1 or document.get("kind") != "saccr-accuracy-contracts":
        return [f"{CONTRACTS}: schema_version must be 1 and kind 'saccr-accuracy-contracts'"]
    contracts = document.get("contracts")
    if not isinstance(contracts, list) or not contracts:
        return [f"{CONTRACTS}: contracts must be a non-empty list"]
    findings: list[str] = []
    owner: dict[str, str] = {}
    for contract in contracts:
        if not isinstance(contract, dict) or set(contract) != CONTRACT_KEYS:
            findings.append(f"{CONTRACTS}: a contract must have exactly {sorted(CONTRACT_KEYS)}")
            continue
        family, before = contract["family"], len(findings)
        quantities, rounded = contract["quantities"], contract["rounded"]
        calibration, held_out = contract["calibration"], contract["held_out"]
        if not (isinstance(quantities, dict) and quantities and isinstance(rounded, dict)
                and isinstance(calibration, list) and isinstance(held_out, list)):
            findings.append(f"{family}: quantities and rounded must be objects, the case lists lists")
            continue
        for pattern, bound in quantities.items():
            findings += tolerance_errors(f"{family}: {pattern}", bound)
            for other, name in owner.items():
                if matches(pattern, other) or matches(other, pattern):
                    findings.append(f"{family}: {pattern} is already claimed by {name}")
            owner[pattern] = family
        for case, bound in rounded.items():
            findings += tolerance_errors(f"{family}: rounded {case}", bound)
        if len(findings) > before:
            continue
        listed = calibration + held_out
        for case in sorted({c for c in listed if listed.count(c) > 1}):
            findings.append(f"{family}: {case} is listed more than once")
        tested: dict[str, list[dict[str, Any]]] = {}
        for case, data in expected.items():
            outputs = [o for o in data.get("outputs", [])
                       if any(matches(o.get("quantity", ""), p) for p in quantities) and "tolerance" in o]
            if outputs:
                tested[case] = outputs
        for case in sorted(set(listed) - set(expected)):
            findings.append(f"{family}: {case} has no expected file")
        for case in sorted(set(listed) & set(expected) - set(tested)):
            findings.append(f"{family}: {case} has no output of the family")
        for case in sorted(set(tested) - set(listed)):
            findings.append(f"{family}: {case} is neither a calibration nor a held-out case")
        for case in sorted(set(rounded) - set(listed)):
            findings.append(f"{family}: rounded {case} is not a listed case")
        for case, outputs in sorted(tested.items()):
            loose = False
            for output in outputs:
                pattern = next(p for p in quantities if matches(output["quantity"], p))
                claimed = quantities[pattern]
                tolerance = output["tolerance"]
                loose |= not within(tolerance, claimed)
                bound = rounded.get(case, claimed)
                if not within(tolerance, bound):
                    kind = "rounding" if case in rounded else "claimed tolerance"
                    trade = f" {output['trade']}" if output.get("trade") else ""
                    findings.append(f"{family}: {case} {output['quantity']}{trade}: tolerance "
                                    f"{tolerance} is looser than the {kind} {bound}")
            if case in rounded and not loose:
                findings.append(f"{family}: rounded {case} is not needed: every output meets the claimed tolerance")
        if not any(case in tested and case not in rounded for case in held_out):
            findings.append(f"{family}: no held-out case tests the claimed tolerance at full precision")
    return findings


def run_check(root: Path) -> dict[str, Any]:
    document = json.loads((root / CONTRACTS).read_text(encoding="utf-8"))
    expected = {path.stem: json.loads(path.read_text(encoding="utf-8"))
                for path in sorted((root / EXPECTED).glob("*.json"))}
    findings = check(document, expected)
    families = [c.get("family") for c in document.get("contracts") or [] if isinstance(c, dict)]
    return {"schema_version": 1, "status": "fail" if findings else "pass",
            "families": families, "findings": findings}


def markdown(report: dict[str, Any]) -> str:
    head = f"Accuracy contracts: {report['status']} ({', '.join(map(str, report['families']))})\n"
    return head + "".join(f"- {item}\n" for item in report["findings"])


def main() -> int:
    options = parse_report_args(sys.argv[1:], description=__doc__)
    return run_gate(options, build=lambda: run_check(options.root.resolve()), markdown=markdown,
                    errors=(OSError, ValueError, KeyError))


if __name__ == "__main__":
    raise SystemExit(main())
