#!/usr/bin/env python3
"""Check the differences an Excel run measured against the accuracy contracts (#537).

The case runner of the workbook (TEST_CaseRunner.ExpectNumber) prints one line
for each numeric comparison it makes, in the log of its suite::

    MEASURED=cre99-example-1.bcbs; trade=EX1-T1; quantity=adjusted_notional; expected=78694000; actual=78694000; difference=0

``tests/accuracy-contracts.json`` states the tolerance each quantity family
claims, and ``tests/expected`` the cases that test it. For a run of the case
suite that passed, the evidence validator (tools/check_excel_evidence.py) calls
:func:`validate`, which fails when:

- an output of a contract family has no line, or has two;
- a line's expected value is not the one of the expected file;
- the difference of a line is not the one its expected and actual values give;
- the measured difference exceeds the bound of the contract: the claimed
  tolerance, or for a case listed under ``rounded`` the rounding of its
  expected value.

Run as a script on a log, it prints the measured differences by family.
"""

from __future__ import annotations

import argparse
import json
import math
import re
import sys
from pathlib import Path
from typing import Any, NamedTuple

from _gatelib import git_bytes
from check_accuracy_contracts import CONTRACTS as CONTRACTS
from check_accuracy_contracts import EXPECTED as EXPECTED
from check_accuracy_contracts import family_outputs, matches

# The suite whose log carries the measured differences.
ACCURACY_SUITE = "TEST_Cases"
LINE = re.compile(r"^MEASURED=(?P<case>[^;]+); trade=(?P<trade>[^;]*); quantity=(?P<quantity>[^;]+); "
                  r"expected=(?P<expected>\S+); actual=(?P<actual>\S+); difference=(?P<difference>\S+)[ \t]*$")
# A printed number has 15 significant digits: the difference is checked against the values to this.
AGREEMENT = 1e-12


class Key(NamedTuple):
    case: str
    trade: str
    quantity: str


class Measured(NamedTuple):
    expected: float
    actual: float
    difference: float


class Output(NamedTuple):
    family: str
    value: float
    permitted: float


def number(text: str) -> float:
    value = float(text)
    if not math.isfinite(value):
        raise ValueError(f"not a finite number: {text}")
    return value


def parse(log: str) -> tuple[dict[Key, Measured], list[str]]:
    """The measured lines of a log by case, trade and quantity, and the problems found in them."""
    found: dict[Key, Measured] = {}
    problems: list[str] = []
    for line in log.splitlines():
        if not line.startswith("MEASURED="):
            continue
        match = LINE.match(line)
        if match is None:
            problems.append(f"unreadable measured line: {line[:100]}")
            continue
        key = Key(match["case"].strip(), match["trade"].strip(), match["quantity"].strip())
        try:
            value = Measured(number(match["expected"]), number(match["actual"]), number(match["difference"]))
        except ValueError:
            problems.append(f"{key.case} {key.quantity}: a printed number is not a finite number")
            continue
        if key in found:
            problems.append(f"{key.case} {key.trade} {key.quantity}: measured twice")
        found[key] = value
    return found, problems


def permitted(bound: dict[str, Any], value: float) -> float:
    return float(max(bound["absolute"], bound["relative"] * abs(value)))


def outputs(document: dict[str, Any], expected: dict[str, dict[str, Any]]) -> tuple[dict[Key, Output], list[str]]:
    """Every output a contract covers, with the difference its contract permits."""
    required: dict[Key, Output] = {}
    problems: list[str] = []
    for contract in document["contracts"]:
        quantities, rounded = contract["quantities"], contract["rounded"]
        for case, items in family_outputs(quantities, expected).items():
            for item in items:
                claimed = quantities[next(p for p in quantities if matches(item["quantity"], p))]
                key = Key(case, item.get("trade") or "", item["quantity"])
                if key in required:
                    problems.append(f"{case} {key.trade} {key.quantity}: two outputs of the expected file")
                required[key] = Output(contract["family"], float(item["value"]),
                                       permitted(rounded.get(case, claimed), float(item["value"])))
    return required, problems


def line_problems(key: Key, line: Measured, wanted: Output) -> list[str]:
    where = f"{key.case} {key.trade} {key.quantity}".replace("  ", " ")
    scale = max(abs(line.expected), abs(line.actual), 1.0)
    if abs(line.expected - wanted.value) > AGREEMENT * max(abs(wanted.value), 1.0):
        return [f"{where}: expected {line.expected} in the log, {wanted.value} in the expected file"]
    if abs(abs(line.actual - line.expected) - line.difference) > AGREEMENT * scale:
        return [f"{where}: difference {line.difference} is not that of {line.expected} and {line.actual}"]
    if line.difference > wanted.permitted + AGREEMENT * scale:
        return [f"{where} ({wanted.family}): measured difference {line.difference:.6g} exceeds "
                f"the contract's {wanted.permitted:.6g}"]
    return []


def check(document: dict[str, Any], expected: dict[str, dict[str, Any]],
          log: str) -> tuple[list[str], dict[str, dict[str, Any]]]:
    """The problems of a log's measured lines, and per family the count and the largest difference."""
    found, problems = parse(log)
    required, more = outputs(document, expected)
    problems += more
    families: dict[str, dict[str, Any]] = {}
    for key, wanted in sorted(required.items()):
        line = found.get(key)
        if line is None:
            problems.append(f"{key.case} {key.trade} {key.quantity}: no measured difference in the log".replace("  ", " "))
            continue
        problems += line_problems(key, line, wanted)
        row = families.setdefault(wanted.family, {"outputs": 0, "largest": 0.0, "case": ""})
        row["outputs"] += 1
        if line.difference >= row["largest"]:
            row["largest"], row["case"] = line.difference, f"{key.case} {key.quantity}"
    return problems, families


def load(root: Path, sha: str | None = None) -> tuple[dict[str, Any], dict[str, dict[str, Any]]]:
    """The contracts and the expected files, from the working tree or from one commit."""
    def read(path: str) -> bytes:
        if sha is None:
            return (root / path).read_bytes()
        result = git_bytes(root, "show", f"{sha}:{path}")
        if result.returncode:
            raise ValueError(f"candidate does not contain {path}")
        return result.stdout

    document = json.loads(read(CONTRACTS).decode("utf-8"))
    if sha is None:
        names = sorted(path.name for path in (root / EXPECTED).glob("*.json"))
    else:
        listing = git_bytes(root, "ls-tree", "--name-only", sha, EXPECTED + "/")
        if listing.returncode:
            raise ValueError("cannot list the expected files of the candidate")
        names = sorted(Path(item).name for item in listing.stdout.decode("utf-8").splitlines()
                       if item.endswith(".json"))
    expected = {Path(name).stem: json.loads(read(f"{EXPECTED}/{name}").decode("utf-8")) for name in names}
    return document, expected


def validate(root: Path, sha: str, log: str) -> None:
    """Raise ValueError when the log of a passing case suite does not meet the contracts of ``sha``."""
    document, expected = load(root, sha)
    problems, _ = check(document, expected, log)
    if problems:
        more = f" (and {len(problems) - 3} more)" if len(problems) > 3 else ""
        raise ValueError("measured differences: " + "; ".join(problems[:3]) + more)


def table(families: dict[str, dict[str, Any]], document: dict[str, Any]) -> str:
    claims = {c["family"]: ", ".join(f"{q} {b['absolute']:g} / {b['relative']:g}"
                                     for q, b in c["quantities"].items()) for c in document["contracts"]}
    rows = ["| Family | Outputs | Largest difference | Where | Claimed (absolute / relative) |",
            "| --- | --- | --- | --- | --- |"]
    for family, row in families.items():
        rows.append(f"| {family} | {row['outputs']} | {row['largest']:.3g} | {row['case']} | {claims[family]} |")
    return "\n".join(rows) + "\n"


def main(arguments: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path.cwd())
    parser.add_argument("--log", type=Path, required=True, help="the log of the case suite")
    options = parser.parse_args(arguments)
    try:
        document, expected = load(options.root.resolve())
        problems, families = check(document, expected, options.log.read_text(encoding="utf-8-sig"))
    except (OSError, ValueError, KeyError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2
    print(table(families, document), end="")
    for item in problems:
        print(f"- {item}")
    return 1 if problems else 0


if __name__ == "__main__":
    raise SystemExit(main())
