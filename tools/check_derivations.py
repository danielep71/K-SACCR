#!/usr/bin/env python3
"""Run every derivation in tests/derivations and compare it with the expected
values it produced (#537).

Each script under ``tests/derivations/`` defines ``expected()``, which returns
the outputs of each expected file it derives, keyed by ``(fixture, regime)``.
An output key is a quantity, followed by a space and a trade ID when the
quantity is per trade, such as ``maturity_factor T1``; a number is compared
with the output's ``value``, a string with its ``text``.

The gate fails when:

- a script cannot be run or defines no ``expected()``;
- a derived expected file is missing, or a derived value differs from the
  committed one by more than four units in the last place, the most the
  platform's math library can move a result; anything a deliberate change
  makes differs by far more;
- a derived output is not in its expected file;
- an expected output cites a script under ``tests/derivations/`` that does
  not derive it;
- an ``illustrative`` output names no derivation.
"""

from __future__ import annotations

import importlib.util
import json
import math
import sys
from pathlib import Path
from typing import Any

from _gatelib import parse_report_args, run_gate

DERIVATIONS = "tests/derivations"
EXPECTED = "tests/expected"
ULPS = 4


def load_derivations(root: Path) -> tuple[dict[str, dict[tuple[str, str], dict[str, Any]]], list[str]]:
    """Run each script's ``expected()``; return its outputs by script path."""
    derived: dict[str, dict[tuple[str, str], dict[str, Any]]] = {}
    findings: list[str] = []
    sys.dont_write_bytecode = True  # Leave no __pycache__ beside the scripts.
    for path in sorted((root / DERIVATIONS).glob("*.py")):
        name = path.relative_to(root).as_posix()
        spec = importlib.util.spec_from_file_location(f"derivation_{path.stem}", path)
        if spec is None or spec.loader is None:
            findings.append(f"{name}: cannot be loaded")
            continue
        try:
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)
            produce = getattr(module, "expected", None)
            if not callable(produce):
                findings.append(f"{name}: defines no expected()")
                continue
            derived[name] = produce()
        except Exception as error:  # A broken script is a finding, not a crash.
            findings.append(f"{name}: cannot run: {type(error).__name__}: {error}")
    return derived, findings


def split_key(key: str) -> tuple[str, str | None]:
    quantity, _, trade = key.partition(" ")
    return quantity, trade or None


def same(derived: Any, output: dict[str, Any]) -> bool:
    if isinstance(derived, str):
        return output.get("text") == derived and "value" not in output
    if isinstance(derived, bool) or not isinstance(derived, (int, float)):
        return False
    committed = output.get("value")
    if isinstance(committed, bool) or not isinstance(committed, (int, float)):
        return False
    derived, committed = float(derived), float(committed)
    return bool(abs(derived - committed) <= ULPS * math.ulp(max(abs(derived), abs(committed))))


def shown(output: dict[str, Any]) -> str:
    return repr(output["text"]) if "text" in output else repr(output.get("value"))


def cited_script(output: dict[str, Any]) -> str | None:
    derivation = (output.get("reference") or {}).get("derivation") or ""
    first = derivation.split(",")[0].split()[0] if derivation.strip() else ""
    return first if first.startswith(DERIVATIONS + "/") else None


def check(derived: dict[str, dict[tuple[str, str], dict[str, Any]]],
          expected: dict[str, dict[str, Any]], scripts: set[str] | None = None) -> list[str]:
    """Compare derived outputs with expected files keyed by repository path;
    ``scripts`` names every script present, including those that failed."""
    scripts = set(derived) if scripts is None else scripts
    findings: list[str] = []
    claimed: set[tuple[str, str, str | None, str]] = set()
    for script, files in sorted(derived.items()):
        for (fixture, regime), outputs in sorted(files.items()):
            name = f"{EXPECTED}/{fixture}.{regime.lower()}.json"
            data = expected.get(name)
            if data is None:
                findings.append(f"{script}: derives {name}, which does not exist")
                continue
            if data.get("fixture") != fixture or data.get("regime") != regime:
                findings.append(f"{name}: fixture or regime differs from {script}")
            committed = {(o.get("quantity"), o.get("trade")): o for o in data.get("outputs", [])}
            for key, value in sorted(outputs.items()):
                quantity, trade = split_key(key)
                output = committed.get((quantity, trade))
                label = f"{name}: {key}"
                if output is None:
                    findings.append(f"{label}: derived by {script} but not in the expected file")
                    continue
                claimed.add((name, quantity, trade, script))
                if cited_script(output) != script:
                    findings.append(f"{label}: derived by {script} but cites "
                                    f"{cited_script(output) or 'no derivation script'}")
                if not same(value, output):
                    findings.append(f"{label}: {script} gives {value!r}, the file {shown(output)}")
    for name, data in sorted(expected.items()):
        for output in data.get("outputs", []):
            reference = output.get("reference") or {}
            key = output.get("quantity") + (f" {output['trade']}" if output.get("trade") else "")
            if reference.get("class") == "illustrative" and not str(reference.get("derivation") or "").strip():
                findings.append(f"{name}: {key}: illustrative output names no derivation")
            cited = cited_script(output)
            if cited and cited not in scripts:
                findings.append(f"{name}: {key}: cites {cited}, which does not exist")
            elif cited in derived and (name, output.get("quantity"), output.get("trade"), cited) not in claimed:
                findings.append(f"{name}: {key}: cites {cited}, which does not derive it")
    return findings


def run_check(root: Path) -> dict[str, Any]:
    derived, findings = load_derivations(root)
    scripts = {path.relative_to(root).as_posix() for path in (root / DERIVATIONS).glob("*.py")}
    expected = {path.relative_to(root).as_posix(): json.loads(path.read_text(encoding="utf-8"))
                for path in sorted((root / EXPECTED).glob("*.json"))}
    findings += check(derived, expected, scripts)
    compared = sum(len(outputs) for files in derived.values() for outputs in files.values())
    return {"schema_version": 1, "status": "fail" if findings else "pass",
            "scripts": sorted(derived), "derived_outputs": compared, "findings": findings}


def markdown(report: dict[str, Any]) -> str:
    head = (f"Derivations: {report['status']} ({len(report['scripts'])} scripts, "
            f"{report['derived_outputs']} derived outputs)\n")
    return head + "".join(f"- {item}\n" for item in report["findings"])


def main() -> int:
    options = parse_report_args(sys.argv[1:], description=__doc__)
    return run_gate(options, build=lambda: run_check(options.root.resolve()), markdown=markdown,
                    errors=(OSError, ValueError))


if __name__ == "__main__":
    raise SystemExit(main())
