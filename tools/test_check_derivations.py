"""The derivation gate fails on a change to a script or to its expected file (#537)."""
from __future__ import annotations

import json
import math
import shutil
import tempfile
import unittest
from collections.abc import Callable
from pathlib import Path
from typing import Any

from check_derivations import run_check

ROOT = Path(__file__).resolve().parents[1]
CCP = "tests/derivations/ccp_trade_exposure.py"
CCP_FILE = "tests/expected/illustrative-ccp-retained-vm.crr.json"


class DerivationGateTests(unittest.TestCase):
    def setUp(self) -> None:
        self.root = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.root)
        for folder in ("tests/derivations", "tests/expected"):
            shutil.copytree(ROOT / folder, self.root / folder,
                            ignore=shutil.ignore_patterns("__pycache__"))

    def edit_text(self, name: str, old: str, new: str) -> None:
        path = self.root / name
        text = path.read_text(encoding="utf-8")
        self.assertIn(old, text)
        path.write_text(text.replace(old, new, 1), encoding="utf-8")

    def edit_json(self, name: str, change: Callable[[list[dict[str, Any]]], object]) -> None:
        path = self.root / name
        data = json.loads(path.read_text(encoding="utf-8"))
        change(data["outputs"])
        path.write_text(json.dumps(data, indent=2), encoding="utf-8")

    def findings(self) -> list[str]:
        report = run_check(self.root)
        findings: list[str] = report["findings"]
        self.assertEqual(report["status"], "fail" if findings else "pass")
        return findings

    def assertFinding(self, text: str) -> None:
        findings = self.findings()
        self.assertTrue(any(text in item for item in findings), findings)

    def output(self, outputs: list[dict[str, Any]], quantity: str,
               trade: str | None = None) -> dict[str, Any]:
        return next(o for o in outputs if o["quantity"] == quantity and o.get("trade") == trade)

    def test_committed_files_pass(self) -> None:
        self.assertEqual(self.findings(), [])
        self.assertEqual(len(run_check(self.root)["scripts"]), 3)

    def test_changed_script_constant(self) -> None:
        self.edit_text(CCP, "ALPHA = 1.4", "ALPHA = 1.5")
        self.assertFinding(f"{CCP_FILE}: exposure_value: {CCP} gives")

    def test_changed_script_rule(self) -> None:
        self.edit_text("tests/derivations/settled_to_market.py", '"CO_OIL_GAS": 0.18', '"CO_OIL_GAS": 0.40')
        self.assertFinding("illustrative-settled-to-market-periodic.bcbs.json: add_on.commodity")

    def test_changed_expected_value(self) -> None:
        def change(outputs: list[dict[str, Any]]) -> None:
            self.output(outputs, "maturity_factor_margined", "T2")["value"] *= 1 + 1e-12
        self.edit_json(CCP_FILE, change)
        self.assertFinding(f"{CCP_FILE}: maturity_factor_margined T2: {CCP} gives")

    def test_changed_expected_text(self) -> None:
        self.edit_json(CCP_FILE, lambda outputs: self.output(outputs, "cap_applied").update(text="N"))
        self.assertFinding(f"{CCP_FILE}: cap_applied: {CCP} gives 'Y', the file 'N'")

    def test_last_place_rounding_is_tolerated(self) -> None:
        def change(outputs: list[dict[str, Any]]) -> None:
            output = self.output(outputs, "exposure_value")
            output["value"] = math.nextafter(output["value"], math.inf)
        self.edit_json(CCP_FILE, change)
        self.assertEqual(self.findings(), [])

    def test_derived_output_missing_from_file(self) -> None:
        self.edit_json(CCP_FILE, lambda outputs: outputs.remove(self.output(outputs, "multiplier")))
        self.assertFinding(f"{CCP_FILE}: multiplier: derived by {CCP} but not in the expected file")

    def test_missing_expected_file(self) -> None:
        (self.root / CCP_FILE).unlink()
        self.assertFinding(f"{CCP}: derives {CCP_FILE}, which does not exist")

    def test_citation_without_derivation(self) -> None:
        def change(outputs: list[dict[str, Any]]) -> None:
            extra = json.loads(json.dumps(self.output(outputs, "multiplier")))
            extra["quantity"] = "aggregate_add_on"
            outputs.append(extra)
        self.edit_json(CCP_FILE, change)
        self.assertFinding(f"{CCP_FILE}: aggregate_add_on: cites {CCP}, which does not derive it")

    def test_citation_of_unknown_script(self) -> None:
        def change(outputs: list[dict[str, Any]]) -> None:
            reference = self.output(outputs, "netting_set_status")["reference"]
            reference["derivation"] = "tests/derivations/missing.py, a script"
        self.edit_json(CCP_FILE, change)
        self.assertFinding("netting_set_status: cites tests/derivations/missing.py, which does not exist")
        self.assertFinding(f"netting_set_status: derived by {CCP} but cites tests/derivations/missing.py")

    def test_illustrative_without_derivation(self) -> None:
        name = "tests/expected/illustrative-ir-swap-unmargined.crr.json"
        self.edit_json(name, lambda outputs: outputs[0]["reference"].update(derivation=""))
        self.assertFinding(f"{name}: replacement_cost: illustrative output names no derivation")

    def test_script_without_expected_function(self) -> None:
        self.edit_text(CCP, "def expected()", "def derived_values()")
        self.assertFinding(f"{CCP}: defines no expected()")

    def test_broken_script(self) -> None:
        self.edit_text(CCP, "VALUATION = date(2026, 9, 30)", "VALUATION = undefined_name")
        self.assertFinding(f"{CCP}: cannot run: NameError")


if __name__ == "__main__":
    unittest.main()
