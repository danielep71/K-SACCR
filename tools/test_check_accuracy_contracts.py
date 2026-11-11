"""The accuracy contracts agree with the expected files' tolerances (#537)."""
from __future__ import annotations

import copy
import json
import unittest
from pathlib import Path

from check_accuracy_contracts import CONTRACTS, check, run_check

ROOT = Path(__file__).resolve().parents[1]


class AccuracyContractTests(unittest.TestCase):
    def setUp(self) -> None:
        self.document = json.loads((ROOT / CONTRACTS).read_text(encoding="utf-8"))
        self.expected = {path.stem: json.loads(path.read_text(encoding="utf-8"))
                         for path in sorted((ROOT / "tests/expected").glob("*.json"))}

    def contract(self, family: str) -> dict:
        return next(c for c in self.document["contracts"] if c["family"] == family)

    def output(self, case: str, quantity: str) -> dict:
        return next(o for o in self.expected[case]["outputs"] if o["quantity"] == quantity)

    def assertFinding(self, text: str) -> None:
        findings = check(self.document, self.expected)
        self.assertTrue(any(text in item for item in findings), findings)

    def test_committed_contracts_pass(self) -> None:
        report = run_check(ROOT)
        self.assertEqual(report["findings"], [])
        self.assertEqual(report["families"], ["EAD", "RC", "PFE", "add-ons"])

    def test_looser_expected_tolerance(self) -> None:
        case = "illustrative-ccp-retained-vm.crr"
        self.output(case, "exposure_value")["tolerance"]["relative"] = 1e-6
        self.assertFinding(f"EAD: {case} exposure_value: tolerance")

    def test_looser_tolerance_than_the_rounding(self) -> None:
        self.output("cre99-example-4.bcbs", "replacement_cost")["tolerance"]["absolute"] = 5000
        self.assertFinding("RC: cre99-example-4.bcbs replacement_cost: tolerance")

    def test_tighter_claim_than_the_expected_files(self) -> None:
        self.contract("PFE")["quantities"]["multiplier"]["absolute"] = 1e-12
        self.assertFinding("PFE: illustrative-ccp-retained-vm.crr multiplier: tolerance")

    def test_unlisted_case(self) -> None:
        self.contract("RC")["held_out"].remove("illustrative-ir-swap-unmargined.crr")
        self.assertFinding("RC: illustrative-ir-swap-unmargined.crr is neither a calibration nor a held-out case")

    def test_case_listed_twice(self) -> None:
        self.contract("RC")["calibration"].append("illustrative-ir-swap-unmargined.crr")
        self.assertFinding("RC: illustrative-ir-swap-unmargined.crr is listed more than once")

    def test_unknown_case_and_case_without_family_output(self) -> None:
        self.contract("RC")["held_out"] += ["missing-case.crr", "illustrative-other-risks.crr"]
        self.assertFinding("RC: missing-case.crr has no expected file")
        self.assertFinding("RC: illustrative-other-risks.crr has no output of the family")

    def test_rounded_entry_not_needed(self) -> None:
        self.contract("RC")["rounded"]["illustrative-ir-swap-unmargined.crr"] = {"absolute": 1, "relative": 0}
        self.assertFinding("RC: rounded illustrative-ir-swap-unmargined.crr is not needed")

    def test_no_full_precision_held_out_case(self) -> None:
        contract = self.contract("RC")
        contract["calibration"] += contract["held_out"]
        contract["held_out"] = []
        self.assertFinding("RC: no held-out case tests the claimed tolerance at full precision")

    def test_overlapping_families(self) -> None:
        self.contract("add-ons")["quantities"]["add_on.equity"] = {"absolute": 1e-6, "relative": 1e-9}
        self.assertFinding("add-ons: add_on.equity is already claimed by add-ons")

    def test_format(self) -> None:
        document = copy.deepcopy(self.document)
        document["kind"] = "other"
        self.assertTrue(check(document, self.expected))
        self.contract("RC")["quantities"]["replacement_cost"] = {"absolute": -1, "relative": 0}
        self.assertFinding("RC: replacement_cost: tolerance must be")


if __name__ == "__main__":
    unittest.main()
