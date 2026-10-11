"""Offline fixtures for the numerical test-case validator."""
import copy
import json
import shutil
import tempfile
import unittest
from pathlib import Path
from typing import Any

import check_test_cases as cases

ROOT = Path(__file__).resolve().parents[1]
NETTING_SET = {
    "id": "NS1", "margined": False, "clearing_role": "none", "remargin_period_business_days": None,
    "large_or_illiquid": False, "margin_disputes": False, "mpor_override_business_days": None,
    "variation_margin_net_amount": 0, "independent_collateral_net_amount": 0,
    "threshold_amount": 0, "minimum_transfer_amount": 0, "alpha_factor": None,
}
TRADE = {
    "id": "T1", "asset_class": "interest_rate", "sub_class": None, "risk_factor": "EUR",
    "instrument": "linear", "direction": "long", "option_type": None, "nature": "standard",
    "hedging_set_label": None, "notional_amount": 10000, "market_value_amount": 30,
    "start_date": None, "end_date": "2031-01-02", "maturity_date": "2031-01-02",
    "option_expiry_date": None, "underlying_price": None, "strike_price": None,
    "lambda_price": None, "attachment_rate": None, "detachment_rate": None, "description": None,
}
FIXTURE = {
    "schema_version": 1, "kind": "saccr-fixture", "id": "sample-case", "synthetic": True,
    "description": "Synthetic sample", "category": "nominal", "calculation_currency": "EUR",
    "valuation_date": "2026-01-02", "parameters": {}, "netting_set": NETTING_SET, "trades": [TRADE],
}
PUBLISHED = {"class": "published", "source": "BCBS-279", "locator": "Annex 4, example 1",
             "derivation": "Printed figure", "derived_by": "Author", "reviewed_by": None,
             "date": "2026-10-06"}
EXPECTED: dict[str, Any] = {
    "schema_version": 1, "kind": "saccr-expected", "fixture": "sample-case", "regime": "BCBS",
    "outputs": [
        {"quantity": "exposure_value", "rule": "CRE52.1", "value": 569, "unit": "EUR",
         "tolerance": {"absolute": 0.5, "relative": 0}, "reference": PUBLISHED, "catalogue": "T04"},
        {"quantity": "trade_status", "rule": "CRE52", "trade": "T1", "text": "OK",
         "reference": PUBLISHED},
    ],
}


class TestCaseValidatorTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        (self.root / "docs/methodology").mkdir(parents=True)
        shutil.copy(ROOT / cases.REGISTER, self.root / cases.REGISTER)
        self.fixture: dict[str, Any] = copy.deepcopy(FIXTURE)
        self.expected: dict[str, Any] = copy.deepcopy(EXPECTED)

    def findings(self, fixture_name: str = "sample-case", expected_name: str = "sample-case.bcbs") -> list[str]:
        for folder, name, data in ((cases.FIXTURES, fixture_name, self.fixture),
                                   (cases.EXPECTED, expected_name, self.expected)):
            target = self.root / folder / f"{name}.json"
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_text(json.dumps(data))
        result: list[str] = cases.run_check(self.root)["findings"]
        return result

    def assertFinding(self, fragment: str, findings: list[str]) -> None:
        self.assertTrue(any(fragment in f for f in findings), findings)

    def test_repository_cases_pass(self) -> None:
        report = cases.run_check(ROOT)
        self.assertEqual(report["findings"], [])
        self.assertGreaterEqual(report["classes"]["published"], 14)

    def test_valid_sample_passes(self) -> None:
        self.assertEqual(self.findings(), [])

    def test_fixture_envelope(self) -> None:
        self.fixture["synthetic"] = False
        self.fixture["category"] = "typical"
        findings = self.findings()
        self.assertFinding("synthetic must be true", findings)
        self.assertFinding("category must be one of", findings)
        self.fixture = copy.deepcopy(FIXTURE)
        self.assertFinding("id must equal", self.findings(fixture_name="other-name"))

    def test_fields_need_their_unit_and_type(self) -> None:
        self.fixture["trades"][0]["notional"] = 10000
        self.fixture["trades"][0]["maturity_date"] = "2031-02-30"
        self.fixture["trades"][0]["direction"] = "up"
        del self.fixture["netting_set"]["threshold_amount"]
        findings = self.findings()
        self.assertFinding("unknown field 'notional'", findings)
        self.assertFinding("maturity_date: invalid date", findings)
        self.assertFinding("direction: invalid direction", findings)
        self.assertFinding("missing field 'threshold_amount'", findings)

    def test_parameters(self) -> None:
        self.fixture["parameters"] = {"DaysPerYear": 360}
        self.assertEqual(self.findings(), [])
        self.fixture["parameters"] = {"Alpha": 1, "DaysPerYear": 0}
        findings = self.findings()
        self.assertFinding("unknown parameter 'Alpha'", findings)
        self.assertFinding("parameters.DaysPerYear", findings)
        self.fixture["parameters"] = []
        self.assertFinding("parameters: must be an object", self.findings())
        del self.fixture["parameters"]
        self.assertFinding("missing field 'parameters'", self.findings())

    def test_duplicate_trade_ids(self) -> None:
        self.fixture["trades"].append(copy.deepcopy(TRADE))
        self.assertFinding("duplicate trade id 'T1'", self.findings())

    def test_expected_names_its_fixture_and_regime(self) -> None:
        self.expected["regime"] = "CRR"
        self.assertFinding("match the file name", self.findings())
        self.expected["regime"] = "BCBS"
        self.assertFinding("fixture must name", self.findings(expected_name="missing.bcbs"))

    def test_trade_quantities_need_a_known_trade(self) -> None:
        self.expected["outputs"][1]["trade"] = "T9"
        self.expected["outputs"][0]["trade"] = "T1"
        findings = self.findings()
        self.assertFinding("needs 'trade' naming a trade", findings)
        self.assertFinding("takes no 'trade'", findings)

    def test_measure_forms(self) -> None:
        self.expected["outputs"][0]["unit"] = "USD"
        self.expected["outputs"][0]["tolerance"] = {"absolute": -1, "relative": 0}
        self.expected["outputs"][1]["value"] = 1
        findings = self.findings()
        self.assertFinding("unit must be EUR", findings)
        self.assertFinding("tolerance needs non-negative", findings)
        self.assertFinding("exactly one of value, text or error", findings)

    def test_text_and_numeric_quantities_are_not_interchangeable(self) -> None:
        self.expected["outputs"][0] = {"quantity": "exposure_value", "rule": "CRE52.1", "text": "569",
                                       "reference": PUBLISHED}
        self.expected["outputs"][1] = {"quantity": "trade_status", "rule": "CRE52", "trade": "T1",
                                       "value": 1, "unit": "1",
                                       "tolerance": {"absolute": 0, "relative": 0},
                                       "reference": PUBLISHED}
        findings = self.findings()
        self.assertFinding("is not a text quantity", findings)
        self.assertFinding("is a text quantity", findings)

    def test_rule_and_quantity_names(self) -> None:
        self.expected["outputs"][0]["rule"] = "Article 274"
        self.expected["outputs"][0]["quantity"] = "ead"
        findings = self.findings()
        self.assertFinding("rule must look like", findings)
        self.assertFinding("unknown quantity 'ead'", findings)

    def test_quantity_specific_units(self) -> None:
        units = {
            "exposure_value": "EUR", "replacement_cost": "EUR",
            "potential_future_exposure": "EUR", "aggregate_add_on": "EUR",
            "adjusted_notional": "EUR", "multiplier": "1",
            "supervisory_delta": "1", "supervisory_factor": "1",
            "margin_period_of_risk": "business_days",
            **{f"add_on.{asset}": "EUR" for asset in cases.ASSET_CLASSES},
        }
        for quantity, unit in units.items():
            for candidate in ("EUR", "USD", "1", "price", "business_days"):
                with self.subTest(quantity=quantity, unit=candidate):
                    output = dict(EXPECTED["outputs"][0], quantity=quantity, unit=candidate)
                    findings = cases.check_measure("output", output, self.fixture)
                    self.assertEqual(bool(findings), candidate != unit, findings)

    def test_lambda_unit_follows_referenced_trade(self) -> None:
        self.fixture["trades"].append(dict(TRADE, id="T2", asset_class="commodity"))
        for trade, unit in (("T1", "1"), ("T2", "price")):
            output = dict(EXPECTED["outputs"][0], quantity="lambda_shift", trade=trade, unit=unit)
            self.assertEqual(cases.check_measure("output", output, self.fixture), [])
            output["unit"] = "price" if unit == "1" else "1"
            self.assertFinding("unit must be", cases.check_measure("output", output, self.fixture))

    def test_non_finite_numbers_are_rejected(self) -> None:
        for number in (float("nan"), float("inf"), float("-inf")):
            with self.subTest(number=number):
                self.assertFalse(cases.is_number(number))
                self.fixture = copy.deepcopy(FIXTURE)
                self.expected = copy.deepcopy(EXPECTED)
                self.fixture["trades"][0]["market_value_amount"] = number
                self.assertFinding("not valid UTF-8 JSON", self.findings())
                self.fixture = copy.deepcopy(FIXTURE)
                self.expected["outputs"][0]["value"] = number
                self.assertFinding("not valid UTF-8 JSON", self.findings())
                self.expected["outputs"][0]["value"] = 1
                self.expected["outputs"][0]["tolerance"]["absolute"] = number
                self.assertFinding("not valid UTF-8 JSON", self.findings())

    def test_overflow_in_optional_reference_is_rejected_at_decode(self) -> None:
        self.expected["outputs"][0]["reference"]["derived_by"] = "OVERFLOW"
        for token in ("1e999", "-1e999"):
            with self.subTest(token=token):
                self.findings()
                target = self.root / cases.EXPECTED / "sample-case.bcbs.json"
                target.write_text(target.read_text().replace('"OVERFLOW"', token))
                self.assertFinding("non-finite JSON number", cases.run_check(self.root)["findings"])
        self.assertEqual(cases.finite_json_float("1.7976931348623157e308"), float("1.7976931348623157e308"))

    def test_numeric_overflow_is_rejected(self) -> None:
        self.assertTrue(cases.check_value("value", "number", json.loads("1e999")))
        self.assertTrue(cases.is_number(10 ** 1000))
        self.assertFalse(cases.is_number(True))

    def test_reference_classes(self) -> None:
        self.expected["outputs"][0]["reference"] = dict(PUBLISHED, locator="", source="WIKI")
        self.expected["outputs"][1]["reference"] = dict(PUBLISHED, **{"class": "independent",
                                                                     "reviewed_by": "Author"})
        findings = self.findings()
        self.assertFinding("a published reference needs 'locator'", findings)
        self.assertFinding("source 'WIKI' is not in the source register", findings)
        self.assertFinding("reviewed by someone other than its author", findings)

    def test_illustrative_prefix_matches_the_classes(self) -> None:
        illustrative = {"class": "illustrative", "derivation": "Format example"}
        for output in self.expected["outputs"]:
            output["reference"] = illustrative
        self.assertFinding("must start with illustrative-", self.findings())
        self.fixture["id"] = "illustrative-sample"
        self.expected["fixture"] = "illustrative-sample"
        self.expected["outputs"][0]["reference"] = PUBLISHED
        self.assertFinding("may hold only illustrative values",
                           self.findings("illustrative-sample", "illustrative-sample.bcbs"))

    def test_catalogue_ids_are_unique(self) -> None:
        self.expected["outputs"][1]["catalogue"] = "T04"
        self.assertFinding("catalogue T04 also used", self.findings())

    def test_fixture_without_expected_file(self) -> None:
        self.findings()
        (self.root / cases.EXPECTED / "sample-case.bcbs.json").unlink()
        self.assertFinding("no expected file", cases.run_check(self.root)["findings"])


if __name__ == "__main__":
    unittest.main()
