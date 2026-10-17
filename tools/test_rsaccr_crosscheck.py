"""Tests for tools/rsaccr_crosscheck.py: the translation and comparison, without R."""

from __future__ import annotations

import copy
import json
import sys
import tempfile
import unittest
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent))

import rsaccr_crosscheck as xc  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent


def fixture(name: str) -> dict[str, Any]:
    loaded: dict[str, Any] = json.loads((ROOT / "tests" / "fixtures" / f"{name}.json").read_text(encoding="utf-8"))
    return loaded


class TranslateTests(unittest.TestCase):
    def test_margined_example_becomes_one_counterparty_with_csa(self) -> None:
        netting_set, notes = xc.translate_fixture(fixture("cre99-example-5"))
        self.assertEqual(notes, [])
        classes = [t["class"] for t in netting_set["trades"]]
        self.assertEqual(classes, ["IRDSwap", "IRDSwap", "IRDSwaption", "Commodity", "Commodity", "Commodity"])
        swaption = netting_set["trades"][2]["args"]
        self.assertAlmostEqual(swaption["Si"], 365 / 365)
        self.assertEqual((swaption["OptionType"], swaption["BuySell"], swaption["Counterparty"]),
                         ("Put", "Buy", "EX5"))
        csa = netting_set["csa"]
        self.assertEqual((csa["remargin_freq"], csa["mta"], csa["nica"], csa["threshold"]), (5, 5000, 150000, 0))
        self.assertEqual(csa["trade_groups"], ["Commodity", "IRD"])
        self.assertEqual([c["type"] for c in netting_set["collateral"]], ["VariationMargin", "ICA"])

    def test_variation_margin_is_entered_even_when_zero(self) -> None:
        data = fixture("cre99-example-5")
        data["netting_set"]["variation_margin_net_amount"] = 0
        data["netting_set"]["independent_collateral_net_amount"] = 0
        netting_set, _ = xc.translate_fixture(data)
        self.assertEqual(netting_set["collateral"],
                         [{"id": "csa_EX5_vm", "csa_id": "csa_EX5", "type": "VariationMargin", "amount": 0}])

    def test_remargining_period_defaults_to_one_business_day(self) -> None:
        data = fixture("cre99-example-5")
        for period in (None, 0):
            data["netting_set"]["remargin_period_business_days"] = period
            netting_set, _ = xc.translate_fixture(data)
            self.assertEqual(netting_set["csa"]["remargin_freq"], 1)

    def test_days_per_year_and_commodity_hedging_set(self) -> None:
        netting_set, _ = xc.translate_fixture(fixture("cre99-example-3"))
        first = netting_set["trades"][0]["args"]
        self.assertAlmostEqual(first["Ei"], 0.75)
        self.assertEqual((first["SubClass"], first["commodity_type"]), ("Energy", "CRUDE OIL"))

    def test_credit_quality_step_maps_to_rating_with_same_factor(self) -> None:
        netting_set, _ = xc.translate_fixture(fixture("illustrative-credit-quality-step"))
        self.assertEqual(netting_set["trades"][0]["args"]["SubClass"], "A")

    def test_matured_trade_is_left_out_with_a_note(self) -> None:
        netting_set, notes = xc.translate_fixture(fixture("illustrative-partly-rejected"))
        self.assertEqual([t["id"] for t in netting_set["trades"]], ["PR-1"])
        self.assertEqual(len(notes), 1)

    def test_unsupported_cases_raise(self) -> None:
        with self.assertRaisesRegex(xc.Unsupported, "commodity option"):
            xc.translate_fixture(fixture("illustrative-negative-commodity-price"))
        with self.assertRaisesRegex(xc.Unsupported, "IRBucketOffset"):
            xc.translate_fixture(fixture("illustrative-ir-sum-of-absolutes"))
        data = fixture("cre99-example-1")
        data["trades"][0]["maturity_date"] = "2037-09-27"
        with self.assertRaisesRegex(xc.Unsupported, "maturity differs"):
            xc.translate_fixture(data)
        data = fixture("cre99-example-5")
        data["netting_set"]["large_or_illiquid"] = True
        with self.assertRaisesRegex(xc.Unsupported, "MPOR"):
            xc.translate_fixture(data)
        data = fixture("cre99-example-1")
        data["netting_set"]["alpha_factor"] = 1.0
        with self.assertRaisesRegex(xc.Unsupported, "alpha"):
            xc.translate_fixture(data)


class CompareTests(unittest.TestCase):
    def test_rows_use_the_expected_tolerance_and_skip_text(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            (root / "tests" / "expected").mkdir(parents=True)
            reference = {"class": "illustrative"}
            outputs = [
                {"quantity": "exposure_value", "value": 100.0,
                 "tolerance": {"absolute": 1, "relative": 0}, "reference": reference},
                {"quantity": "replacement_cost", "value": 50.0,
                 "tolerance": {"absolute": 0, "relative": 0.01}, "reference": reference},
                {"quantity": "supervisory_delta", "trade": "T1", "value": 0.5,
                 "tolerance": {"absolute": 0, "relative": 0}, "reference": reference},
                {"quantity": "netting_set_status", "text": "VALID"},
            ]
            (root / "tests" / "expected" / "case.crr.json").write_text(
                json.dumps({"fixture": "case", "outputs": outputs}), encoding="utf-8")
            result = {"fixture": "case", "EAD": 100.9, "RC": 50.6, "add_on_by_class": {},
                      "trades": {"T1": {"supervisory_delta": 0.5, "adjusted_notional": 1.0}}}
            rows = xc.compare(root, [copy.deepcopy(result)])
        self.assertEqual([(r["quantity"], r["agrees"]) for r in rows],
                         [("exposure_value", True), ("replacement_cost", False), ("supervisory_delta", True)])


if __name__ == "__main__":
    unittest.main()
