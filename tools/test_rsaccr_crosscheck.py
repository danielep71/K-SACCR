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
        netting_set = xc.translate_fixture(fixture("cre99-example-5"))
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
        netting_set = xc.translate_fixture(data)
        self.assertEqual(netting_set["collateral"],
                         [{"id": "csa_EX5_vm", "csa_id": "csa_EX5", "type": "VariationMargin", "amount": 0}])

    def test_remargining_period_defaults_to_one_business_day(self) -> None:
        data = fixture("cre99-example-5")
        for period in (None, 0):
            data["netting_set"]["remargin_period_business_days"] = period
            netting_set = xc.translate_fixture(data)
            self.assertEqual(netting_set["csa"]["remargin_freq"], 1)

    def test_days_per_year_and_commodity_hedging_set(self) -> None:
        netting_set = xc.translate_fixture(fixture("cre99-example-3"))
        first = netting_set["trades"][0]["args"]
        self.assertAlmostEqual(first["Ei"], 0.75)
        self.assertEqual((first["SubClass"], first["commodity_type"]), ("Energy", "CRUDE OIL"))

    def test_credit_quality_step_maps_to_rating_with_same_factor(self) -> None:
        netting_set = xc.translate_fixture(fixture("illustrative-credit-quality-step"))
        self.assertEqual(netting_set["trades"][0]["args"]["SubClass"], "A")

    def test_credit_distinct_maturity_is_unsupported(self) -> None:
        data = fixture("illustrative-credit-quality-step")
        data["trades"][0]["maturity_date"] = "2037-09-27"
        with self.assertRaisesRegex(xc.Unsupported, "maturity differs"):
            xc.translate_fixture(data)

    def test_matured_trade_makes_the_netting_set_not_run(self) -> None:
        with self.assertRaisesRegex(xc.Unsupported, "PR-2 has matured"):
            xc.translate_fixture(fixture("illustrative-partly-rejected"))

    def test_reference_with_two_sub_classes_is_not_run(self) -> None:
        data = fixture("cre99-example-2")
        data["trades"][1]["risk_factor"] = " firm a "
        with self.assertRaisesRegex(xc.Unsupported, "two sub-classes"):
            xc.translate_fixture(data)

    def test_reversed_fx_pair_is_put_in_engine_order_with_its_direction_flipped(self) -> None:
        data = fixture("illustrative-other-risks")
        data["trades"] = data["trades"][:1]
        data["trades"][0].update({"asset_class": "foreign_exchange", "risk_factor": "usd/eur"})
        args = xc.translate_fixture(data)["trades"][0]["args"]
        self.assertEqual((args["ccyPair"], args["BuySell"]), ("EUR/USD", "Sell"))

    def test_references_are_trimmed_and_upper_case(self) -> None:
        data = fixture("cre99-example-2")
        data["trades"][0]["risk_factor"] = " firm a "
        self.assertEqual(xc.translate_fixture(data)["trades"][0]["args"]["RefEntity"], "FIRM A")

    def test_sold_options_only_and_expired_options_are_not_run(self) -> None:
        data = fixture("illustrative-negative-ir-forward")
        data["trades"][0]["direction"] = "short"
        with self.assertRaisesRegex(xc.Unsupported, "sold option"):
            xc.translate_fixture(data)
        data = fixture("cre99-example-1")
        data["trades"][2]["start_date"] = data["trades"][2]["option_expiry_date"] = data["valuation_date"]
        with self.assertRaisesRegex(xc.Unsupported, "expires on or before"):
            xc.translate_fixture(data)

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
    def test_combined_tolerances_use_maximum_and_compare_pfe(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            (root / "tests" / "expected").mkdir(parents=True)
            outputs = [{"quantity": "potential_future_exposure", "value": 100.0,
                        "tolerance": {"absolute": 1.0, "relative": 0.02},
                        "reference": {"class": "illustrative"}}]
            (root / "tests" / "expected" / "case.crr.json").write_text(
                json.dumps({"fixture": "case", "outputs": outputs}), encoding="utf-8")
            for value, agrees in [(102.0, True), (102.5, False), (97.5, False)]:
                with self.subTest(value=value):
                    rows, _ = xc.compare(root, [{"fixture": "case", "PFE": value}])
                    self.assertEqual(len(rows), 1)
                    self.assertEqual(rows[0]["quantity"], "potential_future_exposure")
                    self.assertEqual(rows[0]["agrees"], agrees)

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
            rows, _ = xc.compare(root, [copy.deepcopy(result)])
        self.assertEqual([(r["quantity"], r["agrees"]) for r in rows],
                         [("exposure_value", True), ("replacement_cost", False), ("supervisory_delta", True)])


    def test_bcbs_expectations_with_a_regime_difference_are_left_out(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            for folder in ("expected", "fixtures"):
                (root / "tests" / folder).mkdir(parents=True)
            data = fixture("illustrative-other-risks")
            (root / "tests" / "fixtures" / "illustrative-other-risks.json").write_text(
                json.dumps(data), encoding="utf-8")
            outputs = [{"quantity": "exposure_value", "value": 1232.0,
                        "tolerance": {"absolute": 0, "relative": 0}, "reference": {"class": "illustrative"}}]
            for regime in ("CRR", "BCBS"):
                (root / "tests" / "expected" / f"illustrative-other-risks.{regime.lower()}.json").write_text(
                    json.dumps({"fixture": "illustrative-other-risks", "regime": regime, "outputs": outputs}),
                    encoding="utf-8")
            rows, left_out = xc.compare(root, [{"fixture": "illustrative-other-risks", "EAD": 1232.0}])
        self.assertEqual([r["case"] for r in rows], ["illustrative-other-risks.crr"])
        self.assertEqual([c["case"] for c in left_out], ["illustrative-other-risks.bcbs"])


class RegimeDifferenceTests(unittest.TestCase):
    def test_recorded_differences_are_detected(self) -> None:
        self.assertIsNone(xc.regime_difference(fixture("cre99-example-1")))
        self.assertIn("other risks", xc.regime_difference(fixture("illustrative-other-risks")) or "")
        self.assertIn("climatic", xc.regime_difference(fixture("illustrative-climatic-commodity")) or "")
        self.assertIn("credit quality", xc.regime_difference(fixture("illustrative-credit-quality-step")) or "")
        self.assertIn("lambda", xc.regime_difference(fixture("illustrative-negative-ir-forward")) or "")
        data = fixture("illustrative-ir-swap-unmargined")
        data["trades"][0]["end_date"] = data["trades"][0]["maturity_date"] = "2027-01-02"
        self.assertIn("bucket", xc.regime_difference(data) or "")
        data = fixture("cre99-example-1")
        data["trades"][2]["lambda_price"] = 0.01
        self.assertIn("lambda", xc.regime_difference(data) or "")

    def test_equity_option_with_a_distinct_expiry_is_not_run(self) -> None:
        data = fixture("illustrative-negative-commodity-price")
        trade = data["trades"][0]
        trade.update({"asset_class": "equity", "sub_class": "SINGLE", "underlying_price": 100,
                      "strike_price": 100, "option_expiry_date": "2026-12-31"})
        with self.assertRaisesRegex(xc.Unsupported, "expiry differs"):
            xc.translate_fixture(data)
        trade["option_expiry_date"] = trade["maturity_date"]
        self.assertEqual(xc.translate_fixture(data)["trades"][0]["class"], "EquityOptionSingle")
        trade["lambda_price"] = 5
        with self.assertRaisesRegex(xc.Unsupported, "entered lambda"):
            xc.translate_fixture(data)
        trade["lambda_price"] = None
        trade["option_expiry_date"] = None
        with self.assertRaisesRegex(xc.Unsupported, "without expiry"):
            xc.translate_fixture(data)

    def test_settled_to_market_trade_takes_the_time_to_its_reset(self) -> None:
        data = fixture("illustrative-settled-to-market-daily")
        data["trades"] = data["trades"][:1]
        netting_set = xc.translate_fixture(data)
        self.assertAlmostEqual(netting_set["trades"][0]["args"]["Ei"], 1 / 365)
        # The swap's E must stay the contractual end, which the package cannot
        # separate from M; the call's expiry differs from the reset.
        with self.assertRaisesRegex(xc.Unsupported, "maturity differs"):
            xc.translate_fixture(fixture("illustrative-settled-to-market-periodic"))
        with self.assertRaisesRegex(xc.Unsupported, "expiry differs"):
            xc.translate_fixture(fixture("illustrative-settled-to-market-daily"))

    def test_reset_dates_the_engine_rejects_are_not_run(self) -> None:
        for settlement, reset in (("ordinary", "2026-10-01"), ("settled_to_market", None),
                                  ("settled_to_market", "2026-09-30"), ("settled_to_market", "2026-12-19")):
            with self.subTest(settlement=settlement, reset=reset):
                data = fixture("illustrative-settled-to-market-daily")
                data["trades"] = data["trades"][:1]
                data["trades"][0].update(settlement_type=settlement, next_reset_date=reset)
                with self.assertRaisesRegex(xc.Unsupported, "reset date"):
                    xc.translate_fixture(data)

    def test_trade_ending_on_or_before_its_start_is_not_run(self) -> None:
        data = fixture("illustrative-ir-swap-unmargined")
        data["trades"][0]["start_date"] = data["trades"][0]["end_date"]
        with self.assertRaisesRegex(xc.Unsupported, "ends on or before"):
            xc.translate_fixture(data)


if __name__ == "__main__":
    unittest.main()
