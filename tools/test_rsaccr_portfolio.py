"""The synthetic portfolio of the R cross-check and its comparison with a workbook's evidence (#352)."""
from __future__ import annotations

import copy
import datetime as dt
import json
import unittest
from typing import Any

import csv_contract
import rsaccr_crosscheck as crosscheck
import rsaccr_portfolio as portfolio
from rsaccr_portfolio import ROOT


def package_results(generated: dict[str, Any]) -> dict[str, Any]:
    """Package figures for every netting set, as tools/rsaccr_run.R writes them."""
    results = []
    for index, fixture in enumerate(generated["netting_sets"], start=1):
        results.append({"netting_set": fixture["id"], "fixture": fixture["id"], "RC": 1000.0 * index,
                        "add_on": 2500.5 * index, "multiplier": 0.987654321, "PFE": 2469.75 * index,
                        "EAD": 1.4 * (1000.0 + 2469.75) * index,
                        "add_on_by_class": {"IRD": 1500.5 * index, "FX": 1000.0 * index}})
    return {"versions": {"R": "4.3.3", "SACCR": "3.4"}, "results": results}


def evidence(generated: dict[str, Any], package: dict[str, Any], **changes: str) -> str:
    """An evidence CSV whose Results equal the package figures."""
    header = ["NettingSetID", "Counterparty", "Regime", "Margined", "Trades used", "RC", "AddOn IR", "AddOn FX",
              "AddOn CR", "AddOn EQ", "AddOn CO", "AddOn OT", "AddOn aggregate", "Multiplier", "PFE", "EAD",
              "Status"]
    run = {"format": "K-SACCR evidence 1", "results_status": "CURRENT", "reporting_date": "2026-09-30",
           "reporting_ccy": "EUR"}
    run.update({k: v for k, v in changes.items() if k in run})
    lines = [f"RUN;{key};{value}" for key, value in run.items()]
    lines.append("RESULTS;" + ";".join(header))
    by_id = {r["netting_set"]: r for r in package["results"]}
    for fixture in generated["netting_sets"]:
        r = by_id[fixture["id"]]
        classes = r["add_on_by_class"]
        values = [fixture["id"], "x", "CRR", "N", str(len(fixture["trades"])), repr(r["RC"]),
                  repr(classes.get("IRD", 0.0)), repr(classes.get("FX", 0.0)), "0", "", "0", "0",
                  repr(r["add_on"]), repr(r["multiplier"]), repr(r["PFE"]), repr(r["EAD"]), "VALID"]
        lines.append("RESULTS;" + ";".join(values))
    lines.append("RESULTS;TOTAL CRR;;CRR;;;;;;;;;;;;;1;VALID")
    lines.append("CHECKS;Severity;Message")
    return "\n".join(lines) + "\n"


class GenerationTests(unittest.TestCase):
    def setUp(self) -> None:
        self.generated = portfolio.generate(352, 40)

    def test_deterministic(self) -> None:
        self.assertEqual(json.dumps(portfolio.generate(352, 40)), json.dumps(self.generated))
        self.assertNotEqual(json.dumps(portfolio.generate(353, 40)), json.dumps(self.generated))
        self.assertEqual(len(portfolio.generate(352, 5)["netting_sets"]), 5)

    def test_every_netting_set_is_one_the_package_runs(self) -> None:
        for fixture in self.generated["netting_sets"]:
            with self.subTest(netting_set=fixture["id"]):
                translated = crosscheck.translate_fixture(fixture)
                self.assertEqual(len(translated["trades"]), len(fixture["trades"]))
                self.assertEqual(fixture["trades"][0]["instrument"], "linear")

    def test_avoids_the_package_differences(self) -> None:
        valuation = dt.date(2026, 9, 30)
        classes = set()
        for fixture in self.generated["netting_sets"]:
            ns = fixture["netting_set"]
            self.assertEqual(ns["clearing_role"], "none")
            self.assertFalse(ns["large_or_illiquid"] or ns["margin_disputes"])
            for trade in fixture["trades"]:
                classes.add((trade["asset_class"], trade["instrument"]))
                days = (dt.date.fromisoformat(trade["maturity_date"]) - valuation).days
                self.assertGreaterEqual(days, 30, trade["id"])
                if trade["asset_class"] == "interest_rate":
                    self.assertNotEqual(days, 365, trade["id"])
                if trade["instrument"] == "option":
                    self.assertGreaterEqual(min(trade["underlying_price"], trade["strike_price"]), 0.01)
        # Every asset class appears, with options for interest rates and equity.
        self.assertTrue({"interest_rate", "foreign_exchange", "credit", "equity", "commodity", "other"}
                        <= {c for c, _ in classes})
        self.assertTrue({("interest_rate", "option"), ("equity", "option")} <= classes)
        self.assertTrue(any(f["netting_set"]["margined"] for f in self.generated["netting_sets"]))
        # Two trades at most on one commodity in a netting set.
        for fixture in self.generated["netting_sets"]:
            names = [t["risk_factor"] for t in fixture["trades"] if t["asset_class"] == "commodity"]
            self.assertTrue(all(names.count(n) <= 2 for n in names), fixture["id"])

    def test_csv_files_follow_the_contract(self) -> None:
        contract = csv_contract.load_contract(ROOT)
        files = portfolio.csv_files(self.generated, contract)
        records, findings = csv_contract.validate(files[portfolio.TR_FILE].encode(), "Trades", contract)
        self.assertEqual(findings, [])
        self.assertEqual(len(records), sum(len(f["trades"]) for f in self.generated["netting_sets"]))
        records, _ = csv_contract.validate(files[portfolio.NS_FILE].encode(), "NettingSets", contract)
        self.assertEqual({r["regime_override"] for r in records}, {"CRR"})
        broken = copy.deepcopy(self.generated)
        broken["netting_sets"][0]["trades"][0]["notional_amount"] = None
        with self.assertRaisesRegex(ValueError, "break the input contract"):
            portfolio.csv_files(broken, contract)

    def test_field_format(self) -> None:
        self.assertEqual([portfolio.field(v) for v in (None, True, False, 12000000.0, -0.5, 1e-05, 7, "X")],
                         ["", "Y", "N", "12000000", "-0.5", "0.00001", "7", "X"])


class ComparisonTests(unittest.TestCase):
    def setUp(self) -> None:
        self.generated = portfolio.generate(352, 4)
        self.package = package_results(self.generated)

    def rows(self, text: str) -> list[dict[str, Any]]:
        return portfolio.compare(self.generated, self.package, text)

    def test_identical_figures_agree(self) -> None:
        rows = self.rows(evidence(self.generated, self.package))
        self.assertTrue(all(row["agrees"] for row in rows), [r for r in rows if not r["agrees"]])
        self.assertEqual(len(rows), 4 * 14)
        self.assertIn("RESULT compared=56; differ=0", portfolio.summary(self.generated, self.package, rows))

    def test_a_difference_beyond_the_tolerance_differs(self) -> None:
        text = evidence(self.generated, self.package)
        first = self.package["results"][0]
        # 1e-9 relative is inside the tolerance; 1e-6 relative is not.
        inside = self.rows(text.replace(repr(first["EAD"]), repr(first["EAD"] * (1 + 5e-10)), 1))
        self.assertTrue(all(row["agrees"] for row in inside))
        outside = self.rows(text.replace(repr(first["EAD"]), repr(first["EAD"] * (1 + 1e-6)), 1))
        self.assertEqual([(r["netting_set"], r["quantity"]) for r in outside if not r["agrees"]],
                         [(first["netting_set"], "EAD")])
        csv_text = portfolio.comparison_csv(outside)
        self.assertIn(f"{first['netting_set']};EAD;", csv_text)
        self.assertTrue(csv_text.startswith("netting_set;quantity;workbook;package;difference;agrees;classification"))

    def test_status_regime_and_trades_used(self) -> None:
        ns_id = self.generated["netting_sets"][1]["id"]
        text = evidence(self.generated, self.package)
        line = next(ln for ln in text.splitlines() if ln.startswith(f"RESULTS;{ns_id};"))
        changed = line.replace(";VALID", ";INVALID").replace(";CRR;", ";BCBS;")
        rows = self.rows(text.replace(line, changed))
        self.assertEqual(sorted(r["quantity"] for r in rows if not r["agrees"]), ["Regime", "Status"])

    def test_missing_and_extra_netting_sets(self) -> None:
        text = evidence(self.generated, self.package)
        ns_id = self.generated["netting_sets"][0]["id"]
        line = next(ln for ln in text.splitlines() if ln.startswith(f"RESULTS;{ns_id};"))
        rows = self.rows(text.replace(line + "\n", "") + line.replace(ns_id, "EXTRA-1") + "\n")
        self.assertEqual(sorted((r["netting_set"], r["workbook"]) for r in rows if not r["agrees"]),
                         [("EXTRA-1", "not in the portfolio"), (ns_id, "missing")])

    def test_evidence_of_another_run_is_refused(self) -> None:
        for change in ({"reporting_date": "2026-06-30"}, {"reporting_ccy": "USD"},
                       {"results_status": "OUT OF DATE"}):
            with self.subTest(change=change), self.assertRaises(portfolio.EvidenceError):
                self.rows(evidence(self.generated, self.package, **change))
        with self.assertRaisesRegex(portfolio.EvidenceError, "no RESULTS"):
            self.rows("RUN;format;K-SACCR evidence 1\n")


if __name__ == "__main__":
    unittest.main()
