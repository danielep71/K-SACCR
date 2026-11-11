"""The synthetic portfolio of the R cross-check and its comparison with a workbook's evidence (#352)."""
from __future__ import annotations

import copy
import datetime as dt
import json
import tempfile
import unittest
from pathlib import Path
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
    return {"versions": {"R": "4.3.3", "SACCR": "3.4"}, "results": results,
            "portfolio_sha256": portfolio.fingerprint(generated)}


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
                if trade["asset_class"] == "foreign_exchange":
                    # The engine rejects a notional currency that is not in the pair.
                    self.assertIn("EUR", trade["risk_factor"].split("/"), trade["id"])
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

    def test_package_results_of_another_portfolio_are_refused(self) -> None:
        # Seeds 352 and 361 both give RS-001 five trades with one netting set:
        # only the fingerprint tells the two portfolios apart.
        other = portfolio.generate(361, 1)
        mine = portfolio.generate(352, 1)
        self.assertEqual([f["id"] for f in other["netting_sets"]], [f["id"] for f in mine["netting_sets"]])
        package = package_results(mine)
        text = evidence(mine, package)
        self.assertTrue(all(r["agrees"] for r in portfolio.compare(mine, package, text)))
        with self.assertRaisesRegex(portfolio.EvidenceError, "another portfolio"):
            portfolio.compare(other, package, text)
        del package["portfolio_sha256"]
        with self.assertRaisesRegex(portfolio.EvidenceError, "another portfolio"):
            portfolio.compare(mine, package, text)

    def test_generate_removes_earlier_results(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            out = Path(temp)
            for name in portfolio.DOWNSTREAM:
                (out / name).write_text("stale", encoding="utf-8")
            self.assertEqual(portfolio.main(["generate", "--out", temp, "--netting-sets", "2"]), 0)
            self.assertEqual(sorted(p.name for p in out.iterdir()),
                             sorted(["portfolio.json", portfolio.NS_FILE, portfolio.TR_FILE]))


class CaseTests(unittest.TestCase):
    def setUp(self) -> None:
        self.generated = portfolio.generate(352, 4)
        self.package = package_results(self.generated)
        rows = portfolio.compare(self.generated, self.package, evidence(self.generated, self.package))
        self.comparison = portfolio.comparison_csv(rows)
        self.ids = [f["id"] for f in self.generated["netting_sets"]]

    def test_agreeing_netting_sets_become_independent_cases(self) -> None:
        files = portfolio.cases(self.generated, self.package, self.comparison, self.ids[:2], "2026-11-11")
        self.assertEqual(sorted(files), sorted(
            [f"tests/fixtures/rsaccr-portfolio-{i.lower()}.json" for i in self.ids[:2]] +
            [f"tests/expected/rsaccr-portfolio-{i.lower()}.crr.json" for i in self.ids[:2]]))
        expected = files[f"tests/expected/rsaccr-portfolio-{self.ids[0].lower()}.crr.json"]
        self.assertEqual(expected["regime"], "CRR")
        self.assertEqual({o["reference"]["class"] for o in expected["outputs"]}, {"independent"})
        self.assertEqual({o["reference"]["source"] for o in expected["outputs"]}, {"IMPL-R-SACCR"})
        quantities = [o["quantity"] for o in expected["outputs"]]
        self.assertEqual(quantities, ["replacement_cost", "add_on.interest_rate", "add_on.foreign_exchange",
                                      "aggregate_add_on", "multiplier", "potential_future_exposure",
                                      "exposure_value"])
        self.assertEqual(next(o for o in expected["outputs"] if o["quantity"] == "multiplier")["unit"], "1")
        fixture = files[f"tests/fixtures/rsaccr-portfolio-{self.ids[0].lower()}.json"]
        self.assertEqual(fixture["id"], f"rsaccr-portfolio-{self.ids[0].lower()}")
        self.assertEqual(fixture["netting_set"]["id"], self.ids[0])

    def test_refusals(self) -> None:
        first = self.ids[0]
        differing = self.comparison.replace(f"{first};EAD;", f"{first};EAD;x", 1)
        rows = [r.split(";") for r in differing.splitlines()]
        rows = [r[:5] + ["N"] + r[6:] if r[0] == first and r[1] == "EAD" else r for r in rows]
        differing = "\n".join(";".join(r) for r in rows) + "\n"
        capped = copy.deepcopy(self.package)
        capped["results"][0]["cap_applied"] = True
        stale = copy.deepcopy(self.package)
        stale["portfolio_sha256"] = "0" * 64
        for package, comparison, chosen, message in (
                (self.package, differing, [first], "did not agree"),
                (capped, self.comparison, [first], "cap binds"),
                (stale, self.comparison, [first], "another portfolio"),
                (self.package, self.comparison, ["RS-999"], "not a netting set")):
            with self.subTest(message=message), self.assertRaisesRegex(portfolio.EvidenceError, message):
                portfolio.cases(self.generated, package, comparison, chosen, "2026-11-11")

    def test_committed_cases_are_the_generated_netting_sets(self) -> None:
        generated = {f["id"]: f for f in portfolio.generate(352, 40)["netting_sets"]}
        committed = sorted((ROOT / "tests" / "fixtures").glob("rsaccr-portfolio-*.json"))
        self.assertTrue(committed)
        for path in committed:
            with self.subTest(case=path.stem):
                data = json.loads(path.read_text(encoding="utf-8"))
                ns_id = data["netting_set"]["id"]
                self.assertEqual(data, portfolio.case_fixture(generated[ns_id], 352))
                expected = json.loads((ROOT / "tests" / "expected" / f"{path.stem}.crr.json")
                                      .read_text(encoding="utf-8"))
                self.assertEqual({o["reference"]["class"] for o in expected["outputs"]}, {"independent"})


if __name__ == "__main__":
    unittest.main()
