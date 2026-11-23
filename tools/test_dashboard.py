"""The Dashboard's specification and the reference of its figures (#675)."""
from __future__ import annotations

import copy
import unittest
from typing import Any

import dashboard as db
from dashboard import CheckRow, NettingSetInput, NettingSetResult, Run, RunInfo, TradeLeg

ADDONS = dict.fromkeys(db.ASSET_CLASSES, 0.0)
INFO = RunInfo("2026-09-30", "2026-11-23 10:00:00", "EUR", "CRR", "REGULATORY CRE52-CRR-2026-10-09", "3851427D")


def result(ns_id: str, regime: str, ead: float, *, v: float = 0.0, rc: float = 0.0, pfe: float = 0.0,
           margined: bool = False, status: str = "VALID", counterparty: str = "", cap: bool = False,
           add_ons: dict[str, float] | None = None) -> NettingSetResult:
    return NettingSetResult(ns_id, counterparty or f"CP {ns_id}", regime, margined, v, rc, add_ons or dict(ADDONS),
                            pfe, ead, cap, status)


def portfolio() -> Run:
    """Two regimes; an out-of-the-money set (A2), one holding collateral (A1), an invalid set (A3)."""
    results = [
        # A1: V 500, collateral held 200, so RC 300.
        result("A1", "CRR", 1400.0, v=500.0, rc=300.0, pfe=700.0, margined=True, counterparty="Bank X",
               add_ons={**ADDONS, "IR": 500.0, "FX": 200.0}),
        # A2: out of the money.
        result("A2", "CRR", 560.0, v=-100.0, rc=0.0, pfe=400.0, counterparty="Bank X",
               add_ons={**ADDONS, "EQ": 400.0}),
        result("A3", "CRR", 0.0, status="INVALID - clearing role MULTI", counterparty="Corp Y"),
        result("B1", "BCBS", 2800.0, v=1000.0, rc=1000.0, pfe=1000.0, counterparty="Bank Z", cap=True,
               add_ons={**ADDONS, "CO": 1000.0}),
        result("B2", "BCBS", 0.0, status="INCOMPLETE - EAD withheld"),
    ]
    inputs = {"A1": NettingSetInput("CI", "N"), "A2": NettingSetInput("", "CM"), "A3": NettingSetInput("NFC", "MULTI"),
              "B1": NettingSetInput("OFC", "CCP", "Y"), "B2": NettingSetInput("HH", "N")}
    legs = [
        TradeLeg("T1", "A1", "IR", "OK", 1000.0, 700.0, 0.5, "CRR"),
        TradeLeg("T2", "A1", "FX", "OK", 3000.0, -200.0, 3.0, "CRR"),
        TradeLeg("T3", "A2", "EQ", "OK", 2000.0, -100.0, 7.0, "CRR"),
        TradeLeg("T4", "A3", "IR", "OK", 9000.0, 900.0, 2.0, "CRR"),
        TradeLeg("T5", "A3", "IR", "EXCLUDED", 0.0, 0.0, 0.0, ""),
        TradeLeg("T6", "B1", "CO", "OK", 5000.0, 1000.0, 1.0, "BCBS"),
        TradeLeg("T7", "B2", "CO", "EXCLUDED", 0.0, 0.0, 0.0, "BCBS"),
    ]
    checks = [CheckRow("ERROR", "NS-ROLE"), CheckRow("WARNING", "TR-DATE"), CheckRow("WARNING", "TR-DATE"),
              CheckRow("ERROR", "TR-AC"), CheckRow("INFO", "RUN-WHATIF")]
    return Run(INFO, results, inputs, legs, checks, ["MATCH", "MATCH", "DIFFERENCE", "WITHHELD"], None)


class SpecificationTest(unittest.TestCase):
    def test_committed_specification_is_sound(self) -> None:
        self.assertEqual(db.check_spec(db.load_spec()), [])

    def test_documentation_defines_every_block(self) -> None:
        text = (db.ROOT / "docs" / "workbook" / "DASHBOARD.md").read_text(encoding="utf-8")
        missing = [b["name"] for b in db.load_spec()["blocks"] if f"`{b['name']}`" not in text]
        self.assertEqual(missing, [])

    def test_findings(self) -> None:
        spec = db.load_spec()
        broken = copy.deepcopy(spec)
        broken["blocks"][1]["name"] = broken["blocks"][0]["name"]
        broken["blocks"][2]["rows"][0]["kind"] = "colour"
        broken["blocks"][3]["rows"][1]["key"] = broken["blocks"][3]["rows"][0]["key"]
        broken["scales"]["steps"].pop()
        findings = db.check_spec(broken)
        self.assertEqual(len(findings), 4, findings)
        self.assertIn("duplicate name", " ".join(findings))
        self.assertIn("unknown kind colour", " ".join(findings))


class RuleTest(unittest.TestCase):
    def test_scale_steps(self) -> None:
        self.assertEqual([db.scale_step(a) for a in ([0.0], [999.0], [1000.0], [-2.5e6], [7e9], [1e15])],
                         [0, 0, 1, 2, 3, 3])

    def test_percent_rounds_halves_away_from_zero(self) -> None:
        self.assertEqual(db.percent_text(0.38455), "38.5%")
        self.assertEqual(db.percent_text(0.0), "0.0%")
        self.assertEqual(db.percent_text(1.0), "100.0%")

    def test_reading(self) -> None:
        self.assertEqual(db.reading(1000.0, 600.0, 400.0),
                         "Netting removes 40.0% of gross positive value; collateral a further 20.0%.")
        self.assertEqual(db.reading(1000.0, 600.0, 700.0),
                         "Netting removes 40.0% of gross positive value; collateral and margin terms add 10.0%.")
        self.assertEqual(db.reading(0.0, 0.0, 0.0), "No trade has a positive market value.")


class FigureTest(unittest.TestCase):
    def setUp(self) -> None:
        self.spec = db.load_spec()
        self.blocks = db.compute(portfolio(), self.spec)

    def row(self, block: str, key: str) -> list[Any]:
        keys = [r["key"] for r in next(b for b in self.spec["blocks"] if b["name"] == block)["rows"]]
        return self.blocks[block][keys.index(key)]

    def test_every_block_has_its_shape(self) -> None:
        self.assertEqual(db.shape_problems(self.blocks, self.spec), [])

    def test_scales(self) -> None:
        # Largest exposure amount 2800 -> thousand; largest notional 5000 -> thousand.
        self.assertEqual(self.blocks["Dashboard_Scale"], [[1000, "EUR thousand"], [1000, "EUR thousand"]])

    def test_key_figures_keep_regimes_apart(self) -> None:
        self.assertEqual(self.row("Dashboard_KeyFigures", "ead"), [1.96, 2.8])
        self.assertEqual(self.row("Dashboard_KeyFigures", "netting_sets"), [3, 2])
        self.assertEqual(self.row("Dashboard_KeyFigures", "valid"), [2, 1])
        self.assertEqual(self.row("Dashboard_KeyFigures", "invalid"), [1, 0])
        self.assertEqual(self.row("Dashboard_KeyFigures", "incomplete"), [0, 1])
        self.assertEqual(self.row("Dashboard_KeyFigures", "margined"), [1, 0])
        self.assertEqual(self.row("Dashboard_KeyFigures", "cap_applied"), [0, 1])
        self.assertEqual(self.row("Dashboard_KeyFigures", "counterparties"), [2, 2])
        self.assertEqual(self.row("Dashboard_Counts", "counterparties"), [4])

    def test_counts_and_checks(self) -> None:
        self.assertEqual(self.blocks["Dashboard_Counts"], [[7], [5], [2], [0], [4], [2], [2], [1]])
        self.assertEqual(self.blocks["Dashboard_DataQuality"], [[2], [2]])
        self.assertEqual(self.blocks["Dashboard_Codes"][:3],
                         [["TR-DATE", "WARNING", 2], ["NS-ROLE", "ERROR", 1], ["TR-AC", "ERROR", 1]])
        self.assertEqual(self.blocks["Dashboard_Codes"][3], [None, None, None])

    def test_breakdowns_take_valid_netting_sets_only(self) -> None:
        self.assertEqual(self.row("Dashboard_Sector", "CI"), [1.4, 0.0])
        self.assertEqual(self.row("Dashboard_Sector", "NC"), [0.56, 0.0])
        self.assertEqual(self.row("Dashboard_Sector", "NFC"), [0.0, 0.0])
        self.assertEqual(self.row("Dashboard_ClearingRole", "CM"), [0.56, 0.0])
        self.assertEqual(self.row("Dashboard_ClearingRole", "CCP_Q"), [0.0, 2.8])
        self.assertEqual(self.row("Dashboard_AddOn", "EQ"), [0.4, 0.0])
        self.assertEqual(self.blocks["Dashboard_Composition"], [[0.3, 1.0], [1.1, 1.0], [1.4, 0.0], [0.56, 2.8]])

    def test_ranking_and_concentration(self) -> None:
        top = self.blocks["Dashboard_TopCRR"]
        self.assertEqual(top[0][:3], ["A1", "Bank X", 1.4])
        self.assertAlmostEqual(top[0][3], 1400 / 1960)
        self.assertEqual(top[2], [None] * 4)
        share, herfindahl = self.row("Dashboard_Concentration", "top10_share"), \
            self.row("Dashboard_Concentration", "herfindahl")
        self.assertEqual(share, [1.0, 1.0])
        self.assertAlmostEqual(herfindahl[0], (1400 / 1960) ** 2 + (560 / 1960) ** 2)
        self.assertEqual(herfindahl[1], 1.0)

    def test_maturity_and_asset_classes(self) -> None:
        self.assertEqual(self.blocks["Dashboard_Maturity"], [[1.0, 0.0], [3.0, 5.0], [2.0, 0.0]])
        ir = self.blocks["Dashboard_AssetClass"][0]
        self.assertEqual(ir, [1, 1.0, 0, 0.0])

    def test_notional_to_ead(self) -> None:
        path = {key: self.row("Dashboard_NotionalToEAD", key)
                for key in ("notional", "gross_positive", "netted", "rc", "pfe", "ead")}
        self.assertEqual(path["notional"][::2], [6.0, 5.0])
        self.assertEqual(path["gross_positive"][::2], [0.7, 1.0])
        self.assertEqual(path["netted"][::2], [0.5, 1.0])
        self.assertEqual(path["rc"][::2], [0.3, 1.0])
        self.assertAlmostEqual(path["ead"][1], 1960 / 6000)
        self.assertEqual(self.blocks["Dashboard_Reading"][0],
                         ["Netting removes 28.6% of gross positive value; collateral a further 28.6%."])
        self.assertEqual(self.blocks["Dashboard_Reading"][1],
                         ["Netting removes 0.0% of gross positive value; collateral a further 0.0%."])

    def test_reconciliation(self) -> None:
        self.assertEqual(self.blocks["Dashboard_Reconciliation"],
                         [["Y", "N"], [2, 0], [1, 0], [0, 0], [1, 0]])

    def test_empty_portfolio(self) -> None:
        blocks = db.compute(Run(INFO, [], {}, [], [CheckRow("INFO", "RUN-CLEAN")]), self.spec)
        self.assertEqual(db.shape_problems(blocks, self.spec), [])
        self.assertEqual(blocks["Dashboard_Scale"], [[1, "EUR"], [1, "EUR"]])
        self.assertEqual(blocks["Dashboard_Concentration"], [[None, None], [None, None]])
        self.assertEqual(blocks["Dashboard_Reading"], [[""], [""]])
        self.assertEqual(blocks["Dashboard_Counts"][-3:], [[0], [0], [0]])


if __name__ == "__main__":
    unittest.main()
