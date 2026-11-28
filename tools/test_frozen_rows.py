"""The frozen rows of the workbook template keep the same height from sheet to sheet (#777).

Moving from one sheet to another must not make the view jump (owner, 2026-11-25): the sheets of a group
freeze rows of the same total height, and the tables of the output sheets start their column headers
at the same point, with the same height, whatever the row they are on.
"""
from __future__ import annotations

import re
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "src" / "workbook" / "SACCR_Template.xlsx"
TITLE_BAND = [24.0, 26.1, 30.0]   # K-SACCR, the sheet title, the description
HEADER_HEIGHT = 63.75             # the column header of an output table
HEADER_TOP = 110.1                # where the first table or column header starts
GROUPS = {
    "no frozen header": (("Home", "Params", "NettingSets", "Trades", "Dashboard", "ReconSummary", "FormulaTest",
                          "README"), 120.75),  # row 4 holds the bar's top block: four inputs since #831
    "tables": (("Results", "HedgingSets", "Buckets", "TradeCalc", "Checks", "TestCatalogue", "Basel_vs_CRR"), 173.85),
    "staging sheets": (("Import_NettingSets", "Import_Trades"), 235.35),
    "bank sheets": (("Counterparties", "BankNettingSets", "BankTrades"), 205.5),  # loaded from a CSV file (#831)
    "reconciliation sheets": (("Reconciliation", "TradeReconciliation"), 269.85),
}
OUTPUT_HEADERS = {"Results": 6, "HedgingSets": 5, "Buckets": 5, "TradeCalc": 5, "Checks": 5, "TestCatalogue": 6,
                  "Basel_vs_CRR": 5}


class FrozenRowTests(unittest.TestCase):
    sheets: dict[str, str]

    @classmethod
    def setUpClass(cls) -> None:
        with zipfile.ZipFile(TEMPLATE) as package:
            workbook = package.read("xl/workbook.xml").decode("utf-8")
            rels = package.read("xl/_rels/workbook.xml.rels").decode("utf-8")
            targets = {re.search(r'Id="([^"]+)"', item).group(1): re.search(r'Target="([^"]+)"', item).group(1)  # type: ignore[union-attr]
                       for item in re.findall(r"<Relationship [^>]*/>", rels)}
            cls.sheets = {name: package.read("xl/" + targets[rid].replace("/xl/", "")).decode("utf-8")
                          for name, rid in re.findall(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', workbook)}

    def heights(self, name: str, rows: int) -> list[float]:
        found = []
        for row in range(1, rows + 1):
            element = re.search(r'<row r="%d"([^>]*)>' % row, self.sheets[name])
            height = re.search(r'\bht="([\d.]+)"', element.group(1)) if element else None
            found.append(float(height.group(1)) if height else 15.0)
        return found

    def frozen(self, name: str) -> int:
        return int(re.search(r'<pane\b[^>]*\bySplit="(\d+)"', self.sheets[name]).group(1))  # type: ignore[union-attr]

    def test_every_sheet_belongs_to_a_group(self) -> None:
        grouped = {name for names, _ in GROUPS.values() for name in names}
        self.assertEqual(sorted(set(self.sheets) - grouped), [])

    def test_every_sheet_of_a_group_freezes_rows_of_the_same_height(self) -> None:
        for group, (names, total) in GROUPS.items():
            for name in names:
                with self.subTest(group=group, sheet=name):
                    self.assertAlmostEqual(sum(self.heights(name, self.frozen(name))), total, places=2)

    def test_the_title_band_is_the_same_on_every_sheet_with_a_frozen_band_of_four_rows_or_more(self) -> None:
        for names, _ in GROUPS.values():
            for name in names:
                with self.subTest(sheet=name):
                    self.assertEqual(self.heights(name, 3), TITLE_BAND)

    def test_the_tables_start_their_header_at_the_same_point_with_the_same_height(self) -> None:
        for name, row in OUTPUT_HEADERS.items():
            heights = self.heights(name, row)
            with self.subTest(sheet=name):
                self.assertEqual(self.frozen(name), row)
                self.assertAlmostEqual(sum(heights[:-1]), HEADER_TOP, places=2)
                self.assertEqual(heights[-1], HEADER_HEIGHT)


if __name__ == "__main__":
    unittest.main()
