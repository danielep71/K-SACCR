"""Tests of the filter buttons of the five output sheets of the workbook template (#754)."""

from __future__ import annotations

import re
import unittest
import zipfile
from pathlib import Path

from test_template_buttons import sheet_parts

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "src" / "workbook" / "SACCR_Template.xlsx"
# sheet: the header row, and the first and last column, of its table
OUTPUT_SHEETS = {"Results": (6, "C", "AD"), "HedgingSets": (5, "C", "P"), "Buckets": (5, "C", "M"),
                 "TradeCalc": (5, "C", "AA"), "Checks": (5, "C", "J")}
LAST_ROW = 1048576  # the last row of a sheet: a run of any size is inside the filter


class OutputFilterTests(unittest.TestCase):
    def setUp(self) -> None:
        self.package = zipfile.ZipFile(TEMPLATE)
        self.addCleanup(self.package.close)
        self.parts = sheet_parts(self.package)
        self.workbook = self.package.read("xl/workbook.xml").decode("utf-8")
        self.order = re.findall(r'<sheet name="([^"]+)" sheetId="\d+"', self.workbook)

    def sheet(self, name: str) -> str:
        return self.package.read(self.parts[name]).decode("utf-8")

    def test_each_output_sheet_has_its_filter_over_the_header_and_the_table(self) -> None:
        for name, (row, first, last) in OUTPUT_SHEETS.items():
            ref = f"{first}{row}:{last}{LAST_ROW}"
            xml = self.sheet(name)
            self.assertEqual(re.findall(r"<autoFilter [^>]*>", xml), [f'<autoFilter ref="{ref}"/>'], name)
            database = (f'<definedName name="_xlnm._FilterDatabase" localSheetId="{self.order.index(name)}" hidden="1">'
                        f"{name}!${first}${row}:${last}${LAST_ROW}</definedName>")
            self.assertIn(database, self.workbook, name)

    def test_the_filter_sits_on_a_header_row_of_text(self) -> None:
        strings = [re.sub(r"<[^>]+>", "", item) for item in
                   re.findall(r"<si>(.*?)</si>", self.package.read("xl/sharedStrings.xml").decode("utf-8"), re.S)]
        for name, (row, first, last) in OUTPUT_SHEETS.items():
            xml = self.sheet(name)
            cells = re.search(rf'<row r="{row}"[^>]*>(.*?)</row>', xml, re.S)
            assert cells is not None, name
            found = re.findall(r'<c r="([A-Z]+)\d+"[^>]*t="s"[^>]*><v>(\d+)</v>', cells.group(1))
            self.assertEqual((found[0][0], found[-1][0]), (first, last), name)
            self.assertTrue(all(strings[int(index)] for _, index in found), name)  # every header cell holds text

    def test_the_sheets_allow_filtering_and_nothing_else_is_added(self) -> None:
        for name in OUTPUT_SHEETS:
            protection = re.findall(r"<sheetProtection [^>]*/>", self.sheet(name))
            self.assertEqual(protection, ['<sheetProtection sheet="1" objects="1" scenarios="1" autoFilter="0"/>'], name)
        config = (ROOT / "src" / "core" / "CORE_Config.bas").read_text(encoding="cp1252")
        listed = re.search(r'PROTECT_FILTERING\s+As String = _\s*"([^"]+)"', config)
        assert listed is not None
        self.assertEqual(listed.group(1), "|Trades|Basel_vs_CRR|" + "|".join(OUTPUT_SHEETS) + "|")
        outputs = re.search(r'OUTPUT_FILTER_SHEETS\s+As String = "([^"]+)"', config)
        assert outputs is not None
        self.assertEqual(set(outputs.group(1).split("|")), set(OUTPUT_SHEETS))

    def test_an_operation_applies_a_filter_again_without_setting_or_clearing_one(self) -> None:
        main = (ROOT / "src" / "workbook" / "M_Main.bas").read_text(encoding="cp1252")
        helper = main[main.index("Private Sub ReapplyOutputFilters()"):]
        helper = helper[: helper.index("\nEnd Sub")]
        self.assertIn("If Sheet.FilterMode Then", helper)
        self.assertIn("Sheet.AutoFilter.ApplyFilter", helper)
        for forbidden in ("ShowAllData", "AutoFilterMode = False", ".AutoFilter Field"):
            self.assertNotIn(forbidden, helper)
        operation = main[main.index("Private Function ExecuteOperation("):]
        select = operation.index("        End Select\n        ReapplyOutputFilters")
        self.assertGreater(select, 0)  # after the run, the validation and Clear outputs alike


if __name__ == "__main__":
    unittest.main()
