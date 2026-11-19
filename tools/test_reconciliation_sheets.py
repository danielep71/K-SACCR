"""Tests of the reconciliation output sheets of the workbook template (#720).

The two sheets are never protected, carry an AutoFilter on their header row, show the amount
figures in one number format with a format of their own for the factor rows, and name the
difference column an absolute difference.
"""

from __future__ import annotations

import re
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree

from test_template_buttons import sheet_parts

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "src" / "workbook" / "SACCR_Template.xlsx"
MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
# Sheet: (last table column, the measures shown in the factor format)
SHEETS = {
    "Reconciliation": ("N", {"Multiplier"}),
    "TradeReconciliation": ("O", {"M", "Supervisory duration", "Supervisory delta", "MF unmargined", "MF margined"}),
}
FACTOR_FORMAT = "#,##0.000000"


class ReconciliationSheetTests(unittest.TestCase):
    def setUp(self) -> None:
        self.package = zipfile.ZipFile(TEMPLATE)
        self.addCleanup(self.package.close)
        self.parts = sheet_parts(self.package)

    def sheet(self, name: str) -> ElementTree.Element:
        return ElementTree.fromstring(self.package.read(self.parts[name]))

    def test_the_sheets_are_saved_unprotected(self) -> None:
        for name in SHEETS:
            self.assertIsNone(self.sheet(name).find(MAIN + "sheetProtection"), name)

    def test_the_header_row_has_an_autofilter(self) -> None:
        for name, (last, _) in SHEETS.items():
            autofilter = self.sheet(name).find(MAIN + "autoFilter")
            self.assertIsNotNone(autofilter, name)
            assert autofilter is not None
            self.assertEqual(autofilter.attrib["ref"], f"C6:{last}100000", name)
            workbook = self.package.read("xl/workbook.xml").decode("utf-8")
            self.assertRegex(workbook, rf'<definedName name="_xlnm\._FilterDatabase" localSheetId="\d+" hidden="1">'
                                       rf"{name}!\$C\$6:\${last}\$100000</definedName>")

    def test_the_factor_rows_have_their_own_format(self) -> None:
        styles = ElementTree.fromstring(self.package.read("xl/styles.xml"))
        dxfs = styles.find(MAIN + "dxfs")
        assert dxfs is not None
        for name, (_, measures) in SHEETS.items():
            rules = [rule for block in self.sheet(name).findall(MAIN + "conditionalFormatting")
                     if block.attrib["sqref"] == "I7:K100000" for rule in block.findall(MAIN + "cfRule")]
            self.assertEqual(len(rules), 1, name)
            formula = rules[0].findtext(MAIN + "formula") or ""
            self.assertEqual(set(re.findall(r'\$H7="([^"]+)"', formula)), measures, name)
            dxf = list(dxfs)[int(rules[0].attrib["dxfId"])]
            numfmt = dxf.find(MAIN + "numFmt")
            assert numfmt is not None
            self.assertEqual(numfmt.attrib["formatCode"], FACTOR_FORMAT, name)

    def test_the_difference_column_is_an_absolute_difference(self) -> None:
        strings = self.package.read("xl/sharedStrings.xml").decode("utf-8")
        self.assertIn("<t>Absolute difference</t>", strings)
        self.assertNotIn("calculator - bank", strings)

    def test_the_macros_write_the_absolute_difference_and_one_amount_format(self) -> None:
        reconcile = (ROOT / "src" / "core" / "CORE_Reconcile.bas").read_text(encoding="utf-8")
        self.assertIn('"", Abs(Difference), RelativeDifference', reconcile)
        main = (ROOT / "src" / "workbook" / "M_Main.bas").read_text(encoding="utf-8")
        self.assertIn('"@|0|@|@|@|@|#,##0.00|#,##0.00|"', main)
        self.assertNotIn("0.00########", main)

    def test_the_sheets_are_listed_as_never_protected(self) -> None:
        config = (ROOT / "src" / "core" / "CORE_Config.bas").read_text(encoding="utf-8")
        match = re.search(r'PROTECT_NONE\s+As String = "([^"]*)"', config)
        self.assertIsNotNone(match)
        assert match is not None
        self.assertEqual({name for name in match.group(1).split("|") if name}, set(SHEETS))


if __name__ == "__main__":
    unittest.main()
