"""Tests of the opening state of the workbook template (#726).

Every sheet is saved at a zoom of 100%, the two bank sheets default to a relative tolerance of 1%, and
the workbook opens on the K-SACCR Ribbon tab.
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
UI = "{http://schemas.microsoft.com/office/2009/07/customui}"


class OpeningStateTests(unittest.TestCase):
    def setUp(self) -> None:
        self.package = zipfile.ZipFile(TEMPLATE)
        self.addCleanup(self.package.close)
        self.parts = sheet_parts(self.package)

    def test_every_sheet_is_saved_at_100_percent(self) -> None:
        self.assertEqual(len(self.parts), 22)
        for name, part in self.parts.items():
            view = ElementTree.fromstring(self.package.read(part)).find(f"{MAIN}sheetViews/{MAIN}sheetView")
            assert view is not None
            self.assertEqual(view.attrib.get("zoomScale", "100"), "100", name)  # Excel leaves out the default
            self.assertEqual(view.attrib.get("zoomScaleNormal"), "100", name)

    def test_the_bank_sheets_default_to_a_relative_tolerance_of_one_percent(self) -> None:
        for name in ("BankNettingSets", "BankTrades"):
            sheet = ElementTree.fromstring(self.package.read(self.parts[name]))
            cell = next(item for item in sheet.iter(f"{MAIN}c") if item.attrib["r"] == "D6")
            self.assertEqual(float(cell.findtext(f"{MAIN}v") or "nan"), 0.01, name)
            prompts = [item.attrib.get("prompt", "") for item in sheet.iter(f"{MAIN}dataValidation")
                       if item.attrib.get("sqref") == "D6"]
            self.assertEqual(len(prompts), 1, name)
            self.assertIn("1% (0.01)", prompts[0], name)

    def test_the_data_validation_messages_are_within_the_limits_of_excel(self) -> None:
        # Excel refuses to open a file whose input message is over 255 characters, whose error message is
        # over 255, or whose titles are over 32 (the quick check then stops with "Unable to get the Open
        # property of the Workbooks class").
        limits = {"prompt": 255, "error": 255, "promptTitle": 32, "errorTitle": 32}
        for name, part in self.parts.items():
            sheet = ElementTree.fromstring(self.package.read(part))
            for item in sheet.iter(f"{MAIN}dataValidation"):
                for attribute, limit in limits.items():
                    text = item.attrib.get(attribute, "")
                    self.assertLessEqual(len(text), limit,
                                         f"{name} {item.attrib.get('sqref')}: {attribute} has {len(text)} characters")

    def test_the_tab_the_workbook_opens_on_is_the_k_saccr_tab(self) -> None:
        tree = ElementTree.parse(ROOT / "src" / "workbook" / "customUI14.xml").getroot()
        tabs = [item.attrib["id"] for item in tree.iter(UI + "tab")]
        ribbon = (ROOT / "src" / "workbook" / "M_Ribbon.bas").read_text(encoding="utf-8")
        match = re.search(r'RIBBON_TAB_ID\s+As String = "([^"]+)"', ribbon)
        self.assertIsNotNone(match)
        assert match is not None
        self.assertEqual(tabs, [match.group(1)])
        self.assertIn("Ribbon.ActivateTab RIBBON_TAB_ID", ribbon)

    def test_the_tab_is_selected_again_after_an_add_in_took_it(self) -> None:
        ribbon = (ROOT / "src" / "workbook" / "M_Ribbon.bas").read_text(encoding="utf-8")
        load = ribbon[ribbon.index("Public Sub RibbonOnLoad("):ribbon.index("Public Sub ActivateRibbonTab()")]
        self.assertIn("ScheduleTabActivation", load)
        retry = ribbon[ribbon.index("Public Sub ActivateRibbonTab()"):ribbon.index("Public Sub CancelTabActivation()")]
        self.assertIn("mRibbon.ActivateTab RIBBON_TAB_ID", retry)
        self.assertIn("Application.ActiveWorkbook Is ThisWorkbook", retry)  # another workbook keeps its tab
        schedule = ribbon[ribbon.index("Private Sub ScheduleTabActivation()"):]
        self.assertIn("M_Ribbon.ActivateRibbonTab", schedule)
        self.assertIn("Replace(ThisWorkbook.Name", schedule)
        cancel = (ROOT / "src" / "workbook" / "ThisWorkbook.cls").read_text(encoding="utf-8")
        before_close = cancel[cancel.index("Private Sub Workbook_BeforeClose"):]
        self.assertIn("M_Ribbon.CancelTabActivation", before_close[: before_close.index("End Sub")])
        self.assertNotIn("TabActivation", cancel.replace("CancelTabActivation", ""))  # no event but the close touches it

    def test_the_opening_routine_applies_the_zoom(self) -> None:
        main = (ROOT / "src" / "workbook" / "M_Main.bas").read_text(encoding="utf-8")
        auto_open = main[main.index("Public Sub Auto_Open()"):main.index("Public Sub ApplyOpeningView()")]
        self.assertIn("ApplyOpeningView", auto_open)
        self.assertIn("ThisWorkbook.Windows(1).Zoom = 100", main)


if __name__ == "__main__":
    unittest.main()
