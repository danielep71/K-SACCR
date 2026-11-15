"""Tests of the sheet style file and the template's tabs (#706)."""
import json
import unittest
from pathlib import Path

import check_sheet_style as gate

ROOT = Path(__file__).resolve().parents[1]
MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
STYLE = {"schema_version": 1, "description": "x", "font": "Aptos", "vertical_alignment": "center",
         "date_format": "yyyy-mm-dd", "gridlines": False, "tab_colors": {"Input": "BF9000", "Output": "2E75B6"}}


def sheet(tab: str = '<sheetPr><tabColor rgb="FFBF9000"/></sheetPr>',
          view: str = '<sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews>') -> bytes:
    return f'<worksheet xmlns="{MAIN}">{tab}{view}<sheetData/></worksheet>'.encode()


class SheetStyleTests(unittest.TestCase):
    def test_repository_passes(self) -> None:
        report = gate.run_check(ROOT)
        self.assertEqual(report["status"], "pass", report["findings"])

    def test_style_file_needs_one_color_per_role(self) -> None:
        self.assertEqual(gate.check_style(STYLE, {"Input", "Output"}), [])
        findings = gate.check_style(STYLE, {"Input", "Output", "Staging"})
        self.assertTrue(any("'Staging'" in f and "no tab color" in f for f in findings), findings)
        findings = gate.check_style(STYLE, {"Input"})
        self.assertTrue(any("'Output'" in f and "not a role" in f for f in findings), findings)

    def test_style_file_rejects_bad_values(self) -> None:
        bad = json.loads(json.dumps(STYLE))
        bad["tab_colors"]["Input"] = "bf9000"
        bad["gridlines"] = True
        bad["extra"] = 1
        findings = gate.check_style(bad, {"Input", "Output"})
        self.assertTrue(any("six upper-case hex" in f for f in findings), findings)
        self.assertTrue(any("gridlines must be false" in f for f in findings), findings)
        self.assertTrue(any("keys must be" in f for f in findings), findings)

    def test_sheet_needs_its_role_color_and_no_gridlines(self) -> None:
        self.assertEqual(gate.sheet_findings("S", sheet(), "BF9000"), [])
        findings = gate.sheet_findings("S", sheet(), "2E75B6")
        self.assertTrue(any("tab color FFBF9000, expected 2E75B6" in f for f in findings), findings)
        findings = gate.sheet_findings("S", sheet(tab=""), "BF9000")
        self.assertTrue(any("tab color none" in f for f in findings), findings)
        findings = gate.sheet_findings("S", sheet(view='<sheetViews><sheetView workbookViewId="0"/></sheetViews>'),
                                       "BF9000")
        self.assertTrue(any("gridlines" in f for f in findings), findings)


if __name__ == "__main__":
    unittest.main()
