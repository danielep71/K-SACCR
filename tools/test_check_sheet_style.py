"""Tests of the sheet style file and the template's tabs (#706)."""
import json
import unittest
from pathlib import Path

import check_sheet_style as gate

ROOT = Path(__file__).resolve().parents[1]
MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
STYLE = {"schema_version": 1, "description": "x", "font": "Aptos", "vertical_alignment": "center",
         "date_format": "yyyy-mm-dd", "gridlines": False, "tab_colors": {"Input": "BF9000", "Output": "2E75B6"},
         "header_fill": "1F3864", "header_font_size": 11, "header_rows": {"S": [1]}}
# Font 0 plain Aptos, font 1 bold white Aptos 11, font 2 Arial; fill 2 navy, fill 3 blue. Format 0 plain,
# 1 the header look with wrapping, 2 a header on blue, 3 bottom-aligned, 4 Arial.
STYLES = (f'<styleSheet xmlns="{MAIN}"><fonts count="3"><font><sz val="11"/><name val="Aptos"/></font>'
          '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Aptos"/></font>'
          '<font><sz val="11"/><name val="Arial"/></font></fonts>'
          '<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'
          '<fill><patternFill patternType="solid"><fgColor rgb="FF1F3864"/></patternFill></fill>'
          '<fill><patternFill patternType="solid"><fgColor rgb="FF2E75B6"/></patternFill></fill></fills>'
          '<cellXfs count="5"><xf fontId="0" fillId="0"><alignment vertical="center"/></xf>'
          '<xf fontId="1" fillId="2"><alignment vertical="center" wrapText="1"/></xf>'
          '<xf fontId="1" fillId="3"><alignment vertical="center" wrapText="1"/></xf>'
          '<xf fontId="0" fillId="0"/><xf fontId="2" fillId="0"><alignment vertical="center"/></xf>'
          '</cellXfs></styleSheet>').encode()


def sheet(tab: str = '<sheetPr><tabColor rgb="FFBF9000"/></sheetPr>',
          view: str = '<sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews>',
          data: str = "") -> bytes:
    return f'<worksheet xmlns="{MAIN}">{tab}{view}<sheetData>{data}</sheetData></worksheet>'.encode()


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

    def test_shared_formula_stays_on_its_sheet(self) -> None:
        local = '<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">IF(B1="x!y",1,0)</f></c></row>'
        other = '<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">COUNTA(T!$C$6:$C$9)</f></c></row>'
        single = '<row r="1"><c r="A1"><f>COUNTA(T!$C$6:$C$9)</f></c></row>'
        self.assertEqual(gate.sheet_findings("S", sheet(data=local), "BF9000"), [])
        self.assertEqual(gate.sheet_findings("S", sheet(data=single), "BF9000"), [])
        findings = gate.sheet_findings("S", sheet(data=other), "BF9000")
        self.assertTrue(any("S!A1: shared formula refers to another sheet" in f for f in findings), findings)


class LineCountTests(unittest.TestCase):
    def test_estimate_never_counts_fewer_lines_than_excel(self) -> None:
        # Headers of the template and the lines Excel showed in the owner's run of 29b29be (#708).
        seen = [("SF effective (IR / FX / OT)", 9.75, 4), ("SF (effective)", 8.75, 3),
                ("Stand-alone add-on", 10.625, 3), ("Multiplier (cap basis)", 11.75, 2),
                ("Idiosyncratic: sqrt(sum (1-rho^2) A^2)", 17.75, 3), ("Effective notional (IR / FX / OT)", 15.75, 2),
                ("SF (effective)", 9.75, 2), ("Supervisory duration", 15.625, 2), ("NettingSetID", 14.75, 1)]
        for text, width, lines in seen:
            self.assertGreaterEqual(gate.wrapped_lines(text, width, 11), lines, text)

    def test_lines_follow_breaks_long_words_and_font_size(self) -> None:
        self.assertEqual(gate.wrapped_lines("a\nb", 20, 11), 2)
        self.assertEqual(gate.wrapped_lines("remargin_period_business_days", 10, 11), 3)
        self.assertGreater(gate.wrapped_lines("one two three four", 9, 16), gate.wrapped_lines("one two three four", 9, 11))
        self.assertEqual(gate.row_height(1, 11), 18.96)


class HeaderAndStyleTests(unittest.TestCase):
    def setUp(self) -> None:
        self.styles = gate.read_styles(STYLES)

    def test_every_font_and_alignment(self) -> None:
        findings = gate.style_findings(self.styles, "Aptos", "center")
        self.assertIn("src/workbook/SACCR_Template.xlsx: font 2 is 'Arial', not 'Aptos'", findings)
        self.assertIn("src/workbook/SACCR_Template.xlsx: cell format 3 is aligned 'bottom', not 'center'", findings)
        self.assertEqual(len(findings), 2, findings)

    def header(self, style: int, height: str) -> list[str]:
        part = (f'<worksheet xmlns="{MAIN}"><cols><col min="1" max="1" width="9.75"/></cols><sheetData>'
                f'<row r="1"{height}><c r="A1" s="{style}" t="s"><v>0</v></c></row></sheetData></worksheet>').encode()
        return gate.header_findings("S", part, [1], self.styles, ["SF effective (IR / FX / OT)"], STYLE)

    def test_header_look_and_height(self) -> None:
        self.assertEqual(self.header(1, ' ht="64" customHeight="1"'), [])
        findings = self.header(2, ' ht="64" customHeight="1"')
        self.assertTrue(any("is not bold white 11 pt on 1F3864" in f for f in findings), findings)
        findings = self.header(1, ' ht="42" customHeight="1"')
        self.assertTrue(any("header row 1 is 42 pt, its text needs 4 line(s)" in f for f in findings), findings)


if __name__ == "__main__":
    unittest.main()
