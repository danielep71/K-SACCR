"""The specification of the formatted run report against the template and the code (#259)."""
from __future__ import annotations

import json
import re
import unittest
import zipfile
from typing import Any
from xml.etree import ElementTree

import report
import template_layout as layout

ROOT = report.ROOT
TEMPLATE = ROOT / "src" / "workbook" / "SACCR_Template.xlsx"
MAIN = layout.MAIN
PKG = "{http://schemas.openxmlformats.org/package/2006/relationships}"
REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
TOTAL_ROWS = ("TOTAL CRR", "TOTAL BCBS")


def read_vba(path: str) -> str:
    return (ROOT / path).read_text(encoding="cp1252")


class ReportSpecTests(unittest.TestCase):
    spec: dict[str, Any]
    package: zipfile.ZipFile
    sheets: dict[str, str]
    shared: list[str]

    @classmethod
    def setUpClass(cls) -> None:
        cls.spec = report.load_spec()
        cls.package = zipfile.ZipFile(TEMPLATE)
        workbook = ElementTree.fromstring(cls.package.read("xl/workbook.xml"))
        rels = ElementTree.fromstring(cls.package.read("xl/_rels/workbook.xml.rels"))
        targets = {r.attrib["Id"]: r.attrib["Target"] for r in rels.iter(f"{PKG}Relationship")}
        cls.sheets = {s.attrib["name"]: "xl/" + targets[s.attrib[f"{REL}id"]] for s in workbook.iter(f"{MAIN}sheet")}
        strings = ElementTree.fromstring(cls.package.read("xl/sharedStrings.xml"))
        cls.shared = ["".join(t.text or "" for t in si.iter(f"{MAIN}t")) for si in strings.iter(f"{MAIN}si")]

    def section(self, section_id: str) -> dict[str, Any]:
        return next(s for s in self.spec["sections"] if s["id"] == section_id)

    def header_texts(self, name: str) -> list[str]:
        """The header cells of an output table, from its workbook name to the first empty cell."""
        start = layout.named_cell(self.package, name)
        cells: dict[str, str] = {}
        for cell in ElementTree.fromstring(self.package.read(self.sheets[start.sheet])).iter(f"{MAIN}c"):
            value = cell.find(f"{MAIN}v")
            if value is not None and value.text is not None:
                cells[cell.attrib["r"]] = self.shared[int(value.text)] if cell.attrib.get("t") == "s" else value.text
        headers: list[str] = []
        column = layout.column_number(start.column)
        while f"{layout.column_letters(column)}{start.row}" in cells:
            headers.append(cells[f"{layout.column_letters(column)}{start.row}"])
            column += 1
        return headers

    def test_the_tables_have_the_headers_of_their_sheets(self) -> None:
        for section_id in ("results", "checks", "reconciliation", "trade_reconciliation"):
            section = self.section(section_id)
            with self.subTest(section=section_id):
                self.assertEqual([c["header"] for c in section["columns"]], self.header_texts(section["header_name"]))

    def test_the_columns_have_the_number_of_columns_the_evidence_export_reads(self) -> None:
        config = read_vba("src/core/CORE_Config.bas")
        for section_id, constant in (("results", "RS_NCOLS"), ("checks", "CK_NCOLS"),
                                     ("reconciliation", "RC_NCOLS"), ("trade_reconciliation", "TRC_NCOLS")):
            found = re.search(rf"Const {constant}\s+As Long = (\d+)", config)
            assert found is not None, constant
            self.assertEqual(len(self.section(section_id)["columns"]), int(found.group(1)), section_id)

    def test_the_totals_are_columns_of_results(self) -> None:
        results = [c["header"] for c in self.section("results")["columns"]]
        kinds = {c["header"]: c["kind"] for c in self.section("results")["columns"]}
        for column in self.section("totals")["columns"]:
            self.assertIn(column["header"] if column["header"] != "Regime" else "Regime", results)
            self.assertEqual(column["kind"], kinds[column["header"]], column["header"])

    def test_the_total_rows_are_those_the_evidence_export_reads(self) -> None:
        main = read_vba("src/workbook/M_Main.bas")
        for name in TOTAL_ROWS:
            regime = name.split()[1]
            self.assertIn(f'"TOTAL " & RG_{regime}', main)

    def test_the_run_lines_are_lines_of_the_evidence_run_section(self) -> None:
        main = read_vba("src/workbook/M_Main.bas")
        for row in self.section("run")["rows"]:
            if row["key"] in ("units", "results_status"):
                continue
            if row["key"] in ("reconciliation", "trade_reconciliation"):
                self.assertIn('KeyPrefix = IIf(i = RECON_NS, "reconciliation", "trade_reconciliation")', main)
                continue
            with self.subTest(key=row["key"]):
                self.assertIn(f'RunLine("{row["key"]}"', main)

    def test_the_dashboard_blocks_come_from_the_dashboard_specification(self) -> None:
        section = self.section("dashboard")
        self.assertTrue((ROOT / section["blocks_from"]).is_file())
        self.assertEqual(section["charts_in"], ["pdf"])
        dashboard = ElementTree.fromstring(self.package.read(self.sheets["Dashboard"]))
        self.assertIsNotNone(dashboard)
        drawings = [name for name in self.package.namelist() if re.fullmatch(r"xl/charts/chart\d+\.xml", name)]
        self.assertGreaterEqual(len(drawings), 10)

    def test_the_sheets_of_the_report_are_not_named_like_the_sheets_they_come_from_by_accident(self) -> None:
        for section in self.spec["sections"]:
            self.assertLessEqual(len(section["sheet"]), 31)
            self.assertNotRegex(section["sheet"], r"[\[\]:*?/\\]")

    def test_the_report_has_no_limitations_sheet(self) -> None:
        # The owner's verification of 2026-11-26: the limitations stay in the evidence CSV only.
        self.assertNotIn("limitations", [s["id"] for s in self.spec["sections"]])
        main = read_vba("src/workbook/M_Main.bas")
        self.assertNotIn('AddReportSheet(ReportBook, "Limitations")', main)
        self.assertNotIn("ReportLimitationTexts", read_vba("src/core/CORE_Export.bas"))
        self.assertIn('"no-limitations"', read_vba("tests/modules/TEST_Report.bas"))

    def test_the_measures_that_are_not_amounts_are_those_the_reconciliation_does_not_scale(self) -> None:
        source = read_vba("src/core/CORE_Reconcile.bas")

        def constant(name: str) -> str:
            found = re.search(rf'Const {name}\s+As String = ((?:"[^"]*"(?:\s*&\s*_\s*)?)+)', source)
            assert found is not None, name
            return "".join(re.findall(r'"([^"]*)"', found.group(1)))

        for section_id, measures, ratios in (("reconciliation", "NS_MEASURES", "NS_RATIOS"),
                                             ("trade_reconciliation", "TR_MEASURES", "TR_RATIOS")):
            names = constant(measures).split("|")
            positions = [int(p) for p in constant(ratios).strip("|").split("|")]
            with self.subTest(section=section_id):
                self.assertEqual(self.section(section_id)["measure_ratios"], [names[p] for p in positions])

    def vba_constant(self, source: str, name: str) -> str:
        found = re.search(rf'Const {name}\s+As String = ((?:"[^"]*"(?:\s*&\s*_\s*)?)+)', source)
        assert found is not None, name
        return "".join(re.findall(r'"([^"]*)"', found.group(1)))

    def test_the_writer_has_the_kinds_and_measures_of_the_specification(self) -> None:
        main = read_vba("src/workbook/M_Main.bas")
        for section_id, kinds in (("totals", "RPT_TOTALS_KINDS"), ("results", "RPT_RESULTS_KINDS"),
                                  ("checks", "RPT_CHECKS_KINDS"), ("reconciliation", "RPT_RECON_KINDS"),
                                  ("trade_reconciliation", "RPT_TRADE_RECON_KINDS")):
            with self.subTest(section=section_id):
                self.assertEqual(self.vba_constant(main, kinds).split("|"),
                                 [c["kind"] for c in self.section(section_id)["columns"]])
        self.assertEqual(self.vba_constant(main, "RPT_TOTALS_HEADERS").split("|"),
                         [c["header"] for c in self.section("totals")["columns"]])
        self.assertEqual(self.vba_constant(main, "RPT_RECON_RATIOS").split("|"),
                         self.section("reconciliation")["measure_ratios"])
        self.assertEqual(self.vba_constant(main, "RPT_TRADE_RECON_RATIOS").split("|"),
                         self.section("trade_reconciliation")["measure_ratios"])

    def test_the_writer_has_the_formats_and_the_sheets_of_the_specification(self) -> None:
        main = read_vba("src/workbook/M_Main.bas")
        body = main[main.index("Private Function ReportNumberFormat("):]
        body = body[:body.index("End Function")]
        formats = dict(re.findall(r'Case "(\w+)"\s*\r?\n\s*ReportNumberFormat = "([^"]*)"', body))
        for kind in ("amount", "ratio", "count", "date", "datetime"):
            self.assertEqual(formats.get(kind), self.spec["formats"][kind], kind)
        writer = main[main.index("Private Function WriteReport("):]
        writer = writer[:writer.index("End Function")]
        named = re.findall(r'AddReportSheet\(ReportBook, "(\w+)"\)', writer)
        named[named.index("Results") + 2:named.index("Results") + 2] = []  # the reconciliations are named by IIf
        self.assertIn('"Reconciliation", "TradeReconciliation"', writer)
        expected = [s["sheet"] for s in self.spec["sections"]]
        expected.insert(expected.index("Totals"), "Charts")  # the PDF's charts
        self.assertEqual(named, [name for name in expected if name not in ("Reconciliation", "TradeReconciliation")])

    def test_the_run_items_of_the_writer_are_those_of_the_specification(self) -> None:
        main = read_vba("src/workbook/M_Main.bas")
        body = main[main.index("Private Function WriteReportRun("):]
        body = body[:body.index("End Function")]
        labels = re.findall(r'Items\(\d+, 1\) = "([^"]+)"', body)
        labels[12:12] = re.findall(r'"(Reconciliation of netting sets)", "(Reconciliation of trade legs)"', body)[0]
        self.assertEqual(labels, [r["label"] for r in self.section("run")["rows"]])

    def test_the_report_has_the_look_of_the_workbook_sheets(self) -> None:
        # The owner's verification of 2026-11-26: the report takes the font, the
        # header fill and the tab colors of the K-SACCR sheets.
        style = json.loads((ROOT / "docs" / "workbook" / "sheet-style.json").read_text(encoding="utf-8"))
        main = read_vba("src/workbook/M_Main.bas")

        def bgr(constant: str) -> str:
            found = re.search(rf"Private Const {constant}\s+As Long = &H([0-9A-F]{{6}})\b", main)
            self.assertIsNotNone(found, constant)
            text = found.group(1) if found else ""
            return text[4:6] + text[2:4] + text[0:2]  # VBA's BGR as RRGGBB

        self.assertIn(f'Private Const RPT_FONT                As String = "{style["font"]}"', main)
        self.assertEqual(bgr("RPT_HEADER_FILL"), style["header_fill"])
        self.assertEqual(bgr("RPT_TAB_REFERENCE"), style["tab_colors"]["Reference"])
        self.assertEqual(bgr("RPT_TAB_OUTPUT"), style["tab_colors"]["Output"])
        writer = main[main.index("Private Function WriteReportTable("):]
        writer = writer[:writer.index("End Function")]
        self.assertIn(".Font.Color = vbWhite", writer)
        self.assertIn(".Interior.Color = RPT_HEADER_FILL", writer)
        self.assertNotIn("RGB(217, 225, 242)", main)

    def test_the_charts_are_laid_out_in_the_printable_area_of_the_page(self) -> None:
        # A4 landscape is 841.89 by 595.28 points; the Charts sheet is laid out in what the margins leave.
        main = read_vba("src/workbook/M_Main.bas")

        def constant(name: str) -> float:
            found = re.search(rf"Private Const {name}\s+As (?:Double|Long) = ([\d.]+)#?", main)
            self.assertIsNotNone(found, name)
            return float(found.group(1)) if found else 0.0

        self.assertLessEqual(constant("RPT_PRINT_WIDTH"), 841.89 - 2 * constant("RPT_MARGIN_SIDE"))
        self.assertLessEqual(constant("RPT_PRINT_HEIGHT"), 595.28 - 2 * constant("RPT_MARGIN_END"))
        self.assertGreater(constant("RPT_PRINT_WIDTH"), 841.89 - 2 * constant("RPT_MARGIN_SIDE") - 1)
        self.assertGreater(constant("RPT_PRINT_HEIGHT"), 595.28 - 2 * constant("RPT_MARGIN_END") - 1)
        setup = main[main.index("Private Function ApplyReportPageSetup("):]
        setup = setup[:setup.index("End Function")]
        for line in (".LeftMargin = RPT_MARGIN_SIDE", ".TopMargin = RPT_MARGIN_END", ".Zoom = PrintZoom"):
            self.assertIn(line, setup)

    def test_the_vba_suite_checks_every_example_of_the_specification(self) -> None:
        suite = read_vba("tests/modules/TEST_Report.bas")
        for example in self.spec["formats"]["examples"]:
            with self.subTest(example=example):
                self.assertIn(f'"{example["kind"]}", ', suite)
                self.assertIn(f', "{example["text"]}"', suite)

    def test_every_kind_is_a_format(self) -> None:
        formats = self.spec["formats"]
        for kind in report.KINDS:
            self.assertIn(kind, formats)

    def test_the_documentation_names_the_files_of_the_specification(self) -> None:
        text = (ROOT / "docs" / "workbook" / "REPORT.md").read_text(encoding="utf-8")
        for name in ("report.json", "tools/report.py", "tools/test_report_spec.py"):
            self.assertIn(name, text)
        for section in self.spec["sections"]:
            self.assertIn(f"`{section['sheet']}`", text, section["sheet"])


if __name__ == "__main__":
    unittest.main()
