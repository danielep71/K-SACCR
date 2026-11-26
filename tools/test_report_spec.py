"""The specification of the formatted run report against the template and the code (#259)."""
from __future__ import annotations

import re
import unittest
import zipfile
from pathlib import Path
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

    def test_the_limitations_are_those_of_the_evidence_export_and_one_more(self) -> None:
        export = read_vba("src/core/CORE_Export.bas")
        body = export[export.index("Public Function LimitationLines()"):]
        notes = body[body.index("Notes = Array("):body.index("For i = LBound(Notes)")]
        statements = len(re.findall(r'^ {12}"', notes, flags=re.M))  # a statement starts at 12 spaces, its continuation at 16
        self.assertEqual(statements, 7)
        self.assertEqual(self.section("limitations")["statements"], statements + 1)

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
