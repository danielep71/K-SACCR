"""Tests of the editability-matrix and protection gate of the workbook template (#420, #421)."""
import json
import re
import tempfile
import unittest
import zipfile
from pathlib import Path
from typing import Any

import check_workbook_protection as gate

ROOT = Path(__file__).resolve().parents[1]
MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RELS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
# Style 0 is locked, style 1 is unlocked, style 2 is locked with the yellow input fill.
STYLES = (f'<styleSheet xmlns="{MAIN}"><fills count="2"><fill><patternFill patternType="none"/></fill>'
          '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/></patternFill></fill></fills>'
          '<cellXfs count="3"><xf numFmtId="0"/>'
          '<xf numFmtId="0" applyProtection="1"><protection locked="0"/></xf>'
          '<xf numFmtId="0" fillId="1"/></cellXfs></styleSheet>')
# Input sheet: header in row 1, C:D unlocked from row 2 through the column styles.
OPEN_COLUMNS = '<cols><col min="3" max="4" width="9" style="1"/></cols>'
GOOD_ROWS = '<row r="1"><c r="C1" s="0"/><c r="D1" s="0"/></row><row r="2"><c r="B2"/><c r="C2" s="1"/></row>'
LOCKED_ROWS = GOOD_ROWS.replace('s="1"', 's="0"')
SPEC = {"role": "Input", "unlocked_columns": {"columns": "C:D", "first_row": 2}}
PROTECTED = '<sheetProtection sheet="1" objects="1" scenarios="1"/>'
STRUCTURE = '<workbookProtection lockStructure="1"/>'
LISTS = ("PROTECT_SORTING", "PROTECT_FILTERING", "PROTECT_INSERTING_ROWS", "PROTECT_OUTLINING", "PROTECT_NONE")


def write_template(path: Path, sheet: str, name: str = "Sheet", structure: str = STRUCTURE) -> None:
    workbook = (f'<workbook xmlns="{MAIN}" xmlns:r="{RELS}">{structure}<sheets>'
                f'<sheet name="{name}" sheetId="1" r:id="rId1"/></sheets></workbook>')
    rels = ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("xl/workbook.xml", workbook)
        archive.writestr("xl/_rels/workbook.xml.rels", rels)
        archive.writestr("xl/styles.xml", STYLES)
        archive.writestr("xl/worksheets/sheet1.xml", f'<worksheet xmlns="{MAIN}">{sheet}</worksheet>')


def run(sheet: str, spec: dict[str, Any] | None = None, name: str = "Sheet",
        protection: str = PROTECTED, structure: str = STRUCTURE, listed: tuple[str, ...] = ()) -> list[str]:
    """Findings for a one-sheet template; the CORE_Config lists named in listed hold Sheet."""
    with tempfile.TemporaryDirectory() as folder:
        root = Path(folder)
        (root / gate.MATRIX).parent.mkdir(parents=True)
        (root / gate.TEMPLATE).parent.mkdir(parents=True)
        (root / gate.CONFIG).parent.mkdir(parents=True)
        matrix = {"schema_version": 1, "sheets": {"Sheet": spec or SPEC}}
        (root / gate.MATRIX).write_text(json.dumps(matrix), encoding="utf-8")
        config = "".join(f'        Public Const {constant} As String = "|{"Sheet|" if constant in listed else ""}"\r\n'
                         for constant in LISTS)
        (root / gate.CONFIG).write_text(config, encoding="cp1252")
        write_template(root / gate.TEMPLATE, sheet + protection, name, structure)
        findings: list[str] = gate.run_check(root)["findings"]
        return findings


class ProtectionGateTests(unittest.TestCase):
    def test_template_matches_matrix(self) -> None:
        report = gate.run_check(ROOT)
        self.assertEqual(report["findings"], [])
        self.assertEqual(report["status"], "pass")

    def test_open_columns_pass(self) -> None:
        self.assertEqual(run(OPEN_COLUMNS + f"<sheetData>{GOOD_ROWS}</sheetData>"), [])

    def test_locked_input_cell(self) -> None:
        rows = GOOD_ROWS.replace('<c r="C2" s="1"/>', '<c r="C2" s="0"/>')
        findings = run(OPEN_COLUMNS + f"<sheetData>{rows}</sheetData>")
        self.assertEqual(findings, ["Sheet: lock flag differs from the matrix at C2"])

    def test_header_cell_inherits_open_column(self) -> None:
        # D1 is not stored, so it takes the unlocked column style.
        rows = GOOD_ROWS.replace('<c r="D1" s="0"/>', "")
        findings = run(OPEN_COLUMNS + f"<sheetData>{rows}</sheetData>")
        self.assertEqual(findings, ["Sheet: lock flag differs from the matrix at D1"])

    def test_column_style_must_be_unlocked(self) -> None:
        cols = '<cols><col min="3" max="3" width="9" style="1"/></cols>'
        findings = run(cols + f"<sheetData>{GOOD_ROWS}</sheetData>")
        self.assertIn("Sheet: column D is locked below the stored rows", findings)

    def test_unlocked_column_outside_matrix(self) -> None:
        cols = '<cols><col min="3" max="5" width="9" style="1"/></cols>'
        findings = run(cols + f"<sheetData>{GOOD_ROWS}</sheetData>")
        self.assertIn("Sheet: column E is unlocked below the stored rows", findings)

    def test_unlocked_row_format(self) -> None:
        rows = GOOD_ROWS + '<row r="3" s="1" customFormat="1"/>'
        findings = run(OPEN_COLUMNS + f"<sheetData>{rows}</sheetData>")
        self.assertIn("Sheet: row 3 has an unlocked row format; unlock cells or columns instead", findings)

    def test_unlocked_cells(self) -> None:
        spec = {"role": "Test", "unlocked_cells": ["B2:B3"]}
        rows = '<row r="2"><c r="B2" s="1"/></row><row r="3"><c r="B3" s="1"/><c r="C3" s="1"/></row>'
        findings = run(f"<sheetData>{rows}</sheetData>", spec)
        self.assertEqual(findings, ["Sheet: lock flag differs from the matrix at C3"])

    def test_missing_cell_of_unlocked_range(self) -> None:
        spec = {"role": "Test", "unlocked_cells": ["B5"]}
        findings = run("<sheetData/>", spec)
        self.assertEqual(findings, ["Sheet: lock flag differs from the matrix at B5"])

    def test_locked_input_fill(self) -> None:
        rows = GOOD_ROWS.replace('<c r="B2"/>', '<c r="B2" s="2"/>')
        findings = run(OPEN_COLUMNS + f"<sheetData>{rows}</sheetData>")
        self.assertEqual(findings, ["Sheet: a locked cell has the input fill at B2"])

    def test_sheet_names_must_match(self) -> None:
        findings = run(OPEN_COLUMNS + f"<sheetData>{GOOD_ROWS}</sheetData>", name="Other")
        self.assertEqual(findings, [f"{gate.MATRIX}: template sheet Other is missing from the matrix",
                                    f"{gate.MATRIX}: Sheet is not a sheet of the template"])

    def test_sheet_saved_unprotected(self) -> None:
        findings = run(OPEN_COLUMNS + f"<sheetData>{GOOD_ROWS}</sheetData>", protection="")
        self.assertEqual(findings, ["Sheet: the sheet is not saved protected"])

    def test_working_sheet_is_saved_unprotected_and_listed(self) -> None:
        # A sheet of the role Working is never protected (#720): the template saves it open and
        # CORE_Config lists it in PROTECT_NONE.
        spec = {"role": "Working"}
        sheet = f"<sheetData>{LOCKED_ROWS}</sheetData>"
        self.assertEqual(run(sheet, spec, protection="", listed=("PROTECT_NONE",)), [])

    def test_working_sheet_saved_protected(self) -> None:
        spec = {"role": "Working"}
        sheet = f"<sheetData>{LOCKED_ROWS}</sheetData>"
        findings = run(sheet, spec, listed=("PROTECT_NONE",))
        self.assertEqual(findings, ["Sheet: a sheet of the role Working must not be saved protected"])

    def test_working_sheet_must_be_in_protect_none(self) -> None:
        spec = {"role": "Working"}
        sheet = f"<sheetData>{LOCKED_ROWS}</sheetData>"
        findings = run(sheet, spec, protection="")
        self.assertEqual(findings, [f"{gate.CONFIG}: PROTECT_NONE lists [], the matrix ['Sheet']"])

    def test_protect_none_must_not_list_a_protected_sheet(self) -> None:
        findings = run(OPEN_COLUMNS + f"<sheetData>{GOOD_ROWS}</sheetData>", listed=("PROTECT_NONE",))
        self.assertEqual(findings, [f"{gate.CONFIG}: PROTECT_NONE lists ['Sheet'], the matrix []"])

    def test_the_excel_script_keeps_the_protected_sheet_list(self) -> None:
        # The list of the sheets that must be protected is assigned once, at the top of the script
        # (#720): a later assignment of the same variable, for the sheets of the workbook, made the
        # sheets of the role Working count as unprotected.
        script = (ROOT / "tools" / "Check-SACCR-Protection.ps1").read_text(encoding="utf-8")
        self.assertEqual(len(re.findall(r"\$(?:script:)?sheetNames\s*=", script)), 1)
        self.assertIn('[string]$property.Value.role -eq "Working"', script)

    def test_the_excel_script_finds_its_cells_by_name(self) -> None:
        # #756: the cells the script writes are found from the workbook names, so that rows moved in the
        # template need no edit to the script; the alpha it changes is the one named Alpha.
        script = (ROOT / "tools" / "Check-SACCR-Protection.ps1").read_text(encoding="utf-8")
        self.assertNotRegex(script, r'Worksheets\.Item\("(?:Params|NettingSets|Trades|Import_\w+|Bank\w+|Results)"\)'
                                    r'\.Range\("')
        self.assertIn('$alpha = Get-NamedCell $workbook "Alpha"', script)
        self.assertIn("$Book.Names.Item($Name).RefersToRange", script)

    def test_sheet_protection_password(self) -> None:
        protection = PROTECTED.replace("/>", ' algorithmName="SHA-512" hashValue="x" saltValue="y" spinCount="1"/>')
        findings = run(OPEN_COLUMNS + f"<sheetData>{GOOD_ROWS}</sheetData>", protection=protection)
        self.assertEqual(findings, ["Sheet: the protection has a password"])

    def test_sheet_protection_options(self) -> None:
        spec = dict(SPEC, allow=["sorting"])
        protection = PROTECTED.replace("/>", ' formatCells="0" autoFilter="0"/>')
        findings = run(OPEN_COLUMNS + f"<sheetData>{GOOD_ROWS}</sheetData>", spec, protection=protection,
                       listed=("PROTECT_SORTING",))
        self.assertEqual(findings, ["Sheet: the saved protection forbids sorting, against the matrix",
                                    "Sheet: the saved protection allows filtering, against the matrix",
                                    "Sheet: the protection allows formatCells"])

    def test_filtering_needs_autofilter(self) -> None:
        spec = dict(SPEC, allow=["filtering"])
        protection = PROTECTED.replace("/>", ' autoFilter="0"/>')
        sheet = OPEN_COLUMNS + f"<sheetData>{GOOD_ROWS}</sheetData>"
        findings = run(sheet, spec, protection=protection, listed=("PROTECT_FILTERING",))
        self.assertEqual(findings, ["Sheet: filtering is allowed but the sheet has no AutoFilter, "
                                    "and a protected sheet cannot get one"])
        self.assertEqual(run(sheet, spec, protection=protection + '<autoFilter ref="C1:D2"/>',
                             listed=("PROTECT_FILTERING",)), [])

    def test_structure_saved_unprotected(self) -> None:
        findings = run(OPEN_COLUMNS + f"<sheetData>{GOOD_ROWS}</sheetData>", structure="")
        self.assertEqual(findings, [f"{gate.TEMPLATE}: the workbook structure is not saved protected"])

    def test_config_lists(self) -> None:
        spec = dict(SPEC, allow=["outlining"])
        findings = run(OPEN_COLUMNS + f"<sheetData>{GOOD_ROWS}</sheetData>", spec, listed=("PROTECT_FILTERING",))
        self.assertEqual(findings, [
            f"{gate.CONFIG}: PROTECT_FILTERING lists ['Sheet'], the matrix []",
            f"{gate.CONFIG}: PROTECT_OUTLINING lists [], the matrix ['Sheet']"])

    def test_matrix_shape(self) -> None:
        spec = {"role": "Input", "extra": 1, "unlocked_cells": ["B2:A1"], "allow": ["copying"],
                "unlocked_columns": {"columns": "D:C", "first_row": 1}}
        findings = gate.check_matrix({"schema_version": 1, "sheets": {"Sheet": spec}})
        self.assertEqual(findings, [
            f"{gate.MATRIX}: Sheet: unknown key 'extra'",
            f"{gate.MATRIX}: Sheet: reversed range: 'B2:A1'",
            f"{gate.MATRIX}: Sheet: unknown allowed operation 'copying'",
            f"{gate.MATRIX}: Sheet: unlocked_columns.columns must be a column range such as C:S",
            f"{gate.MATRIX}: Sheet: unlocked_columns.first_row must be an integer of at least 2 or a workbook name"])

    def test_the_matrix_names_cells_by_workbook_name(self) -> None:
        # #756: a name, NAME+N rows below it, and a first row given as the name of the header cell
        names = {"Rate": "Params!$E$5", "Table": "Params!$C$9:$E$12", "Header": "Trades!$C$4"}
        spec = {"role": "Input", "unlocked_cells": ["Rate", "Table+2", "B2"],
                "unlocked_columns": {"columns": "C:D", "first_row": "Header"}}
        self.assertEqual(gate.check_matrix({"schema_version": 1, "sheets": {"Sheet": spec}}), [])
        cells, columns, first = gate.expected_unlocked(spec, names)
        self.assertEqual(cells, {(5, 5), (3, 10), (4, 10), (5, 10), (3, 11), (4, 11), (5, 11), (2, 2)})
        self.assertEqual((columns, first), ({3, 4}, 5))
        with self.assertRaisesRegex(ValueError, "no workbook name 'Missing'"):
            gate.expected_unlocked({"unlocked_cells": ["Missing"]}, names)
        bad = gate.check_matrix({"schema_version": 1, "sheets": {"Sheet": {"role": "Input", "unlocked_cells": ["E5 F6"]}}})
        self.assertEqual(bad, [f"{gate.MATRIX}: Sheet: not a cell address or workbook name: 'E5 F6'"])


if __name__ == "__main__":
    unittest.main()
