"""Tests of the header row of the input sheets of the workbook template (#756).

Every column header of NettingSets and Trades carries a note: its title is the header, and its text
says what the field is, the values allowed, what a blank means and the CSV field name. The three input
sheets freeze rows 1 to 4, as Home does, so that the view does not jump from one sheet to the next. NettingSets
and Trades have the header on row 4, the last frozen row, so that it stays in view (#811); Params has it on row
5, below an empty row 4. The header rows are found from their workbook names.
"""

from __future__ import annotations

import json
import re
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree

from template_layout import MAIN, column_letters, column_number, named_cell
from test_template_buttons import sheet_parts

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "src" / "workbook" / "SACCR_Template.xlsx"
CONTRACT = ROOT / "docs" / "csv" / "contract.json"
INPUT_SHEETS = {"Params": "Params_Code", "NettingSets": "NettingSets_Header", "Trades": "Trades_Header"}
HEADER_ROW = {"Params": 5, "NettingSets": 4, "Trades": 4}  # NettingSets and Trades: in the frozen rows (#811)
FROZEN_ROWS = 4  # rows 1 to 4 (owner, 2026-11-25)
# Rows 3 and 4 of NettingSets and Trades, and 3 to 5 of Params: row 4 is raised for the navigation bar's top
# block, which holds four input buttons since #831; row 3 and the header of Params keep 30.
TOP_ROW_HEIGHTS = {3: 30.0, 4: 40.65, 5: 30.0}


class HeaderNoteTests(unittest.TestCase):
    def setUp(self) -> None:
        self.package = zipfile.ZipFile(TEMPLATE)
        self.addCleanup(self.package.close)
        self.parts = sheet_parts(self.package)
        self.strings = ["".join(t.text or "" for t in item.iter(MAIN + "t")) for item in
                        ElementTree.fromstring(self.package.read("xl/sharedStrings.xml")).findall(MAIN + "si")]
        self.contract = json.loads(CONTRACT.read_text(encoding="utf-8"))

    def sheet(self, name: str) -> ElementTree.Element:
        return ElementTree.fromstring(self.package.read(self.parts[name]))

    def notes(self, name: str) -> dict[str, str]:
        """The notes of a sheet: cell reference to text."""
        rels = self.package.read(self.parts[name].replace("worksheets/", "worksheets/_rels/") + ".rels").decode("utf-8")
        target = re.search(r'Type="[^"]*/comments" Target="\.\./(comments\d+\.xml)"', rels)
        self.assertIsNotNone(target, name)
        assert target is not None
        comments = ElementTree.fromstring(self.package.read("xl/" + target.group(1)))
        return {note.attrib["ref"]: "".join(t.text or "" for t in note.iter(MAIN + "t"))
                for note in comments.iter(MAIN + "comment")}

    def test_every_header_has_a_note_titled_by_the_header_that_names_its_csv_field(self) -> None:
        for kind, anchor in (("NettingSets", "NettingSets_Header"), ("Trades", "Trades_Header")):
            spec = self.contract["files"][kind]
            header = named_cell(self.package, anchor)
            notes = self.notes(kind)
            first = column_number(header.column)
            refs = [f"{column_letters(first + i)}{header.row}" for i in range(len(spec["columns"]))]
            self.assertEqual(sorted(notes), sorted(refs), kind)
            for ref, column in zip(refs, spec["columns"]):
                title, _, body = notes[ref].partition("\n")
                self.assertEqual(title, column["sheet_header"], ref)
                self.assertIn(f"CSV field: {column['name']}", body, ref)
                self.assertGreater(len(body), 60, ref)  # says more than the field name

    def test_the_notes_have_their_drawing_on_the_sheet(self) -> None:
        for kind in ("NettingSets", "Trades"):
            sheet = self.package.read(self.parts[kind]).decode("utf-8")
            self.assertIn("<legacyDrawing ", sheet, kind)
            rels = self.package.read(self.parts[kind].replace("worksheets/", "worksheets/_rels/") + ".rels").decode("utf-8")
            vml = re.search(r'Type="[^"]*/vmlDrawing" Target="\.\./drawings/(vmlDrawing\d+\.vml)"', rels)
            assert vml is not None
            shapes = self.package.read("xl/drawings/" + vml.group(1)).decode("utf-8")
            self.assertEqual(shapes.count('ObjectType="Note"'), len(self.notes(kind)), kind)

    def test_the_input_sheets_and_home_freeze_the_four_rows_above_the_header(self) -> None:
        for name, anchor in INPUT_SHEETS.items():
            self.assertEqual(named_cell(self.package, anchor).row, HEADER_ROW[name], name)
            sheet = self.sheet(name)
            pane = sheet.find(f"{MAIN}sheetViews/{MAIN}sheetView/{MAIN}pane")
            assert pane is not None, name
            self.assertEqual((pane.attrib.get("ySplit"), pane.attrib.get("state")), (str(FROZEN_ROWS), "frozen"), name)
            heights = {row.attrib["r"]: float(row.attrib.get("ht", "15")) for row in sheet.iter(MAIN + "row")}
            self.assertEqual([heights.get(str(row)) for row in range(3, HEADER_ROW[name] + 1)],
                             [TOP_ROW_HEIGHTS[row] for row in range(3, HEADER_ROW[name] + 1)], name)
            if HEADER_ROW[name] > FROZEN_ROWS:
                # Params: row 4 is the empty row between the description and the header
                empty = next(row for row in sheet.iter(MAIN + "row") if row.attrib["r"] == "4")
                self.assertEqual([cell.findtext(MAIN + "v") for cell in empty.iter(MAIN + "c")
                                  if cell.find(MAIN + "v") is not None], [], name)
        # the key columns stay in view: Code on Params, the netting-set ID, the trade and netting-set IDs
        panes = {name: self.sheet(name).find(f"{MAIN}sheetViews/{MAIN}sheetView/{MAIN}pane") for name in INPUT_SHEETS}
        self.assertEqual({name: pane.attrib.get("topLeftCell") for name, pane in panes.items() if pane is not None},
                         {"Params": "D5", "NettingSets": "D5", "Trades": "E5"})
        # Home freezes the same four rows, as high as on the input sheets (owner, 2026-11-25)
        frozen = {}
        for name in ("Home", *INPUT_SHEETS):
            sheet = self.sheet(name)
            pane = sheet.find(f"{MAIN}sheetViews/{MAIN}sheetView/{MAIN}pane")
            assert pane is not None, name
            self.assertEqual(pane.attrib.get("ySplit"), str(FROZEN_ROWS), name)
            heights = {row.attrib["r"]: float(row.attrib.get("ht", "15")) for row in sheet.iter(MAIN + "row")}
            frozen[name] = [heights.get(str(row)) for row in range(1, FROZEN_ROWS + 1)]
        self.assertEqual(len({tuple(rows) for rows in frozen.values()}), 1, frozen)

    def test_the_two_staging_sheets_freeze_rows_of_the_same_height(self) -> None:
        # Moving between Import_NettingSets and Import_Trades leaves the view where it was (owner, 2026-11-25).
        frozen = {}
        for name in ("Import_NettingSets", "Import_Trades"):
            sheet = self.sheet(name)
            pane = sheet.find(f"{MAIN}sheetViews/{MAIN}sheetView/{MAIN}pane")
            assert pane is not None, name
            rows = int(pane.attrib.get("ySplit", "0"))
            heights = {row.attrib["r"]: float(row.attrib.get("ht", "15")) for row in sheet.iter(MAIN + "row")}
            frozen[name] = (rows, [heights.get(str(row), 15.0) for row in range(1, rows + 1)])
        self.assertEqual(frozen["Import_NettingSets"], frozen["Import_Trades"])

    def test_the_staging_sheets_show_the_same_headers(self) -> None:
        for kind, anchor in (("NettingSets", "ImportNettingSets_Header"), ("Trades", "ImportTrades_Header")):
            header = named_cell(self.package, anchor)
            row = next(item for item in self.sheet(header.sheet).iter(MAIN + "row") if item.attrib["r"] == str(header.row))
            texts = [self.strings[int(cell.findtext(MAIN + "v") or "0")] for cell in row.iter(MAIN + "c")
                     if cell.attrib.get("t") == "s"]
            self.assertEqual(texts[1:], [c["sheet_header"] for c in self.contract["files"][kind]["columns"]], kind)


if __name__ == "__main__":
    unittest.main()
