"""The Dashboard sheet of the template against its specification (#675)."""
from __future__ import annotations

import posixpath
import re
import unittest
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree

import dashboard as db

TEMPLATE = db.ROOT / "src" / "workbook" / "SACCR_Template.xlsx"
MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PKG = "{http://schemas.openxmlformats.org/package/2006/relationships}"
CHART = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
SHEET = "Dashboard"
REFERENCE = re.compile(r"^\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+)$")


def column_number(letters: str) -> int:
    number = 0
    for letter in letters:
        number = number * 26 + ord(letter) - 64
    return number


def column_letters(number: int) -> str:
    letters = ""
    while number:
        number, rest = divmod(number - 1, 26)
        letters = chr(65 + rest) + letters
    return letters


def formulas_and_values(xml: bytes, shared: list[str]) -> dict[str, tuple[str | None, str | None]]:
    """Each cell of a sheet: its text or value, and its formula."""
    cells: dict[str, tuple[str | None, str | None]] = {}
    for cell in ElementTree.fromstring(xml).iter(f"{MAIN}c"):
        value, formula = cell.find(f"{MAIN}v"), cell.find(f"{MAIN}f")
        text = None
        if value is not None and value.text is not None:
            text = shared[int(value.text)] if cell.attrib.get("t") == "s" else value.text
        cells[cell.attrib["r"]] = (text, formula.text if formula is not None else None)
    return cells


class DashboardSheetTests(unittest.TestCase):
    parts: dict[str, bytes]
    spec: dict[str, Any]
    sheets: list[tuple[str, str]]
    names: dict[str, str]
    part: str
    xml: str
    shared: list[str]
    cells: dict[str, tuple[str | None, str | None]]

    @classmethod
    def setUpClass(cls) -> None:
        with zipfile.ZipFile(TEMPLATE) as package:
            cls.parts = {name: package.read(name) for name in package.namelist()}
        cls.spec = db.load_spec()
        workbook = ElementTree.fromstring(cls.parts["xl/workbook.xml"])
        rels = ElementTree.fromstring(cls.parts["xl/_rels/workbook.xml.rels"])
        targets = {r.attrib["Id"]: r.attrib["Target"] for r in rels.iter(f"{PKG}Relationship")}
        cls.sheets = [(s.attrib["name"], "xl/" + targets[s.attrib[f"{REL}id"]]) for s in workbook.iter(f"{MAIN}sheet")]
        cls.names = {d.attrib["name"]: d.text or "" for d in workbook.iter(f"{MAIN}definedName")}
        cls.part = dict(cls.sheets)[SHEET]
        cls.xml = cls.parts[cls.part].decode("utf-8")
        strings = ElementTree.fromstring(cls.parts["xl/sharedStrings.xml"])
        cls.shared = ["".join(t.text or "" for t in si.iter(f"{MAIN}t")) for si in strings.iter(f"{MAIN}si")]
        cls.cells = formulas_and_values(cls.parts[cls.part], cls.shared)

    def block_range(self, name: str) -> tuple[int, int, int, int]:
        target = self.names.get(name, "")
        self.assertTrue(target.startswith(f"{SHEET}!"), f"{name} is not a name on {SHEET}: {target!r}")
        found = REFERENCE.match(target.split("!", 1)[1])
        assert found is not None, target
        c1, r1, c2, r2 = found.groups()
        return int(r1), column_number(c1), int(r2), column_number(c2)

    def text(self, row: int, column: int) -> str | None:
        return self.cells.get(f"{column_letters(column)}{row}", (None, None))[0]

    def test_sheet_is_the_first_output_with_its_code_name(self) -> None:
        names = [name for name, _ in self.sheets]
        self.assertEqual(names.index(SHEET) + 1, names.index("Results"))
        self.assertIn('codeName="shDashboard"', self.xml)
        self.assertIn('<tabColor rgb="FF2E75B6"/>', self.xml)
        self.assertIn('<sheetProtection sheet="1" objects="1" scenarios="1"/>', self.xml)

    def test_every_block_has_a_name_of_its_shape(self) -> None:
        for block in self.spec["blocks"]:
            with self.subTest(block=block["name"]):
                r1, c1, r2, c2 = self.block_range(block["name"])
                rows = block.get("rank_rows") or len(block["rows"])
                self.assertEqual((r2 - r1 + 1, c2 - c1 + 1), (rows, len(block["columns"])))

    def test_labels_and_headers_follow_the_specification(self) -> None:
        for block in self.spec["blocks"]:
            with self.subTest(block=block["name"]):
                r1, c1, _, _ = self.block_range(block["name"])
                headers = [self.text(r1 - 1, c1 + i) for i in range(len(block["columns"]))]
                if "rank_rows" not in block:
                    labels = [self.text(r1 + i, c1 - 1) for i in range(len(block["rows"]))]
                    self.assertEqual(labels, [r["label"] for r in block["rows"]])
                self.assertEqual(headers, block["columns"])

    def test_the_template_holds_no_figure(self) -> None:
        for block in self.spec["blocks"]:
            r1, c1, r2, c2 = self.block_range(block["name"])
            filled = [f"{column_letters(c)}{r}" for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)
                      if self.cells.get(f"{column_letters(c)}{r}", (None, None)) != (None, None)]
            self.assertEqual(filled, [], block["name"])

    def test_status_shows_the_status_of_results(self) -> None:
        results = formulas_and_values(self.parts[dict(self.sheets)["Results"]], self.shared)
        self.assertIsNotNone(self.cells["C5"][1])
        self.assertEqual(self.cells["C5"][1], results["C5"][1])
        self.assertIn('<conditionalFormatting sqref="C5"><cfRule type="beginsWith"', self.xml)

    def test_no_formula_reads_another_sheet_but_the_status(self) -> None:
        for reference, (_, formula) in self.cells.items():
            if formula is None or reference == "C5":
                continue
            self.assertNotIn("!", formula, reference)

    def test_charts_read_only_the_dashboard(self) -> None:
        rels = self.parts[posixpath.join(posixpath.dirname(self.part), "_rels",
                                         posixpath.basename(self.part) + ".rels")].decode("utf-8")
        found = re.search(r'Target="\.\./drawings/([^"]+)"', rels)
        assert found is not None
        drawing = "xl/drawings/" + found.group(1)
        drawing_rels = self.parts[f"xl/drawings/_rels/{Path(drawing).name}.rels"].decode("utf-8")
        charts = ["xl/charts/" + name for name in re.findall(r'Target="\.\./charts/([^"]+)"', drawing_rels)]
        self.assertEqual(len(charts), 10)
        for chart in charts:
            formulas = [f.text or "" for f in ElementTree.fromstring(self.parts[chart]).iter(f"{CHART}f")]
            self.assertTrue(formulas, chart)
            for formula in formulas:
                self.assertTrue(formula.startswith(f"{SHEET}!$"), f"{chart}: {formula}")

    def test_home_links_the_sheet(self) -> None:
        home = self.parts[dict(self.sheets)["Home"]].decode("utf-8")
        self.assertIn("location=\"'Dashboard'!A1\"", home)


if __name__ == "__main__":
    unittest.main()
