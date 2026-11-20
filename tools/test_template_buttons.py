"""Tests of the buttons of the workbook template and of the Ribbon (#719, #727).

Every command on a sheet is a rounded-rectangle shape with its macro assigned, in the title row of the sheet,
and the template holds no form control any more. The Ribbon holds Clear outputs as a button of its own.
"""

from __future__ import annotations

import re
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "src" / "workbook" / "SACCR_Template.xlsx"
MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"
UI = "{http://schemas.microsoft.com/office/2009/07/customui}"
XDR = "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"

NAVY, BLUE = "1F3864", "DDEBF7"  # BLUE: the light fill of a secondary command, with navy text
CLEAR = (("btn_ClearOutputs", "ClearOutputs", "Clear outputs", BLUE),)
STAGING = (("btn_ImportCsv", "{import}", "Import CSV", NAVY),
           ("btn_ValidateStaging", "ValidateStaging", "Validate staging", NAVY),
           ("btn_CommitImport", "CommitImport", "Commit import", NAVY),
           ("btn_ExportImportSpecification", "ExportImportSpecification", "Export import specifications", BLUE))
# Sheet: (name, macro, label, fill) of each shape, in order from the left.
EXPECTED = {
    "Results": CLEAR, "HedgingSets": CLEAR, "Buckets": CLEAR, "TradeCalc": CLEAR, "Checks": CLEAR,
    "Reconciliation": CLEAR, "TradeReconciliation": CLEAR,
    "FormulaTest": (("btn_RandomizeInputs", "RandomizeFormulaTestInputs", "Randomize inputs", NAVY),
                    ("btn_RestoreDefaults", "RestoreFormulaTestDefaults", "Restore defaults", BLUE)),
    "Import_NettingSets": tuple((a, b.replace("{import}", "ImportNettingSetsCsv"), c, d) for a, b, c, d in STAGING),
    "Import_Trades": tuple((a, b.replace("{import}", "ImportTradesCsv"), c, d) for a, b, c, d in STAGING),
}
# The shapes start no further than this right of the left edge of column C, and clear of the title.
LATEST_START_PX = 480
EMU = 9525
GAP_PX = 8


def sheet_parts(package: zipfile.ZipFile) -> dict[str, str]:
    """The part name of every sheet of the workbook, by sheet name."""
    workbook = ElementTree.fromstring(package.read("xl/workbook.xml"))
    relationships = {item.attrib["Id"]: item.attrib["Target"]
                     for item in ElementTree.fromstring(package.read("xl/_rels/workbook.xml.rels")).iter(REL + "Relationship")}
    return {item.attrib["name"]: "xl/" + relationships[item.attrib[RID]].lstrip("/").removeprefix("xl/")
            for item in workbook.iter(MAIN + "sheet")}


def drawing_of(package: zipfile.ZipFile, part: str) -> str:
    rels = package.read(part.replace("worksheets/", "worksheets/_rels/") + ".rels").decode("utf-8")
    match = re.search(r'Type="[^"]*/drawing" Target="\.\./drawings/(drawing\d+\.xml)"', rels)
    assert match is not None, part
    return "xl/drawings/" + match.group(1)


def shape_name(shape: ElementTree.Element) -> str:
    properties = shape.find(f"{XDR}nvSpPr/{XDR}cNvPr")
    return properties.attrib.get("name", "") if properties is not None else ""


def column_widths(sheet: str) -> dict[int, int]:
    out: dict[int, int] = {}
    for low, high, width in re.findall(r'<col min="(\d+)" max="(\d+)" width="([\d.]+)"', sheet):
        for column in range(int(low), int(high) + 1):
            out[column] = int(float(width) * 7 + 5)
    return out


class ShapeButtonTests(unittest.TestCase):
    def setUp(self) -> None:
        self.package = zipfile.ZipFile(TEMPLATE)
        self.addCleanup(self.package.close)
        self.parts = sheet_parts(self.package)

    def shapes(self, name: str) -> list[ElementTree.Element]:
        drawing = ElementTree.fromstring(self.package.read(drawing_of(self.package, self.parts[name])))
        return [item for item in drawing.iter(XDR + "sp") if shape_name(item).startswith("btn_")]

    def test_the_template_has_no_form_control(self) -> None:
        names = self.package.namelist()
        self.assertEqual([item for item in names if "ctrlProp" in item or item.endswith(".vml")], [])
        self.assertNotIn("ctrlProp", self.package.read("[Content_Types].xml").decode("utf-8"))
        for name, part in self.parts.items():
            sheet = self.package.read(part).decode("utf-8")
            self.assertNotIn("legacyDrawing", sheet, name)
            self.assertNotIn("<controls>", sheet, name)
            drawing = self.package.read(drawing_of(self.package, part)).decode("utf-8")
            self.assertNotRegex(drawing, r'name="Button \d+"', name)

    def test_every_sheet_has_its_shapes_with_their_macros(self) -> None:
        for name, expected in EXPECTED.items():
            found = self.shapes(name)
            self.assertEqual(len(found), len(expected), name)
            for shape, (button, macro, label, fill) in zip(found, expected, strict=True):
                self.assertEqual(shape.attrib["macro"], f"[0]!{macro}", (name, button))
                self.assertEqual(shape_name(shape), button)
                geometry = shape.find(f"{XDR}spPr/{A}prstGeom")
                assert geometry is not None
                self.assertEqual(geometry.attrib["prst"], "roundRect", (name, button))
                color = shape.find(f"{XDR}spPr/{A}solidFill/{A}srgbClr")
                assert color is not None
                self.assertEqual(color.attrib["val"], fill, (name, button))
                self.assertEqual("".join(item.text or "" for item in shape.iter(A + "t")), label, (name, button))

    def test_shapes_sit_in_the_title_row_left_of_the_old_buttons_and_clear_of_the_title(self) -> None:
        strings = [re.sub("<[^>]+>", "", item) for item in
                   re.findall(r"<si>(.*?)</si>", self.package.read("xl/sharedStrings.xml").decode("utf-8"), re.S)]
        for name in EXPECTED:
            sheet = self.package.read(self.parts[name]).decode("utf-8")
            title = strings[int(re.search(r'<c r="B2"[^>]*t="s"[^>]*><v>(\d+)</v>', sheet).group(1))]  # type: ignore[union-attr]
            widths = column_widths(sheet)
            drawing = ElementTree.fromstring(self.package.read(drawing_of(self.package, self.parts[name])))
            groups = [anchor for anchor in drawing.iter(XDR + "oneCellAnchor")
                      if anchor.find(f"{XDR}grpSp/{XDR}nvGrpSpPr/{XDR}cNvPr") is not None
                      and anchor.find(f"{XDR}grpSp/{XDR}nvGrpSpPr/{XDR}cNvPr").attrib["name"] == "btn_Group"]  # type: ignore[union-attr]
            self.assertEqual(len(groups), 1, name)
            anchor = groups[0]
            self.assertEqual(anchor.findtext(f"{XDR}from/{XDR}row"), "1", name)
            column = int(anchor.findtext(f"{XDR}from/{XDR}col") or "0")
            offset = int(anchor.findtext(f"{XDR}from/{XDR}colOff") or "0") // EMU
            start = sum(widths.get(index, 64) for index in range(3, column + 1)) + offset
            self.assertGreaterEqual(start, 12 * len(title) - 17, f"{name}: the shapes cover the title")
            self.assertLessEqual(start, LATEST_START_PX, name)
            # Inside the group the shapes follow one another with the same gap, whatever the column widths.
            edge = None
            for shape in anchor.iter(XDR + "sp"):
                position = shape.find(f"{XDR}spPr/{A}xfrm/{A}off")
                size = shape.find(f"{XDR}spPr/{A}xfrm/{A}ext")
                assert position is not None and size is not None
                left = int(position.attrib["x"]) // EMU
                if edge is not None:
                    self.assertEqual(left - edge, GAP_PX, name)
                edge = left + int(size.attrib["cx"]) // EMU

    def test_shape_identifiers_are_unique_in_each_drawing(self) -> None:
        for name in EXPECTED:
            drawing = self.package.read(drawing_of(self.package, self.parts[name])).decode("utf-8")
            ids = re.findall(r'<xdr:cNvPr id="(\d+)"', drawing)
            self.assertEqual(len(ids), len(set(ids)), name)

    def test_the_macros_exist_and_are_public(self) -> None:
        source = "".join(path.read_text(encoding="cp1252") for path in (ROOT / "src" / "workbook").glob("*.bas"))
        for expected in EXPECTED.values():
            for _, macro, _, _ in expected:
                self.assertRegex(source, rf"(?m)^Public Sub {macro}\(\)", macro)


class RibbonButtonTests(unittest.TestCase):
    def test_clear_outputs_is_a_button_of_the_calculation_group(self) -> None:
        tree = ElementTree.parse(ROOT / "src" / "workbook" / "customUI14.xml").getroot()
        parents = {child: parent for parent in tree.iter() for child in parent}
        buttons = [item for item in tree.iter(UI + "button") if item.attrib["id"] == "ksClearOutputs"]
        self.assertEqual(len(buttons), 1)
        parent = parents[buttons[0]]
        self.assertEqual(parent.tag, UI + "group")
        self.assertEqual(parent.attrib["id"], "ksGrpCalculation")
        self.assertEqual(buttons[0].attrib["onAction"], "RibbonAction")

    def test_the_ribbon_maps_the_button_to_the_macro(self) -> None:
        source = (ROOT / "src" / "workbook" / "M_Ribbon.bas").read_text(encoding="utf-8")
        self.assertRegex(source, r'Case "ksClearOutputs"\s+RibbonMacro = "ClearOutputs"')


if __name__ == "__main__":
    unittest.main()
