"""Tests of the Clear outputs entry points of the workbook template (#719).

Every output sheet carries a form-control button of the macro ``ClearOutputs``, and the Ribbon holds
the command as a button of its own in the Calculation group, not inside a menu.
"""

from __future__ import annotations

import re
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "src" / "workbook" / "SACCR_Template.xlsx"
OUTPUT_SHEETS = ("Results", "HedgingSets", "Buckets", "TradeCalc", "Checks", "Reconciliation",
                 "TradeReconciliation")
MACRO = "ClearOutputs"
MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"
UI = "{http://schemas.microsoft.com/office/2009/07/customui}"


def sheet_parts(package: zipfile.ZipFile) -> dict[str, str]:
    """The part name of every sheet of the workbook, by sheet name."""
    workbook = ElementTree.fromstring(package.read("xl/workbook.xml"))
    relationships = {item.attrib["Id"]: item.attrib["Target"]
                     for item in ElementTree.fromstring(package.read("xl/_rels/workbook.xml.rels")).iter(REL + "Relationship")}
    rid = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    return {item.attrib["name"]: "xl/" + relationships[item.attrib[rid]].lstrip("/").removeprefix("xl/")
            for item in workbook.iter(MAIN + "sheet")}


class OutputSheetButtonTests(unittest.TestCase):
    def setUp(self) -> None:
        self.package = zipfile.ZipFile(TEMPLATE)
        self.addCleanup(self.package.close)
        self.parts = sheet_parts(self.package)

    def test_every_output_sheet_has_the_button(self) -> None:
        for name in OUTPUT_SHEETS:
            part = self.parts[name]
            sheet = self.package.read(part).decode("utf-8")
            self.assertEqual(len(re.findall(rf'macro="\[0\]!{MACRO}"', sheet)), 1, name)
            rels = self.package.read(part.replace("worksheets/", "worksheets/_rels/") + ".rels").decode("utf-8")
            self.assertIn("vmlDrawing", rels, name)
            self.assertIn("ctrlProp", rels, name)

    def test_the_vml_shape_calls_the_same_macro(self) -> None:
        for name in OUTPUT_SHEETS:
            part = self.parts[name]
            rels = self.package.read(part.replace("worksheets/", "worksheets/_rels/") + ".rels").decode("utf-8")
            vml = re.search(r'Target="\.\./drawings/(vmlDrawing\d+\.vml)"', rels)
            self.assertIsNotNone(vml, name)
            assert vml is not None
            shape = self.package.read("xl/drawings/" + vml.group(1)).decode("utf-8")
            self.assertEqual(re.findall(r"<x:FmlaMacro>(.*?)</x:FmlaMacro>", shape), [f"[0]!{MACRO}"], name)
            self.assertIn("Clear outputs", shape, name)

    def test_shape_ids_are_unique_in_the_workbook(self) -> None:
        ids: list[str] = []
        for name in self.package.namelist():
            if name.endswith(".vml"):
                ids += re.findall(r'<v:shape [^>]*id="(_x0000_s\d+)"', self.package.read(name).decode("utf-8"))
        self.assertEqual(len(ids), len(set(ids)), ids)

    def test_the_macro_is_public_in_the_main_module(self) -> None:
        source = (ROOT / "src" / "workbook" / "M_Main.bas").read_text(encoding="utf-8")
        self.assertRegex(source, rf"(?m)^Public Sub {MACRO}\(\)")


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
        self.assertRegex(source, rf'Case "ksClearOutputs"\s+RibbonMacro = "{MACRO}"')


if __name__ == "__main__":
    unittest.main()
