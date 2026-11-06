"""The Ribbon source agrees with its contract, docs/workbook/RIBBON.md (#213).

The RibbonX source src/workbook/customUI14.xml and the contract's tables must
name the same controls with the same callbacks, macros and sheets. Every
macro must be a public procedure of M_Main and every sheet a document module
of the template, except the targets that later issues of the milestone
deliver, which are listed below with their issue.
"""
import json
import re
import unittest
import xml.etree.ElementTree as ElementTree
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
RIBBON = ROOT / "src" / "workbook" / "customUI14.xml"
CONTRACT = ROOT / "docs" / "workbook" / "RIBBON.md"
M_MAIN = ROOT / "src" / "workbook" / "M_Main.bas"
MANIFEST = ROOT / "tools" / "vba-projects.json"
M_RIBBON = ROOT / "src" / "workbook" / "M_Ribbon.bas"

NAMESPACE = "http://schemas.microsoft.com/office/2009/07/customui"
CALLBACKS = {"RibbonOnLoad", "RibbonAction", "RibbonGoTo", "RibbonGetEnabled", "RibbonGetLabel"}
CALLBACK_ATTRIBUTES = ("onLoad", "onAction", "getEnabled", "getLabel")

# Targets delivered by a later issue of the milestone. Remove an entry when its
# issue lands, so the target is then checked like every other.
PLANNED_MACROS: dict[str, str] = {}
PLANNED_SHEETS: dict[str, str] = {}


def local(tag: str) -> str:
    return tag.split("}", 1)[-1]


def ribbon_root() -> ElementTree.Element:
    data = RIBBON.read_bytes()
    if b"<!DOCTYPE" in data or b"<!ENTITY" in data:
        raise AssertionError("Ribbon DTD declarations are forbidden")
    return ElementTree.fromstring(data)


def controls_with(root: ElementTree.Element, attribute: str, value: str) -> list[str]:
    return [element.get("id", "") for element in root.iter() if element.get(attribute) == value]


def contract_table(heading: str) -> list[list[str]]:
    """The body rows of the first table under a level-2 heading of the contract."""
    text = CONTRACT.read_text(encoding="utf-8")
    section = re.split(r"^## ", text, flags=re.M)
    body = next(part for part in section if part.split("\n", 1)[0].strip().endswith(heading))
    rows = [line for line in body.splitlines() if line.startswith("|")]
    return [[cell.strip() for cell in row.strip("|").split("|")] for row in rows[2:]]


def vba_map(function: str) -> dict[str, str]:
    """The Select Case map of a pure M_Ribbon function: control ID to the string it returns."""
    text = M_RIBBON.read_text(encoding="cp1252").replace("\r\n", "\n")
    body = re.search(rf"^Public Function {function}\(.*?^End Function", text, flags=re.M | re.S)
    if not body:
        raise AssertionError(f"M_Ribbon has no function {function}")
    mapping: dict[str, str] = {}
    cases: list[str] = []
    for line in body.group(0).splitlines():
        case = re.match(r'\s*Case ("[^"]+"(?:\s*,\s*"[^"]+")*)\s*$', line)
        result = re.match(rf'\s*{function} = "([^"]*)"', line)
        if case:
            cases = re.findall(r'"([^"]+)"', case.group(1))
        elif result and cases:
            mapping.update({control_id: result.group(1) for control_id in cases})
            cases = []
    return mapping


def code(cell: str) -> str:
    match = re.match(r"`([^`]+)`", cell)
    if not match:
        raise AssertionError(f"expected a code span, found {cell!r}")
    return match.group(1)


class RibbonContractTests(unittest.TestCase):
    def setUp(self) -> None:
        self.root = ribbon_root()

    def test_namespace_and_root(self) -> None:
        self.assertEqual(self.root.tag, f"{{{NAMESPACE}}}customUI")
        self.assertEqual(self.root.get("onLoad"), "RibbonOnLoad")
        tabs = [element for element in self.root.iter() if local(element.tag) == "tab"]
        self.assertEqual([(tab.get("id"), tab.get("label")) for tab in tabs], [("ksTab", "K-SACCR")])

    def test_ids_are_unique_and_prefixed(self) -> None:
        ids = [element.get("id") for element in self.root.iter() if element.get("id")]
        self.assertEqual(len(ids), len(set(ids)), "control IDs must be unique")
        for control_id in ids:
            with self.subTest(control=control_id):
                self.assertRegex(str(control_id), r"^ks[A-Z][A-Za-z]+$")

    def test_only_contract_callbacks(self) -> None:
        for element in self.root.iter():
            for attribute in CALLBACK_ATTRIBUTES:
                value = element.get(attribute)
                if value is not None:
                    with self.subTest(control=element.get("id"), attribute=attribute):
                        self.assertIn(value, CALLBACKS)

    def test_commands_match_contract(self) -> None:
        documented = {code(row[0]): code(row[3]) for row in contract_table("Commands")}
        self.assertEqual(set(controls_with(self.root, "onAction", "RibbonAction")), set(documented))
        public = set(re.findall(r"^Public Sub (\w+)\(\)", M_MAIN.read_text(encoding="cp1252"), flags=re.M))
        for control_id, macro in documented.items():
            with self.subTest(control=control_id, macro=macro):
                self.assertTrue(macro in public or macro in PLANNED_MACROS,
                                f"{macro} is not a public M_Main macro")

    def test_every_command_has_an_enabled_state(self) -> None:
        # The primary button of a split button takes its state from the split button.
        parents = {child: parent for parent in self.root.iter() for child in parent}
        for control_id in controls_with(self.root, "onAction", "RibbonAction"):
            element = next(e for e in self.root.iter() if e.get("id") == control_id)
            parent = parents[element]
            owner = parent if local(parent.tag) == "splitButton" else element
            with self.subTest(control=control_id):
                self.assertEqual(owner.get("getEnabled"), "RibbonGetEnabled")

    def test_navigation_matches_contract(self) -> None:
        documented = {code(row[1]): code(row[3]) for row in contract_table("Navigation")}
        self.assertEqual(set(controls_with(self.root, "onAction", "RibbonGoTo")), set(documented))
        manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
        sheets = {Path(path).stem for path in manifest["document_modules"]} - {"ThisWorkbook"}
        targets = set(documented.values())
        for sheet in targets:
            with self.subTest(sheet=sheet):
                self.assertTrue(sheet in sheets or sheet in PLANNED_SHEETS, f"{sheet} is not a sheet of the template")
        self.assertEqual(sheets - targets, set(), "every sheet needs a Go to entry")

    def test_vba_command_map_matches_contract(self) -> None:
        documented = {code(row[0]): code(row[3]) for row in contract_table("Commands")}
        self.assertEqual(vba_map("RibbonMacro"), documented)
        text = M_RIBBON.read_text(encoding="cp1252")
        dispatched = set(re.findall(r"^\s*M_Main\.(\w+)\s*$", text, flags=re.M))
        self.assertEqual(dispatched, set(documented.values()) - set(PLANNED_MACROS),
                         "RibbonAction must call every mapped macro, and only those")

    def test_vba_sheet_map_matches_contract(self) -> None:
        documented = {code(row[1]): code(row[3]) for row in contract_table("Navigation")}
        self.assertEqual(vba_map("RibbonSheet"), documented)

    def test_status_matches_contract(self) -> None:
        documented = {code(row[0]) for row in contract_table("Status")}
        labels = {element.get("id") for element in self.root.iter() if local(element.tag) == "labelControl"}
        self.assertEqual(labels, documented)
        self.assertEqual(set(controls_with(self.root, "getLabel", "RibbonGetLabel")), documented)

    def test_top_level_keytips_are_unique(self) -> None:
        for container in self.root.iter():
            if local(container.tag) not in {"tab", "menu"}:
                continue
            keytips: list[str] = []
            for group_or_item in container:
                items = list(group_or_item) if local(group_or_item.tag) == "group" else [group_or_item]
                keytips += [item.get("keytip", "") for item in items if item.get("keytip")]
            with self.subTest(container=container.get("id")):
                self.assertEqual(len(keytips), len(set(keytips)), f"repeated keytip in {container.get('id')}: {keytips}")


if __name__ == "__main__":
    unittest.main()
