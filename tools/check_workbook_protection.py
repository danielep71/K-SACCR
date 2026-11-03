#!/usr/bin/env python3
"""Check the cell lock flags of the workbook template against its editability
matrix (docs/workbook/protection.json), without running Excel.

A cell is unlocked when the matrix lists it, or when its column is an
unlocked input column and its row is at or below the first data row. Every
other cell is locked. The effective style of a cell is its own style, else
the style of a formatted row, else the style of its column; the check
covers the cells the sheet stores, the rows it formats and the columns to
the last column of the sheet, so a cell that does not exist yet is judged
by the style it would receive. A locked cell may not carry the yellow input
fill. Every sheet and the workbook structure are saved protected, without a
password; each sheet allows sorting, filtering, inserting rows and opening
row groups exactly as the matrix says, filtering only on a sheet with an
AutoFilter, and the PROTECT_ lists of CORE_Config, which
M_Main.ProtectWorkbook applies, say the same."""
from __future__ import annotations

import json
import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import Any, NamedTuple

from _gatelib import parse_report_args, run_gate

MATRIX = "docs/workbook/protection.json"
TEMPLATE = "src/workbook/SACCR_Template.xlsx"
ROLES = {"Instructions", "Input", "Staging", "Output", "Bank figures", "Test", "Reference",
         "Run configuration and reference regulatory parameters"}
KEYS = {"role", "unlocked_cells", "unlocked_columns", "allow", "notes"}
CONFIG = "src/core/CORE_Config.bas"
# What a protected sheet may allow: the sheetProtection flag the template saves ("0" = allowed) and the
# CORE_Config list M_Main.ProtectWorkbook applies. Outlining is not saved with the file.
ALLOW = {"sorting": ("sort", "PROTECT_SORTING"), "filtering": ("autoFilter", "PROTECT_FILTERING"),
         "inserting_rows": ("insertRows", "PROTECT_INSERTING_ROWS"), "outlining": (None, "PROTECT_OUTLINING")}
# Saved protection: protected sheet, objects and scenarios; every other operation stays forbidden.
LOCKED_FLAGS = ("formatCells", "formatColumns", "formatRows", "insertColumns", "insertHyperlinks",
                "deleteColumns", "deleteRows", "pivotTables")
PASSWORD_ATTRIBUTES = ("password", "algorithmName", "hashValue", "saltValue", "spinCount",
                       "workbookPassword", "workbookAlgorithmName", "workbookHashValue")
MAX_COLUMN = 16384
INPUT_FILL = "FFFFF2CC"
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "rel": "http://schemas.openxmlformats.org/package/2006/relationships"}
CELL = re.compile(r"^([A-Z]{1,3})([1-9][0-9]*)$")
COLUMNS = re.compile(r"^([A-Z]{1,3}):([A-Z]{1,3})$")


def column_number(letters: str) -> int:
    number = 0
    for letter in letters:
        number = number * 26 + ord(letter) - 64
    return number


def column_name(number: int) -> str:
    name = ""
    while number:
        number, rest = divmod(number - 1, 26)
        name = chr(65 + rest) + name
    return name


def cell_address(text: str) -> tuple[int, int]:
    match = CELL.match(text)
    if not match:
        raise ValueError(f"not a cell address: {text!r}")
    return column_number(match.group(1)), int(match.group(2))


def range_cells(text: str) -> set[tuple[int, int]]:
    first, _, last = text.partition(":")
    (c1, r1), (c2, r2) = cell_address(first), cell_address(last or first)
    if c2 < c1 or r2 < r1:
        raise ValueError(f"reversed range: {text!r}")
    return {(c, r) for c in range(c1, c2 + 1) for r in range(r1, r2 + 1)}


def check_matrix(matrix: dict[str, Any]) -> list[str]:
    """Shape of the matrix itself."""
    findings = []
    if matrix.get("schema_version") != 1:
        findings.append(f"{MATRIX}: schema_version must be 1")
    sheets = matrix.get("sheets")
    if not isinstance(sheets, dict) or not sheets:
        return findings + [f"{MATRIX}: sheets must be a non-empty object"]
    for name, spec in sheets.items():
        if not isinstance(spec, dict):
            findings.append(f"{MATRIX}: {name}: must be an object")
            continue
        for key in sorted(set(spec) - KEYS):
            findings.append(f"{MATRIX}: {name}: unknown key {key!r}")
        if spec.get("role") not in ROLES:
            findings.append(f"{MATRIX}: {name}: unknown role {spec.get('role')!r}")
        for text in spec.get("unlocked_cells", []):
            try:
                range_cells(text)
            except ValueError as error:
                findings.append(f"{MATRIX}: {name}: {error}")
        for item in spec.get("allow", []):
            if item not in ALLOW:
                findings.append(f"{MATRIX}: {name}: unknown allowed operation {item!r}")
        rule = spec.get("unlocked_columns")
        if rule is not None:
            match = COLUMNS.match(str(rule.get("columns", "")))
            first = rule.get("first_row")
            if not match or column_number(match.group(2)) < column_number(match.group(1)):
                findings.append(f"{MATRIX}: {name}: unlocked_columns.columns must be a column range such as C:S")
            if not isinstance(first, int) or isinstance(first, bool) or first < 2:
                findings.append(f"{MATRIX}: {name}: unlocked_columns.first_row must be an integer of at least 2")
    return findings


def expected_unlocked(spec: dict[str, Any]) -> tuple[set[tuple[int, int]], set[int], int]:
    cells: set[tuple[int, int]] = set()
    for text in spec.get("unlocked_cells", []):
        cells |= range_cells(text)
    rule = spec.get("unlocked_columns")
    if not rule:
        return cells, set(), 0
    first, _, last = rule["columns"].partition(":")
    return cells, set(range(column_number(first), column_number(last) + 1)), rule["first_row"]


class Style(NamedTuple):
    locked: bool
    input_fill: bool


def child(element: ET.Element, path: str) -> ET.Element:
    found = element.find(path, NS)
    if found is None:
        raise ValueError(f"missing element {path}")
    return found


def attribute(element: ET.Element, name: str) -> str:
    value = element.get(name)
    if value is None:
        raise ValueError(f"missing attribute {name} on {element.tag}")
    return value


def cell_styles(archive: zipfile.ZipFile) -> list[Style]:
    """Lock flag (Excel's default is locked) and input fill of each cell format, in index order."""
    root = ET.fromstring(archive.read("xl/styles.xml"))
    fills = [any(color.get("rgb", "").upper() == INPUT_FILL for color in fill.iter(f"{{{NS['m']}}}fgColor"))
             for fill in child(root, "m:fills").findall("m:fill", NS)]
    styles = []
    for xf in child(root, "m:cellXfs").findall("m:xf", NS):
        protection = xf.find("m:protection", NS)
        locked = protection is None or protection.get("locked", "1") not in ("0", "false")
        styles.append(Style(locked, fills[int(xf.get("fillId", "0"))]))
    return styles


def sheet_parts(archive: zipfile.ZipFile) -> dict[str, str]:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {attribute(rel, "Id"): attribute(rel, "Target") for rel in rels.findall("rel:Relationship", NS)}
    parts = {}
    for sheet in child(workbook, "m:sheets").findall("m:sheet", NS):
        target = targets[attribute(sheet, f"{{{NS['r']}}}id")]
        parts[attribute(sheet, "name")] = target.lstrip("/") if target.startswith("/") else "xl/" + target
    return parts


class Sheet(NamedTuple):
    column_style: list[int]
    rows: dict[int, tuple[int | None, dict[int, int]]]

    def style(self, column: int, row: int) -> int:
        """Effective format: the cell's own, else a formatted row's, else the column's."""
        row_style, stored = self.rows.get(row, (None, {}))
        if column in stored:
            return stored[column]
        return self.column_style[column] if row_style is None else row_style


def read_sheet(xml: bytes) -> Sheet:
    root = ET.fromstring(xml)
    column_style = [0] * (MAX_COLUMN + 1)
    cols = root.find("m:cols", NS)
    for col in cols.findall("m:col", NS) if cols is not None else []:
        for number in range(int(attribute(col, "min")), int(attribute(col, "max")) + 1):
            column_style[number] = int(col.get("style", "0"))
    rows: dict[int, tuple[int | None, dict[int, int]]] = {}
    for row in child(root, "m:sheetData").findall("m:row", NS):
        row_style = int(row.get("s", "0")) if row.get("customFormat") in ("1", "true") else None
        stored = {cell_address(attribute(cell, "r"))[0]: int(cell.get("s", "0"))
                  for cell in row.findall("m:c", NS)}
        rows[int(attribute(row, "r"))] = (row_style, stored)
    return Sheet(column_style, rows)


def listed(name: str, problem: str, addresses: list[str]) -> list[str]:
    if not addresses:
        return []
    more = f" and {len(addresses) - 8} more" if len(addresses) > 8 else ""
    return [f"{name}: {problem} at {', '.join(addresses[:8])}{more}"]


def check_sheet(name: str, xml: bytes, styles: list[Style], spec: dict[str, Any]) -> list[str]:
    sheet = read_sheet(xml)
    cells, columns, first_row = expected_unlocked(spec)

    def style(index: int) -> Style:
        if index >= len(styles):
            raise ValueError(f"{name}: style {index} is not defined")
        return styles[index]

    findings = [f"{name}: row {row} has an unlocked row format; unlock cells or columns instead"
                for row, (row_style, _) in sorted(sheet.rows.items())
                if row_style is not None and not style(row_style).locked]
    for column in range(1, MAX_COLUMN + 1):
        if style(sheet.column_style[column]).locked == (column in columns):
            state = "locked" if column in columns else "unlocked"
            findings.append(f"{name}: column {column_name(column)} is {state} below the stored rows")
            break
    stored_columns = [column for _, stored in sheet.rows.values() for column in stored]
    last_column = max([column for column, _ in cells] + list(columns) + stored_columns + [1])
    last_row = max([row for _, row in cells] + list(sheet.rows) + [first_row, 1])
    wrong, yellow = [], []
    for row in range(1, last_row + 1):
        for column in range(1, last_column + 1):
            effective = style(sheet.style(column, row))
            unlocked = (column, row) in cells or (column in columns and row >= first_row)
            if effective.locked == unlocked:
                wrong.append(f"{column_name(column)}{row}")
            elif effective.locked and effective.input_fill:
                yellow.append(f"{column_name(column)}{row}")
    findings += listed(name, "lock flag differs from the matrix", wrong)
    return findings + listed(name, "a locked cell has the input fill", yellow)


def vba_list(root: Path, constant: str) -> set[str]:
    """Sheet names of a CORE_Config list constant ("|A|B|"), over continuation lines."""
    text = (root / CONFIG).read_text(encoding="cp1252")
    match = re.search(rf"Public Const {constant}\s+As String\s*=\s*(.*?)\r?\n(?!\s*\")", text, re.S)
    if not match:
        raise ValueError(f"{CONFIG}: {constant} not found")
    value = "".join(re.findall(r'"([^"]*)"', match.group(1)))
    return {name for name in value.split("|") if name}


def check_config(root: Path, sheets: dict[str, Any]) -> list[str]:
    findings = []
    for item, (_, constant) in ALLOW.items():
        expected = {name for name, spec in sheets.items() if item in spec.get("allow", [])}
        actual = vba_list(root, constant)
        if actual != expected:
            findings.append(f"{CONFIG}: {constant} lists {sorted(actual)}, the matrix {sorted(expected)}")
    return findings


def check_saved_protection(name: str, xml: bytes, spec: dict[str, Any]) -> list[str]:
    """The sheet is saved protected, without a password, allowing exactly what the matrix allows."""
    protection = ET.fromstring(xml).find("m:sheetProtection", NS)
    if protection is None:
        return [f"{name}: the sheet is not saved protected"]
    findings = [f"{name}: the protection has a password" for key in PASSWORD_ATTRIBUTES if protection.get(key)][:1]
    for key in ("sheet", "objects", "scenarios"):
        if protection.get(key) not in ("1", "true"):
            findings.append(f"{name}: the protection must cover {key}")
    allowed = spec.get("allow", [])
    for item, (flag, _) in ALLOW.items():
        if flag is None:
            continue
        saved = protection.get(flag) in ("0", "false")
        if saved != (item in allowed):
            findings.append(f"{name}: the saved protection {'allows' if saved else 'forbids'} {item}, "
                            "against the matrix")
    findings += [f"{name}: the protection allows {flag}" for flag in LOCKED_FLAGS
                 if protection.get(flag) in ("0", "false")]
    if "filtering" in allowed and ET.fromstring(xml).find("m:autoFilter", NS) is None:
        findings.append(f"{name}: filtering is allowed but the sheet has no AutoFilter, "
                        "and a protected sheet cannot get one")
    return findings


def check_structure(archive: zipfile.ZipFile) -> list[str]:
    protection = ET.fromstring(archive.read("xl/workbook.xml")).find("m:workbookProtection", NS)
    if protection is None or protection.get("lockStructure") not in ("1", "true"):
        return [f"{TEMPLATE}: the workbook structure is not saved protected"]
    if any(protection.get(key) for key in PASSWORD_ATTRIBUTES):
        return [f"{TEMPLATE}: the workbook structure protection has a password"]
    return []


def run_check(root: Path) -> dict[str, Any]:
    matrix = json.loads((root / MATRIX).read_text(encoding="utf-8"))
    findings = check_matrix(matrix)
    with zipfile.ZipFile(root / TEMPLATE) as archive:
        parts = sheet_parts(archive)
        styles = cell_styles(archive)
        if not styles or not styles[0].locked:
            findings.append(f"{TEMPLATE}: the default cell format must be locked")
        sheets = matrix.get("sheets", {}) if isinstance(matrix.get("sheets"), dict) else {}
        for name in sorted(set(parts) - set(sheets)):
            findings.append(f"{MATRIX}: template sheet {name} is missing from the matrix")
        for name in sorted(set(sheets) - set(parts)):
            findings.append(f"{MATRIX}: {name} is not a sheet of the template")
        if not findings:
            findings += check_structure(archive) + check_config(root, sheets)
            for name, part in parts.items():
                xml = archive.read(part)
                findings += check_saved_protection(name, xml, sheets[name])
                findings += check_sheet(name, xml, styles, sheets[name])
    return {"schema_version": 1, "status": "fail" if findings else "pass",
            "sheets": len(matrix.get("sheets", {})), "findings": findings}


def markdown_report(report: dict[str, Any]) -> str:
    lines = [f"Workbook protection: {report['status'].upper()} ({report['sheets']} sheet(s) in the matrix)"]
    lines.extend(report["findings"])
    return "\n".join(lines) + "\n"


def main() -> int:
    options = parse_report_args(sys.argv[1:], description=__doc__)
    return run_gate(options, build=lambda: run_check(options.root.resolve()),
                    markdown=markdown_report,
                    errors=(OSError, ValueError, KeyError, json.JSONDecodeError, zipfile.BadZipFile, ET.ParseError))


if __name__ == "__main__":
    raise SystemExit(main())
