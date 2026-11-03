#!/usr/bin/env python3
"""Check the cell lock flags of the workbook template against its editability
matrix (docs/workbook/protection.json), without running Excel.

A cell is unlocked when the matrix lists it, or when its column is an
unlocked input column and its row is at or below the first data row. Every
other cell is locked. The effective style of a cell is its own style, else
the style of a formatted row, else the style of its column; the check
covers the cells the sheet stores, the rows it formats and the columns to
the last column of the sheet, so a cell that does not exist yet is judged
by the style it would receive."""
from __future__ import annotations

import json
import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import Any

from _gatelib import parse_report_args, run_gate

MATRIX = "docs/workbook/protection.json"
TEMPLATE = "src/workbook/SACCR_Template.xlsx"
ROLES = {"Instructions", "Input", "Staging", "Output", "Bank figures", "Test", "Reference",
         "Run configuration and reference regulatory parameters"}
KEYS = {"role", "unlocked_cells", "unlocked_columns", "notes"}
MAX_COLUMN = 16384
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


def locked_styles(archive: zipfile.ZipFile) -> list[bool]:
    """Lock flag of each cell format (cellXfs), in index order; Excel's default is locked."""
    root = ET.fromstring(archive.read("xl/styles.xml"))
    flags = []
    for xf in root.find("m:cellXfs", NS).findall("m:xf", NS):
        protection = xf.find("m:protection", NS)
        flags.append(protection is None or protection.get("locked", "1") not in ("0", "false"))
    return flags


def sheet_parts(archive: zipfile.ZipFile) -> dict[str, str]:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {rel.get("Id"): rel.get("Target") for rel in rels.findall("rel:Relationship", NS)}
    parts = {}
    for sheet in workbook.find("m:sheets", NS).findall("m:sheet", NS):
        target = targets[sheet.get(f"{{{NS['r']}}}id")]
        parts[sheet.get("name")] = target.lstrip("/") if target.startswith("/") else "xl/" + target
    return parts


def check_sheet(name: str, xml: bytes, locked: list[bool], spec: dict[str, Any]) -> list[str]:
    root = ET.fromstring(xml)

    def is_locked(style: int) -> bool:
        if style >= len(locked):
            raise ValueError(f"{name}: style {style} is not defined")
        return locked[style]

    column_style = [0] * (MAX_COLUMN + 1)
    cols = root.find("m:cols", NS)
    for col in cols.findall("m:col", NS) if cols is not None else []:
        for number in range(int(col.get("min")), int(col.get("max")) + 1):
            column_style[number] = int(col.get("style", "0"))
    rows: dict[int, tuple[int | None, dict[int, int]]] = {}
    for row in root.find("m:sheetData", NS).findall("m:row", NS):
        row_style = int(row.get("s", "0")) if row.get("customFormat") in ("1", "true") else None
        cells = {cell_address(cell.get("r"))[0]: int(cell.get("s", "0")) for cell in row.findall("m:c", NS)}
        rows[int(row.get("r"))] = (row_style, cells)
    cells, columns, first_row = expected_unlocked(spec)

    def should_unlock(column: int, row: int) -> bool:
        return (column, row) in cells or (column in columns and row >= first_row)

    findings: list[str] = []
    wrong: list[str] = []
    last_column = max([c for c, _ in cells] + list(columns)
                      + [c for _, (_, stored) in rows.items() for c in stored] + [1])
    last_row = max([r for _, r in cells] + list(rows) + [first_row, 1])
    for row in range(1, last_row + 1):
        row_style, stored = rows.get(row, (None, {}))
        for column in range(1, last_column + 1):
            if column in stored:
                style = stored[column]
            elif row_style is not None:
                style = row_style
            else:
                style = column_style[column]
            if is_locked(style) == should_unlock(column, row):
                wrong.append(f"{column_name(column)}{row}")
        if row_style is not None and not is_locked(row_style):
            findings.append(f"{name}: row {row} has an unlocked row format; unlock cells or columns instead")
    for column in range(1, MAX_COLUMN + 1):
        if is_locked(column_style[column]) == (column in columns):
            state = "locked" if column in columns else "unlocked"
            findings.append(f"{name}: column {column_name(column)} is {state} below the stored rows")
            break
    if wrong:
        state = ", ".join(wrong[:8]) + (f" and {len(wrong) - 8} more" if len(wrong) > 8 else "")
        findings.append(f"{name}: lock flag differs from the matrix at {state}")
    return findings


def run_check(root: Path) -> dict[str, Any]:
    matrix = json.loads((root / MATRIX).read_text(encoding="utf-8"))
    findings = check_matrix(matrix)
    with zipfile.ZipFile(root / TEMPLATE) as archive:
        parts = sheet_parts(archive)
        locked = locked_styles(archive)
        if not locked or not locked[0]:
            findings.append(f"{TEMPLATE}: the default cell format must be locked")
        sheets = matrix.get("sheets", {}) if isinstance(matrix.get("sheets"), dict) else {}
        for name in sorted(set(parts) - set(sheets)):
            findings.append(f"{MATRIX}: template sheet {name} is missing from the matrix")
        for name in sorted(set(sheets) - set(parts)):
            findings.append(f"{MATRIX}: {name} is not a sheet of the template")
        if not findings:
            for name, part in parts.items():
                findings += check_sheet(name, archive.read(part), locked, sheets[name])
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
