#!/usr/bin/env python3
"""Check the CSV input contract (docs/csv/contract.json) against the engine's
expected headers, the template's headers and validation lists, the generated
templates, the VBA constants of the template export and the CSV fixtures,
without running Excel."""
from __future__ import annotations

import json
import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import Any

from _gatelib import parse_report_args, run_gate
from csv_contract import (
    DELIMITER,
    load_contract,
    rule_columns,
    template_findings,
    template_text,
    validate,
)

CONFIG = "src/core/CORE_Config.bas"
TEMPLATE = "src/workbook/SACCR_Template.xlsx"
FIXTURES = "tests/csv"
EXPECTED = "tests/csv/expected.json"
CASES = "tests/csv/cases.json"
ENGINE_HEADERS = {"NettingSets": "NS_HEADERS", "Trades": "TR_HEADERS"}
EXPORT_CONSTANTS = {"NettingSets": ("CSV_NS_FILE", "CSV_NS_HEADER"), "Trades": ("CSV_TR_FILE", "CSV_TR_HEADER")}
KEY_COLUMN = {"NettingSets": "NettingSetID", "Trades": "TradeID"}
# Staging sheet and its header anchor per file (#208); the header row starts with the CSV line column.
STAGING = {"NettingSets": ("Import_NettingSets", "ImportNettingSets_Header"),
           "Trades": ("Import_Trades", "ImportTrades_Header")}
STAGING_KEY = "CSV line"
TYPES = {"id", "text", "code", "flag", "currency", "integer", "number", "date"}
NAME = re.compile(r"^[a-z][a-z0-9]*(?:_[a-z0-9]+)*$")
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "rel": "http://schemas.openxmlformats.org/package/2006/relationships"}


def vba_string(root: Path, constant: str) -> str:
    """The value of a VBA string constant in CORE_Config, concatenated over continuation lines."""
    text = (root / CONFIG).read_text(encoding="cp1252")
    match = re.search(rf"Public Const {constant} As String = (.*?)\r?\n(?!\s*\")", text, re.S)
    if not match:
        raise ValueError(f"{CONFIG}: {constant} not found")
    return "".join(re.findall(r'"((?:[^"]|"")*)"', match.group(1))).replace('""', '"')


def engine_headers(root: Path, constant: str) -> list[str]:
    """The headers of a "|"-separated VBA string constant."""
    return vba_string(root, constant).split("|")


def header_key(text: str) -> str:
    """A header as the engine compares it: case, spaces, tabs and line breaks ignored (HeaderKey)."""
    return re.sub(r"\s+", "", text).upper()


def column_index(reference: str) -> int:
    letters = re.match(r"[A-Z]+", reference)
    if not letters:
        raise ValueError(f"not a cell reference: {reference}")
    index = 0
    for letter in letters.group(0):
        index = index * 26 + ord(letter) - 64
    return index


class Workbook:
    """The header rows and list validations of the template's sheets."""

    def __init__(self, path: Path) -> None:
        self.package = zipfile.ZipFile(path)
        shared = ET.fromstring(self.package.read("xl/sharedStrings.xml"))
        self.strings = ["".join(t.text or "" for t in item.iter(f"{{{NS['m']}}}t"))
                        for item in shared.findall("m:si", NS)]

    def sheet(self, name: str) -> ET.Element:
        book = ET.fromstring(self.package.read("xl/workbook.xml"))
        rels = ET.fromstring(self.package.read("xl/_rels/workbook.xml.rels"))
        sheets = book.find("m:sheets", NS)
        for sheet in sheets if sheets is not None else []:
            if sheet.get("name") == name:
                target = next(r.get("Target") for r in rels if r.get("Id") == sheet.get(f"{{{NS['r']}}}id"))
                return ET.fromstring(self.package.read("xl/" + str(target).lstrip("/").removeprefix("xl/")))
        raise ValueError(f"{TEMPLATE}: sheet {name} not found")

    def cell_text(self, cell: ET.Element) -> str:
        value = cell.find("m:v", NS)
        if cell.get("t") == "s" and value is not None and value.text is not None:
            return self.strings[int(value.text)]
        inline = cell.find("m:is", NS)
        if inline is not None:
            return "".join(t.text or "" for t in inline.iter(f"{{{NS['m']}}}t"))
        return value.text or "" if value is not None else ""

    def header(self, sheet: ET.Element, key: str) -> tuple[int, list[str]]:
        """The header row's first column and texts, found by its key header."""
        _, first, texts = self.header_row(sheet, key)
        return first, texts

    def header_row(self, sheet: ET.Element, key: str) -> tuple[int, int, list[str]]:
        """The header row's number, first column and texts, found by its key header."""
        for row in sheet.iter(f"{{{NS['m']}}}row"):
            cells = self.row_cells(row)
            first = next((i for i, text in sorted(cells.items()) if text == key), None)
            if first is not None:
                last = max(i for i, text in cells.items() if text)
                return int(str(row.get("r"))), first, [cells.get(i, "") for i in range(first, last + 1)]
        raise ValueError(f"{TEMPLATE}: header {key} not found")

    def row_cells(self, row: ET.Element) -> dict[int, str]:
        return {column_index(str(c.get("r"))): self.cell_text(c) for c in row.findall("m:c", NS)}

    def row_texts(self, sheet: ET.Element, number: int, first: int, last: int) -> list[str]:
        """The texts of one row from column first to column last."""
        for row in sheet.iter(f"{{{NS['m']}}}row"):
            if int(str(row.get("r"))) == number:
                cells = self.row_cells(row)
                return [cells.get(i, "") for i in range(first, last + 1)]
        return [""] * (last - first + 1)

    def defined_name(self, name: str) -> str | None:
        book = ET.fromstring(self.package.read("xl/workbook.xml"))
        for item in book.iter(f"{{{NS['m']}}}definedName"):
            if item.get("name") == name and item.get("localSheetId") is None:
                return item.text
        return None

    @staticmethod
    def lists(sheet: ET.Element) -> dict[int, set[str]]:
        """Allowed values of each column's list validation."""
        found: dict[int, set[str]] = {}
        for rule in sheet.iter(f"{{{NS['m']}}}dataValidation"):
            formula = rule.find("m:formula1", NS)
            if rule.get("type") == "list" and formula is not None and formula.text:
                values = set(formula.text.strip('"').split(","))
                for reference in str(rule.get("sqref")).split():
                    found[column_index(reference)] = values
        return found


def check_shape(contract: dict[str, Any]) -> list[str]:
    findings: list[str] = []
    if contract["format"]["delimiter"] != DELIMITER:
        findings.append(f"format: delimiter {contract['format']['delimiter']!r} is not the parser's {DELIMITER!r}")
    for kind, spec in contract["files"].items():
        names = [c["name"] for c in spec["columns"]]
        findings += [f"{kind}: column name {n!r} is not snake_case" for n in names if not NAME.match(n)]
        findings += [f"{kind}: column {n!r} repeats" for n in set(names) if names.count(n) > 1]
        for column in spec["columns"]:
            if column["type"] not in TYPES:
                findings.append(f"{kind}.{column['name']}: unknown type {column['type']!r}")
            if column["type"] == "code" and not column.get("allowed"):
                findings.append(f"{kind}.{column['name']}: a code column needs allowed values")
        if spec["key"] not in names:
            findings.append(f"{kind}: key {spec['key']!r} is not a column")
        for rule in spec.get("row_rules", []):
            findings += [f"{kind}: row rule {rule['id']} names unknown column {n!r}"
                         for n in sorted(rule_columns(rule)) if n not in names]
    return findings


def check_workbook(root: Path, contract: dict[str, Any]) -> list[str]:
    findings: list[str] = []
    workbook = Workbook(root / TEMPLATE)
    for kind, spec in contract["files"].items():
        headers = [header_key(c["sheet_header"]) for c in spec["columns"]]
        engine = [header_key(h) for h in engine_headers(root, ENGINE_HEADERS[kind])]
        read = [h for h in engine if h]
        if headers[:len(read)] != read or any(engine[len(read):]):
            findings.append(f"{kind}: sheet headers differ from {ENGINE_HEADERS[kind]} in {CONFIG}")
        sheet = workbook.sheet(spec["sheet"])
        first, template = workbook.header(sheet, KEY_COLUMN[kind])
        if [header_key(h) for h in template] != headers:
            findings.append(f"{kind}: sheet headers differ from the {spec['sheet']} header row of {TEMPLATE}")
        findings += check_staging(workbook, kind, spec)
        lists = workbook.lists(sheet)
        for offset, column in enumerate(spec["columns"]):
            if column["type"] in ("code", "flag"):
                allowed = set(column.get("allowed", contract["types"]["flag"]["allowed"]))
                if lists.get(first + offset) != allowed:
                    findings.append(f"{kind}.{column['name']}: allowed values {sorted(allowed)} differ from the "
                                    f"template's list {sorted(lists.get(first + offset, set()))}")
    return findings


def check_staging(workbook: Workbook, kind: str, spec: dict[str, Any]) -> list[str]:
    """The staging sheet: the CSV line column, then the sheet headers with the CSV names above them."""
    findings: list[str] = []
    name, anchor = STAGING[kind]
    sheet = workbook.sheet(name)
    number, first, texts = workbook.header_row(sheet, STAGING_KEY)
    headers = [header_key(c["sheet_header"]) for c in spec["columns"]]
    if [header_key(h) for h in texts[1:]] != headers:
        findings.append(f"{kind}: the {name} header row differs from the contract's sheet headers")
    names = workbook.row_texts(sheet, number - 1, first + 1, first + len(spec["columns"]))
    if names != [c["name"] for c in spec["columns"]]:
        findings.append(f"{kind}: the CSV names above the {name} header row differ from the contract's columns")
    expected = f"{name}!${column_letters(first)}${number}"
    if workbook.defined_name(anchor) != expected:
        findings.append(f"{kind}: {anchor} is {workbook.defined_name(anchor)!r}, not {expected}")
    return findings


def column_letters(index: int) -> str:
    letters = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def check_export(root: Path, contract: dict[str, Any]) -> list[str]:
    """The file name and header line that M_Main writes for each blank template."""
    findings: list[str] = []
    for kind, spec in contract["files"].items():
        file_constant, header_constant = EXPORT_CONSTANTS[kind]
        if vba_string(root, file_constant) != spec["template"]:
            findings.append(f"{kind}: {file_constant} in {CONFIG} differs from the template name {spec['template']}")
        if vba_string(root, header_constant) + "\n" != template_text(kind, contract):
            findings.append(f"{kind}: {header_constant} in {CONFIG} differs from the contract's header line")
    return findings


def check_fixtures(root: Path, contract: dict[str, Any]) -> tuple[list[str], int]:
    expected = json.loads((root / EXPECTED).read_text(encoding="utf-8"))["fixtures"]
    files = {p.relative_to(root / FIXTURES).as_posix() for p in (root / FIXTURES).glob("*/*.csv")}
    findings = [f"{FIXTURES}/{name}: no entry in {EXPECTED}" for name in sorted(files - set(expected))]
    findings += [f"{EXPECTED}: {name} does not exist" for name in sorted(set(expected) - files)]
    for name in sorted(files & set(expected)):
        entry = expected[name]
        rows, found = validate((root / FIXTURES / name).read_bytes(), entry["kind"], contract)
        actual = [[f.code, f.line, f.record, f.column] for f in found]
        if actual != entry["findings"] or len(rows) != entry["records"]:
            findings.append(f"{FIXTURES}/{name}: findings or record count differ from {EXPECTED}")
        if name.startswith("valid/") and found:
            findings.append(f"{FIXTURES}/{name}: a valid fixture has findings")
        if name.startswith("invalid/") and not found:
            findings.append(f"{FIXTURES}/{name}: an invalid fixture has no finding")
    return findings, len(files)


def case_bytes(case: dict[str, Any]) -> bytes:
    """The content of an inline case: its hex bytes, or its text as UTF-8."""
    return bytes.fromhex(case["hex"]) if "hex" in case else str(case["text"]).encode("utf-8")


def check_cases(root: Path, contract: dict[str, Any]) -> tuple[list[str], int]:
    cases = json.loads((root / CASES).read_text(encoding="utf-8"))["cases"]
    findings: list[str] = []
    names = [case["name"] for case in cases]
    findings += [f"{CASES}: case {name!r} repeats" for name in sorted(set(names)) if names.count(name) > 1]
    for case in cases:
        rows, found = validate(case_bytes(case), case["kind"], contract,
                               max_records=case.get("max_records"), max_bytes=case.get("max_bytes"))
        actual = [[f.code, f.line, f.record, f.column] for f in found]
        if actual != case["findings"] or len(rows) != case["records"]:
            findings.append(f"{CASES}: case {case['name']!r}: findings or record count differ")
    return findings, len(cases)


def run_check(root: Path) -> dict[str, Any]:
    contract = load_contract(root)
    findings = check_shape(contract)
    findings += check_workbook(root, contract)
    findings += template_findings(root, contract)
    findings += check_export(root, contract)
    fixture_findings, fixtures = check_fixtures(root, contract)
    findings += fixture_findings
    case_findings, cases = check_cases(root, contract)
    findings += case_findings
    columns = {kind: len(spec["columns"]) for kind, spec in contract["files"].items()}
    return {"schema_version": 1, "status": "fail" if findings else "pass", "columns": columns,
            "fixtures": fixtures, "cases": cases, "findings": findings}


def markdown_report(report: dict[str, Any]) -> str:
    columns = ", ".join(f"{kind} {count}" for kind, count in report["columns"].items())
    lines = [f"CSV contract: {report['status'].upper()} (columns: {columns}; {report['fixtures']} fixture(s), {report['cases']} case(s))"]
    lines.extend(report["findings"])
    return "\n".join(lines) + "\n"


def main() -> int:
    options = parse_report_args(sys.argv[1:], description=__doc__)
    return run_gate(options, build=lambda: run_check(options.root.resolve()),
                    markdown=markdown_report, errors=(OSError, ValueError, KeyError, zipfile.BadZipFile))


if __name__ == "__main__":
    raise SystemExit(main())
