#!/usr/bin/env python3
"""Reference implementation of the K-SACCR CSV input contract (docs/csv/contract.json).

It parses a Trades or NettingSets CSV file exactly as the contract specifies,
reports every finding with its stable code, physical line and record number,
and writes the blank templates. The workbook importer (#207) follows the same
contract; this module is what the contract tests and the gate run.
"""
from __future__ import annotations

import argparse
import json
import math
import re
import sys
import unicodedata
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any

CONTRACT = "docs/csv/contract.json"
TEMPLATES = "docs/csv/templates"
DOCUMENT = "docs/csv/CSV_CONTRACT.md"
BOM = b"\xef\xbb\xbf"
DELIMITER = ";"
NUMBER = re.compile(r"^-?(?:0|[1-9][0-9]*)(?:\.[0-9]+)?(?:[eE][+-]?[0-9]+)?$")
INTEGER = re.compile(r"^-?(?:0|[1-9][0-9]*)$")
DATE = re.compile(r"^[0-9]{4}-[0-9]{2}-[0-9]{2}$")
FORMULA_START = ("=", "+", "-", "@", "\t", "\r")
# The magnitudes a worksheet cell can hold: larger values overflow, smaller non-zero ones underflow.
EXCEL_LARGEST = 9.99999999999999e307
EXCEL_SMALLEST = 2.2250738585072014e-308


@dataclass(frozen=True)
class Record:
    """One logical CSV record: its fields and the physical line it starts on."""
    fields: list[str]
    line: int


@dataclass(frozen=True)
class Finding:
    """One contract violation; record 0 is the header, column is "" when none applies."""
    code: str
    line: int
    record: int
    column: str
    message: str

    def as_dict(self) -> dict[str, Any]:
        return {"code": self.code, "line": self.line, "record": self.record, "column": self.column,
                "message": self.message}


def load_contract(root: Path) -> dict[str, Any]:
    contract: dict[str, Any] = json.loads((root / CONTRACT).read_text(encoding="utf-8"))
    return contract


def decode(data: bytes) -> tuple[str | None, list[Finding]]:
    """The file text without its BOM, or None with a CSV-ENCODING finding."""
    if data.startswith(BOM):
        data = data[len(BOM):]
    try:
        text = data.decode("utf-8")
    except UnicodeDecodeError as error:
        return None, [Finding("CSV-ENCODING", 0, 0, "", f"not valid UTF-8 at byte {error.start}")]
    if "\x00" in text:
        return None, [Finding("CSV-ENCODING", text.count("\n", 0, text.index("\x00")) + 1, 0, "",
                              "contains a NUL character")]
    return text, []


class _Parser:
    """RFC 4180 records with ";" between fields and CRLF or LF separators; quotes only around whole fields."""

    def __init__(self, text: str) -> None:
        self.text = text
        self.pos = 0
        self.line = 1

    def end_of_line(self) -> int:
        """Length of the record separator at the current position: 2, 1, or 0 for none."""
        if self.text.startswith("\r\n", self.pos):
            return 2
        return 1 if self.text.startswith("\n", self.pos) else 0

    def quoted(self) -> str:
        start_line = self.line
        self.pos += 1
        parts: list[str] = []
        while True:
            close = self.text.find('"', self.pos)
            if close < 0:
                raise SyntaxError(f"quoted field opened on line {start_line} is not closed")
            parts.append(self.text[self.pos:close])
            self.line += self.text.count("\n", self.pos, close)
            if self.text.startswith('""', close):
                parts.append('"')
                self.pos = close + 2
                continue
            self.pos = close + 1
            following = self.text[self.pos:self.pos + 1]
            if following not in ("", DELIMITER) and not self.end_of_line():
                raise SyntaxError("text after a closing quote")
            return "".join(parts)

    def plain(self) -> str:
        start = self.pos
        while self.pos < len(self.text) and self.text[self.pos] not in DELIMITER + "\n" \
                and not self.text.startswith("\r\n", self.pos):
            char = self.text[self.pos]
            if char == '"':
                raise SyntaxError("quote inside an unquoted field")
            if char == "\r":
                raise SyntaxError("carriage return that does not end a line")
            self.pos += 1
        return self.text[start:self.pos]

    def record(self) -> list[str]:
        fields: list[str] = []
        while True:
            fields.append(self.quoted() if self.text.startswith('"', self.pos) else self.plain())
            if self.text.startswith(DELIMITER, self.pos):
                self.pos += 1
                continue
            separator = self.end_of_line()
            self.pos += separator
            if separator:
                self.line += 1
            return fields


def parse(text: str) -> tuple[list[Record], list[Finding]]:
    """All records of text; parsing stops at the first syntax error, which is reported."""
    parser = _Parser(text)
    records: list[Record] = []
    while parser.pos < len(text):
        start_line = parser.line
        try:
            records.append(Record(parser.record(), start_line))
        except SyntaxError as error:
            return records, [Finding("CSV-SYNTAX", start_line, len(records), "", str(error.msg))]
    return records, []


def check_header(header: Record, names: list[str]) -> list[Finding]:
    findings: list[Finding] = []
    seen: set[str] = set()
    for name in header.fields:
        if name in seen:
            findings.append(Finding("CSV-HEADER-DUPLICATE", header.line, 0, name, f"column {name!r} repeats"))
        elif name not in names:
            findings.append(Finding("CSV-HEADER-UNKNOWN", header.line, 0, name, f"column {name!r} is not in the contract"))
        seen.add(name)
    findings += [Finding("CSV-HEADER-MISSING", header.line, 0, name, f"column {name!r} is missing")
                 for name in names if name not in seen]
    return findings


def significant_digits(text: str) -> int:
    """Digits of a well-formed number that a double must hold; leading and trailing zeros add none."""
    mantissa = re.split(r"[eE]", text.lstrip("-"))[0]
    return len(mantissa.replace(".", "").strip("0"))


def check_text(value: str, column: dict[str, Any], types: dict[str, Any]) -> tuple[str, str] | None:
    kind = column["type"]
    limit = types.get(kind, {}).get("max_length")
    if kind in ("id", "code", "flag", "currency", "integer", "number", "date") and value != value.strip():
        return "CSV-TYPE", "leading or trailing spaces"
    if value.startswith(FORMULA_START) and kind not in ("number", "integer"):
        return "CSV-FORMULA", "text starts with a formula character"
    if limit is not None and len(value) > limit:
        return "CSV-LENGTH", f"{len(value)} characters, at most {limit}"
    if kind == "id" and any(char in value for char in types["id"]["forbidden"]):
        return "CSV-TYPE", "identifier contains | or #"
    line_breaks = "\n\r" if kind == "text" else ""
    if any(unicodedata.category(char) == "Cc" and char not in line_breaks for char in value):
        return "CSV-TYPE", "control character"
    return None


def check_range(number: float, column: dict[str, Any]) -> tuple[str, str] | None:
    if "min" in column and number < column["min"]:
        return "CSV-RANGE", f"below {column['min']}"
    if "min_exclusive" in column and number <= column["min_exclusive"]:
        return "CSV-RANGE", f"not above {column['min_exclusive']}"
    if "max" in column and number > column["max"]:
        return "CSV-RANGE", f"above {column['max']}"
    return None


def check_value(value: str, column: dict[str, Any], types: dict[str, Any]) -> tuple[str, str] | None:
    """The finding of one non-blank field, as (code, message), or None."""
    kind = column["type"]
    problem = check_text(value, column, types)
    if problem:
        return problem
    if kind in ("code", "flag"):
        allowed = column.get("allowed", types["flag"]["allowed"] if kind == "flag" else [])
        return None if value in allowed else ("CSV-ALLOWED", f"{value!r} is not one of {', '.join(allowed)}")
    if kind == "currency":
        return None if re.match(types["currency"]["pattern"], value) else ("CSV-TYPE", "not a 3-letter code")
    if kind in ("integer", "number"):
        if not (INTEGER if kind == "integer" else NUMBER).match(value):
            return "CSV-TYPE", f"{value!r} is not {'an integer' if kind == 'integer' else 'a number'}"
        if significant_digits(value) > 15:
            return "CSV-TYPE", "more than 15 significant digits"
        number = float(value)
        if not math.isfinite(number) or abs(number) > EXCEL_LARGEST or \
                (significant_digits(value) and abs(number) < EXCEL_SMALLEST):
            return "CSV-RANGE", "outside the magnitudes a worksheet cell can hold"
        return check_range(number, column)
    if kind == "date":
        return check_date(value, types["date"])
    return None


def check_date(value: str, limits: dict[str, str]) -> tuple[str, str] | None:
    if not DATE.match(value):
        return "CSV-TYPE", f"{value!r} is not a yyyy-mm-dd date"
    try:
        day = date.fromisoformat(value)
    except ValueError:
        return "CSV-TYPE", f"{value!r} is not a calendar date"
    if not date.fromisoformat(limits["min"]) <= day <= date.fromisoformat(limits["max"]):
        return "CSV-RANGE", f"outside {limits['min']} to {limits['max']}"
    return None


def check_row_rules(row: dict[str, str], rules: list[dict[str, Any]],
                    flawed: set[str]) -> list[tuple[str, str]]:
    """(column, message) for each rule the row breaks; a rule naming a flawed field is not judged."""
    problems: list[tuple[str, str]] = []
    for rule in rules:
        if flawed.intersection(rule_columns(rule)):
            continue
        if "when" in rule:
            active = all(row.get(name) == value for name, value in rule["when"].items())
            names = rule["required"] if active else rule["otherwise_blank"]
            for name in names:
                if active and row.get(name, "") == "":
                    problems.append((name, f"{rule['id']}: required when {rule['when']}"))
                elif not active and row.get(name, "") != "":
                    problems.append((name, f"{rule['id']}: must be blank unless {rule['when']}"))
        for key in ("after", "above"):
            if key in rule:
                later, earlier = rule[key]
                if row.get(later) and row.get(earlier) and _ordered(row[later], row[earlier], key) is False:
                    problems.append((later, f"{rule['id']}: {later} must be {key} {earlier}"))
    return problems


def rule_columns(rule: dict[str, Any]) -> set[str]:
    """Every column a row rule reads: its condition, its targets and its ordered pair."""
    return set(rule.get("when", {})) | set(rule.get("required", [])) | set(rule.get("otherwise_blank", [])) \
        | set(rule.get("after", [])) | set(rule.get("above", []))


def _ordered(later: str, earlier: str, key: str) -> bool | None:
    """Whether later > earlier, or None when either field is not of its type."""
    try:
        if key == "after":
            return date.fromisoformat(later) > date.fromisoformat(earlier)
        return float(later) > float(earlier)
    except ValueError:
        return None


def check_record(record: Record, index: int, names: list[str], columns: dict[str, dict[str, Any]],
                 spec: dict[str, Any], types: dict[str, Any]) -> list[Finding]:
    row = dict(zip(names, record.fields))
    if all(value == "" for value in record.fields):
        return [Finding("CSV-BLANK-RECORD", record.line, index, "", "every field is blank")]
    findings: list[Finding] = []
    for name, value in row.items():
        column = columns[name]
        if value == "":
            if column["required"]:
                findings.append(Finding("CSV-REQUIRED", record.line, index, name, "required field is blank"))
            continue
        problem = check_value(value, column, types)
        if problem:
            findings.append(Finding(problem[0], record.line, index, name, problem[1]))
    flawed = {f.column for f in findings}
    findings += [Finding("CSV-ROW-RULE", record.line, index, name, message)
                 for name, message in check_row_rules(row, spec.get("row_rules", []), flawed)
                 if name not in flawed]
    return findings


def check_keys(records: list[tuple[int, Record, dict[str, str]]], spec: dict[str, Any]) -> list[Finding]:
    findings: list[Finding] = []
    first: dict[str, int] = {}
    reserved = {key.upper() for key in spec.get("reserved_keys", [])}
    for index, record, row in records:
        key = row.get(spec["key"], "")
        if not key:
            continue
        folded = key.upper()
        if folded in reserved:
            findings.append(Finding("CSV-RESERVED-KEY", record.line, index, spec["key"], f"{key!r} is reserved"))
        elif folded in first:
            findings.append(Finding("CSV-DUPLICATE-KEY", record.line, index, spec["key"],
                                    f"{key!r} repeats record {first[folded]}"))
        else:
            first[folded] = index
    return findings


def validate(data: bytes, kind: str, contract: dict[str, Any],
             max_records: int | None = None) -> tuple[list[dict[str, str]], list[Finding]]:
    """The records of a CSV file of kind (Trades or NettingSets) as dictionaries, and all findings.

    The records are returned only when there is no finding: a file is accepted or
    rejected as a whole, never in part.
    """
    spec = contract["files"][kind]
    names = [column["name"] for column in spec["columns"]]
    columns = {column["name"]: column for column in spec["columns"]}
    text, findings = decode(data)
    if text is None:
        return [], findings
    records, findings = parse(text)
    if not records:
        return [], findings or [Finding("CSV-EMPTY", 1, 0, "", "the file has no header record")]
    header, body = records[0], records[1:]
    # A header error stops reading: nothing found in the body is reported with it.
    header_findings = check_header(header, names)
    if header_findings:
        return [], header_findings
    limit = spec["max_records"] if max_records is None else max_records
    if len(body) > limit:
        findings.append(Finding("CSV-TOO-MANY-RECORDS", body[limit].line, limit + 1, "",
                                f"{len(body)} records, at most {limit}"))
        return [], findings
    rows: list[tuple[int, Record, dict[str, str]]] = []
    for index, record in enumerate(body, start=1):
        if record.fields == [""]:
            findings.append(Finding("CSV-BLANK-RECORD", record.line, index, "", "empty line"))
            continue
        if len(record.fields) != len(header.fields):
            findings.append(Finding("CSV-FIELD-COUNT", record.line, index, "",
                                    f"{len(record.fields)} fields, the header has {len(header.fields)}"))
            continue
        ordered = Record([dict(zip(header.fields, record.fields))[name] for name in names], record.line)
        findings += check_record(ordered, index, names, columns, spec, contract["types"])
        rows.append((index, record, dict(zip(names, ordered.fields))))
    findings += check_keys(rows, spec)
    findings.sort(key=lambda f: (f.record, f.line, f.column, f.code))
    return ([row for _, _, row in rows] if not findings else []), findings


def template_text(kind: str, contract: dict[str, Any]) -> str:
    """The blank template of kind: the header record only, ending in LF."""
    return DELIMITER.join(column["name"] for column in contract["files"][kind]["columns"]) + "\n"


def describe_type(column: dict[str, Any], types: dict[str, Any]) -> str:
    kind = str(column["type"])
    if kind == "code":
        return "code: " + ", ".join(f"`{value}`" for value in column["allowed"])
    if kind == "flag":
        return "flag: `Y`, `N`"
    limits = [f"{label} {column[key]}" for key, label in (("min", ">="), ("min_exclusive", ">"), ("max", "<="))
              if key in column]
    if kind in ("id", "text"):
        limits.append(f"<= {types[kind]['max_length']} characters")
    return kind + (" (" + ", ".join(limits) + ")" if limits else "")


def field_table(kind: str, contract: dict[str, Any]) -> str:
    """The Markdown table of the columns of kind, as CSV_CONTRACT.md shows it."""
    lines = ["| CSV column | Sheet header | Type | Required | Blank means | Description | Example |",
             "| --- | --- | --- | --- | --- | --- | --- |"]
    for column in contract["files"][kind]["columns"]:
        cells = [f"`{column['name']}`", column["sheet_header"].replace("|", "\\|"),
                 describe_type(column, contract["types"]), "yes" if column["required"] else "no",
                 column.get("blank", ""), column["description"],
                 f"`{column['example']}`" if column["example"] else ""]
        lines.append("| " + " | ".join(cells) + " |")
    return "\n".join(lines) + "\n"


def with_tables(text: str, contract: dict[str, Any]) -> str:
    """text with each generated field table replaced by the contract's current one."""
    for kind in contract["files"]:
        begin, end = f"<!-- BEGIN GENERATED: {kind} -->\n", f"<!-- END GENERATED: {kind} -->"
        start, stop = text.find(begin), text.find(end)
        if start < 0 or stop < start:
            raise ValueError(f"{DOCUMENT}: markers of the {kind} table not found")
        text = text[:start + len(begin)] + field_table(kind, contract) + text[stop:]
    return text


def template_findings(root: Path, contract: dict[str, Any]) -> list[str]:
    findings: list[str] = []
    for kind, spec in contract["files"].items():
        path = root / TEMPLATES / spec["template"]
        if not path.is_file() or path.read_bytes() != template_text(kind, contract).encode("ascii"):
            findings.append(f"{TEMPLATES}/{spec['template']} is not what the contract generates; "
                            "run python tools/csv_contract.py --write-templates")
    document = (root / DOCUMENT).read_text(encoding="utf-8")
    if with_tables(document, contract) != document:
        findings.append(f"{DOCUMENT}: field tables differ from the contract; "
                        "run python tools/csv_contract.py --write-templates")
    return findings


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--kind", choices=("Trades", "NettingSets"), help="Kind of the file to validate")
    parser.add_argument("file", nargs="?", type=Path, help="CSV file to validate")
    parser.add_argument("--write-templates", action="store_true", help="Write the blank templates and the field tables of CSV_CONTRACT.md")
    args = parser.parse_args()
    contract = load_contract(args.root)
    if args.write_templates:
        for kind, spec in contract["files"].items():
            path = args.root / TEMPLATES / spec["template"]
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_bytes(template_text(kind, contract).encode("ascii"))
            print(f"wrote {path.relative_to(args.root).as_posix()}")
        document = args.root / DOCUMENT
        document.write_text(with_tables(document.read_text(encoding="utf-8"), contract), encoding="utf-8")
        print(f"wrote the field tables of {DOCUMENT}")
        return 0
    if not args.kind or not args.file:
        parser.error("give --kind and a file, or --write-templates")
    rows, findings = validate(args.file.read_bytes(), args.kind, contract)
    for finding in findings:
        print(f"line {finding.line}, record {finding.record}, {finding.column or '-'}: "
              f"{finding.code}: {finding.message}")
    print(f"{'ACCEPTED' if not findings else 'REJECTED'}: {len(rows)} records, {len(findings)} findings")
    return 0 if not findings else 1


if __name__ == "__main__":
    sys.exit(main())
