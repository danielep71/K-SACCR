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
from typing import Any, Callable

CONTRACT = "docs/csv/contract.json"
TEMPLATES = "docs/csv/templates"
DOCUMENT = "docs/csv/CSV_CONTRACT.md"
BOM = b"\xef\xbb\xbf"
DELIMITER = ";"
OTHER_DELIMITERS = {",": "','", "\t": "a tab"}  # Delimiters of other CSV dialects
# Findings that leave the file unreadable as records: the rest are about values and can be
# corrected on a staging sheet (#208).
STRUCTURAL = frozenset({"CSV-ENCODING", "CSV-SYNTAX", "CSV-EMPTY", "CSV-DELIMITER", "CSV-HEADER-DUPLICATE",
                        "CSV-HEADER-MISSING", "CSV-HEADER-UNKNOWN", "CSV-FIELD-COUNT",
                        "CSV-TOO-MANY-RECORDS", "CSV-TOO-LARGE"})
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
    if len(header.fields) == 1:
        # One field that holds another dialect's delimiter: a file of that dialect, not one unknown column.
        for other, shown in OTHER_DELIMITERS.items():
            if other in header.fields[0]:
                return [Finding("CSV-DELIMITER", header.line, 0, "",
                                f"the header is one field that contains {shown}; the delimiter is ';'")]
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


def fold(key: str) -> str:
    """A key in upper case, one character at a time, as the engine's UCase$ compares IDs: a character
    whose upper case is longer, such as the German sharp s, is kept as it is."""
    return "".join(char.upper() if len(char.upper()) == 1 else char for char in key)


def check_keys(records: list[tuple[int, Record, dict[str, str]]], spec: dict[str, Any]) -> list[Finding]:
    findings: list[Finding] = []
    first: dict[str, int] = {}
    reserved = {fold(key) for key in spec.get("reserved_keys", [])}
    for index, record, row in records:
        key = row.get(spec["key"], "")
        if not key:
            continue
        folded = fold(key)
        if folded in reserved:
            findings.append(Finding("CSV-RESERVED-KEY", record.line, index, spec["key"], f"{key!r} is reserved"))
        elif folded in first:
            findings.append(Finding("CSV-DUPLICATE-KEY", record.line, index, spec["key"],
                                    f"{key!r} repeats record {first[folded]}"))
        else:
            first[folded] = index
    return findings


def readable(findings: list[Finding]) -> bool:
    """No finding is structural: every record was read and has the header's fields."""
    return not any(f.code in STRUCTURAL for f in findings)


def validate_records(data: bytes, kind: str, contract: dict[str, Any],
                     max_records: int | None = None,
                     max_bytes: int | None = None,
                     keep_records: bool = False) -> tuple[list[tuple[int, dict[str, str]]], list[Finding]]:
    """As validate, each record with the physical line it starts on."""
    spec = contract["files"][kind]
    names = [column["name"] for column in spec["columns"]]
    columns = {column["name"]: column for column in spec["columns"]}
    byte_limit = contract["format"]["max_bytes"] if max_bytes is None else max_bytes
    if len(data) > byte_limit:
        return [], [Finding("CSV-TOO-LARGE", 0, 0, "", f"{len(data)} bytes, at most {byte_limit}")]
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
        findings.sort(key=lambda f: (f.record, f.line, f.column, f.code))
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
        if any(ordered.fields):  # A blank record is reported, never kept (#467).
            rows.append((index, record, dict(zip(names, ordered.fields))))
    findings += check_keys(rows, spec)
    findings.sort(key=lambda f: (f.record, f.line, f.column, f.code))
    keep = not findings or (keep_records and readable(findings))
    return ([(record.line, row) for _, record, row in rows] if keep else []), findings


def validate(data: bytes, kind: str, contract: dict[str, Any],
             max_records: int | None = None,
             max_bytes: int | None = None,
             keep_records: bool = False) -> tuple[list[dict[str, str]], list[Finding]]:
    """The records of a CSV file of kind (Trades or NettingSets) as dictionaries, and all findings.

    The records are returned only when there is no finding: a file is accepted or
    rejected as a whole, never in part. With keep_records, a readable file's records
    are returned with its value findings, for staging (#208).
    """
    rows, findings = validate_records(data, kind, contract, max_records, max_bytes, keep_records)
    return [row for _, row in rows], findings


def check_staged(kind: str, rows: list[tuple[int, dict[str, str]]], contract: dict[str, Any],
                 max_records: int | None = None) -> list[Finding]:
    """The findings about staged records (#209): each field, the rules between fields and the keys, as a
    file's records are checked, for rows read back from a staging sheet.

    rows holds (CSV line, record) in staging order, the line 0 for a row added by hand, a column missing
    from a record being blank; a finding's record is the 1-based position in rows. More rows than the
    contract allows is one finding, and nothing else is checked.
    """
    spec = contract["files"][kind]
    names = [column["name"] for column in spec["columns"]]
    columns = {column["name"]: column for column in spec["columns"]}
    limit = spec["max_records"] if max_records is None else max_records
    if len(rows) > limit:
        return [Finding("CSV-TOO-MANY-RECORDS", rows[limit][0], limit + 1, "", f"{len(rows)} rows, at most {limit}")]
    findings: list[Finding] = []
    kept: list[tuple[int, Record, dict[str, str]]] = []
    for position, (line, row) in enumerate(rows, start=1):
        record = Record([row.get(name, "") for name in names], line)
        findings += check_record(record, position, names, columns, spec, contract["types"])
        if any(record.fields):
            kept.append((position, record, dict(zip(names, record.fields))))
    findings += check_keys(kept, spec)
    findings.sort(key=lambda f: (f.record, f.line, f.column, f.code))
    return findings


Adder = Callable[[str, int, str, str], None]


def is_nonzero_number(text: str) -> bool:
    """Whether text is a number with a non-zero digit before its exponent: the test the VBA reader can make
    without converting, so that both agree on a value too small for a double."""
    return bool(NUMBER.match(text)) and any(char in "123456789" for char in re.split("[eE]", text)[0])


def check_portfolio(trades: list[tuple[int, dict[str, str]]], netting_sets: list[str],
                    contract: dict[str, Any]) -> list[Finding]:
    """The findings between records (#209): each leg's netting set, and the legs of each economic trade,
    judged as CORE_Engine judges them (the unknown-netting-set check and CheckEconomicTrades).

    trades holds (CSV line, record) in staging order, the line 0 when the row has none; a finding's
    record is the 1-based position in trades. netting_sets holds the IDs the legs may refer to.
    """
    codes = {rule["id"]: rule["code"] for rule in contract["portfolio_rules"]}
    known = {fold(key) for key in netting_sets}
    findings: list[Finding] = []

    def add(rule: str, position: int, column: str, message: str) -> None:
        findings.append(Finding(codes[rule], trades[position - 1][0], position, column, message))

    groups: dict[str, list[int]] = {}
    for position, (_, row) in enumerate(trades, start=1):
        if not row["trade_id"]:
            continue
        if row["netting_set_id"] and fold(row["netting_set_id"]) not in known:
            add("netting-set-known", position, "netting_set_id",
                f"netting set {row['netting_set_id']!r} is not on NettingSets")
        economic = row["economic_trade_id"]
        groups.setdefault("E|" + fold(economic) if economic else "L|" + fold(row["trade_id"]), []).append(position)
    for key, legs in groups.items():
        rows = [trades[position - 1][1] for position in legs]
        if key.startswith("E|"):
            check_economic_trade(legs, rows, add)
            continue
        for position, row in zip(legs, rows):
            if fold(row["mtm_leg"]) == "N":
                add("mtm-leg", position, "mtm_leg",
                    "a trade without an Economic trade ID carries its own MtM, so MtM leg must be Y or blank")
    findings.sort(key=lambda f: (f.record, f.line, f.column, f.code))
    return findings


def check_economic_trade(legs: list[int], rows: list[dict[str, str]], add: Adder) -> None:
    """The rules between the legs of one economic trade, each leg compared with the first as the engine does."""
    name, first = rows[0]["economic_trade_id"], rows[0]
    settlement = fold(first["settlement_type"]) or "N"
    for position, row in zip(legs[1:], rows[1:]):
        if fold(row["netting_set_id"]) != fold(first["netting_set_id"]):
            add("same-netting-set", position, "netting_set_id",
                f"economic trade {name}: its legs are in different netting sets")
        if (fold(row["settlement_type"]) or "N") != settlement:
            add("same-settlement", position, "settlement_type",
                f"economic trade {name}: its legs have different settlement types")
        elif row["next_reset_date"] != first["next_reset_date"]:
            add("same-settlement", position, "next_reset_date",
                f"economic trade {name}: its legs have different next reset dates")
    mtm_legs = sum(fold(row["mtm_leg"]) == "Y" for row in rows)
    for position, row in zip(legs, rows):
        if mtm_legs != 1:
            add("mtm-leg", position, "mtm_leg",
                f"economic trade {name}: exactly one leg must be the MtM leg (MtM leg = Y), found {mtm_legs}")
        if fold(row["mtm_leg"]) != "Y" and is_nonzero_number(row["market_value_amount"]):
            add("mtm-value", position, "market_value_amount",
                f"economic trade {name}: leg {row['trade_id']} is not the MtM leg, so its MtM must be 0")
    classes: dict[str, list[int]] = {}
    for index, row in enumerate(rows):
        classes.setdefault(fold(row["asset_class"]), []).append(index)
    for asset_class, members in classes.items():
        if len(members) > 1 and any(fold(rows[i]["instrument"]) != "OPTION" for i in members):
            for i in members:
                add("asset-class", legs[i], "asset_class",
                    f"economic trade {name}: {len(members)} legs in asset class {asset_class} that are not all options")



def normalize_pair(text: str) -> tuple[str, str] | None:
    """The two currencies of an FX pair in alphabetical order, as CORE_Engine.NormalizePair reads it: the
    separators "/", "-", " " and "." removed, six letters left, two different currencies; None otherwise."""
    letters = text.upper()
    for separator in ("/", "-", " ", "."):
        letters = letters.replace(separator, "")
    if len(letters) != 6 or letters[:3] == letters[3:]:
        return None
    first, second = sorted((letters[:3], letters[3:]))
    return first, second


def factor_key(asset_class: str, sub_class: str) -> str:
    """The supervisory-factor key of a trade, as CORE_Engine.FactorKey builds it: IR, FX and OT have one
    factor each, the other classes one per sub-class, such as CR_AAA."""
    return asset_class if asset_class in ("IR", "FX", "OT") else f"{asset_class}_{sub_class}"


def check_context(trades: list[tuple[int, dict[str, str]]], netting_sets: dict[str, str],
                  context: dict[str, Any], contract: dict[str, Any], factor_only: bool = False) -> list[Finding]:
    """The findings that depend on Params (#209): the reporting date, the FX table, the supervisory-factor
    table and the regime, judged as CORE_Engine judges them in a run.

    trades holds (CSV line, record) in staging order, as for check_portfolio. netting_sets maps each netting
    set the legs may refer to onto its regime override, "" for none. context holds the reporting date
    (as_of, yyyy-mm-dd), the reporting_currency, the Params regime, the currencies of the FX table and the
    factors, each key mapped onto its regime, "" for both. A field that is blank or not of its type is not
    judged here: its own finding comes from check_staged. With factor_only, only the factor and regime
    rules are judged, for the live legs when only netting sets are staged.
    """
    codes = {rule["id"]: rule["code"] for rule in contract["context_rules"]}
    columns = {column["name"]: column for column in contract["files"]["Trades"]["columns"]}
    types = contract["types"]
    as_of = date.fromisoformat(context["as_of"])
    currencies = {fold(code) for code in context["currencies"]} | {fold(context["reporting_currency"])}
    factors = {fold(key): fold(regimes) for key, regimes in context["factors"].items()}
    regimes = {fold(key): fold(override) for key, override in netting_sets.items()}
    findings: list[Finding] = []

    def add(rule: str, position: int, column: str, message: str) -> None:
        findings.append(Finding(codes[rule], trades[position - 1][0], position, column, message))

    for position, (_, row) in enumerate(trades, start=1):
        if not row.get("trade_id", ""):
            continue

        def usable(name: str) -> str:
            """The field when it is filled and of its type, else ""."""
            value = row.get(name, "")
            return value if value and check_value(value, columns[name], types) is None else ""

        def has_flawed(name: str) -> bool:
            """Whether the field is filled but not of its type, so that check_staged reports it."""
            return bool(row.get(name, "")) and not usable(name)

        if not factor_only:
            check_dates(position, usable, as_of, context["as_of"], add)
            check_currencies(position, usable, currencies, context["reporting_currency"], add)
        check_factor(position, usable, has_flawed, factors, regimes, fold(context["regime"]), add)
    findings.sort(key=lambda f: (f.record, f.line, f.column, f.code))
    return findings


def check_dates(position: int, usable: Callable[[str], str], as_of: date, as_of_text: str, add: Adder) -> None:
    """The maturity, option expiry and next reset date against the reporting date."""
    matured = False
    for name in ("maturity_date", "end_date", "option_expiry_date"):
        if usable(name):
            if date.fromisoformat(usable(name)) <= as_of:
                matured = True
                add("matured", position, name, f"trade has matured (maturity <= reporting date): {name} "
                                               f"{usable(name)}, reporting date {as_of_text}")
            break
    expiry = usable("option_expiry_date")
    if not matured and fold(usable("instrument")) == "OPTION" and expiry and date.fromisoformat(expiry) <= as_of:
        add("expired", position, "option_expiry_date",
            f"option has expired: option_expiry_date {expiry}, reporting date {as_of_text}")
    reset = usable("next_reset_date")
    if fold(usable("settlement_type")) == "STM" and reset and date.fromisoformat(reset) <= as_of:
        add("reset", position, "next_reset_date", "next reset date must be after the reporting date (the MtM is the "
                                                  f"value after that day's settlement): next_reset_date {reset}, "
                                                  f"reporting date {as_of_text}")


def check_currencies(position: int, usable: Callable[[str], str], currencies: set[str], reporting: str,
                     add: Adder) -> None:
    """The notional and MtM currencies, and an FX pair's currencies, against the FX table."""
    notional = fold(usable("notional_currency"))
    if notional and notional not in currencies:
        add("currency", position, "notional_currency", f"no FX rate for notional currency '{notional}'")
    mtm = fold(usable("market_value_currency"))
    if mtm and mtm not in currencies:
        add("currency", position, "market_value_currency", f"no FX rate for MtM currency '{mtm}'")
    pair = normalize_pair(usable("risk_factor")) if fold(usable("asset_class")) == "FX" else None
    if pair:
        missing = [code for code in pair if code not in currencies]
        if missing:
            add("pair", position, "risk_factor", f"no FX rate for currency '{missing[0]}' of the pair {'/'.join(pair)}")


def check_factor(position: int, usable: Callable[[str], str], has_flawed: Callable[[str], bool],
                 factors: dict[str, str], regimes: dict[str, str], params_regime: str, add: Adder) -> None:
    """The supervisory-factor row of the trade, and its regime and OT's against the netting set's. A blank
    sub-class gives a key such as CR_, which no row has, as in a run."""
    asset_class = fold(usable("asset_class"))
    sub_class = fold(usable("sub_class"))
    if not asset_class or (asset_class not in ("IR", "FX", "OT") and has_flawed("sub_class")):
        return
    key = factor_key(asset_class, sub_class)
    column = "asset_class" if asset_class in ("IR", "FX", "OT") else "sub_class"
    if key not in factors:
        add("factor", position, column, f"no supervisory factor for '{key}' (check sub-class)")
    netting_set = fold(usable("netting_set_id"))
    if netting_set not in regimes:
        return
    regime = regimes[netting_set] if regimes[netting_set] in ("CRR", "BCBS") else params_regime
    if asset_class == "OT" and regime != "CRR":
        add("regime", position, "asset_class",
            "asset class OT (other risks) exists only under CRR (Art. 277(1)(f), 280f); under BCBS map the trade to "
            "the class of its primary risk driver (CRE52)")
    elif factors.get(key, "") not in ("", regime):
        add("regime", position, column,
            f"'{key}' is defined only for regime {factors[key]} (netting set uses {regime})")

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
