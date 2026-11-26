#!/usr/bin/env python3
"""Reference of the formatted run report (#259).

``docs/workbook/report.json`` lists the sections of the report that the Export
report command writes, their columns and formats, and the rules that decide
what the PDF shows. This module holds the rules that are computations, so that
``docs/workbook/REPORT.md`` is executable and the VBA writer can be checked
against it:

- the text of an amount, a ratio, a count, a date and a date with time, which
  a PDF shows as text because the numbers of a PDF would otherwise follow the
  regional settings of the machine that writes it;
- Checks by code: the summary of the codes and the order of the rows;
- the rows of a reconciliation that the PDF keeps.

Usage::

    python tools/report.py --check

checks the specification: the sections, the kinds and the examples.
"""

from __future__ import annotations

import argparse
import json
import sys
from datetime import date, datetime, timedelta
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path
from typing import Any, Iterable, Sequence

ROOT = Path(__file__).resolve().parent.parent
SPEC = ROOT / "docs" / "workbook" / "report.json"
KINDS = {"text", "count", "amount", "ratio", "date", "datetime"}
LAYOUTS = {"key_value", "blocks", "table", "summary_and_table", "list"}
SEVERITIES = ("ERROR", "WARNING", "INFO")
NO_CODE = "(none)"
MATCH = "MATCH"
EXCEL_EPOCH = date(1899, 12, 30)
SECONDS_PER_DAY = 86400


class ReportError(ValueError):
    """The specification or a value breaks the rules of the report."""


def load_spec() -> dict[str, Any]:
    return json.loads(SPEC.read_text(encoding="utf-8"))  # type: ignore[no-any-return]


def significant(value: float) -> Decimal:
    """The value to 15 significant digits, the text CORE_Export.NumberText writes."""
    return Decimal(format(value, ".15g"))


def grouped(whole: int) -> str:
    text = str(whole)
    parts = []
    while len(text) > 3:
        parts.insert(0, text[-3:])
        text = text[:-3]
    parts.insert(0, text)
    return ",".join(parts)


def number_text(value: float, decimals: int, thousands: bool) -> str:
    """Round half away from zero on the 15-digit form; no sign for a value that rounds to zero."""
    step = Decimal(1).scaleb(-decimals)
    rounded = significant(value).copy_abs().quantize(step, rounding=ROUND_HALF_UP)
    whole, _, fraction = f"{rounded:f}".partition(".")
    text = grouped(int(whole)) if thousands else whole
    if decimals:
        text += "." + fraction.ljust(decimals, "0")
    return "-" + text if value < 0 and rounded != 0 else text


def date_text(serial: float) -> str:
    """The date of an Excel serial number of the 1900 system, from 1 March 1900."""
    if serial < 61:
        raise ReportError(f"a date before 1 March 1900 is not reported: {serial}")
    day = EXCEL_EPOCH + timedelta(days=int(serial))
    return f"{day.year:04d}-{day.month:02d}-{day.day:02d}"


def datetime_text(serial: float) -> str:
    """The date and the time to the second, rounded half up from the 15-digit form of the serial."""
    if serial < 61:
        raise ReportError(f"a date before 1 March 1900 is not reported: {serial}")
    total = int((significant(serial) * SECONDS_PER_DAY).quantize(Decimal(1), rounding=ROUND_HALF_UP))
    day, seconds = divmod(total, SECONDS_PER_DAY)
    return f"{date_text(day)} {seconds // 3600:02d}:{seconds % 3600 // 60:02d}:{seconds % 60:02d}"


def cell_text(kind: str, value: Any) -> str:
    """What a PDF shows for a value of a column of this kind; a blank stays blank."""
    if value is None or value == "":
        return ""
    if kind == "amount":
        return number_text(float(value), 2, True)
    if kind == "ratio":
        return number_text(float(value), 6, False)
    if kind == "count":
        return number_text(float(value), 0, False)
    if kind == "date":
        return date_text(float(value))
    if kind == "datetime":
        return datetime_text(float(value))
    if kind == "text":
        return str(value)
    raise ReportError(f"unknown kind {kind!r}")


def severity_rank(severity: str) -> int:
    return SEVERITIES.index(severity) if severity in SEVERITIES else len(SEVERITIES)


def check_code(row: Sequence[Any], code_index: int) -> str:
    code = str(row[code_index]).strip() if row[code_index] is not None else ""
    return code or NO_CODE


def checks_by_code(rows: Iterable[Sequence[Any]], severity_index: int = 0,
                   code_index: int = 6) -> list[tuple[str, str, int]]:
    """The codes of the rows of Checks, with their highest severity and their number of rows."""
    found: dict[str, tuple[int, int]] = {}
    for row in rows:
        code = check_code(row, code_index)
        rank = severity_rank(str(row[severity_index]).strip().upper())
        best, count = found.get(code, (len(SEVERITIES), 0))
        found[code] = (min(best, rank), count + 1)
    ordered = sorted(found.items(), key=lambda item: (item[1][0], item[0]))
    return [(code, SEVERITIES[rank] if rank < len(SEVERITIES) else "", count)
            for code, (rank, count) in ordered]


def checks_in_code_order(rows: Sequence[Sequence[Any]], severity_index: int = 0,
                         code_index: int = 6) -> list[Sequence[Any]]:
    """The rows of Checks grouped by code, in the order of the summary, each code's rows as on the sheet."""
    position = {code: i for i, (code, _, _) in enumerate(checks_by_code(rows, severity_index, code_index))}
    return sorted(rows, key=lambda row: position[check_code(row, code_index)])


def pdf_reconciliation_rows(rows: Sequence[Sequence[Any]], result_index: int) -> tuple[list[Sequence[Any]], int]:
    """The rows the PDF shows, those whose Result is not MATCH, and the number of MATCH rows left out."""
    kept = [row for row in rows if str(row[result_index]).strip().upper() != MATCH]
    return kept, len(rows) - len(kept)


def omitted_note(omitted: int) -> str:
    """The line that says how many rows the PDF leaves out; blank when none."""
    if omitted == 0:
        return ""
    return f"{omitted} row{'s' if omitted != 1 else ''} with the result {MATCH} {'are' if omitted != 1 else 'is'} " \
           "left out; the workbook version of the report has every row."


def check_spec(spec: dict[str, Any]) -> list[str]:
    """The problems of the specification; none when it is sound."""
    problems: list[str] = []
    if spec.get("schema") != "k-saccr-report" or spec.get("version") != 1:
        problems.append("the schema or the version is not k-saccr-report 1")
    ids: list[str] = []
    sheets: list[str] = []
    for section in spec["sections"]:
        ids.append(section["id"])
        sheets.append(section["sheet"])
        if section["layout"] not in LAYOUTS:
            problems.append(f"{section['id']}: unknown layout {section['layout']!r}")
        for key in ("columns", "summary_columns"):
            headers = [column["header"] for column in section.get(key, [])]
            if len(set(headers)) != len(headers):
                problems.append(f"{section['id']}: duplicate headers in {key}")
            for column in section.get(key, []):
                if column["kind"] not in KINDS:
                    problems.append(f"{section['id']}: unknown kind {column['kind']!r}")
        for row in section.get("rows", []):
            if row["kind"] not in KINDS:
                problems.append(f"{section['id']}: unknown kind {row['kind']!r}")
    if len(set(ids)) != len(ids):
        problems.append("duplicate section ids")
    if len(set(sheets)) != len(sheets) or any(len(name) > 31 for name in sheets):
        problems.append("sheet names must be distinct and at most 31 characters")
    for example in spec["formats"]["examples"]:
        found = cell_text(example["kind"], example["value"])
        if found != example["text"]:
            problems.append(f"example {example['kind']} {example['value']}: {found!r}, expected {example['text']!r}")
    return problems


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--check", action="store_true", help="check the specification")
    options = parser.parse_args(argv)
    if not options.check:
        parser.print_help()
        return 0
    problems = check_spec(load_spec())
    for problem in problems:
        print(problem, file=sys.stderr)
    print("report: " + ("FAIL" if problems else "PASS"))
    return 1 if problems else 0


if __name__ == "__main__":
    raise SystemExit(main())
