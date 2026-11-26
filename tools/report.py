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
- the rows of a reconciliation that the PDF keeps: those that are not MATCH for
  netting sets, the largest DIFFERENCE rows for trade legs;
- the rows of Checks the PDF keeps, at most a number of each code.

Usage::

    python tools/report.py --check

checks the specification: the sections, the kinds and the examples.
"""

from __future__ import annotations

import argparse
import json
import sys
from datetime import date, timedelta
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path
from typing import Any, Iterable, Sequence

ROOT = Path(__file__).resolve().parent.parent
SPEC = ROOT / "docs" / "workbook" / "report.json"
KINDS = {"text", "count", "amount", "ratio", "date", "datetime"}
COLUMN_KINDS = KINDS | {"measure"}
LAYOUTS = {"key_value", "blocks", "table", "summary_and_table", "list"}
SEVERITIES = ("ERROR", "WARNING", "INFO")
NO_CODE = "(none)"
MATCH = "MATCH"
DIFFERENCE = "DIFFERENCE"
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
    parts: list[str] = []
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
    """What a PDF shows for a value of a column of this kind; a blank stays blank.

    A logical value is TRUE or FALSE and a text stays as it is, also in a column of numbers
    (the "n/a" of a cap that does not apply), as CORE_Export.ReportCellText writes them.
    """
    if value is None or value == "":
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, str):
        return value
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
        return format(float(value), ".15g").upper()  # NumberText: Str$ with up to 15 significant digits
    raise ReportError(f"unknown kind {kind!r}")


def format_text(value: Any, number_format: str) -> str:
    """A value as the number format of its cell shows it, with invariant separators.

    The Dashboard's figures keep the format of their cell on the sheet in the PDF:
    a format with a year or a day code is a date (with an hour code, a date and time);
    General, text and an empty format give the 15-digit text; otherwise the decimals
    are the digit placeholders after the point of the first section, a comma before
    the point groups thousands, and a percent sign multiplies by 100.
    """
    if value is None or value == "" or isinstance(value, (bool, str)):
        return cell_text("text", value)
    section = number_format.upper().split(";", 1)[0]
    if "YY" in section or "DD" in section:
        return datetime_text(float(value)) if "H" in section else date_text(float(value))
    if section in ("", "GENERAL", "@"):
        return cell_text("text", value)
    point = section.find(".")
    decimals = 0
    if point >= 0:
        for character in section[point + 1:]:
            if character not in "0#":
                break
            decimals += 1
        thousands = "," in section[:point]
    else:
        thousands = "," in section
    if "%" in section:
        return number_text(float(value) * 100, decimals, thousands) + "%"
    return number_text(float(value), decimals, thousands)


def measure_kind(measure: str, ratios: Iterable[str]) -> str:
    """The format of a figure of a reconciliation row: ratio for a measure that is not an amount, else amount."""
    return "ratio" if measure in set(ratios) else "amount"


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


def plural(count: int, one: str, many: str) -> str:
    return f"{count} {one if count == 1 else many}"


def absolute(value: Any) -> float:
    """The absolute value of a figure of a reconciliation; 0 for a blank or a text."""
    try:
        return abs(float(value))
    except (TypeError, ValueError):
        return 0.0


def pdf_trade_rows(rows: Sequence[Sequence[Any]], result_index: int, difference_index: int,
                   limit: int) -> tuple[list[Sequence[Any]], int, int]:
    """The rows of the reconciliation of trade legs the PDF shows: DIFFERENCE rows, largest absolute difference
    first and ties in their order on the sheet, at most limit of them; with the number of DIFFERENCE rows and
    the number of rows left out."""
    differences = [row for row in rows if str(row[result_index]).strip().upper() == DIFFERENCE]
    kept = sorted(differences, key=lambda row: -absolute(row[difference_index]))[:limit]
    return kept, len(differences), len(rows) - len(kept)


def differences_note(shown: int, differences: int, left_out: int) -> str:
    """The line under the reconciliation of trade legs in the PDF; blank when nothing is left out."""
    if left_out == 0:
        return ""
    rest = (f"the other {plural(left_out, 'row is', 'rows are')} left out, counted by result above; "
            "the workbook version of the report has every row.")
    if differences == 0:
        return f"No row has the result {DIFFERENCE}; " + rest
    if shown == differences:
        what = "The row" if shown == 1 else f"The {shown} rows"
        return f"{what} with the result {DIFFERENCE} {'is' if shown == 1 else 'are'} shown; " + rest
    return (f"The {shown} rows with the result {DIFFERENCE} and the largest absolute difference are shown, "
            f"of {differences}; " + rest)


def pdf_checks_rows(rows: Sequence[Sequence[Any]], per_code: int, severity_index: int = 0,
                    code_index: int = 6) -> tuple[list[Sequence[Any]], int]:
    """The rows of Checks the PDF shows, in code order, at most per_code of each code; and how many are left out."""
    seen: dict[str, int] = {}
    kept = []
    for row in checks_in_code_order(rows, severity_index, code_index):
        code = check_code(row, code_index)
        seen[code] = seen.get(code, 0) + 1
        if seen[code] <= per_code:
            kept.append(row)
    return kept, len(rows) - len(kept)


def checks_note(per_code: int, left_out: int) -> str:
    """The line under the title of the rows of Checks in the PDF; blank when nothing is left out."""
    if left_out == 0:
        return ""
    return (f"At most {per_code} rows of each code are shown; {plural(left_out, 'row is', 'rows are')} left out; "
            "the workbook version of the report has every row.")


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
                if column["kind"] not in COLUMN_KINDS:
                    problems.append(f"{section['id']}: unknown kind {column['kind']!r}")
        measures = [c["header"] for c in section.get("columns", []) if c["kind"] == "measure"]
        if measures and not section.get("measure_ratios"):
            problems.append(f"{section['id']}: columns of kind measure need measure_ratios")
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
