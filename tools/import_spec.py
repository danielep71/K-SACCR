#!/usr/bin/env python3
"""Generate the CSV import specification, a standalone HTML page, from the CSV contract (#473).

The page describes the NettingSets and Trades files for a user who has only the
workbook: the file format, the header, the value types, the blank and empty-file
rules, the limits, every column, the rules between fields, between records and
against Params, the finding codes, and the synthetic example records of the
contract as CSV text to copy. Everything it states about the files comes from
docs/csv/contract.json; the prose around it says how the reader applies it.

python tools/import_spec.py --write writes docs/csv/K-SACCR_Import_Specification.html;
without --write it exits 1 when that file differs from what the contract generates.
The check_csv_contract gate runs the same comparison and validates the examples.
"""
from __future__ import annotations

import argparse
import html
import sys
from pathlib import Path
from typing import Any

from csv_contract import (
    DELIMITER,
    FORMULA_START,
    Finding,
    check_context,
    check_portfolio,
    load_contract,
    validate,
)

SPECIFICATION = "docs/csv/K-SACCR_Import_Specification.html"
MIN_EXAMPLES, MAX_EXAMPLES = 3, 5
DATE_COLUMNS = ("maturity_date", "end_date", "option_expiry_date", "next_reset_date")
CURRENCY_COLUMNS = ("notional_currency", "market_value_currency")

STYLE = """
:root { --ink: #1b2733; --muted: #4a5a68; --line: #c9d3dd; --head: #e8f0fa; --accent: #1f5aa6; --code: #f4f6f8; }
* { box-sizing: border-box; }
body { margin: 0; background: #ffffff; color: var(--ink);
       font: 15px/1.5 "Segoe UI", "Helvetica Neue", Arial, sans-serif; }
header, nav, main, footer { max-width: 1180px; margin: 0 auto; padding: 0 20px; }
header { padding-top: 24px; }
h1 { font-size: 26px; margin: 0 0 6px; color: var(--accent); }
h2 { font-size: 21px; margin: 32px 0 10px; padding-bottom: 4px; border-bottom: 2px solid var(--head); }
h3 { font-size: 17px; margin: 22px 0 8px; }
p, ul { margin: 8px 0; }
.lead { color: var(--muted); }
nav ol { columns: 2; margin: 12px 0; padding-left: 22px; }
code { background: var(--code); padding: 0 3px; border-radius: 3px; font-family: Consolas, "Courier New", monospace; font-size: 0.9em; }
.wide { overflow-x: auto; }
table { border-collapse: collapse; width: 100%; margin: 10px 0 16px; font-size: 13.5px; }
th, td { border: 1px solid var(--line); padding: 5px 7px; text-align: left; vertical-align: top; }
th { background: var(--head); }
td.num { text-align: right; white-space: nowrap; }
pre.csv { background: var(--code); border: 1px solid var(--line); padding: 10px 12px; margin: 8px 0 16px;
          font: 13px/1.45 Consolas, "Courier New", monospace; white-space: pre-wrap; overflow-wrap: anywhere;
          user-select: all; }
footer { color: var(--muted); font-size: 13px; padding-top: 24px; padding-bottom: 32px; }
@page { size: A4 landscape; margin: 12mm; }
@media print {
  body { font-size: 11px; }
  nav { display: none; }
  header, main, footer { max-width: none; padding: 0; }
  h2 { break-after: avoid; }
  table { font-size: 9.5px; }
  pre.csv { font-size: 9.5px; }
  tr, pre.csv { break-inside: avoid; }
  .wide { overflow: visible; }
}
""".strip("\n")


def esc(text: str) -> str:
    return html.escape(text, quote=True)


def code(text: str) -> str:
    return f"<code>{esc(text)}</code>"


def listed(values: list[str]) -> str:
    return ", ".join(code(value) for value in values)


def csv_field(value: str) -> str:
    """A field as the contract writes it: quoted when it holds the delimiter, a quote or a line break."""
    if any(char in value for char in (DELIMITER, '"', "\n", "\r")):
        return '"' + value.replace('"', '""') + '"'
    return value


def example_text(kind: str, contract: dict[str, Any]) -> str:
    """The example file of kind: its header and its example records, each line ending in LF."""
    spec = contract["files"][kind]
    names = [column["name"] for column in spec["columns"]]
    lines = [DELIMITER.join(names)]
    lines += [DELIMITER.join(csv_field(record.get(name, "")) for name in names) for record in spec["examples"]]
    return "\n".join(lines) + "\n"


def example_rows(kind: str, contract: dict[str, Any]) -> list[tuple[int, dict[str, str]]]:
    """The example records of kind as (CSV line, record) with every column, for the rules between records."""
    names = [column["name"] for column in contract["files"][kind]["columns"]]
    return [(line, {name: record.get(name, "") for name in names})
            for line, record in enumerate(contract["files"][kind]["examples"], start=2)]


def describe(findings: list[Finding]) -> list[str]:
    return [f"line {f.line}, record {f.record}, {f.column or '-'}: {f.code}: {f.message}" for f in findings]


def example_findings(contract: dict[str, Any], context: dict[str, Any]) -> list[str]:
    """Why the example records would not import together into a workbook with the given Params context:
    their number, each file as the reader validates it, the rules between records and those against Params."""
    findings: list[str] = []
    for kind, spec in contract["files"].items():
        count = len(spec.get("examples", []))
        if not MIN_EXAMPLES <= count <= MAX_EXAMPLES:
            findings.append(f"{kind}: {count} example records, the specification shows {MIN_EXAMPLES} to {MAX_EXAMPLES}")
            continue
        names = {column["name"] for column in spec["columns"]}
        for position, record in enumerate(spec["examples"], start=1):
            unknown = sorted(set(record) - names)
            if unknown:
                findings.append(f"{kind}: example {position} names unknown column(s) {unknown}")
        _, file_findings = validate(example_text(kind, contract).encode("utf-8"), kind, contract)
        findings += [f"{kind} example: {text}" for text in describe(file_findings)]
    if findings:
        return findings
    trades = example_rows("Trades", contract)
    netting_sets = {row["netting_set_id"]: row["regime_override"] for _, row in example_rows("NettingSets", contract)}
    findings += [f"Trades example: {text}" for text in describe(check_portfolio(trades, list(netting_sets), contract))]
    findings += [f"Trades example: {text}" for text in
                 describe(check_context(trades, netting_sets, context, contract))]
    return findings


def type_text(column: dict[str, Any], types: dict[str, Any]) -> str:
    kind = str(column["type"])
    limits = [f"{label} {column[key]}" for key, label in (("min", "&ge;"), ("min_exclusive", "&gt;"), ("max", "&le;"))
              if key in column]
    if kind in ("id", "text"):
        limits.append(f"at most {types[kind]['max_length']} characters")
    return esc(kind) + (" (" + ", ".join(limits) + ")" if limits else "")


def allowed_text(column: dict[str, Any], types: dict[str, Any]) -> str:
    if column["type"] == "code":
        return listed(column["allowed"])
    if column["type"] == "flag":
        return listed(types["flag"]["allowed"])
    if column["type"] == "currency":
        return "three capital letters, such as " + code("EUR")
    return ""


def field_table(kind: str, contract: dict[str, Any]) -> str:
    rows = ['<div class="wide"><table>', "<thead><tr><th>#</th><th>CSV column (exact header)</th><th>Sheet column</th>"
            "<th>Meaning</th><th>Type and limits</th><th>Required</th><th>Allowed values</th><th>Blank means</th>"
            "<th>Example</th></tr></thead>", "<tbody>"]
    for position, column in enumerate(contract["files"][kind]["columns"], start=1):
        cells = [f'<td class="num">{position}</td>', f"<td>{code(column['name'])}</td>",
                 f"<td>{esc(column['sheet_header'])}</td>", f"<td>{esc(column['description'])}</td>",
                 f"<td>{type_text(column, contract['types'])}</td>",
                 f"<td>{'yes' if column['required'] else 'no'}</td>",
                 f"<td>{allowed_text(column, contract['types'])}</td>",
                 f"<td>{esc(column.get('blank', ''))}</td>",
                 f"<td>{code(column['example']) if column['example'] else ''}</td>"]
        rows.append("<tr>" + "".join(cells) + "</tr>")
    rows += ["</tbody>", "</table></div>"]
    return "\n".join(rows)


def row_rule_text(rule: dict[str, Any]) -> str:
    if "when" in rule:
        condition = " and ".join(f"{code(name)} is {code(value)}" for name, value in rule["when"].items())
        return (f"When {condition}, {listed(rule['required'])} {'is' if len(rule['required']) == 1 else 'are'} "
                f"required; otherwise {listed(rule['otherwise_blank'])} "
                f"{'is' if len(rule['otherwise_blank']) == 1 else 'are'} blank.")
    if "after" in rule:
        later, earlier = rule["after"]
        return f"When both are given, {code(later)} is after {code(earlier)}."
    upper, lower = rule["above"]
    return f"When both are given, {code(upper)} is above {code(lower)}."


def rule_table(rules: list[dict[str, Any]]) -> str:
    rows = ["<table>", "<thead><tr><th>Rule</th><th>Finding</th><th>Requirement</th></tr></thead>", "<tbody>"]
    rows += [f"<tr><td>{code(rule['id'])}</td><td>{code(rule['code'])}</td><td>{esc(rule['description'])}</td></tr>"
             for rule in rules]
    rows += ["</tbody>", "</table>"]
    return "\n".join(rows)


def assumptions(contract: dict[str, Any]) -> str:
    """What the Trades examples need from Params, worked out from the examples themselves."""
    rows = [row for _, row in example_rows("Trades", contract)]
    dates = sorted(row[name] for row in rows for name in DATE_COLUMNS if row[name])
    currencies = {row[name] for row in rows for name in CURRENCY_COLUMNS if row[name]}
    currencies |= {part for row in rows if row["asset_class"] == "FX" for part in row["risk_factor"].split("/")}
    keys = sorted({row["asset_class"] + ("_" + row["sub_class"] if row["sub_class"] else "") for row in rows})
    return (f"The trades suppose a reporting date before {code(dates[0])}, a rate in the FX table for "
            f"{listed(sorted(currencies))}, and a supervisory-factor row for {listed(keys)}, as the workbook "
            "template has. With other Params, a validation may report a matured trade or a missing rate.")


def section(anchor: str, title: str, body: str) -> str:
    return f'<section id="{anchor}">\n<h2>{esc(title)}</h2>\n{body}\n</section>'


def overview(contract: dict[str, Any]) -> str:
    rows = ["<table>", "<thead><tr><th>File</th><th>Sheet</th><th>Key column</th><th>Records per file</th>"
            "<th>Blank template</th></tr></thead>", "<tbody>"]
    for kind, spec in contract["files"].items():
        rows.append(f"<tr><td>{esc(kind)}</td><td>{esc(spec['sheet'])}</td><td>{code(spec['key'])}</td>"
                    f'<td class="num">{spec["max_records"]:,}</td><td>{code(spec["template"])}</td></tr>')
    rows += ["</tbody>", "</table>"]
    return "\n".join([
        "<p>Each file carries the inputs of one sheet: a record of <code>NettingSets</code> is one netting set, a "
        "record of <code>Trades</code> one leg. <strong>Import CSV</strong> on the staging sheet "
        "<code>Import_NettingSets</code> or <code>Import_Trades</code> reads a file; <strong>Validate staging</strong> "
        "checks the staged rows of both against each other and against Params; <strong>Commit import</strong> then "
        "replaces the input sheet. <strong>Export CSV template</strong> on <code>NettingSets</code> and "
        "<code>Trades</code> writes the blank template, the header alone.</p>",
        *rows,
    ])


def file_format(contract: dict[str, Any]) -> str:
    fmt = contract["format"]
    starts = ", ".join(code(char) for char in FORMULA_START if char.isprintable()) + ", a tab or a carriage return"
    return "\n".join([
        "<ul>",
        f"<li><strong>Encoding.</strong> {esc(fmt['encoding'])}. A file that is not valid UTF-8, or that contains a NUL "
        "character, is rejected. Save the file as UTF-8 text; any language may be used in a text field.</li>",
        f"<li><strong>Delimiter.</strong> Fields are separated by a semicolon, {code(fmt['delimiter'])}. A file whose "
        "header is one field holding commas or tabs was saved with another delimiter and is rejected.</li>",
        f"<li><strong>Quoting.</strong> A field that contains {code(fmt['delimiter'])}, a double quote or a line break "
        f"is enclosed in double quotes, {code(fmt['quote'])}, and a quote inside it is written twice: "
        f"{code(chr(34) + 'A ' + chr(34) * 2 + 'B' + chr(34) * 2 + chr(34))} reads {code('A ' + chr(34) + 'B' + chr(34))}. "
        "Quotes enclose a whole field only: a quote inside an unquoted field, text after a closing quote, or a quoted "
        "field left open at the end of the file is a syntax error.</li>",
        f"<li><strong>Line endings.</strong> Records are separated by {esc(fmt['record_separator'])}; the last record "
        "may end without one. A carriage return anywhere else outside a quoted field is an error.</li>",
        "<li><strong>Multi-line fields.</strong> A quoted field may span several lines; only text fields accept a line "
        "break. A finding names the line on which its record starts and the record number, counted from 1 after the "
        "header.</li>",
        f"<li><strong>Numbers.</strong> Decimal point {code(fmt['decimal_separator'])}, no thousands separator, no "
        f"{code('+')} sign, no leading zero ({code('0.5')} is a number, {code('007.5')} is not), an optional exponent "
        f"({code('1.25e6')}); at most {fmt['max_significant_digits']} "
        f"significant digits, so that a value is never rounded when it is read. {code('1,5')} and {code('1.000,00')} "
        "are rejected.</li>",
        f"<li><strong>Dates.</strong> {code(fmt['date_format'])}, from {code(contract['types']['date']['min'])} to "
        f"{code(contract['types']['date']['max'])}; {code('31/03/2027')} and serial numbers are rejected.</li>",
        "<li><strong>Spaces.</strong> Spaces belong to the field. A value of a type other than text with a leading or "
        "trailing space is rejected, never trimmed.</li>",
        f"<li><strong>No formulas.</strong> Text that starts with {starts} is rejected, so that no field becomes a "
        "formula on a sheet. A negative number is a value, not a formula.</li>",
        "<li><strong>Excel.</strong> Excel's own CSV save writes numbers and dates in the format of the computer's "
        "locale, such as a decimal comma or a day-first date, and in some locales separates fields with commas, so "
        "such a file is rejected. Write the file with a program that follows these rules.</li>",
        "</ul>",
    ])


def header_rules() -> str:
    return ("<p>The first record is the header. It names every column of the file exactly once, in any order, with "
            "the exact lower-case names below; names are case-sensitive. A missing, unknown or repeated column rejects "
            "the file before any record is read. There are no optional columns: an optional field is a column whose "
            "value may be blank. The column set is the contract version: a file of another version is never mapped by "
            "guessing what a column means.</p>")


def value_types(contract: dict[str, Any]) -> str:
    types = contract["types"]
    rows = [
        ("id", "Text, read as text, so leading zeros are kept (" + code("00123") + ")",
         f"At most {types['id']['max_length']} characters; no " + listed(list(types["id"]["forbidden"]))
         + "; no control character. Keys are compared case-insensitively."),
        ("text", "Text", f"At most {types['text']['max_length']} characters; line breaks allowed, other control "
                         "characters not."),
        ("code", "One of the listed values", "Exact spelling and case: " + code("Linear") + ", not " + code("linear")
         + "."),
        ("flag", listed(types["flag"]["allowed"]), "Exact; " + listed(["y", "Yes", "TRUE", "1"]) + " are rejected."),
        ("currency", "Three capital letters", code("EUR") + ", not " + code("eur") + "."),
        ("integer", code("-?digits"), "No sign " + code("+") + ", no leading zero, no decimal point."),
        ("number", code("-?digits[.digits][e±digits]"), "As under Numbers above; a non-zero magnitude from "
         "2.2250738585072014e-308 to 9.99999999999999e307."),
        ("date", code(contract["format"]["date_format"]), "A calendar date within the range above."),
    ]
    lines = ["<table>", "<thead><tr><th>Type</th><th>Written as</th><th>Rules</th></tr></thead>", "<tbody>"]
    lines += [f"<tr><td>{code(name)}</td><td>{written}</td><td>{rules}</td></tr>" for name, written, rules in rows]
    lines += ["</tbody>", "</table>"]
    return "\n".join(lines)


def blanks_and_limits(contract: dict[str, Any]) -> str:
    limit_rows = "".join(f'<tr><td>{esc(kind)}</td><td class="num">{spec["max_records"]:,}</td></tr>'
                         for kind, spec in contract["files"].items())
    max_bytes = contract["format"]["max_bytes"]
    return "\n".join([
        "<ul>",
        "<li>A <strong>blank field</strong> is an empty one. A required field must not be blank; an optional blank "
        "field takes the meaning given under <em>Blank means</em>, which is the sheet's.</li>",
        "<li>A <strong>blank record</strong>, whose fields are all empty, and an <strong>empty line</strong> are "
        "errors.</li>",
        "<li>A <strong>header-only file</strong> is valid and holds no record. Committing it never clears the input "
        "sheet: a staging sheet without rows leaves its input sheet as it is.</li>",
        "<li>A file is accepted or rejected <strong>as a whole</strong>: one finding rejects every record, and no "
        "record is dropped or truncated.</li>",
        "</ul>",
        "<table><thead><tr><th>File</th><th>Records per file</th></tr></thead><tbody>" + limit_rows
        + "</tbody></table>",
        f"<p>A file of more than {max_bytes:,} bytes ({max_bytes // 1048576} MiB), a byte-order mark included, is "
        "rejected from its size alone. These limits are the import's. The sheets' drop-down lists cover only the rows "
        "a user types into, and the calculation reads every row of the input sheets, so they limit neither the import "
        "nor the calculation.</p>",
    ])


def netting_sets_section(contract: dict[str, Any]) -> str:
    spec = contract["files"]["NettingSets"]
    return "\n".join([
        f"<p>One record per netting set. {code(spec['key'])} is the key: unique in the file, compared "
        f"case-insensitively, and not one of {listed(spec['reserved_keys'])}, which name the totals of Results. Every "
        "trade names its netting set by this key.</p>",
        "<h3>Columns</h3>",
        field_table("NettingSets", contract),
        "<h3>Rules between the fields of a record</h3>",
        "<ul>" + "".join(f"<li>{row_rule_text(rule)}</li>" for rule in spec.get("row_rules", [])) + "</ul>",
        "<p>A clearing member's trade exposure to a central counterparty is a netting set with "
        f"{code('clearing_role')} {code('CCP')}; a client of a clearing member enters {code('CLIENT')}, and a "
        f"clearing member facing its client {code('CM')}. {code('MULTI')} declares a multi-level client structure, "
        "which the file accepts and the calculation rejects.</p>",
        "<h3>Example</h3>",
        f'<pre class="csv" id="example-nettingsets">{esc(example_text("NettingSets", contract))}</pre>',
    ])


def trades_section(contract: dict[str, Any]) -> str:
    spec = contract["files"]["Trades"]
    return "\n".join([
        f"<p>One record per <strong>leg</strong>. {code(spec['key'])} is the key: unique in the file and compared "
        f"case-insensitively. Each leg names its netting set in {code('netting_set_id')}, which must be a netting set "
        f"of the {code('NettingSets')} file, or of the sheet when only trades are imported.</p>",
        "<p>A trade with one risk driver is one record and carries its own MtM: "
        f"{code('economic_trade_id')} is blank and {code('mtm_leg')} is {code('Y')} or blank. A trade with several "
        f"risk drivers, such as a cross-currency swap, is one record per leg, all with the same "
        f"{code('economic_trade_id')} and {code('netting_set_id')}: exactly one leg has {code('mtm_leg')} "
        f"{code('Y')} and carries the whole MtM, and every other leg has {code('mtm_leg')} {code('N')} and an MtM of "
        f"{code('0')}, so that the MtM is counted once.</p>",
        f"<p>A trade settled to market has {code('settlement_type')} {code('STM')} and its "
        f"{code('next_reset_date')}; it belongs to an unmargined netting set.</p>",
        "<h3>Columns</h3>",
        field_table("Trades", contract),
        "<h3>Rules between the fields of a record</h3>",
        "<ul>" + "".join(f"<li>{row_rule_text(rule)}</li>" for rule in spec.get("row_rules", [])) + "</ul>",
        "<h3>Rules between records</h3>",
        "<p>Judged by <strong>Validate staging</strong> on the staged legs together with the netting sets they refer "
        "to, as the calculation judges them.</p>",
        rule_table(contract["portfolio_rules"]),
        "<h3>Rules against Params</h3>",
        "<p>Judged by <strong>Validate staging</strong> against the reporting date, the FX table, the "
        f"supervisory-factor table and the regime of each netting set, its {code('regime_override')} or else the "
        "Params regime.</p>",
        rule_table(contract["context_rules"]),
        "<h3>Example</h3>",
        f"<p>{assumptions(contract)}</p>",
        f'<pre class="csv" id="example-trades">{esc(example_text("Trades", contract))}</pre>',
    ])


def examples_note() -> str:
    return ("<p>The examples are synthetic and hold no real portfolio data. The two belong together: each leg's "
            "netting set is in the NettingSets example. To try them, select the text of an example, copy it into a "
            "text editor, and save it as UTF-8 with the extension <code>.csv</code>; then import both files and "
            "validate them.</p>")


def findings_table(contract: dict[str, Any]) -> str:
    rows = ["<table>", "<thead><tr><th>Finding</th><th>Meaning</th></tr></thead>", "<tbody>"]
    rows += [f"<tr><td>{code(name)}</td><td>{esc(text)}</td></tr>" for name, text in contract["errors"].items()]
    rows += ["</tbody>", "</table>"]
    return "\n".join(rows)


def render(contract: dict[str, Any]) -> str:
    """The specification page of the contract."""
    version = contract["contract_version"]
    sections = [
        ("files", "Files", overview(contract)),
        ("format", "File format", file_format(contract)),
        ("header", "Header", header_rules()),
        ("types", "Value types", value_types(contract)),
        ("blanks", "Blank fields, empty files and limits", blanks_and_limits(contract)),
        ("nettingsets", "NettingSets file", netting_sets_section(contract)),
        ("trades", "Trades file", trades_section(contract)),
        ("examples", "Using the examples", examples_note()),
        ("findings", "Findings", "<p>Every finding of the import and of the validation has one of these codes.</p>\n"
                                 + findings_table(contract)),
    ]
    navigation = "".join(f'<li><a href="#{anchor}">{esc(title)}</a></li>' for anchor, title, _ in sections)
    return "\n".join([
        "<!DOCTYPE html>",
        '<html lang="en">',
        "<head>",
        '<meta charset="utf-8">',
        '<meta name="viewport" content="width=device-width, initial-scale=1">',
        f"<title>K-SACCR CSV import specification, contract version {version}</title>",
        f"<style>\n{STYLE}\n</style>",
        "</head>",
        "<body>",
        "<header>",
        "<h1>K-SACCR CSV import specification</h1>",
        f'<p class="lead">CSV contract version {version}, for the NettingSets and Trades files whose blank templates '
        f"end in <code>_v{version}</code>. The contract version changes only when a column is added, removed or "
        "renamed, not with each K-SACCR release.</p>",
        "</header>",
        f"<nav><ol>{navigation}</ol></nav>",
        "<main>",
        *(section(anchor, title, body) for anchor, title, body in sections),
        "</main>",
        "<footer>Generated from the CSV contract of K-SACCR. It works offline and needs no other file.</footer>",
        "</body>",
        "</html>",
        "",
    ])


def specification_findings(root: Path, contract: dict[str, Any]) -> list[str]:
    path = root / SPECIFICATION
    if not path.is_file() or path.read_text(encoding="utf-8") != render(contract):
        return [f"{SPECIFICATION} is not what the contract generates; run python tools/import_spec.py --write"]
    return []


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--write", action="store_true", help=f"Write {SPECIFICATION}")
    args = parser.parse_args()
    contract = load_contract(args.root)
    if args.write:
        path = args.root / SPECIFICATION
        path.write_bytes(render(contract).encode("utf-8"))
        print(f"wrote {SPECIFICATION}")
        return 0
    findings = specification_findings(args.root, contract)
    for finding in findings:
        print(finding)
    return 1 if findings else 0


if __name__ == "__main__":
    sys.exit(main())
