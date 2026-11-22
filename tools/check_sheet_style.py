#!/usr/bin/env python3
"""Check the sheet style file (docs/workbook/sheet-style.json) and the
workbook template against it, without running Excel (#706, #708).

Every role of the editability matrix has one tab color in the style file;
every sheet of the template has the tab color of its role and its gridlines
turned off. Every font of the template is the style's font and every cell
format is vertically centered. Every cell of a table-header row of the style
file is bold, white, of the header font size, on the header fill, and the row
is tall enough for the header's wrapped text, by a line count that is the same
on every host (wrapped_lines), within the pixel by which Excel rounds a row
height when it saves. A header row is given by the workbook name of a cell of
it wherever the template has one, so that a row moved in Excel needs no edit
to the style file (#756); a header-size exception is a cell or NAME:COLUMN.
The font rule covers the fonts that cell formats use: Excel adds the font of
its notes (Tahoma) to the styles when it saves a sheet with notes. No shared formula of the template refers to
another sheet: Excel never writes one and refuses to open a file that holds
one (#708). What Excel displays after a run or Clear
outputs (cells showing ####, dates) is checked by the sheet review of the
quick Excel run, tools/Verify-SACCR-Branch.ps1 -Quick."""
from __future__ import annotations

import json
import math
import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import Any, NamedTuple

from _gatelib import parse_report_args, run_gate
from check_workbook_protection import MATRIX, NS, TEMPLATE, column_number, sheet_parts
from template_layout import REFERENCE, defined_names

STYLE = "docs/workbook/sheet-style.json"
KEYS = {"schema_version", "description", "font", "vertical_alignment", "date_format", "gridlines", "tab_colors",
        "header_fill", "header_font_size", "header_rows", "header_size_exceptions"}
OPTIONAL_KEYS = {"header_size_exceptions"}
COLOR = re.compile(r"^[0-9A-F]{6}$")
CELL = re.compile(r"^([A-Z]{1,3})([1-9][0-9]*)$")
WHITE = "FFFFFF"
STRING_LITERAL = re.compile(r'"(?:[^"]|"")*"')
# Points of row height per line of text, per point of font size, and the padding of a row: 11 pt text
# takes 15 pt a line, as Excel's default row, and N lines N times that, as Excel's autofit gives them:
# two lines of 11 pt fit the 30 pt header rows of the input sheets (#756). The line count of
# wrapped_lines is generous on purpose: a row a little too tall costs nothing, one too short cuts its
# last line.
LINE_FACTOR = 1.36
ROW_PADDING = 0.0
# Excel saves a row height rounded down to a pixel, 0.75 pt at 96 dpi: 63.84 pt is saved as 63.75 pt.
PIXEL = 0.75
NAME = re.compile(r"^[A-Za-z_][A-Za-z0-9_.]*$")
NAMED_CELL = re.compile(r"^(?P<name>[A-Za-z_][A-Za-z0-9_.]*):(?P<column>[A-Z]{1,3})$")


def wrapped_lines(text: str, width: float, size: float) -> int:
    """Lines a text takes when wrapped in a column of the given width (Excel units) at the given font size.

    A line holds at most one character per width unit at 11 pt, fewer at larger sizes; a text breaks after
    a space or a hyphen and at its line breaks, and a word longer than a line breaks inside. The estimate
    never counts fewer lines than Excel shows for the headers of the template (#708)."""
    capacity = max(1, math.floor(width * 11.0 / size))
    lines = 0
    for paragraph in text.split("\n"):
        words = re.findall(r"[^ -]+-?|-", paragraph.strip()) or [""]
        used = 0
        lines += 1
        for word in words:
            need = len(word) if used == 0 else used + 1 + len(word)
            if need <= capacity:
                used = need
                continue
            if used > 0:
                lines += 1
            extra, rest = divmod(len(word), capacity)
            lines += extra - (1 if rest == 0 and extra > 0 else 0)
            used = rest if rest else capacity
    return lines


def row_height(lines: int, size: float) -> float:
    """The height in points a row needs for so many lines of text of the given size."""
    return round(lines * size * LINE_FACTOR + ROW_PADDING, 2)


class Styles(NamedTuple):
    fonts: list[dict[str, Any]]
    fills: list[str]
    formats: list[dict[str, Any]]


def read_styles(part: bytes) -> Styles:
    root = ET.fromstring(part)
    fonts = []
    for font in root.findall("m:fonts/m:font", NS):
        size = font.find("m:sz", NS)
        name = font.find("m:name", NS)
        color = font.find("m:color", NS)
        fonts.append({"name": name.get("val") if name is not None else "",
                      "size": float(size.get("val", "11")) if size is not None else 11.0,
                      "bold": font.find("m:b", NS) is not None,
                      "color": (color.get("rgb", "") if color is not None else "").upper()[-6:]})
    fills = []
    for fill in root.findall("m:fills/m:fill", NS):
        colour = fill.find("m:patternFill/m:fgColor", NS)
        fills.append((colour.get("rgb", "") if colour is not None else "").upper()[-6:])
    formats = []
    for xf in root.findall("m:cellXfs/m:xf", NS):
        alignment = xf.find("m:alignment", NS)
        formats.append({"font": int(xf.get("fontId", "0")), "fill": int(xf.get("fillId", "0")),
                        "vertical": alignment.get("vertical", "bottom") if alignment is not None else "bottom",
                        "wrap": alignment is not None and alignment.get("wrapText") in ("1", "true")})
    return Styles(fonts, fills, formats)


def size_exception_findings(exceptions: Any) -> list[str]:
    """Findings on the header cells that have a font size of their own."""
    if isinstance(exceptions, dict) and all(
            isinstance(cells, dict) and all((CELL.match(ref) or NAMED_CELL.match(ref)) and isinstance(size, (int, float))
                                            and size > 0 for ref, size in cells.items()) for cells in exceptions.values()):
        return []
    return [f"{STYLE}: header_size_exceptions must map sheets to cells (or NAME:COLUMN) and their header font size"]


def named_row(sheet: str, name: str, names: dict[str, str]) -> int:
    """The row of the first cell of a workbook name on the sheet; raises ValueError otherwise."""
    match = REFERENCE.match(names.get(name, ""))
    if match is None or match["sheet"] != sheet:
        raise ValueError(f"{STYLE}: {sheet}: {name!r} is not a workbook name of a cell of the sheet")
    return int(match["row"])


def resolved_rows(sheet: str, rows: list[Any], names: dict[str, str]) -> list[int]:
    """The header rows of a sheet: row numbers as given, workbook names as the row of their first cell."""
    return [row if isinstance(row, int) else named_row(sheet, row, names) for row in rows]


def resolved_exceptions(sheet: str, cells: dict[str, Any], names: dict[str, str]) -> dict[str, float]:
    """The header-size exceptions of a sheet as cell references: NAME:COLUMN is that column on the name's row."""
    out: dict[str, float] = {}
    for reference, size in cells.items():
        match = NAMED_CELL.match(reference)
        out[f"{match['column']}{named_row(sheet, match['name'], names)}" if match else reference] = float(size)
    return out


def check_style(style: Any, roles: set[str]) -> list[str]:
    """Findings on the style file itself: its keys, and one color per role."""
    if not isinstance(style, dict):
        return [f"{STYLE}: not a JSON object"]
    findings = []
    if set(style) - OPTIONAL_KEYS != KEYS - OPTIONAL_KEYS or not set(style) <= KEYS:
        findings.append(f"{STYLE}: keys must be {sorted(KEYS)} (optional: {sorted(OPTIONAL_KEYS)}), found {sorted(style)}")
    if style.get("schema_version") != 1:
        findings.append(f"{STYLE}: schema_version must be 1")
    if style.get("gridlines") is not False:
        findings.append(f"{STYLE}: gridlines must be false")
    for key in ("font", "vertical_alignment", "date_format"):
        if not isinstance(style.get(key), str) or not style.get(key):
            findings.append(f"{STYLE}: {key} must be a non-empty string")
    if not isinstance(style.get("header_fill"), str) or not COLOR.match(style.get("header_fill", "")):
        findings.append(f"{STYLE}: header_fill must be six upper-case hex digits")
    if not isinstance(style.get("header_font_size"), (int, float)) or style.get("header_font_size", 0) <= 0:
        findings.append(f"{STYLE}: header_font_size must be a positive number")
    findings += size_exception_findings(style.get("header_size_exceptions", {}))
    headers = style.get("header_rows")
    if not isinstance(headers, dict) or not all(
            isinstance(rows, list) and rows and all((isinstance(r, int) and not isinstance(r, bool) and r > 0) or
                                                    (isinstance(r, str) and NAME.match(r)) for r in rows)
            for rows in headers.values()):
        findings.append(f"{STYLE}: header_rows must map sheets to lists of row numbers or workbook names")
    colors = style.get("tab_colors")
    if not isinstance(colors, dict):
        return findings + [f"{STYLE}: tab_colors must be an object"]
    for role in sorted(roles - set(colors)):
        findings.append(f"{STYLE}: role {role!r} of {MATRIX} has no tab color")
    for role in sorted(set(colors) - roles):
        findings.append(f"{STYLE}: tab color for {role!r}, which is not a role of {MATRIX}")
    for role, color in sorted(colors.items()):
        if not isinstance(color, str) or not COLOR.match(color):
            findings.append(f"{STYLE}: tab color of {role!r} must be six upper-case hex digits, found {color!r}")
    return findings


def style_findings(styles: Styles, font: str, vertical: str) -> list[str]:
    """Every font a cell format uses is the style's font and every cell format has the style's vertical alignment."""
    used = {f["font"] for f in styles.formats}
    findings = [f"{TEMPLATE}: font {i} is {f['name']!r}, not {font!r}"
                for i, f in enumerate(styles.fonts) if f["name"] != font and i in used]
    findings += [f"{TEMPLATE}: cell format {i} is aligned {f['vertical']!r}, not {vertical!r}"
                 for i, f in enumerate(styles.formats) if f["vertical"] != vertical]
    return findings


def sheet_findings(name: str, part: bytes, expected: str | None) -> list[str]:
    """Findings on one template sheet: its tab color and its gridlines."""
    root = ET.fromstring(part)
    findings = []
    tab = root.find("m:sheetPr/m:tabColor", NS)
    color = tab.get("rgb", "") if tab is not None else ""
    if expected is not None and color.upper()[-6:] != expected:
        findings.append(f"{TEMPLATE}: {name}: tab color {color or 'none'}, expected {expected} for its role")
    views = root.findall("m:sheetViews/m:sheetView", NS)
    if not views or any(view.get("showGridLines") not in ("0", "false") for view in views):
        findings.append(f"{TEMPLATE}: {name}: gridlines are not turned off")
    for cell in root.iterfind("m:sheetData/m:row/m:c", NS):
        formula = cell.find("m:f", NS)
        if formula is not None and formula.get("t") == "shared" and "!" in STRING_LITERAL.sub("", formula.text or ""):
            findings.append(f"{TEMPLATE}: {name}!{cell.get('r', '')}: shared formula refers to another sheet, "
                            "which Excel refuses to open; give each cell its own formula")
    return findings


def column_widths(root: ET.Element) -> dict[int, float]:
    widths = {}
    for column in root.findall("m:cols/m:col", NS):
        for number in range(int(column.get("min", "1")), int(column.get("max", "1")) + 1):
            widths[number] = float(column.get("width", "9.140625"))
    return widths


def cell_text(cell: ET.Element, strings: list[str]) -> str:
    value = cell.find("m:v", NS)
    if cell.get("t") == "s" and value is not None and value.text:
        return strings[int(value.text)]
    if cell.get("t") == "inlineStr":
        return "".join(t.text or "" for t in cell.iter(f"{{{NS['m']}}}t"))
    return ""


def header_findings(name: str, part: bytes, rows: list[int], styles: Styles, strings: list[str],
                    style: dict[str, Any], exceptions: dict[str, float] | None = None) -> list[str]:
    """Every text cell of a header row has the header look, and the row is tall enough for it."""
    exceptions = exceptions if exceptions is not None else {
        ref: float(size) for ref, size in style.get("header_size_exceptions", {}).get(name, {}).items()}
    root = ET.fromstring(part)
    widths = column_widths(root)
    fill, size = style["header_fill"], float(style["header_font_size"])
    findings = []
    for number in rows:
        row = root.find(f"m:sheetData/m:row[@r='{number}']", NS)
        if row is None:
            findings.append(f"{TEMPLATE}: {name}: header row {number} does not exist")
            continue
        needed = 1
        texts = 0
        for cell in row.findall("m:c", NS):
            text = cell_text(cell, strings)
            if not text.strip():
                continue
            texts += 1
            reference = cell.get("r", "")
            fmt = styles.formats[int(cell.get("s", "0"))]
            font = styles.fonts[fmt["font"]]
            cell_size = float(exceptions.get(reference, size))
            if styles.fills[fmt["fill"]] != fill or not font["bold"] or font["color"] != WHITE or font["size"] != cell_size:
                findings.append(f"{TEMPLATE}: {name}!{reference}: header is not bold white {cell_size:g} pt on {fill}")
            match = CELL.match(reference)
            width = widths.get(column_number(match.group(1)), 9.140625) if match else 9.140625
            needed = max(needed, wrapped_lines(text, width, size) if fmt["wrap"] else 1)
        if texts == 0:
            findings.append(f"{TEMPLATE}: {name}: header row {number} holds no text")
        height = float(row.get("ht", "15"))
        if height + PIXEL < row_height(needed, size):
            findings.append(f"{TEMPLATE}: {name}: header row {number} is {height:g} pt, its text needs "
                            f"{needed} line(s), {row_height(needed, size):g} pt")
    return findings


def shared_strings(archive: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    return ["".join(t.text or "" for t in item.iter(f"{{{NS['m']}}}t")) for item in root.findall("m:si", NS)]


def run_check(root: Path) -> dict[str, Any]:
    matrix = json.loads((root / MATRIX).read_text(encoding="utf-8"))
    style = json.loads((root / STYLE).read_text(encoding="utf-8"))
    sheets = matrix.get("sheets", {})
    roles = {spec.get("role") for spec in sheets.values() if isinstance(spec, dict)}
    findings = check_style(style, {role for role in roles if isinstance(role, str)})
    if findings:
        return {"schema_version": 1, "status": "fail", "sheets": 0, "findings": findings}
    colors = style["tab_colors"]
    with zipfile.ZipFile(root / TEMPLATE) as archive:
        parts = sheet_parts(archive)
        styles = read_styles(archive.read("xl/styles.xml"))
        strings = shared_strings(archive)
        names = defined_names(archive)
        findings += style_findings(styles, style["font"], style["vertical_alignment"])
        for name in sorted(set(style["header_rows"]) - set(parts)):
            findings.append(f"{STYLE}: header rows for {name}, which is not a sheet of the template")
        for name, part in parts.items():
            role = sheets.get(name, {}).get("role") if isinstance(sheets.get(name), dict) else None
            expected = colors.get(role) if isinstance(role, str) else None
            data = archive.read(part)
            findings += sheet_findings(name, data, expected)
            if name in style["header_rows"]:
                try:
                    rows = resolved_rows(name, style["header_rows"][name], names)
                    exceptions = resolved_exceptions(name, style.get("header_size_exceptions", {}).get(name, {}), names)
                except ValueError as error:
                    findings.append(str(error))
                    continue
                findings += header_findings(name, data, rows, styles, strings, style, exceptions)
    return {"schema_version": 1, "status": "fail" if findings else "pass", "sheets": len(parts),
            "findings": findings}


def markdown_report(report: dict[str, Any]) -> str:
    lines = [f"Sheet style: {report['status'].upper()} ({report['sheets']} sheet(s) in the template)"]
    lines.extend(report["findings"])
    return "\n".join(lines) + "\n"


def main() -> int:
    options = parse_report_args(sys.argv[1:], description=__doc__)
    return run_gate(options, build=lambda: run_check(options.root.resolve()),
                    markdown=markdown_report,
                    errors=(OSError, ValueError, KeyError, json.JSONDecodeError, zipfile.BadZipFile, ET.ParseError))


if __name__ == "__main__":
    raise SystemExit(main())
