"""Where things are on the workbook template, read from its workbook names (#756).

The tests and gates find a header row, a table or a parameter cell from the workbook
name that the VBA code also reads (NettingSets_Header, Params_Code, FX_Table, AsOfDate,
...), never from a row number, so that the owner can move rows in Excel without a
change to the code or the tests. Excel keeps the names on their cells when rows move.

The helpers also read a formula as Excel may save it: a cell of a shared formula holds
no text of its own, and its formula is the master's, moved by the offset of the cell.
"""
from __future__ import annotations

import re
import zipfile
from typing import NamedTuple
from xml.etree import ElementTree

MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REFERENCE = re.compile(r"^'?(?P<sheet>[^'!]+)'?!\$?(?P<column>[A-Z]+)\$?(?P<row>\d+)(?::\$?[A-Z]+\$?\d+)?$")
# A cell reference in a formula: not part of a name, a function or a sheet name, nor inside a string.
FORMULA_REFERENCE = re.compile(r"(?<![A-Za-z0-9_.!$'\"])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(!])")
STRING = re.compile(r'"(?:[^"]|"")*"')


class NamedCell(NamedTuple):
    """The top-left cell of a workbook name."""

    sheet: str
    column: str
    row: int


def column_number(letters: str) -> int:
    number = 0
    for letter in letters:
        number = number * 26 + ord(letter) - 64
    return number


def column_letters(number: int) -> str:
    letters = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def defined_names(package: zipfile.ZipFile) -> dict[str, str]:
    """Every workbook-level name of the template and what it refers to."""
    root = ElementTree.fromstring(package.read("xl/workbook.xml"))
    return {name.attrib["name"]: (name.text or "") for name in root.iter(MAIN + "definedName")
            if "localSheetId" not in name.attrib}


def named_cell(package: zipfile.ZipFile, name: str) -> NamedCell:
    """The top-left cell of a workbook name; raises KeyError for an unknown name."""
    reference = defined_names(package)[name]
    match = REFERENCE.match(reference)
    if match is None:
        raise ValueError(f"{name} does not refer to cells: {reference}")
    return NamedCell(match["sheet"], match["column"], int(match["row"]))


def shift_formula(formula: str, rows: int, columns: int) -> str:
    """The formula moved by rows and columns, as Excel moves the relative parts of its references."""
    def move(match: re.Match[str]) -> str:
        column_fixed, letters, row_fixed, row = match.groups()
        new_letters = letters if column_fixed else column_letters(column_number(letters) + columns)
        new_row = row if row_fixed else str(int(row) + rows)
        return f"{column_fixed}{new_letters}{row_fixed}{new_row}"

    pieces: list[str] = []
    last = 0
    for literal in STRING.finditer(formula):
        pieces.append(FORMULA_REFERENCE.sub(move, formula[last:literal.start()]))
        pieces.append(literal.group(0))
        last = literal.end()
    pieces.append(FORMULA_REFERENCE.sub(move, formula[last:]))
    return "".join(pieces)


def formula_text(sheet: ElementTree.Element, cell: ElementTree.Element) -> str:
    """The formula of a cell, "" if it has none; a cell of a shared formula gets the master's, moved."""
    formula = cell.find(MAIN + "f")
    if formula is None:
        return ""
    if formula.text or formula.attrib.get("t") != "shared":
        return formula.text or ""
    index = formula.attrib.get("si")
    for master in sheet.iter(MAIN + "f"):
        if master.attrib.get("t") == "shared" and master.attrib.get("si") == index and master.text:
            origin = re.match(r"([A-Z]+)(\d+)", master.attrib.get("ref", ""))
            target = re.match(r"([A-Z]+)(\d+)", cell.attrib.get("r", ""))
            if origin is None or target is None:
                return ""
            return shift_formula(master.text, int(target[2]) - int(origin[2]),
                                 column_number(target[1]) - column_number(origin[1]))
    return ""
