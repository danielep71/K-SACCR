#!/usr/bin/env python3
"""Draw the sheet navigation bar into the workbook template, or check it (#615).

Every sheet of src/workbook/SACCR_Template.xlsx carries a vertical bar of
linked shapes in column A, listing every sheet in the sections of the Ribbon's
Go to menu (src/workbook/customUI14.xml), with the menu's tooltips:

* column A is widened and filled with the title band's colour; column B stays
  the 4-wide gutter, so no table, header or input cell moves;
* the Inputs section forms a top block anchored in rows 1 to 3, which are
  frozen on every sheet, so it never scrolls; it fits the rows as they are
  (row 3 would be raised only if a sheet's rows 1 to 3 were too short);
* the other sections form a second block anchored on the first row below the
  frozen rows, in the menu's order except Start (Home), which comes last; it
  scrolls with the rows;
* column A is frozen, so both blocks stay in view when the sheet scrolls
  horizontally;
* each button carries an internal hyperlink to its sheet, so the bar works
  with macros disabled; the current sheet is highlighted and not linked; the
  shapes are not printed.

The package is edited as XML, so every other part of the template (the
Ribbon, form controls, validation extensions, protection) is kept byte for
byte. Every step is idempotent: the bar is removed and drawn again, so
``--write`` on a written template changes nothing.

``--write`` rewrites the template. Without it the script draws the bar on a
copy of the committed template and fails when the result differs from the
committed template, or when a sheet lacks its bar or lists the wrong sheets.
"""
from __future__ import annotations

import argparse
import math
import os
import posixpath
import re
import sys
import tempfile
import zipfile
from collections.abc import Callable, Iterator
from pathlib import Path
from xml.sax.saxutils import escape

TEMPLATE = "src/workbook/SACCR_Template.xlsx"
CONTRACT = "src/workbook/customUI14.xml"
CONFIG = "src/core/CORE_Config.bas"           # names the groups for the VBA (#620)
FIXED_TIME = (1980, 1, 1, 0, 0, 0)

# ------------------------------------------------------------------ design ----
NAV_COL_WIDTH = "21.7"           # characters of the default font; about 157 px
PX = 9525                        # EMU per pixel at 96 dpi
BAR_X = 8                        # left margin of the bar inside column A, px
BAR_W = 141                      # bar width, px
PAD_TOP = 6
PAD_BOTTOM = 3
HEAD_H = 17
BTN_H = 21
BTN_GAP = 2
SECTION_GAP = 7
STRIP_W = 3
TOP_ROWS = 3                     # rows frozen on every sheet

FONT = "Aptos"
HEAD_COLOR = "FFC000"            # gold of the K-SACCR wordmark
BTN_FILL = "262626"
BTN_TEXT = "D9D9D9"
CUR_FILL = "265A99"              # title band blue
CUR_TEXT = "FFFFFF"
CUR_STRIP = "FFC000"
DARK_STRIP = "8FAADC"            # used when a tab colour would vanish on the button

GROUP_TOP = "NavBarTop"
GROUP_REST = "NavBar"

NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_DRAWING = NS_R + "/drawing"
REL_HYPERLINK = NS_R + "/hyperlink"
CT_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml"
EMPTY_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
              '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
              '</Relationships>')
EMPTY_DRAWING = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                 '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
                 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"></xdr:wsDr>')
GUTTER_COL = '<col min="2" max="2" width="4" customWidth="1"/>'
ORIGINAL_COLS = '<col min="1" max="2" width="4" customWidth="1"/>'
ANCHOR = re.compile(r"<xdr:oneCellAnchor>(?:(?!</xdr:oneCellAnchor>).)*?</xdr:oneCellAnchor>", re.S)
RELATIONSHIP = re.compile(r"<Relationship\b[^>]*/>")
Block = list[tuple[str, list[tuple[str, str]]]]


def need(pattern: str, text: str, flags: int = 0) -> re.Match[str]:
    """re.search that fails with the pattern named, instead of returning None."""
    found = re.search(pattern, text, flags)
    if found is None:
        raise ValueError(f"not found: {pattern}")
    return found


def attribute(element: str, name: str) -> str | None:
    found = re.search(r'\b%s="([^"]*)"' % name, element)
    return found.group(1) if found else None


# ---------------------------------------------------------------- contract ----
def sections(contract: str) -> Block:
    """The Go to menu: [(section title, [(sheet name, tooltip), ...]), ...]."""
    menu = re.search(r'<menu id="ksGoMenu"[^>]*>(.*?)</menu>', contract, re.S)
    if not menu:
        raise ValueError(f"{CONTRACT}: no ksGoMenu")
    found: Block = []
    for element in re.findall(r"<(?:menuSeparator|button)\b[^>]*/>", menu.group(1)):
        if element.startswith("<menuSeparator"):
            found.append((unescape(attribute(element, "title") or ""), []))
        elif found:
            found[-1][1].append((unescape(attribute(element, "label") or ""),
                                 unescape(attribute(element, "supertip") or "")))
    return found


def unescape(text: str) -> str:
    return (text.replace("&apos;", "'").replace("&quot;", '"').replace("&lt;", "<")
            .replace("&gt;", ">").replace("&amp;", "&"))


def blocks(menu: Block) -> tuple[Block, Block]:
    """The bar's two blocks: Inputs in the frozen rows; the rest, with Start (Home) moved last."""
    if len(menu) < 2:
        raise ValueError(f"{CONTRACT}: the Go to menu needs at least the Start and Inputs sections")
    return [menu[1]], menu[2:] + [menu[0]]


# --------------------------------------------------------------- package -----
class Package:
    def __init__(self, path: Path):
        with zipfile.ZipFile(path) as package:
            self.infos = package.infolist()
            self.entries = {info.filename: package.read(info.filename) for info in self.infos}
        self.added: list[str] = []

    def text(self, name: str) -> str:
        return self.entries[name].decode("utf-8")

    def has(self, name: str) -> bool:
        return name in self.entries

    def put(self, name: str, text: str) -> None:
        if name not in self.entries:
            self.added.append(name)
        self.entries[name] = text.encode("utf-8")

    def write(self, path: Path) -> None:
        descriptor, temporary = tempfile.mkstemp(dir=path.parent, suffix=".xlsx")
        os.close(descriptor)
        try:
            with zipfile.ZipFile(temporary, "w") as output:
                for info in self.infos:
                    output.writestr(info, self.entries[info.filename], compress_type=info.compress_type)
                for name in sorted(self.added):
                    info = zipfile.ZipInfo(name, date_time=FIXED_TIME)
                    info.compress_type = zipfile.ZIP_DEFLATED
                    output.writestr(info, self.entries[name])
            os.replace(temporary, path)
        finally:
            if os.path.exists(temporary):
                os.remove(temporary)


def relationships(xml: str) -> dict[str, tuple[str, str]]:
    found = {}
    for element in RELATIONSHIP.findall(xml):
        found[need(r'Id="([^"]+)"', element).group(1)] = (
            need(r'Type="([^"]+)"', element).group(1), need(r'Target="([^"]+)"', element).group(1))
    return found


def next_id(xml: str) -> str:
    numbers = [int(n) for n in re.findall(r'Id="rId(\d+)"', xml)]
    return "rId%d" % (max(numbers, default=0) + 1)


def rels_path(part: str) -> str:
    return posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")


def sheets_of(package: Package) -> list[tuple[str, str]]:
    workbook = package.text("xl/workbook.xml")
    targets = relationships(package.text("xl/_rels/workbook.xml.rels"))
    found = []
    for element in re.findall(r"<sheet\b[^>]*/>", workbook):
        name = unescape(need(r'\bname="([^"]+)"', element).group(1))
        target = targets[need(r'\br:id="([^"]+)"', element).group(1)][1]
        found.append((name, posixpath.normpath(posixpath.join("xl", target.lstrip("/")))
                      if not target.startswith("/") else target.lstrip("/")))
    return found


# ---------------------------------------------------------------- drawing ----
def shape(shape_id: int, name: str, x: int, y: int, w: int, h: int, fill: str | None, text: str = "",
          color: str = "", bold: bool = False, size: int = 900, link: str | None = None, tip: str = "",
          anchor: str = "ctr", inset: int = 0, spacing: int = 0) -> str:
    click = ""
    if link:
        click = '<a:hlinkClick xmlns:r="%s" r:id="%s" tooltip="%s"/>' % (NS_R, link, escape(tip, {'"': "&quot;"}))
    fill_xml = '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>' % fill if fill else "<a:noFill/>"
    body = ""
    if text:
        run = ('<a:rPr lang="en-GB" sz="%d" b="%d"%s><a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
               '<a:latin typeface="%s"/><a:cs typeface="%s"/></a:rPr>') % (
                   size, 1 if bold else 0, ' spc="%d"' % spacing if spacing else "", color, FONT, FONT)
        body = ('<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="none" lIns="%d" tIns="0" '
                'rIns="0" bIns="%d" rtlCol="0" anchor="%s"/><a:lstStyle/><a:p><a:pPr algn="l"/>'
                '<a:r>%s<a:t>%s</a:t></a:r></a:p></xdr:txBody>') % (
                    inset * PX, 2 * PX if anchor == "b" else 0, anchor, run, escape(text))
    return ('<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="%d" name="%s">%s</xdr:cNvPr>'
            '<xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>%s<a:ln><a:noFill/></a:ln></xdr:spPr>%s</xdr:sp>') % (
                shape_id, name, click, x, y, w, h, fill_xml, body)


def block_height(block: Block) -> int:
    height = PAD_TOP
    for index, (_, items) in enumerate(block):
        height += (SECTION_GAP if index else 0) + HEAD_H + len(items) * (BTN_H + BTN_GAP) - BTN_GAP
    return height + PAD_BOTTOM


def strip_colour(rgb: str) -> str:
    red, green, blue = (int(rgb[i:i + 2], 16) for i in (0, 2, 4))
    return DARK_STRIP if 0.2126 * red + 0.7152 * green + 0.0722 * blue < 70 else rgb


def group(name: str, block: Block, sheet: str, ids: Iterator[int], link_for: Callable[[str], str],
          colours: dict[str, str], row: int, y0: int) -> str:
    """One block of the bar, anchored at the top of `row` (0-based), whose top lies `y0` EMU down the sheet."""
    x0, children, y = BAR_X * PX, [], PAD_TOP
    for index, (title, items) in enumerate(block):
        if index:
            y += SECTION_GAP
        children.append(shape(next(ids), "Nav_Section_" + re.sub(r"\W", "", title.title()), x0, y0 + y * PX,
                              BAR_W * PX, HEAD_H * PX, None, title.upper(), HEAD_COLOR, True, 700,
                              anchor="b", inset=4, spacing=60))
        y += HEAD_H
        for target, tip in items:
            current = target == sheet
            link = None if current else link_for(target)
            children.append(shape(next(ids), "Nav_" + target, x0, y0 + y * PX, BAR_W * PX, BTN_H * PX,
                                  CUR_FILL if current else BTN_FILL, target, CUR_TEXT if current else BTN_TEXT,
                                  current, 900, link, tip, inset=12))
            children.append(shape(next(ids), "Nav_" + target + "_Mark", x0, y0 + y * PX, STRIP_W * PX, BTN_H * PX,
                                  CUR_STRIP if current else strip_colour(colours.get(target, "1F3864"))))
            y += BTN_H + BTN_GAP
    height = (y - BTN_GAP + PAD_BOTTOM) * PX
    xml = ('<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="%d" name="%s" descr="Sheet navigation"/><xdr:cNvGrpSpPr/>'
           '</xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/>'
           '<a:chOff x="%d" y="%d"/><a:chExt cx="%d" cy="%d"/></a:xfrm></xdr:grpSpPr>%s</xdr:grpSp>') % (
               next(ids), name, x0, y0, BAR_W * PX, height, x0, y0, BAR_W * PX, height, "".join(children))
    return ('<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>%d</xdr:colOff><xdr:row>%d</xdr:row>'
            '<xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="%d" cy="%d"/>%s'
            '<xdr:clientData fPrintsWithSheet="0"/></xdr:oneCellAnchor>') % (x0, row, BAR_W * PX, height, xml)


def strip_bar(drawing: str, drawing_rels: str) -> tuple[str, str]:
    """Remove the bar's anchors and the hyperlinks only they used."""
    drawing = ANCHOR.sub(lambda m: "" if re.search(r'name="(%s|%s)"' % (GROUP_TOP, GROUP_REST), m.group(0))
                         else m.group(0), drawing)
    used = set(re.findall(r'r:id="([^"]+)"', drawing))
    for element in RELATIONSHIP.findall(drawing_rels):
        rid = need(r'Id="([^"]+)"', element).group(1)
        if REL_HYPERLINK in element and rid not in used:
            drawing_rels = drawing_rels.replace(element, "")
    return drawing, drawing_rels


# ------------------------------------------------------------------ sheets ----
def column_letter(index: int) -> str:
    letters = ""
    while index:
        index, rest = divmod(index - 1, 26)
        letters = chr(65 + rest) + letters
    return letters


def cell_position(reference: str) -> tuple[int, int]:
    match = need(r"^([A-Z]+)(\d+)", reference)
    column = 0
    for letter in match.group(1):
        column = column * 26 + ord(letter) - 64
    return column, int(match.group(2))


def freeze(xml: str) -> tuple[str, int]:
    """Freeze column A and rows 1 to 3 at least; return the sheet and its frozen rows."""
    view = need(r"<sheetView\b[^>]*?(/>|>(.*?)</sheetView>)", xml, re.S)
    inner = view.group(2) or ""
    pane = re.search(r"<pane\b[^>]*/>", inner)
    def split(name: str) -> int:
        value = attribute(pane.group(0), name) if pane else None
        return int(float(value)) if value else 0

    xsplit, ysplit = max(split("xSplit"), 1), max(split("ySplit"), TOP_ROWS)
    selections = re.findall(r"<selection\b[^>]*/>", inner)
    active = None
    for selection in selections:
        if 'pane="bottomRight"' in selection:
            active = re.search(r'activeCell="([A-Z]+\d+)"', selection)
    if active is None and selections:
        active = re.search(r'activeCell="([A-Z]+\d+)"', selections[-1])
    first = column_letter(xsplit + 1) + str(ysplit + 1)
    cell = active.group(1) if active else first
    column, row = cell_position(cell)
    if column <= xsplit or row <= ysplit:
        cell = first
    new = ('<pane xSplit="%d" ySplit="%d" topLeftCell="%s" activePane="bottomRight" state="frozen"/>'
           '<selection pane="topRight" activeCell="%s1" sqref="%s1"/>'
           '<selection pane="bottomLeft" activeCell="A%d" sqref="A%d"/>'
           '<selection pane="bottomRight" activeCell="%s" sqref="%s"/>') % (
               xsplit, ysplit, first, column_letter(xsplit + 1), column_letter(xsplit + 1),
               ysplit + 1, ysplit + 1, cell, cell)
    rest = re.sub(r"<pane\b[^>]*/>|<selection\b[^>]*/>", "", inner)
    opening = need(r"^<sheetView\b[^>]*?(?=/?>)", view.group(0)).group(0)
    xml = xml[:view.start()] + opening + ">" + new + rest + "</sheetView>" + xml[view.end():]
    return xml, ysplit


def raise_row(xml: str, row: int, points: float) -> str:
    element = re.search(r'<row r="%d"[^>]*?/?>' % row, xml)
    if element is None:
        raise ValueError("no row %d" % row)
    text = element.group(0)
    height = re.search(r'\bht="([\d.]+)"', text)
    if height and float(height.group(1)) >= points:
        return xml
    value = ("%.2f" % points).rstrip("0").rstrip(".")
    if height:
        new = text.replace(height.group(0), 'ht="%s"' % value)
    else:
        new = re.sub(r"(/?>)$", ' ht="%s" customHeight="1"\\1' % value, text)
    if "customHeight" not in new:
        new = re.sub(r"(/?>)$", ' customHeight="1"\\1', new)
    return xml[:element.start()] + new + xml[element.end():]


def row_offset_points(xml: str, rows: int) -> float:
    default = float(need(r'defaultRowHeight="([\d.]+)"', xml).group(1))
    heights = {int(r): float(h) for r, h in re.findall(r'<row r="(\d+)"[^>]*?\bht="([\d.]+)"', xml)}
    hidden = {int(r) for r in re.findall(r'<row r="(\d+)"[^>]*?\bhidden="1"', xml)}
    return sum(heights.get(r, default) for r in range(1, rows + 1) if r not in hidden)


def first_visible_row(xml: str, after: int) -> int:
    hidden = {int(r) for r in re.findall(r'<row r="(\d+)"[^>]*?\bhidden="1"', xml)}
    row = after + 1
    while row in hidden:
        row += 1
    return row


# --------------------------------------------------------------- the build ----
def build(package: Package, contract: str) -> list[str]:
    """Draw the bar on every sheet of the package; return what was done per sheet."""
    menu = sections(contract)
    sheets = sheets_of(package)
    listed = [name for _, items in menu for name, _ in items]
    names = [name for name, _ in sheets]
    if listed != names:
        raise ValueError(f"{CONTRACT}: the Go to menu lists {listed}, the template's sheets are {names}")
    top, rest = blocks(menu)
    top_points = math.ceil(block_height(top) * 0.75 * 4) / 4

    styles = package.text("xl/styles.xml")
    fills = re.findall(r"<fill>.*?</fill>|<fill/>", need(r"<fills[^>]*>(.*?)</fills>", styles, re.S).group(1))
    band = next(i for i, f in enumerate(fills) if '<fgColor theme="1"/>' in f)
    style_xml = '<xf numFmtId="0" fontId="0" fillId="%d" borderId="0" xfId="0" applyFill="1"/>' % band
    xfs = need(r'<cellXfs count="(\d+)">(.*?)</cellXfs>', styles, re.S)
    existing = re.findall(r"<xf\b[^>]*?(?:/>|>.*?</xf>)", xfs.group(2), re.S)
    if style_xml in existing:
        nav_style = existing.index(style_xml)
    else:
        nav_style = len(existing)
        styles = styles.replace(xfs.group(0), '<cellXfs count="%d">%s%s</cellXfs>' % (
            nav_style + 1, xfs.group(2), style_xml), 1)
        package.put("xl/styles.xml", styles)

    colours: dict[str, str] = {}
    for name, part in sheets:
        found = re.search(r'<tabColor rgb="FF([0-9A-F]{6})"', package.text(part))
        colours[name] = found.group(1) if found else "1F3864"

    types = package.text("[Content_Types].xml")
    report: list[str] = []
    for name, part in sheets:
        xml = package.text(part)
        sheet_rels_path = rels_path(part)
        sheet_rels = package.text(sheet_rels_path) if package.has(sheet_rels_path) else EMPTY_RELS

        nav_col = '<col min="1" max="1" width="%s" style="%d" customWidth="1"/>' % (NAV_COL_WIDTH, nav_style)
        if ORIGINAL_COLS in xml:
            xml = xml.replace(ORIGINAL_COLS, nav_col + GUTTER_COL, 1)
        elif (nav_col + GUTTER_COL) not in xml:
            raise ValueError(f"{name}: columns A and B are not the 4-wide margins the bar expects")
        for row in re.findall(r'<row r="(\d+)"[^>]*customFormat="1"[^>]*?(?<!/)>', xml):
            opening = need(r'<row r="%s"[^>]*>' % row, xml)
            if not xml[opening.end():].startswith('<c r="A%s"' % row):
                xml = xml[:opening.end()] + '<c r="A%s" s="%d"/>' % (row, nav_style) + xml[opening.end():]

        xml, frozen = freeze(xml)
        needed = top_points - row_offset_points(xml, TOP_ROWS - 1)
        xml = raise_row(xml, TOP_ROWS, needed)
        rest_row = first_visible_row(xml, frozen)

        element = re.search(r'<drawing r:id="([^"]+)"/>', xml)
        if element:
            drawing_part = posixpath.normpath(posixpath.join(
                posixpath.dirname(part), relationships(sheet_rels)[element.group(1)][1]))
            drawing = package.text(drawing_part)
        else:
            number = 1
            while package.has("xl/drawings/drawing%d.xml" % number):
                number += 1
            drawing_part, drawing = "xl/drawings/drawing%d.xml" % number, EMPTY_DRAWING
            rid = next_id(sheet_rels)
            sheet_rels = sheet_rels.replace("</Relationships>", '<Relationship Id="%s" Type="%s" '
                                            'Target="../drawings/%s"/></Relationships>' % (
                                                rid, REL_DRAWING, posixpath.basename(drawing_part)))
            tag = '<drawing r:id="%s"/>' % rid
            follower = re.search(r"<(legacyDrawing|legacyDrawingHF|picture|oleObjects|controls|"
                                 r"webPublishItems|tableParts|extLst)\b", xml)
            xml = xml[:follower.start()] + tag + xml[follower.start():] if follower \
                else xml.replace("</worksheet>", tag + "</worksheet>")
            types = types.replace("</Types>", '<Override PartName="/%s" ContentType="%s"/></Types>' % (
                drawing_part, CT_DRAWING))
            package.put(sheet_rels_path, sheet_rels)
        drawing_rels_path = rels_path(drawing_part)
        drawing_rels = package.text(drawing_rels_path) if package.has(drawing_rels_path) else EMPTY_RELS
        drawing, drawing_rels = strip_bar(drawing, drawing_rels)

        state = {"rels": drawing_rels}

        def link_for(target: str) -> str:
            rid = next_id(state["rels"])
            state["rels"] = state["rels"].replace("</Relationships>", '<Relationship Id="%s" Type="%s" '
                                                  'Target="#\'%s\'!A1"/></Relationships>' % (
                                                      rid, REL_HYPERLINK, target.replace("'", "''")))
            return rid

        start = max([int(i) for i in re.findall(r'<xdr:cNvPr id="(\d+)"', drawing)], default=1) + 1
        ids = iter(range(start, start + 10000))
        rest_y = int(round(row_offset_points(xml, rest_row - 1) * 12700))
        anchors = (group(GROUP_TOP, top, name, ids, link_for, colours, 0, 0)
                   + group(GROUP_REST, rest, name, ids, link_for, colours, rest_row - 1, rest_y))
        if 'xmlns:a="' not in drawing.split(">", 2)[1]:
            drawing = drawing.replace("<xdr:wsDr ", '<xdr:wsDr xmlns:a="http://schemas.openxmlformats.org/'
                                                    'drawingml/2006/main" ', 1)
        drawing = drawing.replace("</xdr:wsDr>", anchors + "</xdr:wsDr>", 1)

        package.put(part, xml)
        package.put(drawing_part, drawing)
        package.put(drawing_rels_path, state["rels"])
        report.append("%-20s frozen rows %d, second block from row %d" % (name, frozen, rest_row))
    package.put("[Content_Types].xml", types)
    return report


# ------------------------------------------------------------------- check ----
def bar_problems(package: Package, contract: str) -> list[str]:
    """Each sheet holds both blocks, linking every other sheet in the menu's order."""
    top, rest = blocks(sections(contract))
    order = [name for _, items in top + rest for name, _ in items]
    found: list[str] = []
    for name, part in sheets_of(package):
        sheet_rels = package.text(rels_path(part)) if package.has(rels_path(part)) else EMPTY_RELS
        element = re.search(r'<drawing r:id="([^"]+)"/>', package.text(part))
        if not element:
            found.append(f"{name}: no navigation bar")
            continue
        drawing_part = posixpath.normpath(posixpath.join(posixpath.dirname(part),
                                                         relationships(sheet_rels)[element.group(1)][1]))
        drawing = package.text(drawing_part)
        links = relationships(package.text(rels_path(drawing_part))) if package.has(rels_path(drawing_part)) else {}
        for group_name in (GROUP_TOP, GROUP_REST):
            if drawing.count('name="%s"' % group_name) != 1:
                found.append(f"{name}: {group_name} missing or repeated")
        targets = [need(r"^#'(.*)'!A1$", links[rid][1]).group(1).replace("''", "'")
                   for rid in re.findall(r'<a:hlinkClick [^>]*r:id="([^"]+)"', drawing)
                   if rid in links and links[rid][0] == REL_HYPERLINK and links[rid][1].startswith("#")]
        expected = [sheet for sheet in order if sheet != name]
        if targets != expected:
            found.append(f"{name}: the bar links {targets}, expected {expected}")
    return found


def config_problems(root: Path) -> list[str]:
    """CORE_Config names the two groups as this tool draws them, for the VBA that hides them."""
    config = root / CONFIG
    if not config.is_file():
        return [f"{CONFIG} is missing"]
    text = config.read_text(encoding="cp1252")
    found = []
    for constant, expected in (("NAV_GROUP_TOP", GROUP_TOP), ("NAV_GROUP_REST", GROUP_REST)):
        match = re.search(r'Public Const %s\s+As String = "([^"]*)"' % constant, text)
        if not match or match.group(1) != expected:
            found.append(f'{CONFIG}: {constant} must be "{expected}", the name tools/navigation_bar.py draws')
    return found


def problems(root: Path) -> list[str]:
    template, source = root / TEMPLATE, root / CONTRACT
    if not source.is_file():
        return [f"{CONTRACT} is missing"]
    contract = source.read_text(encoding="utf-8")
    committed = Package(template)
    found: list[str] = config_problems(root)
    try:
        found += bar_problems(committed, contract)
    except (ValueError, KeyError, AttributeError) as error:
        found.append(f"{TEMPLATE}: {error}")
    rebuilt = Package(template)
    try:
        build(rebuilt, contract)
    except (ValueError, KeyError, AttributeError) as error:
        return found + [f"{TEMPLATE}: {error}"]
    changed = sorted(name for name in rebuilt.entries if committed.entries.get(name) != rebuilt.entries[name])
    if changed:
        found.append(f"{TEMPLATE}: the navigation bar is not current in {', '.join(changed)}; "
                     "run tools/navigation_bar.py --write")
    return found


def write(root: Path) -> list[str]:
    template = root / TEMPLATE
    package = Package(template)
    report = build(package, (root / CONTRACT).read_text(encoding="utf-8"))
    package.write(template)
    return report


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--write", action="store_true", help="draw the bar into the template")
    args = parser.parse_args()
    root = args.root.resolve()
    if args.write:
        for line in write(root):
            print(line)
    found = problems(root)
    for problem in found:
        print(problem, file=sys.stderr)
    if not found:
        print(f"{TEMPLATE}: the navigation bar matches {CONTRACT}")
    return 1 if found else 0


if __name__ == "__main__":
    sys.exit(main())
