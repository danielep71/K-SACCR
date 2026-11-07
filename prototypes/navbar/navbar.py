"""Add a vertical navigation bar to every sheet of the K-SACCR template.

Prototype. Works on the package XML directly so that every part the
template carries (RibbonX, form controls, data-validation extensions,
protection) survives untouched.

Per sheet:
  * column A becomes the sidebar: width NAV_COL_WIDTH, filled with the
    row-1 band colour (theme dk1), so the bar and the top band form one frame;
  * a group of shapes, anchored in column A on the first unfrozen row
    ("move but don't size"), lists every sheet in the Ribbon's six sections;
    each button links to its sheet (no VBA), the current sheet is highlighted;
  * column A is frozen, so the bar stays visible when scrolling right.

Usage: python navbar.py <template.xlsx> <customUI14.xml> <out.xlsx>
"""
import os
import re
import shutil
import sys
import tempfile
import zipfile
from xml.sax.saxutils import escape

SRC, CUSTOM_UI, OUT = sys.argv[1], sys.argv[2], sys.argv[3]

# ---------------------------------------------------------------- design ----
NAV_COL_WIDTH = 21.7          # characters; about 157 px with Aptos 11
MDW = 7                        # maximum digit width of Aptos 11, px
PX = 9525                      # EMU per pixel at 96 dpi
PT = 12700                     # EMU per point

BAR_X = 8                      # left margin inside column A, px
BAR_W = 141                    # bar width, px
PAD_TOP = 10
HEAD_H = 17
BTN_H = 21
BTN_GAP = 2
SECTION_GAP = 7
STRIP_W = 3

FONT = "Aptos"
HEAD_COLOR = "FFC000"          # gold of the K-SACCR wordmark
BTN_FILL = "262626"            # black, lighter 15 %
BTN_TEXT = "D9D9D9"
CUR_FILL = "265A99"            # title band blue
CUR_TEXT = "FFFFFF"
CUR_STRIP = "FFC000"

SECTIONS = [
    ("START", ["Home"]),
    ("INPUTS", ["Params", "NettingSets", "Trades"]),
    ("STAGING", ["Import_NettingSets", "Import_Trades"]),
    ("OUTPUTS", ["Results", "HedgingSets", "Buckets", "TradeCalc", "Checks"]),
    ("RECONCILIATION", ["BankNettingSets", "Reconciliation", "BankTrades", "TradeReconciliation"]),
    ("REFERENCE & TESTS", ["TestCatalogue", "FormulaTest", "Basel_vs_CRR", "README"]),
]

NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_DRAWING = NS_R + "/drawing"
REL_HYPERLINK = NS_R + "/hyperlink"
CT_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml"


def col_px(width):
    return int(((256 * width + int(128 / MDW)) / 256) * MDW)


def lighten_if_dark(rgb):
    """Accent strips sit on a near-black button: lift colours that would vanish."""
    r, g, b = (int(rgb[i:i + 2], 16) for i in (0, 2, 4))
    if 0.2126 * r + 0.7152 * g + 0.0722 * b < 70:
        return "8FAADC"        # light tint of the navy 1F3864
    return rgb


# ------------------------------------------------------------- read package -
work = tempfile.mkdtemp()
with zipfile.ZipFile(SRC) as z:
    order = z.namelist()
    z.extractall(work)


def rd(p):
    with open(os.path.join(work, p), encoding="utf-8") as f:
        return f.read()


def wr(p, s):
    full = os.path.join(work, p)
    os.makedirs(os.path.dirname(full), exist_ok=True)
    with open(full, "w", encoding="utf-8") as f:
        f.write(s)


def rel_targets(rels_xml):
    out = {}
    for t in re.findall(r"<Relationship [^>]*>", rels_xml):
        out[re.search(r'Id="([^"]+)"', t).group(1)] = (
            re.search(r'Type="([^"]+)"', t).group(1),
            re.search(r'Target="([^"]+)"', t).group(1))
    return out


def next_rid(rels_xml):
    ids = [int(n) for n in re.findall(r'Id="rId(\d+)"', rels_xml)]
    return "rId%d" % (max(ids, default=0) + 1)


workbook = rd("xl/workbook.xml")
wb_rels = rel_targets(rd("xl/_rels/workbook.xml.rels"))
sheets = []
for name, rid in re.findall(r'<sheet name="([^"]+)" sheetId="\d+"(?: state="\w+")? r:id="([^"]+)"', workbook):
    target = wb_rels[rid][1].lstrip("/")
    target = target if target.startswith("xl/") else "xl/" + target
    sheets.append((name, target))
names_in_book = [n for n, _ in sheets]
listed = [n for _, items in SECTIONS for n in items]
assert sorted(listed) == sorted(names_in_book), (set(listed) ^ set(names_in_book))

# tooltips: the Ribbon's own supertips, so the two never disagree in wording
ui = open(CUSTOM_UI, encoding="utf-8").read()
tips = {}
for label, tip in re.findall(r'<button id="ksGo\w*" label="([^"]+)" onAction="RibbonGoTo" supertip="([^"]+)"', ui):
    tips[label] = tip.replace("&apos;", "'").replace("&amp;", "&")

# tab colours
tab = {}
for name, path in sheets:
    m = re.search(r'<tabColor rgb="FF([0-9A-F]{6})"', rd(path))
    tab[name] = m.group(1) if m else "1F3864"

# --------------------------------------------------------------- styles -----
styles = rd("xl/styles.xml")
fills = re.findall(r"<fill>.*?</fill>|<fill/>", re.search(r"<fills[^>]*>(.*?)</fills>", styles, re.S).group(1))
band_fill = next(i for i, f in enumerate(fills) if '<fgColor theme="1"/>' in f)
cellxfs = re.search(r'<cellXfs count="(\d+)">', styles)
nav_xf = int(cellxfs.group(1))
styles = styles.replace(cellxfs.group(0), '<cellXfs count="%d">' % (nav_xf + 1), 1)
styles = styles.replace(
    "</cellXfs>",
    '<xf numFmtId="0" fontId="0" fillId="%d" borderId="0" xfId="0" applyFill="1"/></cellXfs>' % band_fill, 1)
wr("xl/styles.xml", styles)

ctypes = rd("[Content_Types].xml")


# --------------------------------------------------------------- shapes -----
def row_heights(xml):
    default = float(re.search(r'defaultRowHeight="([\d.]+)"', xml).group(1))
    hts = {int(r): float(h) for r, h in re.findall(r'<row r="(\d+)"[^>]*?\bht="([\d.]+)"', xml)}
    hidden = {int(r) for r in re.findall(r'<row r="(\d+)"[^>]*?\bhidden="1"', xml)}
    return default, hts, hidden


def shape(sid, name, x, y, w, h, fill, text="", color=None, bold=False, size=900,
          link=None, tip="", anchor="ctr", lins=0, spc=0, caps=False):
    hl = ""
    if link:
        hl = '<a:hlinkClick xmlns:r="%s" r:id="%s" tooltip="%s"/>' % (NS_R, link, escape(tip, {'"': "&quot;"}))
    fill_xml = '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>' % fill if fill else "<a:noFill/>"
    body = ""
    if text:
        rpr = '<a:rPr lang="en-GB" sz="%d" b="%d"%s%s><a:solidFill><a:srgbClr val="%s"/></a:solidFill>' \
              '<a:latin typeface="%s"/><a:cs typeface="%s"/></a:rPr>' % (
                  size, 1 if bold else 0, ' spc="%d"' % spc if spc else "", ' cap="all"' if caps else "",
                  color, FONT, FONT)
        body = ('<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="none" lIns="%d" tIns="0" '
                'rIns="0" bIns="%d" rtlCol="0" anchor="%s"/><a:lstStyle/><a:p><a:pPr algn="l"/>'
                '<a:r>%s<a:t>%s</a:t></a:r></a:p></xdr:txBody>') % (
                    lins * PX, 2 * PX if anchor == "b" else 0, anchor, rpr, escape(text))
    return ('<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="%d" name="%s">%s</xdr:cNvPr>'
            '<xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>%s<a:ln><a:noFill/></a:ln></xdr:spPr>%s</xdr:sp>') % (
                sid, name, hl, x, y, w, h, fill_xml, body)


def build_bar(sheet_name, ids, new_link, x0, y0):
    """Return (group xml, height px)."""
    kids, y = [], PAD_TOP
    for s_i, (title, items) in enumerate(SECTIONS):
        if s_i:
            y += SECTION_GAP
        kids.append(shape(next(ids), "Nav_Section_%d" % (s_i + 1), x0, y0 + y * PX, BAR_W * PX, HEAD_H * PX,
                          None, title, HEAD_COLOR, True, 700, anchor="b", lins=4, spc=60))
        y += HEAD_H
        for target in items:
            current = target == sheet_name
            link = None if current else new_link(target)
            tip = tips.get(target, target)
            kids.append(shape(next(ids), "Nav_" + target, x0, y0 + y * PX, BAR_W * PX, BTN_H * PX,
                              CUR_FILL if current else BTN_FILL, target,
                              CUR_TEXT if current else BTN_TEXT, current, 900, link, tip, lins=12))
            kids.append(shape(next(ids), "Nav_" + target + "_Mark", x0, y0 + y * PX, STRIP_W * PX, BTN_H * PX,
                              CUR_STRIP if current else lighten_if_dark(tab[target]), link=link, tip=tip))
            y += BTN_H + BTN_GAP
    height = y - BTN_GAP + PAD_TOP
    grp = ('<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="%d" name="NavBar" descr="Sheet navigation"/><xdr:cNvGrpSpPr/>'
           '</xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/>'
           '<a:chOff x="%d" y="%d"/><a:chExt cx="%d" cy="%d"/></a:xfrm></xdr:grpSpPr>%s</xdr:grpSp>') % (
               next(ids), x0, y0, BAR_W * PX, height * PX, x0, y0, BAR_W * PX, height * PX, "".join(kids))
    return grp, height


# ---------------------------------------------------------------- sheets ----
report = []
for sheet_name, path in sheets:
    xml = rd(path)
    rels_path = os.path.join(os.path.dirname(path), "_rels", os.path.basename(path) + ".rels")
    rels = rd(rels_path) if os.path.exists(os.path.join(work, rels_path)) else \
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' \
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'

    # column A: sidebar width and fill; B keeps the 4-wide gutter
    old = '<col min="1" max="2" width="4" customWidth="1"/>'
    assert old in xml, sheet_name
    xml = xml.replace(old, '<col min="1" max="1" width="%s" style="%d" customWidth="1"/>'
                           '<col min="2" max="2" width="4" customWidth="1"/>' % (NAV_COL_WIDTH, nav_xf), 1)
    # rows carrying their own style would override the column fill in A
    for m in list(re.finditer(r'(<row r="(\d+)"[^>]*customFormat="1"[^>]*>)', xml)):
        xml = xml.replace(m.group(1), m.group(1) + '<c r="A%s" s="%d"/>' % (m.group(2), nav_xf), 1)

    # panes: freeze column A as well, reset the saved scroll position
    pane = re.search(r"<pane [^>]*/>", xml)
    ysplit = int(re.search(r'ySplit="(\d+)"', pane.group(0)).group(1)) if pane and "ySplit" in pane.group(0) else 0
    xsplit = int(re.search(r'xSplit="(\d+)"', pane.group(0)).group(1)) if pane and "xSplit" in pane.group(0) else 0
    first_col = "ABCDEFGHIJ"[max(xsplit, 1)]
    if ysplit and xsplit:
        new_pane = '<pane xSplit="%d" ySplit="%d" topLeftCell="%s%d" activePane="bottomRight" state="frozen"/>' % (
            xsplit, ysplit, first_col, ysplit + 1)
        xml = xml.replace(pane.group(0), new_pane, 1)
    elif ysplit:
        sel = re.findall(r"<selection [^>]*/>", xml)
        act = re.search(r'activeCell="([A-Z]+)(\d+)"', sel[-1]) if sel else None
        cell = "%s%s" % (act.group(1), act.group(2)) if act and act.group(1) != "A" else "B%d" % (ysplit + 1)
        new_pane = ('<pane xSplit="1" ySplit="%d" topLeftCell="B%d" activePane="bottomRight" state="frozen"/>'
                    '<selection pane="topRight" activeCell="B1" sqref="B1"/>'
                    '<selection pane="bottomLeft" activeCell="A%d" sqref="A%d"/>'
                    '<selection pane="bottomRight" activeCell="%s" sqref="%s"/>') % (
                        ysplit, ysplit + 1, ysplit + 1, ysplit + 1, cell, cell)
        for s in sel:
            xml = xml.replace(s, "", 1)
        xml = xml.replace(pane.group(0), new_pane, 1)
    # xsplit only (FormulaTest): column A already frozen

    # anchor row: first visible unfrozen row
    default_h, hts, hidden = row_heights(xml)
    anchor_row = ysplit + 1
    while anchor_row in hidden:
        anchor_row += 1
    y0 = int(sum(hts.get(r, default_h) for r in range(1, anchor_row)) * PT)
    x0 = BAR_X * PX

    # drawing part: reuse the sheet's own, or create one
    drawing_el = re.search(r'<drawing r:id="([^"]+)"/>', xml)
    if drawing_el:
        dpath_rel = rel_targets(rels)[drawing_el.group(1)][1]
        dpath = os.path.normpath(os.path.join(os.path.dirname(path), dpath_rel)).replace("\\", "/")
        dxml = rd(dpath)
    else:
        n = 1
        while os.path.exists(os.path.join(work, "xl/drawings/drawing%d.xml" % n)):
            n += 1
        dpath = "xl/drawings/drawing%d.xml" % n
        dxml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
                'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"></xdr:wsDr>')
        rid = next_rid(rels)
        rels = rels.replace("</Relationships>", '<Relationship Id="%s" Type="%s" Target="../drawings/%s"/>'
                            "</Relationships>" % (rid, REL_DRAWING, os.path.basename(dpath)))
        el = '<drawing r:id="%s"/>' % rid
        m = re.search(r"<(legacyDrawing|legacyDrawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b", xml)
        xml = xml[:m.start()] + el + xml[m.start():] if m else xml.replace("</worksheet>", el + "</worksheet>")
        ctypes = ctypes.replace("</Types>", '<Override PartName="/%s" ContentType="%s"/></Types>' % (dpath, CT_DRAWING))
    drels_path = os.path.join(os.path.dirname(dpath), "_rels", os.path.basename(dpath) + ".rels")
    drels = rd(drels_path) if os.path.exists(os.path.join(work, drels_path)) else \
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' \
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'

    def new_link(target, _state={}):
        nonlocal_drels = _state
        rid = next_rid(nonlocal_drels.get("x", drels))
        nonlocal_drels["x"] = nonlocal_drels.get("x", drels).replace(
            "</Relationships>", '<Relationship Id="%s" Type="%s" Target="#\'%s\'!A1"/></Relationships>' % (
                rid, REL_HYPERLINK, target))
        return rid

    link_state = {}
    ids_start = max([int(i) for i in re.findall(r'<xdr:cNvPr id="(\d+)"', dxml)], default=1) + 1
    ids = iter(range(ids_start, ids_start + 1000))
    grp, height = build_bar(sheet_name, ids, lambda t: new_link(t, link_state), x0, y0)
    drels = link_state["x"]

    # anchor: one cell, move but don't size, not printed
    row_idx = anchor_row - 1
    anchor = ('<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>%d</xdr:colOff><xdr:row>%d</xdr:row>'
              '<xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="%d" cy="%d"/>%s'
              '<xdr:clientData fPrintsWithSheet="0"/></xdr:oneCellAnchor>') % (
                  x0, row_idx, BAR_W * PX, height * PX, grp)
    if "xmlns:a=" not in dxml.split(">", 2)[1]:
        dxml = dxml.replace("<xdr:wsDr ", '<xdr:wsDr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ', 1)
    dxml = dxml.replace("</xdr:wsDr>", anchor + "</xdr:wsDr>", 1)

    wr(path, xml)
    wr(rels_path, rels)
    wr(dpath, dxml)
    wr(drels_path, drels)
    report.append("%-20s anchor row %3d  bar %3d px  drawing %s" % (sheet_name, anchor_row, height, dpath))

wr("[Content_Types].xml", ctypes)

# ------------------------------------------------------------- write out ----
if os.path.exists(OUT):
    os.remove(OUT)
written = set()
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    for name in order:                          # keep the original part order
        if not name.endswith("/"):
            z.write(os.path.join(work, name), name)
            written.add(name)
    for root, _, files in os.walk(work):
        for f in files:
            name = os.path.relpath(os.path.join(root, f), work).replace("\\", "/")
            if name not in written:
                z.write(os.path.join(root, f), name)
shutil.rmtree(work)
print("\n".join(report))
print("written", OUT)
