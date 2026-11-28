#!/usr/bin/env python3
"""Draw the icons of the K-SACCR Ribbon tab (#803), or check that they are drawn.

Each icon is a navy rounded tile, the header fill 1F3864 of the workbook's tables
(docs/workbook/sheet-style.json), with a white glyph: the style the owner chose on
2026-11-26. src/workbook/customUI14.xml names an icon with image="ksiGlyphSize", where
Size is Large, for a large button, or Small, for a menu entry or a normal button.
Every icon is drawn at twice its size on the Ribbon (64 px for Large, 32 px for
Small), so that it stays sharp on a display scaled up to 200%; Office scales it to
the control. The files are src/workbook/ribbon-icons/ksiGlyphSize.png;
tools/ribbon_template.py embeds them in the template.

``--write`` draws every icon the Ribbon source names, and deletes the files it does
not name. Without it the script lists the icons and fails when one is missing or
left over. Drawing needs Pillow; checking does not. The drawings are compared by the
Ribbon gate byte for byte with the template, not with a new drawing, as another
version of Pillow may write other bytes.
"""
from __future__ import annotations

import argparse
import math
import re
import sys
from collections.abc import Callable
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
SOURCE = "src/workbook/customUI14.xml"
FOLDER = "src/workbook/ribbon-icons"
SIZES = {"Large": 64, "Small": 32}       # pixels: twice the 32 and 16 of the Ribbon
NAVY = (31, 56, 100, 255)                 # 1F3864
WHITE = (255, 255, 255, 255)
S = 512                                   # the drawing's own scale
W = 30                                    # stroke width at that scale
IMAGE = re.compile(r'\bimage="(ksi[A-Za-z]+?)(Large|Small)"')

Draw = Any


def referenced(root: Path) -> list[tuple[str, str]]:
    """The (glyph, size) pairs the Ribbon source names, in order, without repeats."""
    found = IMAGE.findall((root / SOURCE).read_text(encoding="utf-8"))
    return list(dict.fromkeys(found))


def file_name(glyph: str, size: str) -> str:
    return f"{glyph}{size}.png"


# ------------------------------------------------------------------- glyphs ----
def line(d: Draw, points: list[tuple[float, float]], width: int = W) -> None:
    """A white stroke with round joints and round ends."""
    d.line(points, fill=WHITE, width=width, joint="curve")
    for x, y in (points[0], points[-1]):
        d.ellipse((x - width / 2, y - width / 2, x + width / 2, y + width / 2), fill=WHITE)


def run(d: Draw, image: Any) -> None:
    d.polygon([(190, 140), (190, 372), (370, 256)], fill=WHITE)


def checklist(d: Draw, image: Any) -> None:
    for y in (170, 256, 342):
        line(d, [(130, y), (152, y + 22), (192, y - 22)], 24)
        d.rounded_rectangle((232, y - 14, 392, y + 14), radius=14, fill=WHITE)


def clear(d: Draw, image: Any) -> None:
    from PIL import Image, ImageDraw  # type: ignore[import-not-found, unused-ignore]
    block = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    b = ImageDraw.Draw(block)
    b.rounded_rectangle((176, 196, 356, 316), radius=22, outline=WHITE, width=W)
    b.rectangle((176 + W // 2, 196 + W // 2, 246, 316 - W // 2), fill=WHITE)
    image.alpha_composite(block.rotate(35, center=(266, 256), resample=Image.BICUBIC))
    line(d, [(140, 386), (372, 386)], 24)


def tray(d: Draw) -> None:
    line(d, [(132, 300), (132, 384), (380, 384), (380, 300)])


def import_file(d: Draw, image: Any) -> None:
    line(d, [(256, 120), (256, 300)])
    line(d, [(186, 236), (256, 306), (326, 236)])
    tray(d)


def export_file(d: Draw, image: Any) -> None:
    line(d, [(256, 300), (256, 120)])
    line(d, [(186, 190), (256, 120), (326, 190)])
    tray(d)


def staging(d: Draw, image: Any) -> None:
    d.rounded_rectangle((150, 116, 362, 396), radius=24, outline=WHITE, width=W)
    line(d, [(196, 262), (240, 306), (318, 210)], 34)


def commit(d: Draw, image: Any) -> None:
    d.ellipse((126, 126, 386, 386), outline=WHITE, width=W)
    line(d, [(190, 262), (238, 310), (326, 210)], 34)


def templates(d: Draw, image: Any) -> None:
    line(d, [(150, 112), (300, 112), (362, 174), (362, 400), (150, 400), (150, 112)])
    line(d, [(300, 112), (300, 174), (362, 174)], 24)
    for y in (240, 296, 352):
        d.rounded_rectangle((196, y - 11, 316, y + 11), radius=11, fill=WHITE)


def report(d: Draw, image: Any) -> None:
    d.rounded_rectangle((140, 112, 372, 400), radius=24, outline=WHITE, width=W)
    for x, h in ((196, 70), (256, 130), (316, 100)):
        d.rounded_rectangle((x - 20, 352 - h, x + 20, 352), radius=8, fill=WHITE)
    d.rounded_rectangle((186, 160, 326, 182), radius=11, fill=WHITE)


def restore(d: Draw, image: Any) -> None:
    d.arc((136, 136, 376, 376), start=-200, end=110, fill=WHITE, width=W)
    angle = math.radians(-200)
    x, y = 256 + 120 * math.cos(angle), 256 + 120 * math.sin(angle)
    d.polygon([(x - 56, y - 6), (x + 44, y - 34), (x + 6, y + 60)], fill=WHITE)


def home(d: Draw, image: Any) -> None:
    d.polygon([(256, 118), (398, 250), (114, 250)], fill=WHITE)
    d.rectangle((160, 240, 352, 394), fill=WHITE)
    d.rectangle((228, 300, 284, 394), fill=NAVY)


def navigation_bar(d: Draw, image: Any) -> None:
    d.rounded_rectangle((118, 140, 394, 372), radius=26, outline=WHITE, width=W)
    d.rounded_rectangle((118, 140, 218, 372), radius=26, fill=WHITE)
    d.rectangle((190, 140, 218, 372), fill=WHITE)
    for y in (196, 246, 296):
        d.rounded_rectangle((140, y - 9, 196, y + 9), radius=9, fill=NAVY)


def table(d: Draw, image: Any) -> None:
    d.rounded_rectangle((116, 136, 396, 376), radius=24, outline=WHITE, width=W)
    d.rounded_rectangle((116, 136, 396, 206), radius=24, fill=WHITE)
    d.rectangle((116, 180, 396, 206), fill=WHITE)
    line(d, [(126, 290), (386, 290)], 22)
    line(d, [(212, 206), (212, 366)], 22)
    line(d, [(304, 206), (304, 366)], 22)


def gear(d: Draw, image: Any) -> None:
    for k in range(8):
        a = math.radians(k * 45)
        cx, cy = 256 + 128 * math.cos(a), 256 + 128 * math.sin(a)
        dx, dy = 34 * math.cos(a + math.pi / 2), 34 * math.sin(a + math.pi / 2)
        ex, ey = 30 * math.cos(a), 30 * math.sin(a)
        d.polygon([(cx - dx - ex, cy - dy - ey), (cx + dx - ex, cy + dy - ey),
                   (cx + dx + ex, cy + dy + ey), (cx - dx + ex, cy - dy + ey)], fill=WHITE)
    d.ellipse((150, 150, 362, 362), fill=WHITE)
    d.ellipse((214, 214, 298, 298), fill=NAVY)


def chart(d: Draw, image: Any) -> None:
    for x, h in ((176, 110), (256, 210), (336, 160)):
        d.rounded_rectangle((x - 30, 368 - h, x + 30, 368), radius=10, fill=WHITE)
    line(d, [(120, 392), (392, 392)], 24)


def calculator(d: Draw, image: Any) -> None:
    d.rounded_rectangle((156, 110, 356, 402), radius=28, outline=WHITE, width=W)
    d.rounded_rectangle((196, 152, 316, 206), radius=10, fill=WHITE)
    for row in range(3):
        for column in range(2):
            x, y = 222 + column * 68, 256 + row * 50
            d.ellipse((x - 18, y - 18, x + 18, y + 18), fill=WHITE)


def bank(d: Draw, image: Any) -> None:
    d.polygon([(256, 112), (400, 200), (112, 200)], fill=WHITE)
    for x in (164, 230, 296, 362):
        d.rounded_rectangle((x - 18 - 4, 222, x + 18 - 4, 352), radius=6, fill=WHITE)
    d.rounded_rectangle((112, 364, 400, 396), radius=8, fill=WHITE)


def compare(d: Draw, image: Any) -> None:
    line(d, [(136, 200), (360, 200)])
    line(d, [(306, 146), (366, 200), (306, 254)])
    line(d, [(376, 316), (152, 316)])
    line(d, [(206, 262), (146, 316), (206, 370)])


def book(d: Draw, image: Any) -> None:
    line(d, [(256, 176), (256, 384)], 24)
    line(d, [(256, 176), (200, 146), (120, 146), (120, 354), (200, 354), (256, 384)])
    line(d, [(256, 176), (312, 146), (392, 146), (392, 354), (312, 354), (256, 384)])


def info(d: Draw, image: Any) -> None:
    d.ellipse((120, 120, 392, 392), outline=WHITE, width=W)
    d.ellipse((236, 166, 276, 206), fill=WHITE)
    d.rounded_rectangle((236, 230, 276, 344), radius=16, fill=WHITE)


def specification(d: Draw, image: Any) -> None:
    line(d, [(196, 176), (124, 256), (196, 336)])
    line(d, [(316, 176), (388, 256), (316, 336)])
    line(d, [(284, 150), (228, 362)], 26)


def card(d: Draw, box: tuple[float, float, float, float], solid: bool = False) -> None:
    """A trade: a card with two lines; solid in front of the cards behind it."""
    left, top, right, bottom = box
    d.rounded_rectangle(box, radius=22, fill=NAVY, outline=WHITE, width=26)
    inner = (left + 40, right - 40)
    for k, y in enumerate((top + (bottom - top) * 0.40, top + (bottom - top) * 0.66)):
        d.rounded_rectangle((inner[0], y - 11, inner[1] - (60 if k else 0), y + 11), radius=11, fill=WHITE)


def trade_cards(d: Draw, top: float) -> None:
    card(d, (136, top, 376, top + (392 - top)))


def netting_cards(d: Draw, top: float) -> None:
    height = 392 - top
    card(d, (196, top, 400, top + height * 0.70))
    card(d, (164, top + height * 0.15, 368, top + height * 0.85))
    card(d, (132, top + height * 0.30, 336, 392))


def roof(d: Draw) -> None:
    d.polygon([(256, 100), (404, 176), (108, 176)], fill=WHITE)


def trades(d: Draw, image: Any) -> None:
    trade_cards(d, 140)


def netting_sets(d: Draw, image: Any) -> None:
    netting_cards(d, 120)


def bank_trades(d: Draw, image: Any) -> None:
    roof(d)
    trade_cards(d, 206)


def bank_netting_sets(d: Draw, image: Any) -> None:
    roof(d)
    netting_cards(d, 200)


def counterparties(d: Draw, image: Any) -> None:
    """Two people, one behind the other: the counterparties (#831)."""
    d.ellipse((268, 118, 372, 222), fill=WHITE)
    d.chord((222, 236, 418, 432), 180, 360, fill=WHITE)
    d.ellipse((128, 158, 244, 274), fill=NAVY, outline=WHITE, width=W)
    d.chord((80, 286, 292, 498), 180, 360, fill=NAVY, outline=WHITE, width=W)
    d.line((80, 392, 292, 392), fill=WHITE, width=W)


GLYPHS: dict[str, Callable[[Draw, Any], None]] = {
    "ksiRun": run, "ksiChecklist": checklist, "ksiClear": clear, "ksiImport": import_file,
    "ksiStaging": staging, "ksiCommit": commit, "ksiTemplates": templates, "ksiEvidence": export_file,
    "ksiReport": report, "ksiRestore": restore, "ksiHome": home, "ksiNavigationBar": navigation_bar,
    "ksiTable": table, "ksiGear": gear, "ksiChart": chart, "ksiCalculator": calculator, "ksiBank": bank,
    "ksiCompare": compare, "ksiBook": book, "ksiInfo": info, "ksiSpecification": specification,
    "ksiTrades": trades, "ksiNettingSets": netting_sets, "ksiBankTrades": bank_trades,
    "ksiBankNettingSets": bank_netting_sets, "ksiCounterparties": counterparties,
}


def draw(glyph: str, pixels: int) -> bytes:
    """The PNG of one icon: the tile and the glyph at the drawing's scale, then reduced."""
    import io

    from PIL import Image, ImageDraw  # type: ignore[import-not-found, unused-ignore]
    image = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(image)
    d.rounded_rectangle((16, 16, S - 16, S - 16), radius=96, fill=NAVY)
    GLYPHS[glyph](d, image)
    out = io.BytesIO()
    image.resize((pixels, pixels), Image.LANCZOS).save(out, format="PNG", optimize=True)
    return out.getvalue()


# ------------------------------------------------------------------- check -----
def problems(root: Path) -> list[str]:
    found: list[str] = []
    wanted = referenced(root)
    for glyph, size in wanted:
        if glyph not in GLYPHS:
            found.append(f"{SOURCE}: image {glyph}{size} has no drawing in tools/ribbon_icons.py")
        if not (root / FOLDER / file_name(glyph, size)).is_file():
            found.append(f"{FOLDER}/{file_name(glyph, size)} is missing; run tools/ribbon_icons.py --write")
    names = {file_name(glyph, size) for glyph, size in wanted}
    folder = root / FOLDER
    if folder.is_dir():
        for path in sorted(folder.iterdir()):
            if path.name not in names:
                found.append(f"{FOLDER}/{path.name} is not named by {SOURCE}")
    return found


def write(root: Path) -> list[str]:
    folder = root / FOLDER
    folder.mkdir(parents=True, exist_ok=True)
    wanted = referenced(root)
    names = {file_name(glyph, size) for glyph, size in wanted}
    for path in folder.iterdir():
        if path.name not in names:
            path.unlink()
    written = []
    for glyph, size in wanted:
        (folder / file_name(glyph, size)).write_bytes(draw(glyph, SIZES[size]))
        written.append(file_name(glyph, size))
    return written


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--write", action="store_true", help="draw the icons the Ribbon source names")
    args = parser.parse_args()
    root = args.root.resolve()
    if args.write:
        print("drew " + ", ".join(write(root)))
    found = problems(root)
    for problem in found:
        print(problem, file=sys.stderr)
    if not found:
        print(f"{FOLDER}: {len(referenced(root))} icons, as {SOURCE} names them")
    return 1 if found else 0


if __name__ == "__main__":
    sys.exit(main())
