#!/usr/bin/env python3
"""Process the documentation screenshots of an Excel run (#616) into docs/assets.

tools/Take-SACCR-Screenshots.ps1 pictures the sheets of the documentation on the owner's
Windows runner, through the Excel quick check workflow in its mode ``screenshots``; the
pictures come in the run's artifact, folder ``screenshots``. This script reduces each
picture to a 256-colour PNG palette without dithering, the processing of docs/assets
(docs/assets/README.md, screenshots), and writes it as docs/assets/NAME.png. The pictures
are not otherwise edited: they are already cut to their rows and columns.

Usage::

    python tools/screenshots.py ARTIFACT/screenshots          # list what would be written
    python tools/screenshots.py ARTIFACT/screenshots --write  # write it

Processing needs Pillow.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
ASSETS = "docs/assets"
PREFIX = "SACCR_"  # the pictures of the documentation; any other file is left out
COLOURS = 256


def pictures(source: Path) -> list[Path]:
    """The pictures of the documentation in a folder, in name order."""
    return sorted(p for p in source.glob(PREFIX + "*.png") if p.is_file())


def palette(source: Path, target: Path) -> tuple[int, int]:
    """Write source as a 256-colour PNG palette without dithering; return its size."""
    from PIL import Image  # type: ignore[import-not-found, unused-ignore]

    with Image.open(source) as image:
        rgb = image.convert("RGB")
        reduced = rgb.quantize(colors=COLOURS, method=Image.Quantize.MEDIANCUT, dither=Image.Dither.NONE)
        target.parent.mkdir(parents=True, exist_ok=True)
        reduced.save(target, optimize=True)
        width, height = reduced.size
        return int(width), int(height)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("source", type=Path, help="the folder screenshots of the run's artifact")
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--write", action="store_true", help="write the pictures into docs/assets")
    args = parser.parse_args(argv)
    found = pictures(args.source)
    if not found:
        print(f"{args.source}: no {PREFIX}*.png picture", file=sys.stderr)
        return 1
    for picture in found:
        target = args.root / ASSETS / picture.name
        if args.write:
            width, height = palette(picture, target)
            print(f"wrote {ASSETS}/{picture.name} ({width} x {height}, {COLOURS} colours)")
        else:
            print(f"would write {ASSETS}/{picture.name}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
