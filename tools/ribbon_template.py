#!/usr/bin/env python3
"""Embed the Ribbon source in the workbook template, or check that it is (#214).

The macro-free template src/workbook/SACCR_Template.xlsx carries the RibbonX
source src/workbook/customUI14.xml as its customUI/customUI14.xml part, with
one package relationship of the Office 2010 extensibility type. Excel keeps
that part through the build's SaveAs .xlsm and through ordinary saves (#213
spike), so every workbook built from the template has the Ribbon.

``--write`` rewrites the template with the part and the relationship, leaving
every other entry as it is. Without it the script checks the template and
fails when the part is missing, differs from the source, or has no
relationship, or when the package holds another customUI part.
"""
from __future__ import annotations

import argparse
import os
import re
import sys
import tempfile
import zipfile
from pathlib import Path

TEMPLATE = "src/workbook/SACCR_Template.xlsx"
SOURCE = "src/workbook/customUI14.xml"
PART = "customUI/customUI14.xml"
RELATIONSHIPS = "_rels/.rels"
RELATIONSHIP_TYPE = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
RELATIONSHIP_ID = "rIdKsRibbon"
RELATIONSHIP = f'<Relationship Id="{RELATIONSHIP_ID}" Type="{RELATIONSHIP_TYPE}" Target="{PART}"/>'
FIXED_TIME = (1980, 1, 1, 0, 0, 0)
RELATIONSHIP_ELEMENT = re.compile(r"<Relationship\b[^>]*/>")


def ribbon_relationships(rels: str) -> list[str]:
    """The relationship elements of the package that point at a customUI part."""
    return [element for element in RELATIONSHIP_ELEMENT.findall(rels) if "/ui/extensibility" in element]


def problems(root: Path) -> list[str]:
    template, source = root / TEMPLATE, root / SOURCE
    if not source.is_file():
        return [f"{SOURCE} is missing"]
    with zipfile.ZipFile(template) as package:
        names = package.namelist()
        found: list[str] = []
        others = [n for n in names if n.lower().startswith("customui/") and n != PART]
        found += [f"{TEMPLATE}: unexpected customUI part {n}" for n in others]
        if PART not in names:
            return found + [f"{TEMPLATE}: no {PART} part; run tools/ribbon_template.py --write"]
        if package.read(PART) != source.read_bytes():
            found.append(f"{TEMPLATE}: {PART} differs from {SOURCE}; run tools/ribbon_template.py --write")
        rels = package.read(RELATIONSHIPS).decode("utf-8")
    if ribbon_relationships(rels) != [RELATIONSHIP]:
        found.append(f"{TEMPLATE}: {RELATIONSHIPS} must hold exactly the relationship {RELATIONSHIP}")
    return found


def write(root: Path) -> None:
    template, source = root / TEMPLATE, root / SOURCE
    data = source.read_bytes()
    with zipfile.ZipFile(template) as package:
        infos = package.infolist()
        entries = {info.filename: package.read(info.filename) for info in infos}
    rels = entries[RELATIONSHIPS].decode("utf-8")
    for element in ribbon_relationships(rels):
        rels = rels.replace(element, "")
    if rels.count("</Relationships>") != 1:
        raise SystemExit(f"{TEMPLATE}: {RELATIONSHIPS} has no single closing element")
    entries[RELATIONSHIPS] = rels.replace("</Relationships>", RELATIONSHIP + "</Relationships>").encode("utf-8")
    descriptor, temporary = tempfile.mkstemp(dir=template.parent, suffix=".xlsx")
    os.close(descriptor)
    try:
        with zipfile.ZipFile(temporary, "w") as output:
            for info in infos:
                if info.filename.lower().startswith("customui/"):
                    continue
                output.writestr(info, entries[info.filename], compress_type=info.compress_type)
            part = zipfile.ZipInfo(PART, date_time=FIXED_TIME)
            part.compress_type = zipfile.ZIP_DEFLATED
            output.writestr(part, data)
        os.replace(temporary, template)
    finally:
        if os.path.exists(temporary):
            os.remove(temporary)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--write", action="store_true", help="embed the source in the template")
    args = parser.parse_args()
    root = args.root.resolve()
    if args.write:
        write(root)
        print(f"wrote {PART} into {TEMPLATE}")
    found = problems(root)
    for problem in found:
        print(problem, file=sys.stderr)
    if not found:
        print(f"{TEMPLATE}: {PART} matches {SOURCE}")
    return 1 if found else 0


if __name__ == "__main__":
    sys.exit(main())
