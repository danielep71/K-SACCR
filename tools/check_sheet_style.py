#!/usr/bin/env python3
"""Check the sheet style file (docs/workbook/sheet-style.json) and the tabs
of the workbook template against it, without running Excel (#706).

Every role of the editability matrix has one tab color in the style file;
every sheet of the template has the tab color of its role and its gridlines
turned off. What Excel displays after a run or Clear outputs (fonts, vertical
alignment, dates, cells showing ####) is checked by the sheet review of the
quick Excel run, tools/Verify-SACCR-Branch.ps1 -Quick."""
from __future__ import annotations

import json
import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import Any

from _gatelib import parse_report_args, run_gate
from check_workbook_protection import MATRIX, NS, TEMPLATE, sheet_parts

STYLE = "docs/workbook/sheet-style.json"
KEYS = {"schema_version", "description", "font", "vertical_alignment", "date_format", "gridlines", "tab_colors"}
COLOR = re.compile(r"^[0-9A-F]{6}$")


def check_style(style: Any, roles: set[str]) -> list[str]:
    """Findings on the style file itself: its keys, and one color per role."""
    if not isinstance(style, dict):
        return [f"{STYLE}: not a JSON object"]
    findings = []
    if set(style) != KEYS:
        findings.append(f"{STYLE}: keys must be {sorted(KEYS)}, found {sorted(style)}")
    if style.get("schema_version") != 1:
        findings.append(f"{STYLE}: schema_version must be 1")
    if style.get("gridlines") is not False:
        findings.append(f"{STYLE}: gridlines must be false")
    for key in ("font", "vertical_alignment", "date_format"):
        if not isinstance(style.get(key), str) or not style.get(key):
            findings.append(f"{STYLE}: {key} must be a non-empty string")
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
    return findings


def run_check(root: Path) -> dict[str, Any]:
    matrix = json.loads((root / MATRIX).read_text(encoding="utf-8"))
    style = json.loads((root / STYLE).read_text(encoding="utf-8"))
    sheets = matrix.get("sheets", {})
    roles = {spec.get("role") for spec in sheets.values() if isinstance(spec, dict)}
    findings = check_style(style, {role for role in roles if isinstance(role, str)})
    colors = style.get("tab_colors", {}) if isinstance(style, dict) else {}
    with zipfile.ZipFile(root / TEMPLATE) as archive:
        parts = sheet_parts(archive)
        for name, part in parts.items():
            role = sheets.get(name, {}).get("role") if isinstance(sheets.get(name), dict) else None
            expected = colors.get(role) if isinstance(colors, dict) and isinstance(role, str) else None
            findings += sheet_findings(name, archive.read(part), expected)
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
