#!/usr/bin/env python3
"""Compare the VBA exported from a built workbook with the repository source.

Part of the source round trip (#34): a workbook built from the template and
the source of one commit, then exported with tools/Export-SACCR-VBA.ps1, must
hold exactly the components that tools/vba-projects.json lists for its mode,
each with the same code and attributes as the source of that commit. Line
endings are ignored, and so are trailing blank lines and the
VERSION/BEGIN/END block of a class file; everything else must match. The
source is read from Git at the candidate commit, not from the working tree.
The tool never runs Excel. See docs/EXCEL_EVIDENCE.md.
"""
from __future__ import annotations

import argparse
import difflib
import hashlib
import json
import re
import sys
from dataclasses import dataclass
from pathlib import Path

from _gatelib import git_bytes

MANIFEST = "tools/vba-projects.json"
TEMPLATE = "src/workbook/SACCR_Template.xlsx"
VBA_ENCODING = "cp1252"
SUFFIXES = (".bas", ".cls", ".frm")
VB_NAME = re.compile(r'^Attribute VB_Name = "([^"]+)"$')
# Lines the VBE writes before the code of a class or form: the VERSION line
# ("VERSION 1.0 CLASS" for a class, "VERSION 5.00" for a form), the
# BEGIN ... END property block (title case and nested in a form) and the
# Attribute lines.
CLASS_HEADER = re.compile(r"^(VERSION \d+\.\d+( CLASS)?|(?i:begin)(\s.*)?|(?i:end)|\s+\S.*)$")


@dataclass(frozen=True)
class Component:
    name: str
    attributes: tuple[str, ...]
    body: tuple[str, ...]


def parse(raw: bytes, label: str) -> Component:
    """Split a VBA file into its VB_Name, its Attribute lines and its code."""
    lines = raw.decode(VBA_ENCODING).replace("\r\n", "\n").replace("\r", "\n").split("\n")
    index = 0
    attributes: list[str] = []
    while index < len(lines):
        line = lines[index]
        if line.startswith("Attribute "):
            attributes.append(line)
        elif not CLASS_HEADER.match(line) or attributes:
            break
        index += 1
    body = lines[index:]
    while body and body[-1] == "":
        body.pop()
    names = [m.group(1) for m in map(VB_NAME.match, attributes) if m]
    if len(names) != 1:
        raise ValueError(f"{label}: expected one Attribute VB_Name line, found {len(names)}")
    return Component(names[0], tuple(attributes), tuple(body))


def committed(root: Path, sha: str, path: str) -> bytes:
    completed = git_bytes(root, "show", f"{sha}:{path}")
    if completed.returncode != 0:
        raise ValueError(f"{path}: not found at {sha}")
    return completed.stdout


def expected_components(root: Path, sha: str, mode: str) -> dict[str, tuple[str, Component]]:
    """The components of ``mode`` at ``sha``, by lower-case VB_Name."""
    manifest = json.loads(committed(root, sha, MANIFEST).decode("utf-8"))
    configuration = manifest["configurations"].get(mode)
    if configuration is None:
        raise ValueError(f"{MANIFEST}: no configuration '{mode}'")
    expected: dict[str, tuple[str, Component]] = {}
    for path in [*configuration["modules"], *manifest["document_modules"]]:
        component = parse(committed(root, sha, path), path)
        expected[component.name.lower()] = (path, component)
    return expected


def exported_components(folder: Path) -> dict[str, tuple[str, Component]]:
    exported: dict[str, tuple[str, Component]] = {}
    for path in sorted(folder.iterdir()):
        if path.suffix.lower() not in SUFFIXES:
            continue
        component = parse(path.read_bytes(), path.name)
        key = component.name.lower()
        if key in exported:
            raise ValueError(f"{path.name}: VB_Name {component.name} exported twice")
        exported[key] = (path.name, component)
    return exported


def first_difference(source: tuple[str, ...], export: tuple[str, ...]) -> str:
    for number, (left, right) in enumerate(zip(source, export), start=1):
        if left != right:
            return f"line {number}: source {left!r}, export {right!r}"
    shorter = min(len(source), len(export))
    return f"line {shorter + 1}: source has {len(source)} lines, export {len(export)}"


def compare(expected: dict[str, tuple[str, Component]],
            exported: dict[str, tuple[str, Component]]) -> list[str]:
    """The differences between the source and the export; empty when they match."""
    findings = [f"missing from the export: {expected[key][1].name} ({expected[key][0]})"
                for key in sorted(set(expected) - set(exported))]
    findings += [f"not in the source for this mode: {exported[key][1].name} ({exported[key][0]})"
                 for key in sorted(set(exported) - set(expected))]
    for key in sorted(set(expected) & set(exported)):
        path, source = expected[key]
        _, export = exported[key]
        if source.name != export.name:
            findings.append(f"{path}: VB_Name {source.name} exported as {export.name}")
        if source.attributes != export.attributes:
            changed = [line for line in difflib.unified_diff(
                source.attributes, export.attributes, lineterm="", n=0)
                if line[:1] in "+-" and line[:3] not in ("+++", "---")]
            findings.append(f"{path}: attributes differ: {'; '.join(changed)}")
        if source.body != export.body:
            findings.append(f"{path}: code differs at {first_difference(source.body, export.body)}")
    return findings


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def main(arguments: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parent.parent)
    parser.add_argument("--export", type=Path, required=True, help="folder written by Export-SACCR-VBA.ps1")
    parser.add_argument("--mode", choices=("production", "development"), required=True)
    parser.add_argument("--candidate-sha", default="HEAD", help="commit the workbook was built from")
    parser.add_argument("--workbook", type=Path, help="built .xlsm, whose SHA-256 is printed")
    options = parser.parse_args(arguments)
    try:
        resolved = git_bytes(options.root, "rev-parse", "--verify", f"{options.candidate_sha}^{{commit}}")
        if resolved.returncode != 0:
            raise ValueError(f"unknown commit: {options.candidate_sha}")
        sha = resolved.stdout.decode().strip()
        if not options.export.is_dir():
            raise ValueError(f"export folder not found: {options.export}")
        expected = expected_components(options.root, sha, options.mode)
        exported = exported_components(options.export)
        findings = compare(expected, exported)
        template = sha256(committed(options.root, sha, TEMPLATE))
    except ValueError as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2
    print(f"Candidate: {sha}")
    print(f"Mode:      {options.mode} ({len(expected)} components expected, {len(exported)} exported)")
    print(f"Template:  {TEMPLATE} sha256 {template}")
    if options.workbook is not None:
        print(f"Workbook:  {options.workbook.name} sha256 {sha256(options.workbook.read_bytes())}")
    for finding in findings:
        print(f"  FAIL {finding}")
    print(f"RESULT={'PASS' if not findings else 'FAIL'}; components={len(expected)}; differences={len(findings)}")
    return 0 if not findings else 1


if __name__ == "__main__":
    sys.exit(main())
