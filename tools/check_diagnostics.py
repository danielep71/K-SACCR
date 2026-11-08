"""Check the stable Checks diagnostic registry and all engine emission sites."""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path
from typing import Any

from _gatelib import parse_report_args, run_gate
from vba_lexer import Tok, logical_lines, tokenize

CONFIG = "src/core/CORE_Config.bas"
REGISTRY = "docs/diagnostics.json"
EMITTERS = {"logmsg", "adderr", "addwarn", "addgrouperror"}
# These private wrappers forward only codes supplied by checked call sites.
FORWARDERS = {("adderr", "logmsg", "DiagnosticCode"),
              ("addwarn", "logmsg", "DiagnosticCode"),
              ("loggroupfindings", "logmsg", "CStr(Finding(0))")}


def check_sources(sources: dict[str, str], registry: dict[str, Any]) -> list[str]:
    findings: list[str] = []
    codes = registry["codes"]
    constants: dict[str, str] = {}
    for code, item in codes.items():
        name = item["constant"]
        if name in constants or not re.fullmatch(r"[A-Z][A-Z0-9]*(?:-[A-Z0-9]+)+", code):
            findings.append(f"duplicate constant or malformed code: {code}")
        if item["status"] not in {"active", "retired"} or not item["description"]:
            findings.append(f"invalid registry entry: {code}")
        constants[name] = code
    declared = dict(re.findall(r'Public Const (\w+)\s+As String = "([A-Z][A-Z0-9-]+)"',
                               sources[CONFIG], re.I))
    for name, code in constants.items():
        if declared.get(name) != code:
            findings.append(f"registry/CORE_Config mismatch: {name} = {code}")
    for name in declared:
        if name.startswith("DIAG_") and name not in constants:
            findings.append(f"unregistered constant: {name}")
    source_findings, used = check_emissions(sources, constants, codes)
    findings.extend(source_findings)
    for name, code in constants.items():
        if codes[code]["status"] == "active" and name not in used:
            findings.append(f"unused active code: {name}")
    return findings


def check_emissions(sources: dict[str, str], constants: dict[str, str],
                    codes: dict[str, Any]) -> tuple[list[str], set[str]]:
    findings: list[str] = []
    used: set[str] = set()
    for path, source in sources.items():
        if path == CONFIG:
            continue
        procedure = ""
        for start, _end, code in logical_lines(source.splitlines()):
            declaration = re.match(r"(?:Public|Private)\s+(?:Sub|Function)\s+(\w+)", code, re.I)
            if declaration:
                procedure = declaration[1].lower()
                continue
            tokens = tokenize(code)
            for token in tokens:
                if token.kind == "id" and token.text in constants:
                    used.add(token.text)
                    if codes[constants[token.text]]["status"] == "retired":
                        findings.append(f"{path}:{start}: retired code {token.text}")
            # Tokens, rather than raw regex, exclude comments and quoted messages.
            for index, token in enumerate(tokens):
                if token.kind != "id" or token.low not in EMITTERS:
                    continue
                argument = first_argument(tokens[index + 1:])
                if argument not in constants and (procedure, token.low, argument) not in FORWARDERS:
                    findings.append(f"{path}:{start}: {token.text} missing/unknown code: {argument}")
    return findings, used


def first_argument(tokens: list[Tok]) -> str:
    depth = 0
    first: list[str] = []
    for arg in tokens:
        if arg.text == "," and depth == 0:
            break
        first.append(arg.text)
        depth += (arg.text == "(") - (arg.text == ")")
    return "".join(first)


def run_check(root: Path) -> dict[str, Any]:
    sources = {path.relative_to(root).as_posix(): path.read_bytes().decode("cp1252")
               for path in (root / "src").rglob("*")
               if path.suffix.lower() in {".bas", ".cls", ".frm"}}
    registry = json.loads((root / REGISTRY).read_text(encoding="utf-8"))
    findings = check_sources(sources, registry)
    return {"schema_version": 1, "status": "fail" if findings else "pass", "findings": findings}


def main() -> int:
    options = parse_report_args(sys.argv[1:], description=__doc__)
    return run_gate(options, build=lambda: run_check(options.root),
                    markdown=lambda report: "Diagnostics: " + report["status"] + "\n" +
                    "\n".join(report["findings"]) + "\n",
                    errors=(OSError, ValueError, KeyError))


if __name__ == "__main__":
    raise SystemExit(main())
