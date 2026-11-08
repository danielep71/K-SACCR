"""Gate resolved VBA layer references, with reviewed #220 edge exceptions."""
from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

from _gatelib import parse_report_args, run_gate
from check_vba_calls import Module, matching_paren, parse_module, parse_params, statements_of
from check_vba_conditionals import logical_units, reachable_sources
from check_vba_identifiers import HOST_OBJECTS, project_sources, references
from vba_lexer import Tok

POLICY = "tools/vba-dependency-exceptions.json"
ALLOWED = {"core": {"core"}, "modules": {"modules", "core", "classes", "host"},
           "workbook": {"workbook", "modules", "classes", "host"},
           "classes": {"classes", "core"}, "forms": {"forms", "modules", "classes", "host"},
           "tests": {"tests", "core", "modules", "workbook", "classes", "forms", "host"},
           "examples": {"examples", "modules", "host"}}


def layer(path: str) -> str:
    parts = Path(path).parts
    return parts[1] if parts[0] == "src" else parts[0]


def symbols(modules: list[Module], *, types_only: bool = False) -> dict[str, set[str]]:
    owners: dict[str, set[str]] = {}
    for module in modules:
        owners.setdefault(module.name.casefold(), set()).add(module.path)
        for name, member in module.members.items():
            if (member.visibility != "private"
                    and (module.kind == "standard" or member.kind in {"enum", "enum_member"})
                    and (not types_only or member.kind in {"type", "enum"})):
                owners.setdefault(name, set()).add(module.path)
    return owners


def event_parameters(tokens: list[Tok]) -> set[str]:
    if not any(token.is_kw("event") for token in tokens[:3]):
        return set()
    opening = next((i for i, token in enumerate(tokens) if token.is_op("(")), None)
    if opening is None:
        return set()
    closing = matching_paren(tokens, opening)
    if closing < 0:
        raise ValueError("Unclosed event parameter list")
    return {p.name.casefold() for p in parse_params(tokens[opening + 1:closing])}


def module_edges(module: Module, source: str, owners: dict[str, set[str]],
                 type_owners: dict[str, set[str]]) -> list[dict[str, Any]]:
    edges = []
    procedures = module.procs + [member.accessors["declare"] for member in module.members.values()
                                 if "declare" in member.accessors]
    local_by_line = {line: set(proc.locals) | {p.name.casefold() for p in proc.params} | {proc.name.casefold()}
                     for proc in procedures for line in [proc.line, *(line for line, _ in proc.statements)]}
    for line, _, kind, code in logical_units(source.splitlines(), comment_aware=True):
        if kind == "directive":
            continue
        for tokens in statements_of(code, comment_aware=True):
            if tokens[0].is_kw("attribute", "option", "rem"):
                continue
            local_names = local_by_line.get(line, set()) | event_parameters(tokens)
            field_name = (tokens[0].low if tokens[0].is_id() and len(tokens) > 1
                          and (tokens[1].is_kw("as") or
                               (tokens[1].is_op("(") and any(t.is_kw("as") for t in tokens)))
                          else None)
            type_tokens = {id(tokens[i + 1]) for i, t in enumerate(tokens[:-1])
                          if t.is_kw("as", "new")}
            for token in references(tokens, include_host_statements=True):
                if field_name is not None and token is tokens[0]:
                    continue
                is_type = id(token) in type_tokens
                own = module.members.get(token.low)
                if token.kind != "host" and ((not is_type and token.low in local_names) or
                        (own is not None and (not is_type or own.kind in {"type", "enum"}))):
                    continue
                targets = ({"host:" + token.low} if token.kind == "host" else
                           (type_owners if is_type else owners).get(token.low, set()))
                if not targets and token.low in HOST_OBJECTS:
                    targets = {"host:" + token.low}
                for target in targets:
                    target_layer = "host" if target.startswith("host:") else layer(target)
                    if target != module.path and target_layer not in ALLOWED.get(layer(module.path), set()):
                        edges.append({"path": module.path, "target": target, "line": line,
                                      "identifier": token.text})
    return edges


def analyze(sources: dict[str, str], exceptions: list[dict[str, str]]) -> dict[str, Any]:
    accepted = {(item["source"], item["target"]) for item in exceptions}
    failures: list[dict[str, Any]] = []
    variants: dict[str, dict[str, str]] = {}
    for path, text in sources.items():
        environments, errors = reachable_sources(path, text, comment_aware=True)
        failures.extend(errors)
        for env, source in environments.items():
            variants.setdefault(env, {})[path] = source
    violations: list[dict[str, Any]] = []
    exemptions: list[dict[str, Any]] = []
    for env, files in variants.items():
        modules = [parse_module(path, text, comment_aware=True) for path, text in files.items()]
        owners = symbols(modules)
        type_owners = symbols(modules, types_only=True)
        for module in modules:
            failures.extend(module.failures)
            for edge in module_edges(module, files[module.path], owners, type_owners):
                target = exemptions if (edge["path"], edge["target"]) in accepted else violations
                target.append({**edge, "environment": env})
    return {"status": "fail" if violations or failures else "pass", "findings": violations,
            "accepted_exceptions": exemptions, "analysis_failures": failures}


def read_exceptions(root: Path) -> list[dict[str, str]]:
    policy = json.loads((root / POLICY).read_text(encoding="utf-8"))
    if policy.get("schema_version") != 1 or not isinstance(policy.get("exceptions"), list):
        raise ValueError("Invalid dependency exception schema")
    result: list[dict[str, str]] = []
    seen: set[tuple[str, str]] = set()
    for entry in policy["exceptions"]:
        if not isinstance(entry, dict) or set(entry) != {"source", "target", "issue", "reason"}:
            raise ValueError("Each exception needs source, target, issue and reason")
        if any(not isinstance(v, str) or not v.strip() for v in entry.values()):
            raise ValueError("Exception fields must be nonempty strings")
        pair = (entry["source"], entry["target"])
        if pair in seen or any(c in entry["source"] + entry["target"] for c in "*?[]"):
            raise ValueError("Duplicate/wildcard exceptions are not allowed")
        if entry["issue"] != "#220":
            raise ValueError("Exceptions must refer to the reviewed #220 deviations")
        seen.add(pair)
        result.append(entry)
    return result


def run_check(root: Path) -> dict[str, Any]:
    exceptions = read_exceptions(root)
    projects = {name: analyze(files, exceptions) for name, files in project_sources(root).items()}
    return {"schema_version": 1, "status": "fail" if any(p["status"] == "fail" for p in projects.values()) else "pass",
            "projects": projects,
            "scope": "Resolved static roots and type references; dynamic targets and external member bodies require review."}


def markdown_report(report: dict[str, Any]) -> str:
    lines = ["## VBA dependency direction", "", f"Status: {report['status']}"]
    for name, project in report["projects"].items():
        lines.append(f"- {name}: {len(project['findings'])} violations; "
                     f"{len(project['accepted_exceptions'])} accepted references; "
                     f"{len(project['analysis_failures'])} analysis failures")
        for finding in project["findings"] + project["analysis_failures"]:
            lines.append("  " + json.dumps(finding, sort_keys=True))
    return "\n".join(lines) + "\n"


def main() -> int:
    options = parse_report_args(sys.argv[1:], description=__doc__)
    return run_gate(options, build=lambda: run_check(options.root), markdown=markdown_report,
                    errors=(OSError, UnicodeError, ValueError, RuntimeError, KeyError))


if __name__ == "__main__":
    raise SystemExit(main())
