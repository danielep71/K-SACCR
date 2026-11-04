"""Write the host.json of an Excel evidence bundle from the logs a certification run left.

tools/Certify-SACCR.ps1 runs the harness and every suite in Excel and writes their
logs, session.log and session.json into the bundle (#535). This script reads them,
with the policy and the source inventory of the candidate, and writes host.json as
docs/EXCEL_EVIDENCE.md describes the record; it then validates the bundle with
check_excel_evidence.py and returns its exit code. It never runs Excel and records
only what the logs and session.json show: a suite without a log is NOT_RUN, a log
without a RESULT line is TIMEOUT, and the counts are the ones the RESULT lines print.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from pathlib import Path
from typing import Any

import check_excel_evidence as evidence

DESCRIPTION = "Write host.json for an Excel evidence bundle and validate it."
SESSION_KEYS = ("operator started_at finished_at excel_version excel_build office_bitness os "
                "os_architecture runtime references macro_policy vba_project_access import "
                "compile workbook_closed")
HARNESS_RESULT = re.compile(r"^RESULT=(PASS|FAIL); completeness=(\w+); cases=(\d+); "
                            r"assertions=(\d+); failures=(\d+); cleanup=(\w+)\s*$", re.MULTILINE)
SUITE_RESULT = re.compile(r"^RESULT=(PASS|FAIL); cases=(\d+); checks=(\d+); failures=(\d+)", re.MULTILINE)


def log_name(entry_point: str) -> str:
    """The log file of an entry point, as TEST_Suite.Certify names it."""
    return "harness.log" if entry_point == "TEST_Harness.RunTests" else entry_point.split(".")[0] + ".log"


def reference(bundle: Path, name: str) -> dict[str, str]:
    return {"path": name, "sha256": hashlib.sha256((bundle / name).read_bytes()).hexdigest()}


def read(bundle: Path, name: str) -> str | None:
    path = bundle / name
    return path.read_bytes().decode("utf-8-sig").replace("\r\n", "\n") if path.is_file() else None


def stage(status: str, detail: str, log: dict[str, str] | None) -> dict[str, Any]:
    return {"status": status, "detail": detail, "log": None if status == "NOT_RUN" else log}


def environment_line(log: str) -> dict[str, str]:
    values = re.findall(r"^ENVIRONMENT=(.*)$", log, re.MULTILINE)
    fields: dict[str, str] = {}
    for item in (values[0].strip().split("; ") if len(values) == 1 else []):
        key, _, value = item.partition("=")
        fields[key] = value
    return fields


def harness_record(policy: dict[str, Any], log: str | None) -> tuple[str, dict[str, Any] | None, str]:
    """Return the regression status the harness log supports, the harness object and its cleanup."""
    if log is None:
        return "NOT_RUN", None, "NOT_RUN"
    found = HARNESS_RESULT.findall(log)
    if len(found) != 1:
        return "TIMEOUT", None, "NOT_RUN"
    verdict, completeness, cases, assertions, failures, cleanup = found[0]
    clean = verdict == "PASS"
    errors = [{"case": case, "status": "PASS" if clean else "NOT_RUN",
               "detail": "Implied by the complete passing harness" if clean
               else "Not inferred: the harness did not pass"} for case in policy["expected_error_cases"]]
    harness = {"entry_point": policy["entry_point"], "cases": int(cases), "assertions": int(assertions),
               "failures": int(failures), "completeness": completeness, "expected_errors": errors}
    return verdict, harness, cleanup


def suite_record(bundle: Path, spec: dict[str, Any], ran: bool) -> dict[str, Any]:
    name = log_name(spec["entry_point"])
    log = read(bundle, name) if ran else None
    item: dict[str, Any] = {"entry_point": spec["entry_point"], "status": "NOT_RUN",
                            "cases": None, "checks": None, "failures": None, "log": None}
    if log is None:
        return item
    item["log"] = reference(bundle, name)
    found = SUITE_RESULT.findall(log)
    if len(found) != 1:
        item["status"] = "TIMEOUT"
        return item
    verdict, cases, checks, failures = found[0]
    item.update(status=verdict, cases=int(cases), checks=int(checks), failures=int(failures))
    return item


def build_record(root: Path, sha: str, bundle: Path) -> dict[str, Any]:
    """The evidence record the bundle's logs and session.json support."""
    evidence.require_commit(root, sha)
    policy = evidence.load_policy(root, sha)
    session = evidence.object_keys(json.loads((bundle / "session.json").read_text(encoding="utf-8-sig")),
                                   SESSION_KEYS, "session.json")
    session_log = reference(bundle, "session.log")
    imported = evidence.object_keys(session["import"], "status detail", "session import")
    compiled = evidence.object_keys(session["compile"], "status detail", "session compile")
    ran = imported["status"] == "PASS" and compiled["status"] == "PASS"
    harness_log = read(bundle, "harness.log") if ran else None
    harness_status, harness, cleanup = harness_record(policy, harness_log)
    suites = [suite_record(bundle, spec, harness_status != "NOT_RUN") for spec in policy["suites"]]
    environment = environment_line(harness_log) if harness_log else {}
    if harness_status in ("PASS", "FAIL"):
        failed = harness_status == "FAIL" or any(item["status"] == "FAIL" for item in suites)
        regression = stage("FAIL" if failed else "PASS",
                           "Certify-SACCR: harness and every suite run once each, on their own",
                           reference(bundle, "harness.log"))
    elif harness_status == "TIMEOUT":
        regression = stage("TIMEOUT", "The harness log has no RESULT line", reference(bundle, "harness.log"))
    else:
        regression = stage("NOT_RUN", "Not run: import or compile did not pass", None)
    closed = session["workbook_closed"] is True
    cleanup_status = "PASS" if closed and cleanup in ("PASS", "NOT_RUN") else "FAIL"
    cleanup_detail = (f"Harness cleanup {cleanup}; " if cleanup != "NOT_RUN" else "") + \
        ("workbook closed" if closed else "workbook not closed")
    summary = None
    if harness_status in ("PASS", "FAIL") and all(item["log"] for item in suites):
        text = read(bundle, "summary.log")
        if text is not None and re.search(r"^RESULT=", text, re.MULTILINE):
            summary = reference(bundle, "summary.log")
    return {
        "schema_version": 2, "repository": policy["repository"], "candidate_sha": sha,
        "execution": "manual", "availability_reason": None,
        "started_at": session["started_at"], "finished_at": session["finished_at"],
        "operator": session["operator"],
        "environment": {
            "excel_version": environment.get("version", session["excel_version"]),
            "excel_build": session["excel_build"],
            "office_bitness": environment.get("office", session["office_bitness"]),
            "os": session["os"], "os_architecture": session["os_architecture"],
            "runtime": environment.get("runtime", session["runtime"]),
            "references": session["references"], "macro_policy": session["macro_policy"],
            "vba_project_access": session["vba_project_access"], "trust_changes": False},
        "sources": evidence.source_inventory(root, sha),
        "stages": {
            "import": stage(imported["status"], imported["detail"], session_log),
            "compile": stage(compiled["status"] if imported["status"] == "PASS" else "NOT_RUN",
                             compiled["detail"], session_log),
            "regression": regression,
            "cleanup": stage(cleanup_status, cleanup_detail, session_log)},
        "harness": harness, "suites": suites, "summary": summary,
    }


def main(arguments: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=DESCRIPTION)
    parser.add_argument("--root", type=Path, default=Path.cwd())
    parser.add_argument("--candidate-sha", required=True)
    parser.add_argument("--bundle", type=Path, required=True,
                        help="the bundle folder holding session.json, session.log and the test logs")
    options = parser.parse_args(arguments)
    target = options.bundle / "host.json"
    if target.exists():
        print(f"ERROR: {target} exists already; a record is never replaced", file=sys.stderr)
        return 2
    try:
        record = build_record(options.root, options.candidate_sha, options.bundle)
    except (ValueError, KeyError, OSError, json.JSONDecodeError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2
    target.write_text(json.dumps(record, indent=2) + "\n", encoding="utf-8")
    report = evidence.evaluate(options.root, options.candidate_sha, target)
    print(evidence.markdown(report))
    return 0 if report["status"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
