#!/usr/bin/env python3
"""Validate a manual Excel evidence bundle against one exact candidate commit.

The validator never runs Excel and cannot authenticate what an operator did. It
checks that the record is complete and consistent, that the declared source
inventory is the exact Git content of the candidate, and that the retained logs
match their digests and the reports of the harness and of every other suite the
policy lists, each run on its own, and of the optional Test_RunAll summary.
With --check-policy it instead checks that the policy's suites and counts match
the test modules of the working tree. See docs/EXCEL_EVIDENCE.md.
Adapted from EXCEL-VBA-PROJECT-TEMPLATE b903fe44: manual execution only, and
the source inventory is derived from the repository layout.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from datetime import datetime
from pathlib import Path, PurePosixPath
from typing import Any

from _gatelib import git_bytes, run_gate

DESCRIPTION = "Validate a manual Excel evidence bundle against one exact candidate commit."
POLICY = ".github/excel-evidence-policy.json"
STAGES = ("import", "compile", "regression", "cleanup")
OUTCOMES = {"import": "IMPORT_FAILED", "compile": "COMPILE_FAILED",
            "regression": "TEST_FAILED", "cleanup": "CLEANUP_FAILED"}
# Imported components: production and tests, never examples (docs/REPOSITORY_STRUCTURE.md).
SOURCE_ROOTS = ("src", "tests")
VBA_SUFFIXES = (".bas", ".cls", ".frm")
# The macro-free workbook that holds the sheets; the build starts from it (issue #58).
WORKBOOK_TEMPLATE = "src/workbook/SACCR_Template.xlsx"
SHA = re.compile(r"[0-9a-f]{40}")
DIGEST = re.compile(r"[0-9a-fA-F]{64}")
SCOPE = ("Checks a retained manual record and its source and log bindings; "
         "does not run Excel or authenticate the operator.")
RECORD_KEYS = ("schema_version repository candidate_sha execution availability_reason "
               "started_at finished_at operator environment sources stages harness suites summary")
POLICY_KEYS = ("schema_version repository entry_point cases assertions expected_error_cases "
               "required_references suites summary_entry_point")
SUITE_POLICY_KEYS = "entry_point header mode_line cases checks restoration first_case"
SUITE_RECORD_KEYS = "entry_point status cases checks failures log"
BITNESS = ("32-bit", "64-bit")
STATUSES = ("PASS", "FAIL", "NOT_RUN", "TIMEOUT")
# One line of the Test_RunAll summary, e.g. "  PASS  TEST_Cases            cases=20; ...  (4,1 s)".
SUMMARY_LINE = re.compile(r"^  (PASS|FAIL)  (TEST_\w+) +(.+?)  \(([0-9.,]+) s\)[ \t]*$", re.MULTILINE)
ENVIRONMENT_KEYS = ("excel_version excel_build office_bitness os os_architecture runtime "
                    "references macro_policy vba_project_access trust_changes")


def require(condition: object, message: str) -> None:
    if not condition:
        raise ValueError(message)


def object_keys(value: Any, keys: str, label: str) -> dict[str, Any]:
    """Return ``value`` when it is an object with exactly ``keys``, else fail."""
    if not (isinstance(value, dict) and set(value) == set(keys.split())):
        raise ValueError(f"{label}: invalid object fields")
    return value


def nonempty(value: Any) -> bool:
    return isinstance(value, str) and bool(value.strip())


def relative(value: Any) -> bool:
    return (nonempty(value) and "\\" not in value and "\0" not in value
            and not PurePosixPath(value).is_absolute() and ".." not in value.split("/")
            and PurePosixPath(value).as_posix() == value and value != ".")


def decode(raw: bytes) -> Any:
    def unique(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            require(key not in result, f"duplicate JSON key: {key}")
            result[key] = value
        return result
    return json.loads(raw.decode("utf-8-sig"), object_pairs_hook=unique)


def string_list(value: Any, label: str) -> list[str]:
    require(isinstance(value, list) and all(nonempty(item) for item in value),
            f"invalid {label}")
    require(len(value) == len(set(value)), f"duplicate {label}")
    return list(value)


def require_commit(root: Path, sha: str) -> None:
    require(bool(SHA.fullmatch(sha)), "candidate must be a full 40-character lowercase SHA")
    found = git_bytes(root, "cat-file", "-e", f"{sha}^{{commit}}")
    require(found.returncode == 0, f"candidate commit {sha} is not in this repository")


def committed(root: Path, sha: str, path: str) -> bytes:
    require(relative(path), f"unsafe path: {path}")
    result = git_bytes(root, "show", f"{sha}:{path}")
    require(result.returncode == 0, f"candidate does not contain {path}")
    return result.stdout


def source_inventory(root: Path, sha: str) -> list[dict[str, str]]:
    """Return the build inputs at ``sha`` with SHA-256 of their exact Git bytes.

    These are the imported VBA components, each form's resource and, when
    tracked, the workbook template the build starts from.
    """
    listing = git_bytes(root, "ls-tree", "-r", "-z", "--name-only", sha, "--", *SOURCE_ROOTS)
    require(listing.returncode == 0, "cannot list candidate source")
    tracked = {item.decode("utf-8") for item in listing.stdout.split(b"\0") if item}
    paths = {path for path in tracked if path.lower().endswith(VBA_SUFFIXES)}
    # A form's binary resource is imported with it, so it is part of the identity.
    for path in tuple(paths):
        if path.lower().endswith(".frm") and path[:-4] + ".frx" in tracked:
            paths.add(path[:-4] + ".frx")
    require(bool(paths), "candidate has no importable VBA source")
    if WORKBOOK_TEMPLATE in tracked:
        paths.add(WORKBOOK_TEMPLATE)
    return [{"path": path, "sha256": hashlib.sha256(committed(root, sha, path)).hexdigest()}
            for path in sorted(paths)]


def count(value: Any, label: str) -> int:
    require(type(value) is int and value >= 0, f"invalid {label}")
    return int(value)


def parse_suite_policy(value: Any) -> dict[str, Any]:
    suite = object_keys(value, SUITE_POLICY_KEYS, "policy suite")
    require(nonempty(suite["entry_point"]) and "." in suite["entry_point"]
            and nonempty(suite["header"]), "policy suite needs an entry point and a header")
    require(isinstance(suite["mode_line"], bool), "policy suite mode_line must be true or false")
    require(count(suite["cases"], "policy suite cases") > 0, "policy suite needs cases")
    checks = object_keys(suite["checks"], " ".join(BITNESS), "policy suite checks")
    require(all(count(checks[key], "policy suite checks") > 0 for key in BITNESS),
            "policy suite needs checks for each bitness")
    require(suite["restoration"] is None
            or (isinstance(suite["restoration"], str)
                and re.fullmatch(r"\w+=\w+", suite["restoration"]) is not None),
            "policy suite restoration must be null or key=value")
    require(suite["first_case"] is None
            or (isinstance(suite["first_case"], str)
                and re.fullmatch(r"[\w.-]+", suite["first_case"]) is not None),
            "policy suite first_case must be null or a case name")
    return suite


def parse_policy(raw: bytes) -> dict[str, Any]:
    policy = object_keys(decode(raw), POLICY_KEYS, "evidence policy")
    require(type(policy["schema_version"]) is int and policy["schema_version"] == 2,
            "unsupported evidence policy schema")
    require(nonempty(policy["repository"]) and nonempty(policy["entry_point"]),
            "policy needs a repository and an entry point")
    require(type(policy["assertions"]) is int and policy["assertions"] > 0,
            "policy needs a positive assertion count")
    cases = string_list(policy["cases"], "policy cases")
    errors = string_list(policy["expected_error_cases"], "policy expected-error cases")
    string_list(policy["required_references"], "policy references")
    require(bool(cases) and set(errors) <= set(cases),
            "expected-error cases must belong to the declared suite")
    require(isinstance(policy["suites"], list) and bool(policy["suites"]), "policy needs suites")
    suites = [parse_suite_policy(suite) for suite in policy["suites"]]
    entries = [suite["entry_point"] for suite in suites]
    require(len(set(entries + [policy["entry_point"]])) == len(entries) + 1,
            "policy entry points must be distinct")
    # Suites that print the same header are told apart by the first case they run.
    headers = [suite["header"] for suite in suites]
    shared = [suite for suite in suites if headers.count(suite["header"]) > 1]
    require(all(suite["first_case"] is not None for suite in shared),
            "suites that share a header need their first case")
    firsts = [suite["first_case"] for suite in shared]
    require(len(set(firsts)) == len(firsts), "suites that share a header need distinct first cases")
    require(nonempty(policy["summary_entry_point"]), "policy needs the summary entry point")
    return policy


def load_policy(root: Path, sha: str) -> dict[str, Any]:
    return parse_policy(committed(root, sha, POLICY))


def validate_identity(value: Any, policy: dict[str, Any], sha: str) -> dict[str, Any]:
    record = object_keys(value, RECORD_KEYS, "evidence record")
    require(type(record["schema_version"]) is int and record["schema_version"] == 2,
            "unsupported evidence record schema")
    require(record["repository"] == policy["repository"], "record repository differs from policy")
    require(record["candidate_sha"] == sha, "record candidate_sha differs from the candidate")
    require(record["execution"] in ("manual", "unavailable"), "execution must be manual or unavailable")
    times = []
    for key in ("started_at", "finished_at"):
        require(nonempty(record[key]), f"{key} is required")
        stamp = datetime.fromisoformat(record[key].replace("Z", "+00:00"))
        require(stamp.utcoffset() is not None, f"{key} needs a timezone")
        times.append(stamp)
    require(times[1] >= times[0], "finished_at precedes started_at")
    return record


def validate_environment(record: dict[str, Any], policy: dict[str, Any]) -> None:
    require(nonempty(record["operator"]), "operator and workstation are required")
    environment = object_keys(record["environment"], ENVIRONMENT_KEYS, "environment")
    require(environment["trust_changes"] is False, "a run must not change macro trust settings")
    require(all(nonempty(value) for key, value in environment.items()
                if key not in ("references", "trust_changes")),
            "environment fields must be nonempty")
    require(environment["office_bitness"] in ("32-bit", "64-bit"), "invalid Office bitness")
    require(environment["os"].casefold().startswith("windows"), "Excel for Windows is required")
    references = string_list(environment["references"], "references")
    require(sorted(references) == sorted(policy["required_references"]),
            "references differ from the required default set")


def retained_log(directory: Path, value: Any) -> str:
    log = object_keys(value, "path sha256", "log reference")
    require(relative(log["path"]), "unsafe log path")
    require(isinstance(log["sha256"], str) and DIGEST.fullmatch(log["sha256"]),
            "log digest must be 64 hexadecimal characters")
    path: Path = directory / log["path"]
    require(not any(item.is_symlink() for item in (path, *path.parents)), "symlinked log path")
    require(path.resolve().is_relative_to(directory.resolve()) and path.is_file(),
            f"missing log: {log['path']}")
    raw = path.read_bytes()
    require(hashlib.sha256(raw).hexdigest() == log["sha256"].lower(),
            f"log digest mismatch: {log['path']}")
    return raw.decode("utf-8-sig").replace("\r\n", "\n")


def validate_stages(record: dict[str, Any], directory: Path) -> tuple[list[str], dict[str, str]]:
    stages = object_keys(record["stages"], " ".join(STAGES), "stages")
    outcomes: list[str] = []
    logs: dict[str, str] = {}
    for name in STAGES:
        stage = object_keys(stages[name], "status detail log", f"{name} stage")
        status = stage["status"]
        require(status in ("PASS", "FAIL", "NOT_RUN", "TIMEOUT"), f"invalid {name} status")
        require(nonempty(stage["detail"]), f"{name} needs a detail")
        if status == "NOT_RUN":
            require(stage["log"] is None, f"NOT_RUN {name} cannot cite a log")
            outcomes.append("INCOMPLETE")
            continue
        logs[name] = retained_log(directory, stage["log"])
        if status != "PASS":
            outcomes.append("EXECUTION_TIMEOUT" if status == "TIMEOUT" else OUTCOMES[name])
    status = {name: stages[name]["status"] for name in STAGES}
    require(status["import"] == "PASS" or status["compile"] == "NOT_RUN",
            "compile cannot run after a failed import")
    require(status["compile"] == "PASS" or status["regression"] == "NOT_RUN",
            "regression cannot run after a failed compile")
    require("import" not in logs or record["candidate_sha"] in logs["import"],
            "import log must name the candidate SHA")
    return list(dict.fromkeys(outcomes)), logs


def single(log: str, key: str, label: str = "harness") -> str:
    values: list[str] = re.findall(rf"^{re.escape(key)}=(.*)$", log, re.MULTILINE)
    require(len(values) == 1, f"{label} log needs exactly one {key}= line")
    return values[0].strip()


def require_candidate(log: str, sha: str, label: str) -> None:
    """Each test log starts with the operator's CANDIDATE= line, so it cannot serve another commit."""
    require(single(log, "CANDIDATE", label) == sha, f"{label} log names another candidate")


def report_cases(record: dict[str, Any], policy: dict[str, Any], log: str) -> list[str]:
    """Check the report header and return its CASE lines, which must follow the policy order."""
    require(re.findall(r"^SACCR TESTS[ \t]*$", log, re.MULTILINE) == ["SACCR TESTS"],
            "harness log must contain exactly one run")
    require(single(log, "MODE") == "NORMAL", "harness log is not a normal-mode run")
    reported: dict[str, str] = {}
    for item in single(log, "ENVIRONMENT").split("; "):
        key, _, value = item.partition("=")
        reported[key] = value
    environment = record["environment"]
    # Application.OperatingSystem, e.g. "Windows (64-bit) NT 10.00"; Excel for Mac is unsupported.
    require(reported.get("os", "").startswith("Windows"),
            "harness ENVIRONMENT os is not Windows")
    for key, field in (("version", "excel_version"), ("office", "office_bitness"),
                       ("runtime", "runtime")):
        require(reported.get(key) == environment[field],
                f"harness ENVIRONMENT {key} differs from the record's {field}")
    cases = [case.strip() for case in re.findall(r"^CASE=(.*)$", log, re.MULTILINE)]
    require(cases == policy["cases"][:len(cases)], "harness CASE lines differ from the policy")
    require(len(cases) == record["harness"]["cases"], "harness CASE lines differ from the record")
    return cases


def validate_report(record: dict[str, Any], policy: dict[str, Any], log: str) -> None:
    """Bind a passing or failed regression to one complete normal-mode harness report."""
    harness = record["harness"]
    require_candidate(log, record["candidate_sha"], "harness")
    report_cases(record, policy, log)
    for field in ("cases", "assertions", "failures"):
        require(single(log, field.upper()) == str(harness[field]),
                f"harness {field.upper()} differs from the record")
    cleanup = single(log, "CLEANUP").split(";")[0]
    require(cleanup in ("PASS", "FAIL"), "harness CLEANUP must be PASS or FAIL")
    # The harness reports COMPLETE only when both expected counts were reached.
    complete = (harness["cases"] == len(policy["cases"])
                and harness["assertions"] == policy["assertions"])
    require(harness["completeness"] == ("COMPLETE" if complete else "INCOMPLETE"),
            "harness completeness contradicts the counts")
    passed = complete and harness["failures"] == 0 and cleanup == "PASS"
    require(single(log, "RESULT") == (
        f"{'PASS' if passed else 'FAIL'}; completeness={harness['completeness']}; "
        f"cases={harness['cases']}; assertions={harness['assertions']}; "
        f"failures={harness['failures']}; cleanup={cleanup}"),
        "harness RESULT contradicts the record")
    regression = record["stages"]["regression"]["status"]
    require(regression == "FAIL" or (complete and harness["failures"] == 0),
            "regression PASS contradicts the harness results")
    # The suites are validated after the harness; read their statuses defensively here.
    suites_failed = isinstance(record["suites"], list) and any(
        isinstance(suite, dict) and suite.get("status") == "FAIL" for suite in record["suites"])
    require(regression == "PASS" or not passed or suites_failed,
            "regression FAIL contradicts a passing report")
    require(cleanup == "PASS" or record["stages"]["cleanup"]["status"] == "FAIL",
            "a harness cleanup failure must be recorded as cleanup FAIL")


def validate_harness(record: dict[str, Any], policy: dict[str, Any], log: str) -> None:
    harness = object_keys(record["harness"], "entry_point cases assertions failures "
                          "completeness expected_errors", "harness")
    require(harness["entry_point"] == policy["entry_point"], "wrong harness entry point")
    for field in ("cases", "assertions", "failures"):
        require(type(harness[field]) is int and harness[field] >= 0, f"invalid harness {field}")
    require(harness["completeness"] in ("COMPLETE", "INCOMPLETE"), "invalid completeness")
    require(isinstance(harness["expected_errors"], list), "expected_errors must be an array")
    observed = []
    for error in harness["expected_errors"]:
        error = object_keys(error, "case status detail", "expected error")
        require(error["status"] in ("PASS", "FAIL", "NOT_RUN") and nonempty(error["detail"]),
                "invalid expected-error result")
        observed.append(error["case"])
    require(observed == policy["expected_error_cases"], "expected-error results differ from policy")
    # Tied to the harness's own counts: another suite's failure makes the regression
    # stage FAIL but says nothing about the harness's expected errors.
    clean = harness["completeness"] == "COMPLETE" and harness["failures"] == 0
    require(not clean or all(error["status"] == "PASS" for error in harness["expected_errors"]),
            "a passing harness needs every expected-error result to pass")
    validate_report(record, policy, log)


def result_fields(result: str, keys: list[str], label: str) -> tuple[str, list[str]]:
    """Split a RESULT line into its verdict and the values of ``keys``, in that order."""
    verdict, *fields = result.split("; ")
    require(verdict in ("PASS", "FAIL"), f"{label} RESULT must be PASS or FAIL")
    pairs = [field.partition("=") for field in fields]
    require([key for key, _, _ in pairs] == keys, f"{label} RESULT fields differ from the policy")
    return verdict, [value for _, _, value in pairs]


def validate_suite_log(spec: dict[str, Any], item: dict[str, Any], log: str,
                       record: dict[str, Any]) -> None:
    """Bind one suite record to a single normal-mode report of that suite."""
    label = spec["entry_point"]
    require_candidate(log, record["candidate_sha"], label)
    require(re.findall(rf"^{re.escape(spec['header'])}[ \t]*$", log, re.MULTILINE) == [spec["header"]],
            f"{label} log must contain exactly one run")
    modes = [mode.strip() for mode in re.findall(r"^MODE=(.*)$", log, re.MULTILINE)]
    require(modes == (["NORMAL"] if spec["mode_line"] else []), f"{label} log is not a normal-mode run")
    cases = [case.strip() for case in re.findall(r"^CASE=(.*)$", log, re.MULTILINE)]
    require(len(cases) == item["cases"], f"{label} CASE lines differ from the record")
    # A shared header does not say which suite ran; its first case does (regime suffix aside).
    first = spec["first_case"]
    require(first is None or not cases or cases[0] == first or cases[0].startswith(first + "."),
            f"{label} log starts with another suite's case")
    restoration = spec["restoration"].split("=") if spec["restoration"] else None
    keys = ["cases", "checks", "failures"] + ([restoration[0]] if restoration else [])
    verdict, values = result_fields(single(log, "RESULT", label), keys, label)
    require(values[:3] == [str(item[key]) for key in ("cases", "checks", "failures")],
            f"{label} RESULT counts differ from the record")
    if item["status"] == "FAIL":
        require(verdict == "FAIL", f"{label} FAIL contradicts a passing report")
        # A failed run may stop early, but never runs more cases than its suite has.
        require(item["cases"] <= spec["cases"], f"{label} reports more cases than the policy has")
        return
    checks = spec["checks"][record["environment"]["office_bitness"]]
    require(verdict == "PASS" and item["failures"] == 0, f"{label} PASS contradicts its report")
    require(item["cases"] == spec["cases"] and item["checks"] == checks,
            f"{label} counts differ from the policy ({spec['cases']} cases, {checks} checks)")
    require(restoration is None or values[3] == restoration[1], f"{label} did not restore its state")


def validate_suite(spec: dict[str, Any], value: Any, record: dict[str, Any],
                   directory: Path) -> tuple[str | None, str | None]:
    """Return the suite's outcome, if it is not a pass, and its retained log."""
    item = object_keys(value, SUITE_RECORD_KEYS, "suite")
    label = spec["entry_point"]
    require(item["entry_point"] == label, "suites differ from the policy")
    require(item["status"] in STATUSES, f"invalid {label} status")
    if item["status"] in ("NOT_RUN", "TIMEOUT"):
        require(all(item[key] is None for key in ("cases", "checks", "failures")),
                f"an unfinished {label} has no results")
        if item["status"] == "NOT_RUN":
            require(item["log"] is None, f"NOT_RUN {label} cannot cite a log")
            return "INCOMPLETE", None
        log = retained_log(directory, item["log"])
        require_candidate(log, record["candidate_sha"], label)
        require(not re.search(r"^RESULT=", log, re.M),
                f"{label} log has a RESULT line; record it as PASS or FAIL, not TIMEOUT")
        return "EXECUTION_TIMEOUT", log
    for key in ("cases", "checks", "failures"):
        count(item[key], f"{label} {key}")
    log = retained_log(directory, item["log"])
    validate_suite_log(spec, item, log, record)
    return ("TEST_FAILED" if item["status"] == "FAIL" else None), log


def validate_suites(record: dict[str, Any], policy: dict[str, Any],
                    directory: Path) -> tuple[list[str], dict[str, str]]:
    """Validate every suite the policy lists, in its order; return outcomes and logs."""
    suites = record["suites"]
    require(isinstance(suites, list), "suites must be an array")
    require([item.get("entry_point") if isinstance(item, dict) else None for item in suites]
            == [spec["entry_point"] for spec in policy["suites"]],
            "suites differ from the policy: every suite is recorded once, in policy order")
    regression = record["stages"]["regression"]["status"]
    require(regression != "NOT_RUN" or all(item["status"] == "NOT_RUN" for item in suites),
            "no suite can run when the regression did not")
    outcomes: list[str] = []
    logs: dict[str, str] = {}
    for spec, item in zip(policy["suites"], suites):
        outcome, log = validate_suite(spec, item, record, directory)
        if outcome:
            outcomes.append(outcome)
        if log is not None:
            logs[spec["entry_point"]] = log
    # A suite that did not finish is INCOMPLETE or EXECUTION_TIMEOUT, not a test failure.
    require(regression != "PASS" or all(item["status"] != "FAIL" for item in suites),
            "regression PASS cannot hide a failed suite")
    return outcomes, logs


def report_detail(log: str, label: str) -> tuple[str, str]:
    """The verdict of a report and its RESULT fields, as Test_RunAll prints them."""
    verdict, _, detail = single(log, "RESULT", label).partition("; ")
    return verdict, detail


def validate_summary(record: dict[str, Any], policy: dict[str, Any], directory: Path,
                     logs: dict[str, str]) -> None:
    """Check the optional Test_RunAll summary against the individual reports."""
    if record["summary"] is None:
        return
    entries = [policy["entry_point"]] + [spec["entry_point"] for spec in policy["suites"]]
    require(set(entries) <= set(logs), "a summary needs every suite's own report")
    log = retained_log(directory, record["summary"])
    require_candidate(log, record["candidate_sha"], "summary")
    require(len(re.findall(r"^SACCR TEST SUMMARY ", log, re.MULTILINE)) == 1,
            "summary log must contain exactly one Test_RunAll run")
    lines = SUMMARY_LINE.findall(log)
    require([name for _, name, _, _ in lines] == [entry.split(".")[0] for entry in entries],
            "summary suites differ from the policy")
    for (status, name, detail, _), entry in zip(lines, entries):
        require((status, detail.strip()) == report_detail(logs[entry], entry),
                f"summary line of {name} contradicts its own report")
    failed = sum(status == "FAIL" for status, _, _, _ in lines)
    verdict, values = result_fields(single(log, "RESULT", "summary"),
                                    ["suites", "passed", "failed", "seconds"], "summary")
    require(values[:3] == [str(len(entries)), str(len(entries) - failed), str(failed)]
            and verdict == ("FAIL" if failed else "PASS"), "summary RESULT contradicts its suite lines")


def validate_unavailable(record: dict[str, Any]) -> list[str]:
    require(nonempty(record["availability_reason"]), "unavailable needs a reason")
    require(all(record[key] is None for key in ("operator", "environment", "harness", "summary"))
            and record["sources"] == [] and record["stages"] == {} and record["suites"] == [],
            "unavailable cannot carry execution claims")
    return ["UNAVAILABLE"]


def normalized(log: str) -> str:
    """A log as the validator reads it: no blank lines, no trailing spaces."""
    return "\n".join(line.rstrip() for line in log.splitlines() if line.strip())


def validate_manual(root: Path, sha: str, policy: dict[str, Any], record: dict[str, Any],
                    directory: Path) -> list[str]:
    require(record["availability_reason"] is None, "a manual run has no unavailability reason")
    validate_environment(record, policy)
    require(record["sources"] == source_inventory(root, sha),
            "source inventory differs from the candidate's exact source")
    outcomes, logs = validate_stages(record, directory)
    regression = record["stages"]["regression"]["status"]
    if regression in ("NOT_RUN", "TIMEOUT"):
        # An unfinished run is non-green; its log is retained and hashed but makes no claims.
        require(record["harness"] is None, "an unfinished regression has no harness results")
        if regression == "TIMEOUT":
            require_candidate(logs["regression"], record["candidate_sha"], "harness")
            require(not re.search(r"^RESULT=", logs["regression"], re.M),
                    "a log with a RESULT line finished; record it as PASS or FAIL, not TIMEOUT")
    else:
        validate_harness(record, policy, logs["regression"])
    suite_outcomes, suite_logs = validate_suites(record, policy, directory)
    if regression in ("PASS", "FAIL"):
        suite_logs[policy["entry_point"]] = logs["regression"]
    validate_summary(record, policy, directory, suite_logs)
    # One log cited for two runs is one run, whatever blank lines or trailing
    # spaces were added to a copy: the validator reads past both.
    runs = {normalized(log) for log in suite_logs.values()}
    require(len(runs) == len(suite_logs), "the same log is cited for two test runs")
    return list(dict.fromkeys(outcomes + suite_outcomes))


def declared(text: str, names: str, position: int) -> set[int]:
    """Counts a test module declares: Const <names> As Long = n, or BeginSuite n, m."""
    values = {int(value) for value in re.findall(rf"Const (?:{names})\s+As Long = (\d+)", text)}
    begin = re.findall(r"BeginSuite (\d+), (\d+)", text)
    return values | {int(pair[position]) for pair in begin}


def checks_by_bitness(text: str) -> dict[str, set[int]]:
    """Check counts a module declares for each bitness: under #If Win64 for 64-bit,
    under its #Else for 32-bit, and outside such a block for both."""
    found: dict[str, set[int]] = {bits: set() for bits in BITNESS}
    branch: str | None = None
    for line in text.splitlines():
        stripped = line.strip()
        if re.match(r"#If\s+Win64\s+Then\b", stripped, re.IGNORECASE):
            branch = "64-bit"
        elif branch and re.match(r"#Else\b", stripped, re.IGNORECASE):
            branch = "32-bit"
        elif branch and re.match(r"#End\s+If\b", stripped, re.IGNORECASE):
            branch = None
        else:
            values = declared(line, "EXPECTED_CHECKS|CHECKS", 1)
            for bits in ([branch] if branch else BITNESS):
                found[bits] |= values
    return found


def module_text(root: Path, entry_point: str) -> str:
    module, _, procedure = entry_point.partition(".")
    path = root / "tests" / "modules" / f"{module}.bas"
    require(path.is_file(), f"{entry_point}: no tests/modules/{module}.bas")
    text = path.read_text(encoding="cp1252")
    require(re.search(rf"^Public Sub {re.escape(procedure)}\b", text, re.MULTILINE) is not None,
            f"{entry_point}: no Public Sub {procedure}")
    return text


def suite_problems(root: Path, spec: dict[str, Any]) -> list[str]:
    """The policy counts of one suite against the counts its module declares."""
    text = module = module_text(root, spec["entry_point"])
    problems = []
    if declared(text, "EXPECTED_CASES|CASES", 0) != {spec["cases"]}:
        problems.append(f"{spec['entry_point']}: policy cases differ from the module")
    # A module that declares no check count (MainState) is held only by the evidence logs.
    for bits, checks in checks_by_bitness(text).items():
        if checks and checks != {spec["checks"][bits]}:
            problems.append(f"{spec['entry_point']}: policy {bits} checks differ from the module")
    # The case-runner suites print the runner's header and MODE line.
    runner = "TEST_CaseRunner.BeginSuite" in text
    if runner:
        text += (root / "tests" / "modules" / "TEST_CaseRunner.bas").read_text(encoding="cp1252")
    if f'"{spec["header"]}"' not in text:
        problems.append(f"{spec['entry_point']}: header {spec['header']!r} not printed by the module")
    if spec["mode_line"] != runner:
        problems.append(f"{spec['entry_point']}: policy mode_line differs from the module")
    # The first case a case-runner module starts in its source is the first it runs.
    first = re.search(r'\b(?:BeginCase|StartCase)\s+"([^"]+)"', module)
    if spec["first_case"] is not None and (first is None or first.group(1) != spec["first_case"]):
        problems.append(f"{spec['entry_point']}: policy first_case differs from the module")
    return problems


def readme_problems(root: Path, policy: dict[str, Any]) -> list[str]:
    """The test counts the root README shows, against the policy (64-bit counts)."""
    rows = [(policy["entry_point"].split(".")[0], len(policy["cases"]), policy["assertions"])]
    rows += [(spec["entry_point"].split(".")[0], spec["cases"], spec["checks"]["64-bit"])
             for spec in policy["suites"]]
    suites = len(rows)
    cases = sum(row[1] for row in rows)
    checks = sum(row[2] for row in rows)
    text = (root / "README.md").read_text(encoding="utf-8")
    problems = []
    # Each display appears once with the policy's values: a stale copy left beside
    # the corrected one would contradict it.
    displays = {
        "Excel suites badge": (r"badge/Excel_suites-([^-?)]+)", f"{suites}%2F{suites}_PASS"),
        "Regression badge": (r"badge/Regression-([^-?)]+)", f"{cases}_cases_%C2%B7_{checks}_checks"),
        "Testing capability": (r"\((\d+ cases, \d+ checks)\)", f"{cases} cases, {checks} checks"),
        "suite table total": (r"^\| \*\*Total\*\* \|(.*)\|$", f" | **{cases}** | **{checks}** "),
        "Test_RunAll result": (r"^RESULT=(\w+; suites=\d+; passed=\d+; failed=\d+);",
                               f"PASS; suites={suites}; passed={suites}; failed=0"),
    }
    for label, (pattern, value) in displays.items():
        displayed = re.findall(pattern, text, re.MULTILINE)
        if displayed != [value]:
            problems.append(f"README.md: {label} must appear once, as {value!r}, not {displayed}")
    # Every row, in order: a second, stale row for one suite is a contradiction.
    table = [(name, int(found_cases), int(found_checks)) for name, found_cases, found_checks
             in re.findall(r"^\| `(TEST_\w+)` \|.*\| (\d+) \| (\d+) \|$", text, re.MULTILINE)]
    if table != rows:
        problems.append("README.md: suite table rows differ from the policy suites and counts")
    # Every suite-result line must parse: one in another format, such as an older
    # summary without timings, would otherwise be skipped while still displayed.
    lines = re.findall(r"^[ \t]*(?:PASS|FAIL)[ \t]+TEST_\w+.*$", text, re.MULTILINE)
    if len(re.findall(r"^SACCR TEST SUMMARY ", text, re.MULTILINE)) != 1 \
            or len(lines) != len(SUMMARY_LINE.findall(text)):
        problems.append("README.md: Test_RunAll output must be one summary whose lines all parse")
    shown = []
    for status, name, detail, _ in SUMMARY_LINE.findall(text):
        found = re.search(r"\bcases=(\d+); (?:checks|assertions)=(\d+); failures=0\b", detail)
        shown.append((name, int(found.group(1)), int(found.group(2))) if status == "PASS" and found else None)
    if shown != rows:
        problems.append("README.md: Test_RunAll output differs from the policy suites and counts")
    return problems


def policy_problems(root: Path) -> list[str]:
    """Check that the working tree's policy matches the test modules, Test_RunAll and README."""
    try:
        policy = parse_policy((root / POLICY).read_bytes())
        harness = module_text(root, policy["entry_point"])
        problems = []
        if (declared(harness, "EXPECTED_CASES", 0) != {len(policy["cases"])}
                or declared(harness, "EXPECTED_ASSERTIONS", 1) != {policy["assertions"]}):
            problems.append(f"{policy['entry_point']}: policy counts differ from the module")
        for spec in policy["suites"]:
            problems += suite_problems(root, spec)
        runner = module_text(root, policy["summary_entry_point"])
        listed = re.findall(r'"(TEST_\w+\.\w+)"', runner.split("suites = Array(", 1)[-1].split(")\n", 1)[0])
        entries = [policy["entry_point"]] + [spec["entry_point"] for spec in policy["suites"]]
        if listed != entries:
            problems.append(f"{policy['summary_entry_point']} runs {listed}, not the policy suites")
        if declared(runner, "SUITE_COUNT", 0) != {len(entries)}:
            problems.append("SUITE_COUNT differs from the number of policy suites")
        return problems + readme_problems(root, policy)
    except (ValueError, OSError) as error:
        return [str(error)]


def evaluate(root: Path, sha: str, evidence: Path) -> dict[str, Any]:
    execution = "unresolved"
    outcomes: list[str] = []
    findings: list[str] = []
    try:
        require_commit(root, sha)
        policy = load_policy(root, sha)
        record = validate_identity(decode(evidence.read_bytes()), policy, sha)
        execution = record["execution"]
        if execution == "unavailable":
            outcomes = validate_unavailable(record)
        else:
            outcomes = validate_manual(root, sha, policy, record, evidence.parent)
    except (ValueError, TypeError, KeyError, AttributeError, OSError) as error:
        outcomes.append("EVIDENCE_INVALID")
        findings.append(str(error))
    return {"status": "fail" if outcomes else "pass", "candidate_sha": sha,
            "execution": execution, "outcomes": outcomes or ["PASS"],
            "findings": findings, "scope": SCOPE}


def markdown(report: dict[str, Any]) -> str:
    lines = [f"# Excel evidence: {report['status'].upper()}", "",
             f"Candidate: `{report['candidate_sha']}`; execution: `{report['execution']}`", "",
             "Outcomes: " + ", ".join(report["outcomes"]), ""]
    lines += [f"- {item}" for item in report["findings"]]
    return "\n".join(lines + ["", report["scope"], ""])


def main(arguments: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=DESCRIPTION)
    parser.add_argument("--root", type=Path, default=Path.cwd())
    parser.add_argument("--candidate-sha")
    parser.add_argument("--evidence", type=Path, help="host.json inside the retained bundle")
    parser.add_argument("--inventory", action="store_true",
                        help="print the candidate's source inventory for the record and exit")
    parser.add_argument("--check-policy", action="store_true",
                        help="check the working tree's policy against the test modules and exit")
    parser.add_argument("--output", type=Path)
    parser.add_argument("--summary", type=Path)
    options = parser.parse_args(arguments)
    if options.check_policy:
        problems = policy_problems(options.root)
        print("\n".join(problems) if problems else "Excel evidence policy: PASS")
        return 1 if problems else 0
    if options.candidate_sha is None:
        parser.error("--candidate-sha is required")
    if options.inventory:
        try:
            require_commit(options.root, options.candidate_sha)
            print(json.dumps(source_inventory(options.root, options.candidate_sha), indent=2))
            return 0
        except (ValueError, OSError) as error:
            print(f"ERROR: {error}", file=sys.stderr)
            return 2
    if options.evidence is None:
        parser.error("--evidence is required unless --inventory is given")
    bundle = options.evidence.resolve().parent
    for path in (options.output, options.summary):
        if path is not None and path.resolve().is_relative_to(bundle):
            parser.error("reports must be written outside the evidence bundle")
    return run_gate(options, build=lambda: evaluate(options.root, options.candidate_sha,
                                                    options.evidence),
                    markdown=markdown, errors=(OSError,))


if __name__ == "__main__":
    raise SystemExit(main())
