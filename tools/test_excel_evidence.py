"""Synthetic Excel evidence records; none of these tests runs Excel."""
from __future__ import annotations

import contextlib
import copy
import hashlib
import io
import json
import shutil
import subprocess
import tempfile
import unittest
from pathlib import Path
from typing import Any

import check_excel_evidence as evidence

ROOT = Path(__file__).resolve().parents[1]
SOURCES = {
    "src/core/CORE_Engine.bas": "Attribute VB_Name = \"CORE_Engine\"\n",
    "src/modules/SACCR_Formulas.bas": "Attribute VB_Name = \"SACCR_Formulas\"\n",
    "src/forms/FSample.frm": "Attribute VB_Name = \"FSample\"\n",
    "src/forms/FSample.frx": "binary resource\n",
    "src/workbook/SACCR_Template.xlsx": "template bytes\n",
    "tests/modules/TEST_Harness.bas": "Attribute VB_Name = \"TEST_Harness\"\n",
    "examples/ExampleUse.bas": "Attribute VB_Name = \"ExampleUse\"\n",
    "README.md": "fixture\n",
}


class ExcelEvidenceTests(unittest.TestCase):
    def setUp(self) -> None:
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.area = Path(temporary.name)
        self.root = self.area / "candidate"
        self.bundle = self.area / "bundle"
        self.bundle.mkdir()
        self.path = self.bundle / "host.json"
        self.git("init", "-q", str(self.root))
        (self.root / ".github").mkdir()
        shutil.copy(ROOT / evidence.POLICY, self.root / evidence.POLICY)
        for name, text in SOURCES.items():
            target = self.root / name
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_text(text)
        self.commit("Synthetic candidate")
        self.sha = self.git("-C", str(self.root), "rev-parse", "HEAD").strip()
        self.policy: dict[str, Any] = json.loads((ROOT / evidence.POLICY).read_text())
        self.session = f"Candidate {self.sha}\nImported 4 components; compiled clean.\n"
        self.log = "\n".join([
            f"CANDIDATE={self.sha}", "SACCR TESTS", "MODE=NORMAL",
            "ENVIRONMENT=host=Microsoft Excel; version=16.0; os=Windows (64-bit) NT 10.00; "
            "office=64-bit; runtime=VBA7+",
            *("CASE=" + case for case in self.policy["cases"]),
            "CASES=4", "ASSERTIONS=6", "FAILURES=0",
            "CLEANUP=PASS; detail=Calculation, DisplayAlerts, EnableEvents and ScreenUpdating unchanged",
            "RESULT=PASS; completeness=COMPLETE; cases=4; assertions=6; failures=0; cleanup=PASS", "",
        ])
        self.record: dict[str, Any] = {
            "schema_version": 2, "repository": "danielep71/VBA-SACCR-Toolkit", "candidate_sha": self.sha,
            "execution": "manual", "availability_reason": None,
            "started_at": "2026-10-06T09:00:00+02:00", "finished_at": "2026-10-06T09:10:00+02:00",
            "operator": "Synthetic operator, synthetic workstation",
            "environment": {"excel_version": "16.0", "excel_build": "16.0.20326.20072",
                            "office_bitness": "64-bit", "os": "Windows 11 Pro 24H2",
                            "os_architecture": "x64", "runtime": "VBA7+",
                            "references": list(self.policy["required_references"]),
                            "macro_policy": "Disable with notification", "vba_project_access": "Off",
                            "trust_changes": False},
            "sources": evidence.source_inventory(self.root, self.sha),
            "stages": {
                "import": self.stage("session.log", self.session),
                "compile": self.stage("session.log", self.session),
                "regression": self.stage("harness.log", self.log),
                "cleanup": self.stage("session.log", self.session),
            },
            "harness": {"entry_point": "TEST_Harness.RunTests", "cases": 4, "assertions": 6,
                        "failures": 0, "completeness": "COMPLETE", "expected_errors": [
                            {"case": "cdo-delta.invalid-tranche", "status": "PASS",
                             "detail": "Implied by the complete passing suite"}]},
            "suites": [self.suite(spec) for spec in self.policy["suites"]],
            "summary": None,
        }

    def git(self, *args: str) -> str:
        return subprocess.run(["git", *args], check=True, capture_output=True, text=True).stdout

    def commit(self, message: str) -> None:
        self.git("-C", str(self.root), "add", "--all")
        self.git("-C", str(self.root), "-c", "user.name=Fixture", "-c", "user.email=fixture@invalid",
                 "-c", "commit.gpgsign=false", "commit", "-q", "-m", message)

    def stage(self, name: str, text: str) -> dict[str, Any]:
        (self.bundle / name).write_text(text)
        return {"status": "PASS", "detail": "Synthetic",
                "log": {"path": name, "sha256": hashlib.sha256(text.encode()).hexdigest()}}

    def suite_log(self, spec: dict[str, Any], *, verdict: str = "PASS", cases: int | None = None,
                  checks: int | None = None, failures: int = 0, restored: bool = True,
                  mode: str = "NORMAL", candidate: str | None = None) -> str:
        """A report as the suite prints it, with the operator's CANDIDATE line first."""
        cases = spec["cases"] if cases is None else cases
        checks = spec["checks"]["64-bit"] if checks is None else checks
        result = f"RESULT={verdict}; cases={cases}; checks={checks}; failures={failures}"
        if spec["restoration"]:
            key, value = spec["restoration"].split("=")
            result += f"; {key}={value if restored else 'FAIL'}"
        names = [f"case-{index}" for index in range(cases)]
        if spec["first_case"] and names:
            names[0] = spec["first_case"] + ".crr"
        return "\n".join([f"CANDIDATE={candidate or self.sha}", spec["header"],
                          *(["MODE=" + mode] if spec["mode_line"] else []),
                          *(f"CASE={name}" for name in names), result, ""])

    def suite(self, spec: dict[str, Any], log: str | None = None, *, status: str = "PASS",
              **counts: int) -> dict[str, Any]:
        name = spec["entry_point"].split(".")[0] + ".log"
        text = self.suite_log(spec) if log is None else log
        (self.bundle / name).write_text(text)
        item = {"entry_point": spec["entry_point"], "status": status, "cases": spec["cases"],
                "checks": spec["checks"]["64-bit"], "failures": 0,
                "log": {"path": name, "sha256": hashlib.sha256(text.encode()).hexdigest()}}
        item.update(counts)
        return item

    def summary(self, lines: dict[str, str] | None = None, result: str | None = None) -> None:
        """Write a Test_RunAll summary consistent with the individual reports, unless overridden."""
        rows = [("TEST_Harness", "completeness=COMPLETE; cases=4; assertions=6; failures=0; cleanup=PASS")]
        for spec in self.policy["suites"]:
            detail = self.suite_log(spec).split("RESULT=PASS; ")[1].splitlines()[0]
            rows.append((spec["entry_point"].split(".")[0], detail))
        text = [f"CANDIDATE={self.sha}", "SACCR TEST SUMMARY 2026-10-16 08:51:30"]
        for name, detail in rows:
            text.append((lines or {}).get(name, f"  PASS  {name:<22}{detail}  (0,5 s)"))
        text += [result or f"RESULT=PASS; suites={len(rows)}; passed={len(rows)}; failed=0; seconds=57,7", ""]
        body = "\n".join(text)
        (self.bundle / "summary.log").write_text(body)
        self.record["summary"] = {"path": "summary.log", "sha256": hashlib.sha256(body.encode()).hexdigest()}

    def spec(self, module: str) -> dict[str, Any]:
        return next(spec for spec in self.policy["suites"] if spec["entry_point"].startswith(module + "."))

    def index(self, module: str) -> int:
        return [spec["entry_point"].split(".")[0] for spec in self.policy["suites"]].index(module)

    def not_run(self) -> dict[str, Any]:
        return {"status": "NOT_RUN", "detail": "Earlier stage failed", "log": None}

    def evaluate(self) -> dict[str, Any]:
        self.path.write_text(json.dumps(self.record))
        return evidence.evaluate(self.root, self.sha, self.path)

    def assertInvalid(self, fragment: str = "") -> None:
        report = self.evaluate()
        self.assertIn("EVIDENCE_INVALID", report["outcomes"], report)
        self.assertIn(fragment, " ".join(report["findings"]), report)

    def test_inventory_is_exact_import_set(self) -> None:
        paths = [item["path"] for item in self.record["sources"]]
        self.assertEqual(paths, ["src/core/CORE_Engine.bas", "src/forms/FSample.frm",
                                 "src/forms/FSample.frx", "src/modules/SACCR_Formulas.bas",
                                 "src/workbook/SACCR_Template.xlsx", "tests/modules/TEST_Harness.bas"])
        expected = hashlib.sha256(SOURCES["tests/modules/TEST_Harness.bas"].encode()).hexdigest()
        self.assertEqual(self.record["sources"][-1]["sha256"], expected)

    def test_valid_manual_record_passes_repeatably(self) -> None:
        first = self.evaluate()
        self.assertEqual(first["status"], "pass", first)
        self.assertEqual(first["outcomes"], ["PASS"])
        self.assertEqual(first, self.evaluate())

    def test_crlf_log_and_uppercase_digest_accepted(self) -> None:
        crlf = self.log.replace("\n", "\r\n")
        self.record["stages"]["regression"] = self.stage("harness.log", crlf)
        log = self.record["stages"]["regression"]["log"]
        log["sha256"] = log["sha256"].upper()
        self.assertEqual(self.evaluate()["status"], "pass")

    def test_wrong_source_identity(self) -> None:
        self.record["sources"][0]["sha256"] = "0" * 64
        self.assertInvalid("source inventory")

    def test_source_changed_after_candidate(self) -> None:
        (self.root / "src/core/CORE_Engine.bas").write_text("changed\n")
        self.commit("Later change")
        self.record["candidate_sha"] = self.git("-C", str(self.root), "rev-parse", "HEAD").strip()
        self.path.write_text(json.dumps(self.record))
        report = evidence.evaluate(self.root, self.record["candidate_sha"], self.path)
        self.assertIn("EVIDENCE_INVALID", report["outcomes"])

    def test_record_for_another_candidate(self) -> None:
        self.record["candidate_sha"] = "f" * 40
        self.assertInvalid("candidate_sha")
        self.record["candidate_sha"] = self.sha
        self.path.write_text(json.dumps(self.record))
        report = evidence.evaluate(self.root, "e" * 40, self.path)
        self.assertIn("not in this repository", " ".join(report["findings"]))

    def test_import_log_must_name_candidate(self) -> None:
        self.record["stages"]["import"] = self.stage("import.log", "Imported something\n")
        self.assertInvalid("candidate SHA")

    def test_environment_fields_required(self) -> None:
        original = copy.deepcopy(self.record)
        for key in self.record["environment"]:
            self.record = copy.deepcopy(original)
            self.record["environment"][key] = True if key == "trust_changes" else ""
            self.assertInvalid()

    def test_references_must_be_the_defaults(self) -> None:
        self.record["environment"]["references"].append("Microsoft Scripting Runtime")
        self.assertInvalid("references")

    def test_harness_environment_must_match_record(self) -> None:
        self.record["environment"]["office_bitness"] = "32-bit"
        self.assertInvalid("office")

    def test_harness_must_report_windows(self) -> None:
        mac = self.log.replace("os=Windows (64-bit) NT 10.00", "os=Macintosh (Intel) Kernel Version 23.0")
        self.record["stages"]["regression"] = self.stage("harness.log", mac)
        self.assertInvalid("os is not Windows")

    def test_incomplete_execution(self) -> None:
        self.record["stages"]["regression"] = self.not_run()
        self.record["harness"] = None
        self.assertInvalid("no suite can run")
        for item in self.record["suites"]:
            item.update(status="NOT_RUN", cases=None, checks=None, failures=None, log=None)
        self.assertEqual(self.evaluate()["outcomes"], ["INCOMPLETE"])
        self.record["harness"] = {"cases": 2}
        self.assertInvalid()

    def test_interrupted_run_cannot_pass(self) -> None:
        self.record["harness"]["completeness"] = "INCOMPLETE"
        self.assertInvalid("contradicts")
        self.record["harness"]["completeness"] = "COMPLETE"
        partial = self.log.replace("CASE=option-delta.repeatability\n", "")
        self.record["stages"]["regression"] = self.stage("harness.log", partial)
        self.assertInvalid("CASE")

    def test_failed_cleanup(self) -> None:
        self.record["stages"]["cleanup"]["status"] = "FAIL"
        self.assertEqual(self.evaluate()["outcomes"], ["CLEANUP_FAILED"])

    def test_harness_cleanup_failure_must_be_recorded(self) -> None:
        failed = self.log.replace("CLEANUP=PASS", "CLEANUP=FAIL").replace(
            "RESULT=PASS", "RESULT=FAIL").replace("cleanup=PASS", "cleanup=FAIL")
        self.record["stages"]["regression"] = self.stage("harness.log", failed)
        self.assertInvalid("cleanup")
        self.record["stages"]["cleanup"]["status"] = "FAIL"
        self.assertEqual(self.evaluate()["outcomes"], ["CLEANUP_FAILED"])

    def failed_regression(self, log: str, **harness: Any) -> None:
        self.record["stages"]["regression"] = self.stage("harness.log", log)
        self.record["stages"]["regression"]["status"] = "FAIL"
        self.record["harness"].update(harness)

    def test_test_failure(self) -> None:
        failed = self.log.replace("FAILURES=0", "FAILURES=1\nFAILURE_DETAILS=replacement-cost.exact: expected=71").replace(
            "RESULT=PASS", "RESULT=FAIL").replace("failures=0", "failures=1")
        self.failed_regression(failed, failures=1)
        self.record["harness"]["expected_errors"][0]["status"] = "FAIL"
        self.assertEqual(self.evaluate()["outcomes"], ["TEST_FAILED"])

    def test_interrupted_failure_keeps_observed_counts(self) -> None:
        cases = self.policy["cases"][:3]
        partial = "\n".join([
            *self.log.splitlines()[:4], *("CASE=" + case for case in cases),
            "CASES=3", "ASSERTIONS=5", "FAILURES=1", "CLEANUP=PASS; detail=unchanged",
            "FAILURE_DETAILS=runner error 11", "RESULT=FAIL; completeness=INCOMPLETE; cases=3; "
            "assertions=5; failures=1; cleanup=PASS", ""])
        self.failed_regression(partial, cases=3, assertions=5, failures=1, completeness="INCOMPLETE")
        self.assertEqual(self.evaluate()["outcomes"], ["TEST_FAILED"])
        self.record["harness"]["completeness"] = "COMPLETE"
        self.assertInvalid("completeness")

    def test_failure_record_must_match_its_log(self) -> None:
        self.failed_regression(self.log, failures=1)
        self.assertInvalid("FAILURES")
        self.failed_regression(self.log, failures=0)
        self.assertInvalid("passing report")
        injected = self.log.replace("MODE=NORMAL", "MODE=INJECTED_FAILURE")
        self.failed_regression(injected, failures=0)
        self.assertInvalid("normal-mode")

    def test_timeout_keeps_log_without_claims(self) -> None:
        for kept in (1, 2, 6, 12):
            partial = "\n".join([*self.log.splitlines()[:kept], ""])
            self.record["stages"]["regression"] = self.stage("harness.log", partial)
            self.record["stages"]["regression"]["status"] = "TIMEOUT"
            self.record["harness"] = None
            self.assertEqual(self.evaluate()["outcomes"], ["EXECUTION_TIMEOUT"], kept)
        self.record["harness"] = {"cases": 2}
        self.assertInvalid("no harness results")
        self.record["harness"] = None
        self.record["stages"]["regression"] = self.stage("harness.log", "SACCR TESTS\n")
        self.record["stages"]["regression"]["status"] = "TIMEOUT"
        self.assertInvalid("exactly one CANDIDATE")
        for finished in (self.log, self.log.replace("RESULT=PASS", "RESULT=FAIL")):
            self.record["stages"]["regression"] = self.stage("harness.log", finished)
            self.record["stages"]["regression"]["status"] = "TIMEOUT"
            self.assertInvalid("not TIMEOUT")

    def test_compile_and_import_failures(self) -> None:
        self.record["stages"]["compile"]["status"] = "FAIL"
        self.assertInvalid("regression cannot run")
        self.record["stages"]["regression"] = self.not_run()
        self.record["harness"] = None
        for item in self.record["suites"]:
            item.update(status="NOT_RUN", cases=None, checks=None, failures=None, log=None)
        self.assertEqual(self.evaluate()["outcomes"], ["COMPILE_FAILED", "INCOMPLETE"])
        self.record["stages"]["import"]["status"] = "TIMEOUT"
        self.assertInvalid("compile cannot run")
        self.record["stages"]["compile"] = self.not_run()
        self.assertEqual(self.evaluate()["outcomes"], ["EXECUTION_TIMEOUT", "INCOMPLETE"])

    def test_injected_failure_or_repeated_runs_rejected(self) -> None:
        for raw in (self.log.replace("MODE=NORMAL", "MODE=INJECTED_FAILURE"), self.log + self.log):
            self.record["stages"]["regression"] = self.stage("harness.log", raw)
            self.assertInvalid()

    def test_counts_and_expected_errors(self) -> None:
        original = copy.deepcopy(self.record)
        for field, value in (("cases", 3), ("assertions", True), ("failures", 1),
                             ("entry_point", "TEST_Harness.RunTestsWithInjectedFailure"),
                             ("expected_errors", [])):
            self.record = copy.deepcopy(original)
            self.record["harness"][field] = value
            self.assertInvalid()
        self.record = copy.deepcopy(original)
        self.record["harness"]["expected_errors"][0]["status"] = "FAIL"
        self.assertInvalid()

    def test_log_tamper_missing_traversal_and_symlink(self) -> None:
        (self.bundle / "harness.log").write_text("changed")
        self.assertInvalid("digest mismatch")
        (self.bundle / "harness.log").unlink()
        self.assertInvalid("missing log")
        self.record["stages"]["regression"]["log"]["path"] = "../outside.log"
        self.assertInvalid("unsafe")
        outside = self.area / "outside.log"
        outside.write_text(self.log)
        (self.bundle / "harness.log").symlink_to(outside)
        self.record["stages"]["regression"]["log"]["path"] = "harness.log"
        self.assertInvalid("symlink")

    def test_unavailable_is_never_pass(self) -> None:
        self.record.update(execution="unavailable", availability_reason="No Excel host",
                           operator=None, environment=None, harness=None, sources=[], stages={},
                           suites=[], summary=None)
        report = self.evaluate()
        self.assertEqual((report["status"], report["outcomes"]), ("fail", ["UNAVAILABLE"]))
        self.record["harness"] = {"cases": 4}
        self.assertInvalid()

    def test_automated_execution_not_supported(self) -> None:
        self.record["execution"] = "automated"
        self.assertInvalid("manual")

    def test_timestamps_need_order_and_timezone(self) -> None:
        self.record["finished_at"] = "2026-10-06T08:00:00+02:00"
        self.assertInvalid("precedes")
        self.record["finished_at"] = "2026-10-06T09:10:00"
        self.assertInvalid("timezone")

    def test_duplicate_keys_rejected(self) -> None:
        text = json.dumps(self.record)
        self.path.write_text(text[:-1] + ', "execution": "manual"}')
        report = evidence.evaluate(self.root, self.sha, self.path)
        self.assertIn("duplicate JSON key", " ".join(report["findings"]))

    def test_harness_log_must_name_the_candidate(self) -> None:
        unbound = self.log.replace(f"CANDIDATE={self.sha}\n", "")
        self.record["stages"]["regression"] = self.stage("harness.log", unbound)
        self.assertInvalid("exactly one CANDIDATE")
        other = self.log.replace(self.sha, "f" * 40)
        self.record["stages"]["regression"] = self.stage("harness.log", other)
        self.assertInvalid("another candidate")

    def test_every_suite_is_required_in_order(self) -> None:
        del self.record["suites"][-1]
        self.assertInvalid("suites differ from the policy")
        self.record["suites"] = [self.suite(spec) for spec in self.policy["suites"]]
        self.record["suites"].reverse()
        self.assertInvalid("policy order")

    def test_suite_log_from_another_candidate(self) -> None:
        spec = self.spec("TEST_Reconcile")
        log = self.suite_log(spec, candidate="e" * 40)
        self.record["suites"][self.index("TEST_Reconcile")] = self.suite(spec, log)
        self.assertInvalid("another candidate")

    def test_stale_counts_are_rejected(self) -> None:
        spec = self.spec("TEST_Reconcile")
        stale = self.suite_log(spec, cases=18, checks=137)
        self.record["suites"][self.index("TEST_Reconcile")] = self.suite(spec, stale, cases=18, checks=137)
        self.assertInvalid("counts differ from the policy")

    def test_suite_record_must_match_its_log(self) -> None:
        spec = self.spec("TEST_Formulas")
        self.record["suites"][self.index("TEST_Formulas")] = self.suite(spec, checks=81)
        self.assertInvalid("RESULT counts differ")
        (self.bundle / "TEST_Formulas.log").write_text("changed")
        self.record["suites"][self.index("TEST_Formulas")]["checks"] = 82
        self.assertInvalid("digest mismatch")

    def test_suite_log_must_be_one_normal_run(self) -> None:
        spec = self.spec("TEST_Cases")
        position = self.index("TEST_Cases")
        for log in (self.suite_log(spec, mode="INJECTED_FAILURE"), self.suite_log(spec) * 2,
                    self.suite_log(spec).replace(spec["header"], "SACCR TESTS")):
            self.record["suites"][position] = self.suite(spec, log)
            self.assertInvalid(spec["entry_point"])

    def test_restoration_failure_cannot_pass(self) -> None:
        spec = self.spec("TEST_MainState")
        log = self.suite_log(spec, restored=False)
        self.record["suites"][self.index("TEST_MainState")] = self.suite(spec, log)
        self.assertInvalid("did not restore")

    def test_failed_suite(self) -> None:
        spec = self.spec("TEST_InputValidation")
        log = self.suite_log(spec, verdict="FAIL", failures=2)
        self.record["suites"][self.index("TEST_InputValidation")] = self.suite(
            spec, log, status="FAIL", failures=2)
        self.assertInvalid("regression PASS cannot hide a failed suite")
        self.record["stages"]["regression"]["status"] = "FAIL"
        self.assertEqual(self.evaluate()["outcomes"], ["TEST_FAILED"])
        self.record["suites"][self.index("TEST_InputValidation")] = self.suite(
            spec, self.suite_log(spec), status="FAIL")
        self.assertInvalid("contradicts a passing report")
        self.record["suites"][self.index("TEST_InputValidation")] = self.suite(
            spec, log, status="FAIL", failures=2)
        self.record["harness"]["expected_errors"][0]["status"] = "FAIL"
        self.assertInvalid("a passing harness needs every expected-error result")

    def test_one_failed_log_cannot_serve_two_suites(self) -> None:
        self.record["stages"]["regression"]["status"] = "FAIL"
        failed = self.suite_log(self.spec("TEST_Cases"), verdict="FAIL", cases=3, checks=5, failures=1)
        # A copy padded with blank lines is still TEST_Cases' report: its first case says so.
        for module, padding in (("TEST_Cases", ""), ("TEST_Invariants", "\n\n")):
            self.record["suites"][self.index(module)] = self.suite(
                self.spec(module), failed + padding, status="FAIL", cases=3, checks=5, failures=1)
        self.assertInvalid("TEST_Invariants.RunInvariantTests log starts with another suite's case")
        # A run that stopped before its first case names no case; a padded copy is still one run.
        empty = self.suite_log(self.spec("TEST_Cases"), verdict="FAIL", cases=0, checks=0, failures=1)
        padded = empty.replace("\n", " \n", 1) + "\n \n"
        for module, log in (("TEST_Cases", empty), ("TEST_Invariants", padded)):
            self.record["suites"][self.index(module)] = self.suite(
                self.spec(module), log, status="FAIL", cases=0, checks=0, failures=1)
        self.assertInvalid("the same log is cited for two test runs")
        too_many = self.suite_log(self.spec("TEST_Invariants"), verdict="FAIL", cases=5, failures=1)
        self.record["suites"][self.index("TEST_Invariants")] = self.suite(
            self.spec("TEST_Invariants"), too_many, status="FAIL", cases=5, failures=1)
        self.assertInvalid("more cases than the policy")

    def test_unfinished_suites(self) -> None:
        position = self.index("TEST_Aggregation")
        self.record["suites"][position].update(status="NOT_RUN", cases=None, checks=None,
                                               failures=None, log=None)
        self.assertEqual(self.evaluate()["outcomes"], ["INCOMPLETE"])
        spec = self.spec("TEST_Aggregation")
        hung = self.suite(spec, f"CANDIDATE={self.sha}\nSACCR CASE TESTS\nCASE=case-0\n", status="TIMEOUT")
        hung.update(cases=None, checks=None, failures=None)
        self.record["suites"][position] = hung
        self.assertEqual(self.evaluate()["outcomes"], ["EXECUTION_TIMEOUT"])
        hung.update(cases=11)
        self.assertInvalid("has no results")
        foreign = self.suite(spec, "CANDIDATE=" + "e" * 40 + "\nSACCR CASE TESTS\n", status="TIMEOUT")
        foreign.update(cases=None, checks=None, failures=None)
        self.record["suites"][position] = foreign
        self.assertInvalid("another candidate")

    def test_summary_must_agree_with_the_reports(self) -> None:
        self.summary()
        self.assertEqual(self.evaluate()["outcomes"], ["PASS"])
        self.summary({"TEST_Reconcile": "  PASS  TEST_Reconcile        cases=18; checks=137; failures=0  (8,7 s)"})
        self.assertInvalid("contradicts its own report")
        self.summary(result="RESULT=PASS; suites=8; passed=7; failed=1; seconds=1,0")
        self.assertInvalid("summary RESULT")
        self.summary({"TEST_Cases": ""})
        self.assertInvalid("summary suites differ")

    def test_policy_matches_the_test_modules(self) -> None:
        self.assertEqual(evidence.policy_problems(ROOT), [])
        copy_root = self.area / "copy"
        shutil.copytree(ROOT / "tests" / "modules", copy_root / "tests" / "modules")
        shutil.copy(ROOT / "README.md", copy_root / "README.md")
        (copy_root / ".github").mkdir()
        policy = copy.deepcopy(self.policy)
        policy["suites"][-1]["cases"] = 18
        policy["suites"][0]["checks"]["64-bit"] = 81
        (copy_root / evidence.POLICY).write_text(json.dumps(policy))
        problems = " ".join(evidence.policy_problems(copy_root))
        self.assertIn("TEST_Reconcile.RunReconcileTests: policy cases differ", problems)
        self.assertIn("TEST_Formulas.RunFormulaTests: policy 64-bit checks differ", problems)
        policy = copy.deepcopy(self.policy)
        policy["suites"][0]["checks"] = {"32-bit": 82, "64-bit": 80}
        (copy_root / evidence.POLICY).write_text(json.dumps(policy))
        swapped = " ".join(evidence.policy_problems(copy_root))
        self.assertIn("policy 32-bit checks differ", swapped)
        self.assertIn("policy 64-bit checks differ", swapped)
        policy = copy.deepcopy(self.policy)
        del policy["suites"][1]
        (copy_root / evidence.POLICY).write_text(json.dumps(policy))
        self.assertIn("not the policy suites", " ".join(evidence.policy_problems(copy_root)))
        policy = copy.deepcopy(self.policy)
        policy["suites"][self.index("TEST_Invariants")]["first_case"] = "invariants-demo-repeat"
        (copy_root / evidence.POLICY).write_text(json.dumps(policy))
        self.assertIn("TEST_Invariants.RunInvariantTests: policy first_case differs",
                      " ".join(evidence.policy_problems(copy_root)))
        policy["suites"][self.index("TEST_Aggregation")]["first_case"] = None
        with self.assertRaisesRegex(ValueError, "share a header need their first case"):
            evidence.parse_policy(json.dumps(policy).encode())
        policy["suites"][self.index("TEST_Aggregation")]["first_case"] = "invariants-demo-repeat"
        with self.assertRaisesRegex(ValueError, "distinct first cases"):
            evidence.parse_policy(json.dumps(policy).encode())

    def test_readme_counts_match_the_policy(self) -> None:
        copy_root = self.area / "copy"
        shutil.copytree(ROOT / "tests" / "modules", copy_root / "tests" / "modules")
        (copy_root / ".github").mkdir()
        shutil.copy(ROOT / evidence.POLICY, copy_root / evidence.POLICY)
        readme = (ROOT / "README.md").read_text(encoding="utf-8")
        for old, new, message in (
                ("Excel_suites-8%2F8_PASS", "Excel_suites-7%2F7_PASS", "Excel suites badge"),
                ("Regression-227_cases", "Regression-226_cases", "Regression badge"),
                ("| 19 | 140 |", "| 18 | 137 |", "suite table rows differ"),
                ("| `TEST_Formulas` |", "| `TEST_Formulas` | stale | 9 | 80 |\n| `TEST_Formulas` |",
                 "suite table rows differ"),
                ("| `TEST_Formulas` |", "| `TEST_Formulas` | stale | **9** | **80** |\n| `TEST_Formulas` |",
                 "suite table rows differ"),
                ("| **227** | **975** |", "| **226** | **975** |", "suite table total"),
                ("TEST_Reconcile        cases=19; checks=140", "TEST_Reconcile        cases=18; checks=137",
                 "Test_RunAll output differs"),
                ("RESULT=PASS; suites=8; passed=8", "RESULT=PASS; suites=7; passed=7", "Test_RunAll result"),
                # A stale copy beside the corrected display contradicts it.
                ("[![Excel suites]", "![stale](https://img.shields.io/badge/Excel_suites-7%2F7_PASS-2ea44f)\n"
                 "[![Excel suites]", "Excel suites badge"),
                ("| **Total** |", "| **Total** | | **226** | **975** |\n| **Total** |", "suite table total"),
                ("RESULT=PASS; suites=8", "RESULT=FAIL; suites=8; passed=7; failed=1; seconds=1\n"
                 "RESULT=PASS; suites=8", "Test_RunAll result"),
                ("Excel_suites-8%2F8_PASS", "Excel_suites-8%2F8_FAIL", "Excel suites badge"),
                ("Eight Excel suites (227", "Seven Excel suites (**226** cases, 975 checks); eight (227",
                 "Testing capability"),
                ("| `TEST_Formulas` |", "   | `TEST_Formulas` | stale | **9** | **80** |\n| `TEST_Formulas` |",
                 "suite table rows differ"),
                ("  PASS  TEST_MainState", "  PASS  TEST_MainState        cases=20; checks=381; failures=0; "
                 "caller_state=RESTORED\n  PASS  TEST_MainState", "lines all parse"),
                ("SACCR TEST SUMMARY 2026", "SACCR TEST SUMMARY 2026-10-13 08:28:48\nSACCR TEST SUMMARY 2026",
                 "lines all parse")):
            with self.subTest(message=message):
                self.assertIn(old, readme)
                (copy_root / "README.md").write_text(readme.replace(old, new), encoding="utf-8")
                self.assertIn(message, " ".join(evidence.policy_problems(copy_root)))
        # Counts in other prose are not the Testing capability.
        prose = readme.replace("## Test suites", "Benchmarks (3 cases and 4 checks) differ.\n\n## Test suites")
        (copy_root / "README.md").write_text(prose, encoding="utf-8")
        self.assertEqual(evidence.policy_problems(copy_root), [])

    def test_cli_exit_codes_inventory_and_output_guard(self) -> None:
        self.path.write_text(json.dumps(self.record))
        base = ["--root", str(self.root), "--candidate-sha", self.sha]
        with contextlib.redirect_stdout(io.StringIO()) as out:
            self.assertEqual(evidence.main(base + ["--inventory"]), 0)
        self.assertEqual(json.loads(out.getvalue()), self.record["sources"])
        report = self.area / "report.json"
        with contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(evidence.main(base + ["--evidence", str(self.path),
                                                   "--output", str(report)]), 0)
        self.assertEqual(json.loads(report.read_text())["outcomes"], ["PASS"])
        self.record["stages"]["cleanup"]["status"] = "FAIL"
        self.path.write_text(json.dumps(self.record))
        with contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(evidence.main(base + ["--evidence", str(self.path)]), 1)
        with contextlib.redirect_stderr(io.StringIO()), self.assertRaises(SystemExit):
            evidence.main(base + ["--evidence", str(self.path),
                                  "--output", str(self.bundle / "report.json")])


if __name__ == "__main__":
    unittest.main()
