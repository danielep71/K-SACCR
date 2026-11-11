"""Synthetic certification bundles for write_excel_evidence.py; none of these tests runs Excel."""
from __future__ import annotations

import contextlib
import io
import json
import shutil
import subprocess
import tempfile
import unittest
from pathlib import Path
from typing import Any

import check_excel_evidence as evidence
import write_excel_evidence as writer

ROOT = Path(__file__).resolve().parents[1]
SOURCES = {
    "src/core/CORE_Engine.bas": "Attribute VB_Name = \"CORE_Engine\"\n",
    "src/workbook/SACCR_Template.xlsx": "template bytes\n",
    "tests/modules/TEST_Harness.bas": "Attribute VB_Name = \"TEST_Harness\"\n",
}


class WriteEvidenceTests(unittest.TestCase):
    def setUp(self) -> None:
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        area = Path(temporary.name)
        self.root = area / "candidate"
        self.bundle = area / "bundle"
        self.bundle.mkdir()
        self.git("init", "-q", str(self.root))
        (self.root / ".github").mkdir()
        shutil.copy(ROOT / evidence.POLICY, self.root / evidence.POLICY)
        for name, text in SOURCES.items():
            target = self.root / name
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_text(text)
        self.git("-C", str(self.root), "add", "--all")
        self.git("-C", str(self.root), "-c", "user.name=Fixture", "-c", "user.email=fixture@invalid",
                 "-c", "commit.gpgsign=false", "commit", "-q", "-m", "Synthetic candidate")
        self.sha = self.git("-C", str(self.root), "rev-parse", "HEAD").strip()
        self.policy: dict[str, Any] = json.loads((ROOT / evidence.POLICY).read_text())
        self.session: dict[str, Any] = {
            "operator": "Synthetic operator, synthetic workstation",
            "started_at": "2026-11-04T09:00:00+01:00", "finished_at": "2026-11-04T09:10:00+01:00",
            "excel_version": "16.0", "excel_build": "16.0.20326.20072", "office_bitness": "64-bit",
            "os": "Windows 11 Pro 24H2", "os_architecture": "x64", "runtime": "VBA7+",
            "references": list(self.policy["required_references"]),
            "macro_policy": "Disable VBA macros with notification", "vba_project_access": "On",
            "import": {"status": "PASS", "detail": "Sync-SACCR-VBA -FromTemplate, development"},
            "compile": {"status": "PASS", "detail": "Compile VBAProject completed"},
            "workbook_closed": True,
        }
        self.write("session.log", f"Candidate {self.sha}\nImported and compiled.\n")
        self.write("harness.log", self.harness_log())
        for spec in self.policy["suites"]:
            self.write(writer.log_name(spec["entry_point"]), self.suite_log(spec))

    def git(self, *args: str) -> str:
        return subprocess.run(["git", *args], check=True, capture_output=True, text=True).stdout

    def write(self, name: str, text: str) -> None:
        # Certify writes CRLF line ends.
        (self.bundle / name).write_bytes(text.replace("\n", "\r\n").encode("utf-8"))

    def harness_log(self, result: str | None = None) -> str:
        lines = [f"CANDIDATE={self.sha}", "SACCR TESTS", "MODE=NORMAL",
                 "ENVIRONMENT=host=Microsoft Excel; version=16.0; os=Windows (64-bit) NT 10.00; "
                 "office=64-bit; runtime=VBA7+",
                 *("CASE=" + case for case in self.policy["cases"]),
                 "CASES=4", "ASSERTIONS=6", "FAILURES=0", "CLEANUP=PASS; detail=unchanged"]
        if result != "none":
            lines.append(result or "RESULT=PASS; completeness=COMPLETE; cases=4; assertions=6; "
                                   "failures=0; cleanup=PASS")
        return "\n".join(lines + [""])

    def suite_log(self, spec: dict[str, Any], verdict: str = "PASS", failures: int = 0) -> str:
        result = (f"RESULT={verdict}; cases={spec['cases']}; checks={spec['checks']['64-bit']}; "
                  f"failures={failures}; restore=PASS")
        names = [f"case-{index}" for index in range(spec["cases"])]
        return "\n".join([f"CANDIDATE={self.sha}", "SUITE=" + spec["entry_point"].split(".")[0], "MODE=NORMAL",
                          *(f"CASE={name}" for name in names), result, ""])

    def summary_log(self) -> str:
        rows = [("TEST_Harness", "completeness=COMPLETE; cases=4; assertions=6; failures=0; cleanup=PASS")]
        for spec in self.policy["suites"]:
            rows.append((spec["entry_point"].split(".")[0],
                         self.suite_log(spec).split("RESULT=PASS; ")[1].splitlines()[0]))
        text = [f"CANDIDATE={self.sha}", "SACCR TEST SUMMARY 2026-11-04 09:05:00"]
        text += [f"  PASS  {name:<22}{detail}  (0,5 s)" for name, detail in rows]
        text.append(f"RESULT=PASS; suites={len(rows)}; passed={len(rows)}; failed=0; seconds=57,7")
        return "\n".join(text + [""])

    def run_writer(self) -> tuple[int, dict[str, Any]]:
        (self.bundle / "session.json").write_text(json.dumps(self.session))
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            code = writer.main(["--root", str(self.root), "--candidate-sha", self.sha,
                                "--bundle", str(self.bundle)])
        record = json.loads((self.bundle / "host.json").read_text()) if (self.bundle / "host.json").exists() else {}
        return code, record

    def outcomes(self) -> list[str]:
        return list(evidence.evaluate(self.root, self.sha, self.bundle / "host.json")["outcomes"])

    def test_complete_passing_run_validates(self) -> None:
        code, record = self.run_writer()
        self.assertEqual(code, 0)
        self.assertEqual(self.outcomes(), ["PASS"])
        self.assertEqual(record["environment"]["office_bitness"], "64-bit")
        self.assertIsNone(record["summary"])

    def test_summary_is_cited_when_written(self) -> None:
        self.write("summary.log", self.summary_log())
        code, record = self.run_writer()
        self.assertEqual((code, record["summary"]["path"]), (0, "summary.log"))

    def test_interrupted_summary_is_not_ignored(self) -> None:
        # Test_RunAll stopped after a failed suite row, before its RESULT line.
        lines = self.summary_log().splitlines()
        lines[3] = lines[3].replace("PASS", "FAIL", 1)
        self.write("summary.log", "\n".join(lines[:-1] + [""]))
        code, record = self.run_writer()
        self.assertEqual((code, record["summary"]["path"]), (1, "summary.log"))
        self.assertNotIn("PASS", self.outcomes())

    def test_harness_cleanup_failure_keeps_its_expected_errors(self) -> None:
        log = self.harness_log("RESULT=FAIL; completeness=COMPLETE; cases=4; assertions=6; failures=0; "
                               "cleanup=FAIL")
        self.write("harness.log", log.replace("CLEANUP=PASS; detail=unchanged", "CLEANUP=FAIL; detail=changed"))
        code, record = self.run_writer()
        self.assertEqual(code, 1)
        self.assertTrue(all(error["status"] == "PASS" for error in record["harness"]["expected_errors"]))
        self.assertNotIn("EVIDENCE_INVALID", self.outcomes())
        self.assertIn("CLEANUP_FAILED", self.outcomes())

    def test_failed_suite_is_a_test_failure(self) -> None:
        spec = self.policy["suites"][2]
        self.write(writer.log_name(spec["entry_point"]), self.suite_log(spec, "FAIL", 1))
        code, record = self.run_writer()
        self.assertEqual(code, 1)
        self.assertEqual(record["stages"]["regression"]["status"], "FAIL")
        self.assertEqual(self.outcomes(), ["TEST_FAILED"])

    def test_missing_suite_log_is_incomplete(self) -> None:
        (self.bundle / writer.log_name(self.policy["suites"][-1]["entry_point"])).unlink()
        code, record = self.run_writer()
        self.assertEqual((code, record["suites"][-1]["status"]), (1, "NOT_RUN"))
        self.assertEqual(self.outcomes(), ["INCOMPLETE"])

    def test_suite_stopped_by_an_error_is_a_timeout(self) -> None:
        spec = self.policy["suites"][0]
        self.write(writer.log_name(spec["entry_point"]),
                   f"CANDIDATE={self.sha}\nSUITE={spec['entry_point'].split('.')[0]}\nERROR=91: Object variable not set\n")
        code, record = self.run_writer()
        self.assertEqual((code, record["suites"][0]["status"]), (1, "TIMEOUT"))
        self.assertEqual(self.outcomes(), ["EXECUTION_TIMEOUT"])

    def test_harness_without_result_is_a_timeout(self) -> None:
        self.write("harness.log", self.harness_log("none"))
        code, record = self.run_writer()
        self.assertEqual((code, record["stages"]["regression"]["status"]), (1, "TIMEOUT"))
        self.assertIn("EXECUTION_TIMEOUT", self.outcomes())

    def test_failed_compile_runs_nothing(self) -> None:
        self.session["compile"] = {"status": "FAIL", "detail": "Compile error in TEST_Suite"}
        code, record = self.run_writer()
        self.assertEqual(code, 1)
        self.assertEqual(record["stages"]["regression"]["status"], "NOT_RUN")
        self.assertTrue(all(item["status"] == "NOT_RUN" for item in record["suites"]))
        self.assertEqual(self.outcomes(), ["COMPILE_FAILED", "INCOMPLETE"])

    def test_workbook_left_open_fails_cleanup(self) -> None:
        self.session["workbook_closed"] = False
        code, _ = self.run_writer()
        self.assertEqual(code, 1)
        self.assertEqual(self.outcomes(), ["CLEANUP_FAILED"])

    def test_edited_log_after_writing_is_rejected(self) -> None:
        self.run_writer()
        with (self.bundle / "harness.log").open("ab") as log:
            log.write(b"edited\r\n")
        self.assertEqual(self.outcomes(), ["EVIDENCE_INVALID"])

    def test_record_is_never_replaced(self) -> None:
        self.run_writer()
        with contextlib.redirect_stderr(io.StringIO()):
            code = writer.main(["--root", str(self.root), "--candidate-sha", self.sha,
                                "--bundle", str(self.bundle)])
        self.assertEqual(code, 2)


if __name__ == "__main__":
    unittest.main()
