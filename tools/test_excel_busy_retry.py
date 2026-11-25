"""The OLE message filter that retries calls Excel rejects while busy (#773).

Every script that starts Excel must register the filter of
tools/Excel-BusyRetry.ps1 before creating Excel.Application, so that a call
Excel rejects as "retry later" (RPC_E_CALL_REJECTED) is retried instead of
stopping the script. The decision rule is checked in PowerShell 7, where the
filter's C# compiles; registering needs Windows and is not run here.
"""

from __future__ import annotations

import re
import shutil
import subprocess
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TOOLS = ROOT / "tools"
HELPER = TOOLS / "Excel-BusyRetry.ps1"
PWSH = shutil.which("pwsh")
CREATES_EXCEL = "New-Object -ComObject Excel.Application"
DOT_SOURCE = '. (Join-Path $PSScriptRoot "Excel-BusyRetry.ps1")'
REGISTER = "[void](Register-ExcelBusyRetry)"


class ScriptsRegisterTheFilter(unittest.TestCase):
    def test_every_script_that_starts_excel_registers_the_filter_first(self) -> None:
        scripts = [path for path in sorted(TOOLS.glob("*.ps1"))
                   if CREATES_EXCEL in path.read_text(encoding="utf-8-sig")]
        self.assertGreaterEqual(len(scripts), 5, [path.name for path in scripts])
        for path in scripts:
            text = path.read_text(encoding="utf-8-sig")
            with self.subTest(script=path.name):
                self.assertIn(DOT_SOURCE, text)
                self.assertIn(REGISTER, text)
                self.assertLess(text.index(DOT_SOURCE), text.index(REGISTER))
                self.assertLess(text.index(REGISTER), text.index(CREATES_EXCEL))

    def test_the_verify_log_reports_the_retries(self) -> None:
        text = (TOOLS / "Verify-SACCR-Branch.ps1").read_text(encoding="utf-8-sig")
        self.assertIn("Write-Log (Get-ExcelBusyRetrySummary)", text)

    def test_the_filter_retries_only_retry_later_within_its_limit(self) -> None:
        text = HELPER.read_text(encoding="utf-8")
        self.assertIn('Guid("00000016-0000-0000-C000-000000000046")', text)
        self.assertRegex(text, r"rejectType == SERVERCALL_RETRYLATER && tickCount < TimeoutMs")
        self.assertRegex(text, r"public const int SERVERCALL_RETRYLATER = 2;")
        self.assertRegex(text, r"public static int TimeoutMs = 120000;")
        self.assertIn(b"\r\n", HELPER.read_bytes())


@unittest.skipUnless(PWSH, "PowerShell 7 (pwsh) is not installed")
class FilterInPowerShell(unittest.TestCase):
    def run_pwsh(self, script: str) -> str:
        command = f". '{HELPER}'; . '{HELPER}'; {script}"
        result = subprocess.run([PWSH or "pwsh", "-NoProfile", "-NonInteractive", "-Command", command],
                                capture_output=True, text=True, timeout=180)
        self.assertEqual(result.returncode, 0, result.stderr)
        return result.stdout.strip()

    def test_the_decision_rule(self) -> None:
        output = self.run_pwsh(
            "$t = [KSaccr.ExcelBusyRetry]; "
            "@($t::Decide(0, 2), $t::Decide(119999, 2), $t::Decide(120000, 2), "
            "$t::Decide(0, 1), $t::Decide(0, 0)) -join ','")
        self.assertEqual(output, "250,250,-1,-1,-1")

    def test_registering_off_an_sta_thread_does_nothing(self) -> None:
        output = self.run_pwsh(
            "if ([Threading.Thread]::CurrentThread.GetApartmentState() -eq 'STA') { 'STA' } "
            "else { [string](Register-ExcelBusyRetry); Unregister-ExcelBusyRetry; Get-ExcelBusyRetrySummary }")
        if output == "STA":
            self.skipTest("this PowerShell runs on an STA thread; registering is checked on Windows")
        lines = output.splitlines()
        self.assertEqual(lines[0], "False")
        self.assertTrue(re.fullmatch(r"Excel busy: 0 call\(s\) retried, 0 abandoned after 120 s", lines[1]),
                        lines[1])


if __name__ == "__main__":
    unittest.main()
