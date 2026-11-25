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

    def test_the_verify_log_reports_the_retries_and_names_the_windows(self) -> None:
        text = (TOOLS / "Verify-SACCR-Branch.ps1").read_text(encoding="utf-8-sig")
        self.assertIn("Write-Log (Get-ExcelBusyRetrySummary)", text)
        self.assertIn("$script:excelProcess = Get-ExcelProcessId $excel", text)
        self.assertLess(text.index("$script:excelProcess = 0"), text.index(CREATES_EXCEL))
        # When a call is abandoned, and when the workbook cannot be closed, the log
        # names the windows Excel shows, a dialog with its text.
        self.assertEqual(text.count("Write-Log (Get-ExcelWindowsSummary $script:excelProcess)"), 2)
        # A dialog's text cannot be read, so it is pictured into the artifact's sheets folder.
        self.assertIn('Write-Log (Save-ExcelDialogPictures $script:excelProcess (Join-Path $Bundle "sheets"))', text)

    def test_the_review_lets_excel_idle_on_its_workbook_before_the_picture_workbook(self) -> None:
        text = (TOOLS / "Review-SACCR-Sheets.ps1").read_text(encoding="utf-8-sig")
        activate = text.index("[void]$Workbook.Activate()")
        idle = text.index("Start-Sleep -Seconds $IdleSecondsBeforePictures")
        add = text.index("$pictureBook = $Excel.Workbooks.Add()")
        self.assertLess(activate, idle)
        self.assertLess(idle, add)
        self.assertIn("$IdleSecondsBeforePictures = 2", text)

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

    def test_the_windows_summary_needs_a_process_and_never_raises(self) -> None:
        output = self.run_pwsh("Get-ExcelWindowsSummary 0; Get-ExcelWindowsSummary 4242")
        lines = output.splitlines()
        self.assertEqual(lines[0], "Excel windows: the process is not known")
        self.assertTrue(lines[1].startswith("Excel windows: "), lines[1])

    def test_dialog_pictures_need_a_process_and_windows_powershell(self) -> None:
        output = self.run_pwsh("Save-ExcelDialogPictures 0 '.'; Save-ExcelDialogPictures 4242 '.'")
        lines = output.splitlines()
        self.assertEqual(lines[0], "Excel dialogs: the process is not known")
        self.assertEqual(lines[1], "Excel dialogs: no picture taken here (System.Drawing is not available)")

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
