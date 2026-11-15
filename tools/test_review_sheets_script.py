"""Offline tests for functions of tools/Review-SACCR-Sheets.ps1 that need no
Excel (#706). They run in PowerShell 7 (pwsh) and are skipped where pwsh is
not installed; fake objects stand in for Excel's ranges."""
from __future__ import annotations

import os
import shutil
import subprocess
import unittest
from pathlib import Path

SCRIPT = Path(__file__).resolve().parent / "Review-SACCR-Sheets.ps1"
PWSH = shutil.which("pwsh")


def run(body: str) -> str:
    code = "$ErrorActionPreference = 'Stop'\n. $env:REVIEW_SCRIPT\n" + body
    completed = subprocess.run([str(PWSH), "-NoProfile", "-NonInteractive", "-Command", code],
                               capture_output=True, text=True, check=False,
                               env={**os.environ, "REVIEW_SCRIPT": str(SCRIPT)})
    if completed.returncode:
        raise AssertionError(completed.stdout + completed.stderr)
    return completed.stdout.strip()


@unittest.skipUnless(PWSH, "PowerShell 7 (pwsh) is not installed")
class ReviewScriptTests(unittest.TestCase):
    def test_colors_convert_both_ways(self) -> None:
        self.assertEqual(run("ConvertTo-ExcelColor 'BF9000'"), str(0xBF + 256 * 0x90))
        self.assertEqual(run("ConvertTo-HexColor (ConvertTo-ExcelColor '2E75B6')"), "2E75B6")
        self.assertEqual(run("ConvertTo-HexColor 0"), "000000")

    def test_blank_values(self) -> None:
        self.assertEqual(run("(Test-Blank $null), (Test-Blank ''), (Test-Blank ([System.DBNull]::Value)), "
                             "(Test-Blank 0), (Test-Blank 'x') -join ','"), "True,True,True,False,False")

    def test_content_bounds_ignore_formatted_empty_cells(self) -> None:
        body = r"""
$a = [System.Array]::CreateInstance([object], @(4, 3), @(1, 1))
$a.SetValue('Header', 1, 1); $a.SetValue('=1+1', 3, 2); $a.SetValue('', 4, 3)
$used = [pscustomobject]@{ Row = 2; Column = 3; Formula = $a }
$sheet = [pscustomobject]@{ UsedRange = $used }
(Get-ContentBound $sheet) -join ','
$one = [pscustomobject]@{ UsedRange = [pscustomobject]@{ Row = 1; Column = 1; Formula = '' } }
(Get-ContentBound $one) -join ','
"""
        self.assertEqual(run(body).splitlines(), ["4,4", "0,0"])

    def test_listing_is_capped(self) -> None:
        body = r"""
$items = New-Object System.Collections.Generic.List[string]
1..11 | ForEach-Object { $items.Add("A$_") }
Format-Listed $items
"""
        self.assertEqual(run(body), "A1, A2, A3, A4, A5, A6, A7, A8 and 3 more")


if __name__ == "__main__":
    unittest.main()
