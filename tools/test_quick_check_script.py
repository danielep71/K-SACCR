"""Offline tests of tools/QuickCheck-Report.ps1 and tools/Run-SACCR-QuickCheck.ps1, the automatic quick
Excel check (#713): the guard against forks, the report, the artifact, the comment upsert and the cleanup of
old bundles. They run in PowerShell 7 (pwsh) and are skipped where it is not installed; no Excel and no
GitHub are needed."""
from __future__ import annotations

import os
import re
import shutil
import subprocess
import tempfile
import unittest
from pathlib import Path

TOOLS = Path(__file__).resolve().parent
FUNCTIONS = TOOLS / "QuickCheck-Report.ps1"
RUNNER = TOOLS / "Run-SACCR-QuickCheck.ps1"
PWSH = shutil.which("pwsh")

SUMMARY = """K-SACCR branch verification (quick): a pull-request check, not a release certification
Branch: feat/713-x
Commit: 0123456789abcdef0123456789abcdef01234567
Mode: quick: a pull-request check, not a release certification
Bundle: D:\\runner\\_work\\K-SACCR\\saccr-verify-0123456
Duration: 6 min 10 s
Result: PASS (quick run; compile, Test_RunAll, the sheet review and the 14 protection checks)
Compile: PASS; Debug > Compile VBAProject completed with no error
Sheet review: RESULT=PASS; sheets=19; states=2; failed_checks=0; pictures=38/38

Test_RunAll (summary.log):
RESULT=PASS; suites=10; passed=10; failed=0; seconds=102,2

Protection checks (protection-checks.txt):
RESULT=PASS; checks=14; failed=0
"""


def run(body: str, **env: str) -> str:
    code = "$ErrorActionPreference = 'Stop'\n. $env:QUICK_FUNCTIONS\n" + body
    completed = subprocess.run([str(PWSH), "-NoProfile", "-NonInteractive", "-Command", code],
                               capture_output=True, text=True, check=False,
                               env={**os.environ, "QUICK_FUNCTIONS": str(FUNCTIONS), **env})
    if completed.returncode:
        raise AssertionError(completed.stdout + completed.stderr)
    return completed.stdout.strip()


@unittest.skipUnless(PWSH, "PowerShell 7 (pwsh) is not installed")
class QuickCheckScriptTests(unittest.TestCase):
    def test_only_branches_of_this_repository_and_runs_by_hand_are_allowed(self) -> None:
        body = r"""
$r = 'danielep71/K-SACCR'
@(
  (Test-QuickRunAllowed 'workflow_dispatch' '' $r),
  (Test-QuickRunAllowed 'pull_request' 'danielep71/K-SACCR' $r),
  (Test-QuickRunAllowed 'pull_request' 'DanieleP71/k-saccr' $r),
  (Test-QuickRunAllowed 'pull_request' 'someone/K-SACCR' $r),
  (Test-QuickRunAllowed 'pull_request' '' $r),
  (Test-QuickRunAllowed 'pull_request' 'danielep71/K-SACCR' ''),
  (Test-QuickRunAllowed 'pull_request_target' 'danielep71/K-SACCR' $r),
  (Test-QuickRunAllowed 'push' 'danielep71/K-SACCR' $r)
) -join ','
"""
        self.assertEqual(run(body), "True,True,True,False,False,False,False,False")

    def test_bundle_is_read_from_the_summary(self) -> None:
        self.assertEqual(run("Get-QuickBundle $env:SUMMARY", SUMMARY=SUMMARY), "D:\\runner\\_work\\K-SACCR\\saccr-verify-0123456")
        self.assertEqual(run("[string]($null -eq (Get-QuickBundle 'no bundle here'))"), "True")

    def test_paths_are_hidden_longest_first_and_case_insensitively(self) -> None:
        body = r"""
$map = @{ 'D:\runner' = '<runner>'; 'D:\runner\_work\K-SACCR\saccr-verify-0123456' = '<bundle>'; '\bob\' = '\<user>\' }
Hide-QuickPaths -Text 'file=d:\RUNNER\_work\K-SACCR\saccr-verify-0123456\x.csv and D:\runner\y and C:\Users\bob\Temp and bobby' -Replacements $map
"""
        self.assertEqual(run(body), r"file=<bundle>\x.csv and <runner>\y and C:\Users\<user>\Temp and bobby")

    def test_a_dollar_sign_in_a_replacement_is_literal(self) -> None:
        self.assertEqual(run("Hide-QuickPaths -Text 'a X b' -Replacements @{ 'X' = '$1$$' }"), "a $1$$ b")

    def test_report_names_the_commit_and_folds_the_details(self) -> None:
        body = r"""
$s = $env:SUMMARY
New-QuickReport -Summary $s -Sha '0123456789abcdef0123456789abcdef01234567' -Passed $true -RunUrl 'https://example.test/run/1' -ArtifactName 'excel-quick-check-1-1'
"""
        report = run(body, SUMMARY=SUMMARY)
        lines = report.splitlines()
        self.assertEqual(lines[0], "<!-- excel-quick-check:0123456789abcdef0123456789abcdef01234567 -->")
        self.assertEqual(lines[1], "### Excel quick check of `0123456`: PASS")
        self.assertNotIn("Bundle:", report)
        self.assertIn("Sheet review: RESULT=PASS", report)
        self.assertIn("<details><summary>Test_RunAll, protection checks and sheet review</summary>", report)
        self.assertLess(report.index("Sheet review: RESULT=PASS"), report.index("<details>"))
        self.assertGreater(report.index("RESULT=PASS; suites=10"), report.index("<details>"))
        self.assertIn("not a release certification", report)
        self.assertIn("[Workflow run](https://example.test/run/1)", report)
        self.assertIn("`excel-quick-check-1-1`", report)

    def test_a_failed_report_says_fail_and_a_long_one_is_cut(self) -> None:
        failed = run("New-QuickReport -Summary 'x' -Sha 'abcdef0' -Passed $false -RunUrl 'u' -ArtifactName ''")
        self.assertIn("### Excel quick check of `abcdef0`: FAIL", failed)
        self.assertNotIn("artifact", failed)
        body = "$long = 'y' * 70000\n(New-QuickReport -Summary $long -Sha 'abcdef0' -Passed $true -RunUrl 'u' -ArtifactName '').Length"
        self.assertLessEqual(int(run(body)), 60200)

    def test_the_artifact_holds_the_pictures_and_the_text_but_no_workbook(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            bundle = Path(folder) / "saccr-verify-0123456"
            (bundle / "sheets").mkdir(parents=True)
            (bundle / "protection").mkdir()
            (bundle / "K-SACCR_Verification.xlsm").write_bytes(b"workbook")
            (bundle / "protection" / "K-SACCR_Production.xlsm").write_bytes(b"workbook")
            (bundle / "verify-summary.txt").write_text("log at D:\\runner\\x\n", encoding="utf-8")
            (bundle / "summary.log").write_text("RESULT=PASS\n", encoding="utf-8")
            picture = bytes(range(256))
            (bundle / "sheets" / "01-Home-run.png").write_bytes(picture)
            destination = Path(folder) / "artifact"
            body = (f"$map = @{{ 'D:\\runner' = '<runner>' }}\n"
                    f"Copy-QuickArtifacts -Bundle '{bundle}' -Destination '{destination}' -Replacements $map")
            self.assertEqual(run(body), "3")
            copied = sorted(str(p.relative_to(destination)).replace("\\", "/") for p in destination.rglob("*") if p.is_file())
            self.assertEqual(copied, ["sheets/01-Home-run.png", "summary.log", "verify-summary.txt"])
            self.assertEqual((destination / "sheets" / "01-Home-run.png").read_bytes(), picture)
            self.assertEqual((destination / "verify-summary.txt").read_text(encoding="utf-8"), "log at <runner>\\x\n")

    def test_an_existing_comment_of_the_commit_is_found(self) -> None:
        body = r"""
$comments = @(
  [pscustomobject]@{ id = 11; body = 'Codex: fine' },
  [pscustomobject]@{ id = 12; body = (Get-QuickMarker 'aaa') + "`nold report" },
  [pscustomobject]@{ id = 13; body = (Get-QuickMarker 'bbb') + "`nother commit" }
)
(Find-QuickReportComment -Comments $comments -Marker (Get-QuickMarker 'bbb'))
[string]($null -eq (Find-QuickReportComment -Comments $comments -Marker (Get-QuickMarker 'ccc')))
[string]($null -eq (Find-QuickReportComment -Comments @() -Marker (Get-QuickMarker 'aaa')))
[string]($null -eq (Find-QuickReportComment -Comments @("") -Marker (Get-QuickMarker 'aaa')))
[string]($null -eq (Find-QuickReportComment -Comments "" -Marker (Get-QuickMarker 'aaa')))
[string]($null -eq (Find-QuickReportComment -Comments @([pscustomobject]@{ id = 1 }) -Marker (Get-QuickMarker 'aaa')))
"""
        # A pull request with no comment yet: Windows PowerShell 5.1 returns "" for [], which has no body.
        self.assertEqual(run(body).splitlines(), ["13", "True", "True", "True", "True", "True"])

    def test_only_old_quick_bundles_are_removed(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            parent = Path(folder)
            for name in ("saccr-verify-abcdef0", "saccr-verify-abcdef1-20261101-120000", "saccr-verify-abcdef2",
                         "saccr-verify-xyz", "saccr-evidence-abcdef0", "other"):
                (parent / name).mkdir()
                (parent / name / "x.txt").write_text("x", encoding="utf-8")
            (parent / "saccr-verify-abcdef3").write_text("a file", encoding="utf-8")
            body = (f"$old = (Get-Date).AddDays(-30)\n"
                    f"foreach ($n in 'saccr-verify-abcdef0','saccr-verify-abcdef1-20261101-120000','saccr-verify-xyz',"
                    f"'saccr-evidence-abcdef0','other') {{ (Get-Item -LiteralPath (Join-Path '{parent}' $n)).LastWriteTime = $old }}\n"
                    f"(Remove-OldQuickBundles -Parent '{parent}' -KeepDays 14 | Sort-Object) -join ','")
            self.assertEqual(run(body), "saccr-verify-abcdef0,saccr-verify-abcdef1-20261101-120000")
            left = sorted(p.name for p in parent.iterdir())
            self.assertEqual(left, ["other", "saccr-evidence-abcdef0", "saccr-verify-abcdef2", "saccr-verify-abcdef3",
                                    "saccr-verify-xyz"])

    def test_the_scripts_parse_and_no_variable_shadows_a_parameter_in_another_case(self) -> None:
        # PowerShell names are case-insensitive: a "$branch" next to a "$Branch" parameter is the parameter.
        body = r"""
$found = @()
foreach ($path in @($env:QUICK_FUNCTIONS, $env:QUICK_RUNNER)) {
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$null, [ref]$errors)
    foreach ($e in $errors) { $found += ($path + ': ' + $e.Message) }
    $scopes = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true))
    foreach ($f in $scopes) {
        $names = @()
        if ($null -ne $f.Body.ParamBlock) { $names = @($f.Body.ParamBlock.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath }) }
        foreach ($v in $f.Body.FindAll({ param($n) $n -is [System.Management.Automation.Language.VariableExpressionAst] }, $true)) {
            $name = $v.VariablePath.UserPath
            foreach ($p in $names) {
                if ($name -ieq $p -and $name -cne $p) { $found += ($f.Name + ': $' + $name + ' / $' + $p) }
            }
        }
    }
    if ($null -ne $ast.ParamBlock) {
        $names = @($ast.ParamBlock.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath })
        foreach ($v in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.VariableExpressionAst] }, $true)) {
            $name = $v.VariablePath.UserPath
            foreach ($p in $names) {
                if ($name -ieq $p -and $name -cne $p) { $found += ('script: $' + $name + ' / $' + $p) }
            }
        }
    }
}
$found -join "`n"
"""
        self.assertEqual(run(body, QUICK_RUNNER=str(RUNNER)), "")


FAKE_VERIFY = r"""
param([string]$Branch, [switch]$Quick, [string]$RepoRoot, [string]$Commit)
$parent = Split-Path -Parent $RepoRoot
$sha = $Commit
$bundle = Join-Path $parent ('saccr-verify-' + $sha.Substring(0, 7))
New-Item -ItemType Directory -Path (Join-Path $bundle 'sheets') -Force | Out-Null
[System.IO.File]::WriteAllBytes((Join-Path $bundle 'sheets\\01-Home-run.png'), [byte[]](1, 2, 3))
[System.IO.File]::WriteAllBytes((Join-Path $bundle 'K-SACCR_Verification.xlsm'), [byte[]](9))
Set-Content -LiteralPath (Join-Path $bundle 'summary.log') -Value 'RESULT=PASS'
$lines = @('K-SACCR branch verification (quick): a pull-request check, not a release certification',
           "Branch: $Branch", "Commit: $sha", "Bundle: $bundle", 'Result: ' + $env:FAKE_RESULT,
           '', 'Test_RunAll (summary.log):', 'RESULT=PASS; suites=10', 'see ' + $bundle + '\\x')
Set-Content -LiteralPath (Join-Path $bundle 'verify-summary.txt') -Value $lines
$lines | ForEach-Object { Write-Output $_ }
exit [int]$env:FAKE_CODE
"""
GIT = ["git", "-c", "user.name=t", "-c", "user.email=t@example.test"]


def run_script(script: Path, arguments: list[str], env: dict[str, str]) -> subprocess.CompletedProcess[str]:
    """Run a script in pwsh and print an exception's message as one plain line: pwsh's own error view wraps
    the message at the terminal's width and colours it, which differs between machines."""
    quoted = " ".join(argument if re.fullmatch(r"-[A-Za-z]+", argument) else "'" + argument.replace("'", "''") + "'"
                      for argument in arguments)
    command = (f"try {{ & '{script}' {quoted} }} catch {{ [Console]::Error.WriteLine('ERROR: ' + $_.Exception.Message); "
               "exit 1 }; exit $LASTEXITCODE")
    return subprocess.run([str(PWSH), "-NoProfile", "-NonInteractive", "-Command", command],
                          capture_output=True, text=True, check=False, env=env)
OTHER = "f" * 40


def make_repo(root: Path, *scripts: str) -> tuple[Path, str]:
    """A git repository of the tools folder with one commit; returns it and the commit."""
    repo = root / "work" / "K-SACCR"
    (repo / "tools").mkdir(parents=True)
    for name in scripts:
        shutil.copy(TOOLS / name, repo / "tools" / name)
    subprocess.run(["git", "init", "-q", str(repo)], check=True)
    subprocess.run([*GIT, "-C", str(repo), "add", "-A"], check=True)
    subprocess.run([*GIT, "-C", str(repo), "commit", "-q", "-m", "x"], check=True)
    sha = subprocess.run(["git", "-C", str(repo), "rev-parse", "HEAD"], capture_output=True, text=True, check=True)
    return repo, sha.stdout.strip()


@unittest.skipUnless(PWSH and shutil.which("git"), "PowerShell 7 (pwsh) or git is not installed")
class QuickCheckRunnerTests(unittest.TestCase):
    """tools/Run-SACCR-QuickCheck.ps1 end to end, with a fake Verify-SACCR-Branch.ps1 in place of Excel."""

    def run_runner(self, folder: Path, branch: str, **env: str) -> tuple[subprocess.CompletedProcess[str], Path, str]:
        repo, sha = make_repo(folder, "Run-SACCR-QuickCheck.ps1", "QuickCheck-Report.ps1")
        (repo / "tools" / "Verify-SACCR-Branch.ps1").write_text(FAKE_VERIFY, encoding="utf-8")
        subprocess.run(["git", "-C", str(repo), "update-index", "--assume-unchanged", "tools/Verify-SACCR-Branch.ps1"],
                       check=False, capture_output=True)
        summary = folder / "step-summary.md"
        head_sha = env.pop("HEAD_SHA", sha)
        completed = run_script(
            repo / "tools" / "Run-SACCR-QuickCheck.ps1",
            ["-Branch", branch, "-Shell", str(PWSH), "-MutexName", "Local\\SACCR-Test-" + folder.name, "-WaitMinutes", "1"],
            {**os.environ, "GITHUB_STEP_SUMMARY": str(summary), "GITHUB_TOKEN": "", "GITHUB_WORKSPACE": str(repo),
             "RUN_URL": "https://example.test/run/7", "ARTIFACT_NAME": "excel-quick-check-7-1",
             "REPOSITORY": "danielep71/K-SACCR", "USERNAME": "", "HEAD_SHA": head_sha, **env})
        return completed, folder / "work", sha

    def test_a_branch_of_this_repository_is_run_and_reported(self) -> None:
        with tempfile.TemporaryDirectory() as name:
            folder = Path(name)
            completed, work, sha = self.run_runner(folder, "feat/713-x", EVENT_NAME="pull_request",
                                                   HEAD_REPO="danielep71/K-SACCR", FAKE_RESULT="PASS", FAKE_CODE="0")
            self.assertEqual(completed.returncode, 0, completed.stdout + completed.stderr)
            report = (folder / "step-summary.md").read_text(encoding="utf-8")
            self.assertIn("### Excel quick check of `" + sha[:7] + "`: PASS", report)
            self.assertIn("<!-- excel-quick-check:" + sha + " -->", report)
            self.assertNotIn("Bundle:", report)
            self.assertNotIn(str(work), report)
            artifact = work / "K-SACCR" / "quick-run-artifact"
            copied = sorted(str(p.relative_to(artifact)).replace("\\", "/") for p in artifact.rglob("*") if p.is_file())
            self.assertEqual(copied, ["sheets/01-Home-run.png", "summary.log", "verify-summary.txt"])
            self.assertNotIn(str(work), (artifact / "verify-summary.txt").read_text(encoding="utf-8"))

    def test_a_failed_check_fails_the_step_and_says_so(self) -> None:
        with tempfile.TemporaryDirectory() as name:
            folder = Path(name)
            completed, _, sha = self.run_runner(folder, "feat/713-x", EVENT_NAME="workflow_dispatch", HEAD_REPO="",
                                                FAKE_RESULT="FAIL", FAKE_CODE="1", HEAD_SHA="")
            self.assertEqual(completed.returncode, 1, completed.stdout + completed.stderr)
            self.assertIn("### Excel quick check of `" + sha[:7] + "`: FAIL",
                          (folder / "step-summary.md").read_text(encoding="utf-8"))

    def test_the_branch_of_a_fork_is_refused_and_nothing_runs(self) -> None:
        with tempfile.TemporaryDirectory() as name:
            folder = Path(name)
            completed, work, sha = self.run_runner(folder, "feat/713-x", EVENT_NAME="pull_request",
                                                   HEAD_REPO="someone/K-SACCR", FAKE_RESULT="PASS", FAKE_CODE="0")
            self.assertNotEqual(completed.returncode, 0)
            self.assertIn("Refused", completed.stdout + completed.stderr)
            self.assertFalse((work / ("saccr-verify-" + sha[:7])).exists())
            self.assertFalse((folder / "step-summary.md").exists())

    def test_a_checkout_that_is_not_the_commit_of_the_event_is_refused(self) -> None:
        # The branch moved on while the run waited: the check must not be attached to a commit that was not tested.
        with tempfile.TemporaryDirectory() as name:
            folder = Path(name)
            completed, work, sha = self.run_runner(folder, "feat/713-x", EVENT_NAME="pull_request",
                                                   HEAD_REPO="danielep71/K-SACCR", FAKE_RESULT="PASS", FAKE_CODE="0",
                                                   HEAD_SHA=OTHER)
            self.assertNotEqual(completed.returncode, 0)
            self.assertIn("not at the commit of the event", completed.stdout + completed.stderr)
            self.assertFalse((work / ("saccr-verify-" + sha[:7])).exists())

    def test_the_commit_is_passed_to_the_quick_check(self) -> None:
        with tempfile.TemporaryDirectory() as name:
            folder = Path(name)
            completed, _, sha = self.run_runner(folder, "feat/713-x", EVENT_NAME="workflow_dispatch", HEAD_REPO="",
                                                FAKE_RESULT="PASS", FAKE_CODE="0", HEAD_SHA="")
            self.assertEqual(completed.returncode, 0, completed.stdout + completed.stderr)
            self.assertIn("Commit: " + sha, completed.stdout)


@unittest.skipUnless(PWSH and shutil.which("git"), "PowerShell 7 (pwsh) or git is not installed")
class VerifyCommitTests(unittest.TestCase):
    """Verify-SACCR-Branch.ps1 -Commit verifies the checked-out commit and nothing else (#713)."""

    def verify(self, repo: Path, commit: str) -> subprocess.CompletedProcess[str]:
        return run_script(repo / "tools" / "Verify-SACCR-Branch.ps1",
                          ["-Branch", "feat/713-x", "-Quick", "-Commit", commit], dict(os.environ))

    def test_another_commit_is_refused_before_anything_runs(self) -> None:
        with tempfile.TemporaryDirectory() as name:
            repo, sha = make_repo(Path(name), "Verify-SACCR-Branch.ps1")
            completed = self.verify(repo, OTHER)
            output = completed.stdout + completed.stderr
            self.assertNotEqual(completed.returncode, 0)
            self.assertIn("nothing was verified", output)
            self.assertFalse(any(p.name.startswith("saccr-verify-") for p in repo.parent.iterdir()))

    def test_a_short_or_upper_case_sha_is_refused(self) -> None:
        with tempfile.TemporaryDirectory() as name:
            repo, sha = make_repo(Path(name), "Verify-SACCR-Branch.ps1")
            for bad in (sha[:7], sha.upper()):
                completed = self.verify(repo, bad)
                self.assertNotEqual(completed.returncode, 0)
                self.assertIn("Not a full commit SHA", completed.stdout + completed.stderr)

    def test_the_checked_out_commit_is_the_one_verified_and_nothing_is_fetched(self) -> None:
        with tempfile.TemporaryDirectory() as name:
            repo, sha = make_repo(Path(name), "Verify-SACCR-Branch.ps1")
            completed = self.verify(repo, sha)
            output = completed.stdout + completed.stderr
            # No Excel here, so the quick check itself fails; what matters is where it started.
            self.assertIn("Verifying the checked-out commit " + sha, output)
            self.assertNotIn("Fetching", output)
            self.assertIn("Commit: " + sha, output)


if __name__ == "__main__":
    unittest.main()
