<#
.SYNOPSIS
Check out a branch of origin, certify it with tools\Certify-SACCR.ps1 and
print a short summary to report.
.DESCRIPTION
For the branch of a pull request, with one command on the Windows host:
1. refuses a working tree with changes to tracked files, as
   Certify-SACCR.ps1 does, and never discards them;
2. fetches the branch from origin and checks it out: a local branch that
   tracks origin is created when there is none, and an existing one is
   fast-forwarded to origin. A local branch with commits origin does not
   have is refused;
3. runs tools\Certify-SACCR.ps1 of that branch, in its own PowerShell, with
   a new bundle folder next to the checkout: saccr-evidence-SHORTSHA, or
   saccr-evidence-SHORTSHA-yyyyMMdd-HHmmss when that folder exists already.
   Certification builds and compiles a development workbook, runs every
   suite and Test_RunAll, runs the 14 protection checks and validates the
   bundle; it may ask the operator to enable content;
4. writes verify-summary.txt in the bundle and prints it: the branch, the
   commit, the Test_RunAll summary, the 14 protection checks and the result.
The exit code is that of Certify-SACCR.ps1: 0 only when the bundle validates
and the 14 protection checks pass. The script never pushes and never changes
the macro setting or the access to the VBA project. Checks a person makes by
eye, such as the Ribbon's labels, stay with the person.
.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Verify-SACCR-Branch.ps1 -Branch chore/218-tooling-cleanup
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Branch,
    [Parameter(Mandatory = $false)]
    [string]$Python = "python",
    [Parameter(Mandatory = $false)]
    [string]$Operator,
    [Parameter(Mandatory = $false)]
    [string]$RepoRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $scriptPath = $MyInvocation.MyCommand.Path
    if ([string]::IsNullOrWhiteSpace($scriptPath)) {
        throw "Cannot determine the script path. Supply -RepoRoot explicitly."
    }
    $RepoRoot = Split-Path -Parent (Split-Path -Parent $scriptPath)
}
$RepoRoot = (Resolve-Path -LiteralPath $RepoRoot).Path
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

function Invoke-GitRaw {
    # Run git in the checkout; returns its exit code and its output, stdout
    # and stderr, as text. Git writes progress and ordinary messages to
    # stderr; under Windows PowerShell 5.1 a redirected stderr line is an
    # error record, which the script's "Stop" preference would turn into a
    # terminating error, so the preference is "Continue" here and only the
    # exit code decides.
    param([string[]]$Arguments)
    $ErrorActionPreference = "Continue"
    $lines = @(& git -C $RepoRoot @Arguments 2>&1 | ForEach-Object { [string]$_ })
    $code = $LASTEXITCODE
    return [pscustomobject]@{ Code = $code; Output = ($lines -join [Environment]::NewLine).Trim() }
}

function Invoke-Git {
    # Run git in the checkout and stop on a failure; returns its output.
    param([string[]]$Arguments, [string]$Failure)
    $result = Invoke-GitRaw $Arguments
    if ($result.Code -ne 0) {
        throw ($Failure + " (git " + ($Arguments -join " ") + ": " + $result.Output + ")")
    }
    return $result.Output
}

function Get-FileText {
    # A bundle file's text, or a note when it was not written.
    param([string]$Path)
    if (Test-Path -LiteralPath $Path) {
        return ([System.IO.File]::ReadAllText($Path)).TrimEnd()
    }
    return "(not written)"
}

# ----------------------------------------------------------------------------
# Branch
# ----------------------------------------------------------------------------
if ($Branch.StartsWith("-")) {
    throw "Not a branch name: $Branch"
}
if ((Invoke-GitRaw @("check-ref-format", "--branch", $Branch)).Code -ne 0) {
    throw "Not a valid branch name: $Branch"
}
$dirty = Invoke-Git @("status", "--porcelain", "--untracked-files=no") "Cannot read the working tree"
if (-not [string]::IsNullOrWhiteSpace($dirty)) {
    throw "The working tree has changes. Commit or stash them first; nothing was checked out."
}

Write-Host "Fetching $Branch from origin..."
[void](Invoke-Git @("fetch", "origin", ("+refs/heads/" + $Branch + ":refs/remotes/origin/" + $Branch)) `
    "Cannot fetch the branch from origin")
if ((Invoke-GitRaw @("show-ref", "--verify", "--quiet", ("refs/heads/" + $Branch))).Code -eq 0) {
    [void](Invoke-Git @("checkout", $Branch) "Cannot check out the branch")
    [void](Invoke-Git @("merge", "--ff-only", ("origin/" + $Branch)) `
        "The local branch has commits origin does not have; it was left as it is")
}
else {
    [void](Invoke-Git @("checkout", "-b", $Branch, "--track", ("origin/" + $Branch)) "Cannot check out the branch")
}
$sha = Invoke-Git @("rev-parse", "HEAD") "Cannot read the checked-out commit"
$remoteSha = Invoke-Git @("rev-parse", ("origin/" + $Branch)) "Cannot read the branch on origin"
if ($sha -ne $remoteSha) {
    throw "The checkout is at $sha, not at origin/$Branch ($remoteSha)."
}
Write-Host "Checked out $Branch at $sha."

# ----------------------------------------------------------------------------
# Certify
# ----------------------------------------------------------------------------
$parent = Split-Path -Parent $RepoRoot
$bundle = Join-Path $parent ("saccr-evidence-" + $sha.Substring(0, 7))
if (Test-Path -LiteralPath $bundle) {
    $bundle = $bundle + "-" + (Get-Date).ToString("yyyyMMdd-HHmmss")
}
$certify = Join-Path $RepoRoot "tools\Certify-SACCR.ps1"
if (-not (Test-Path -LiteralPath $certify)) {
    throw "This branch has no tools\Certify-SACCR.ps1."
}
$arguments = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $certify, "-BundleDir", $bundle,
               "-Python", $Python, "-RepoRoot", $RepoRoot)
if (-not [string]::IsNullOrWhiteSpace($Operator)) {
    $arguments += @("-Operator", $Operator)
}
Write-Host ""
& powershell.exe @arguments
$code = $LASTEXITCODE

# ----------------------------------------------------------------------------
# Summary
# ----------------------------------------------------------------------------
$verdict = "FAIL"
if ($code -eq 0) {
    $verdict = "PASS"
}
$lines = @(
    "K-SACCR branch verification",
    ("Branch: " + $Branch),
    ("Commit: " + $sha),
    ("Bundle: " + $bundle),
    ("Result: " + $verdict + " (Certify-SACCR exit code " + $code + ")"),
    "",
    "Test_RunAll (summary.log):",
    (Get-FileText (Join-Path $bundle "summary.log")),
    "",
    "Protection checks (protection-checks.txt):",
    (Get-FileText (Join-Path $bundle "protection-checks.txt"))
)
$summary = ($lines -join [Environment]::NewLine) + [Environment]::NewLine
if (Test-Path -LiteralPath $bundle) {
    [System.IO.File]::WriteAllText((Join-Path $bundle "verify-summary.txt"), $summary, $utf8NoBom)
}
Write-Host ""
Write-Host "----------------------------------------------------------------------"
Write-Host $summary
Write-Host "----------------------------------------------------------------------"
if (Test-Path -LiteralPath $bundle) {
    Write-Host ("Report the text above; it is also in " + (Join-Path $bundle "verify-summary.txt") + ".")
}
exit $code
