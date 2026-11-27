<#
.SYNOPSIS
Run the quick Excel check of a branch on the self-hosted runner and report it
in the pull request (#713).
.DESCRIPTION
The step of .github\workflows\excel-quick.yml that needs Excel. In order:
1. refuses a run that is neither a branch of this repository nor started by
   hand (Test-QuickRunAllowed), so no code of a fork runs on the machine, and
   a checkout that is not at the commit of the event;
2. removes the bundles of runs older than 14 days next to the checkout;
3. takes the machine's Excel mutex, so that two runs never drive Excel at once,
   and waits for it up to -WaitMinutes;
4. runs tools\Verify-SACCR-Branch.ps1 -Smoke -Commit (a pull request, or a
   run by hand with CHECK_MODE smoke; #801) or -Quick -Commit (a run by hand
   with CHECK_MODE full) on the commit the
   workflow checked out, in its own PowerShell, and releases the mutex. The
   commit is fixed before the wait: a push that moves the branch while the run
   waits changes nothing, so a result is never attached to a commit whose
   workbook was not tested;
5. builds the report from verify-summary.txt, with the machine's paths, user
   and computer name replaced, copies the text results and the sheet pictures
   to -ArtifactDir, writes the job summary and posts the report on the pull
   request: one comment per commit, updated by a re-run.
The exit code is that of the quick check: 0 only when it passed. A failure of
the reporting itself is a warning and does not change it. The report never
replaces the maintainer's verification and is not release evidence.
The environment supplies the event (EVENT_NAME, HEAD_REPO, REPOSITORY, PR_NUMBER,
RUN_URL, ARTIFACT_NAME) and GITHUB_TOKEN; nothing from the event is put into a
command line. A run started by hand supplies CHECK_MODE, full or smoke (#801)
or screenshots (#616: the smoke check, then the documentation screenshots),
and may supply BENCHMARK_SIZES: a full quick check then runs the benchmark too
(-BenchmarkSizes of Verify-SACCR-Branch.ps1, #771), which checks the sizes; a
pull request never does.
.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Run-SACCR-QuickCheck.ps1 -Branch fix/123-slug
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Branch,
    [Parameter(Mandatory = $false)]
    [string]$ArtifactDir,
    [Parameter(Mandatory = $false)]
    [string]$MutexName = "Local\SACCR-ExcelAutomation",
    [Parameter(Mandatory = $false)]
    [int]$WaitMinutes = 40,
    [Parameter(Mandatory = $false)]
    [int]$KeepDays = 14,
    [Parameter(Mandatory = $false)]
    [string]$Shell = "powershell.exe"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
. (Join-Path $PSScriptRoot "QuickCheck-Report.ps1")

function Get-EnvText {
    param([string]$Name)
    $value = [Environment]::GetEnvironmentVariable($Name)
    if ($null -eq $value) {
        return ""
    }
    return $value
}

if ([string]::IsNullOrWhiteSpace($ArtifactDir)) {
    $ArtifactDir = Join-Path $repoRoot "quick-run-artifact"
}

# 1. Fork guard.
if (-not (Test-QuickRunAllowed -EventName (Get-EnvText "EVENT_NAME") -HeadRepository (Get-EnvText "HEAD_REPO") `
        -Repository (Get-EnvText "REPOSITORY"))) {
    throw "Refused: this run is neither a branch of this repository nor started by hand."
}

# The commit the workflow checked out is the one that is verified: the event's own for a pull
# request, the branch tip at checkout for a run by hand.
$pinned = (@(& git -C $repoRoot rev-parse HEAD 2>&1 | ForEach-Object { [string]$_ }) -join "").Trim()
if ($LASTEXITCODE -ne 0 -or $pinned -cnotmatch '^[0-9a-f]{40}$') {
    throw "Cannot read the checked-out commit: $pinned"
}
$expected = Get-EnvText "HEAD_SHA"
if ($expected -and $pinned -ne $expected) {
    throw "The checkout is at $pinned, not at the commit of the event ($expected); nothing was verified."
}

# 2. Old bundles.
$parent = Split-Path -Parent $repoRoot
$removed = @(Remove-OldQuickBundles -Parent $parent -KeepDays $KeepDays)
if ($removed.Count -gt 0) {
    Write-Host ("Removed " + $removed.Count + " bundle(s) older than $KeepDays days.")
}

# 3 and 4. One Excel at a time, then the quick check.
$mutex = New-Object System.Threading.Mutex($false, $MutexName)
$held = $false
$code = 1
$output = @()
# A pull request gets the smoke check; the full quick check runs only by hand (#801). The mode is known
# before anything runs, so that a report without a summary still carries the marker of its mode (#828).
$checkMode = Get-QuickCheckMode -EventName (Get-EnvText "EVENT_NAME") -Mode (Get-EnvText "CHECK_MODE")
try {
    try {
        $held = $mutex.WaitOne([TimeSpan]::FromMinutes($WaitMinutes))
    }
    catch [System.Threading.AbandonedMutexException] {
        # The holder stopped without releasing it; the mutex is ours now.
        $held = $true
    }
    if (-not $held) {
        throw "Excel stayed busy for $WaitMinutes minutes ($MutexName); nothing was run."
    }
    Write-Host "Excel mutex taken: $MutexName"
    $verify = Join-Path $repoRoot "tools\Verify-SACCR-Branch.ps1"
    # Windows PowerShell 5.1 turns a child's stderr line into an error record under "Stop".
    $ErrorActionPreference = "Continue"
    $modeSwitch = "-Quick"
    if ($checkMode -eq "smoke") {
        $modeSwitch = "-Smoke"
    }
    elseif ($checkMode -eq "screenshots") {
        # The smoke check, then the documentation screenshots (#616).
        $modeSwitch = "-Screenshots"
    }
    Write-Host "Check: $checkMode"
    $verifyArguments = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $verify, "-Branch", $Branch, $modeSwitch,
                         "-Commit", $pinned, "-RepoRoot", $repoRoot)
    # The benchmark only in a full quick check run by hand (#771): its timings belong to the machine, not to
    # the pull request.
    $sizes = (Get-EnvText "BENCHMARK_SIZES").Trim()
    if ($sizes -and $checkMode -eq "full") {
        $verifyArguments += @("-BenchmarkSizes", $sizes)
        Write-Host "The benchmark runs too, sizes $sizes"
    }
    $output = @(& $Shell @verifyArguments 2>&1 | ForEach-Object { [string]$_ })
    $code = $LASTEXITCODE
    $ErrorActionPreference = "Stop"
    $output | ForEach-Object { Write-Host $_ }
}
finally {
    if ($held) {
        $mutex.ReleaseMutex()
    }
    $mutex.Dispose()
}

# 5. The report.
try {
    $text = $output -join [Environment]::NewLine
    $bundle = Get-QuickBundle -Text $text
    $summaryFile = $null
    if ($bundle) {
        $summaryFile = Join-Path $bundle "verify-summary.txt"
    }
    if ($summaryFile -and (Test-Path -LiteralPath $summaryFile)) {
        $summary = [System.IO.File]::ReadAllText($summaryFile)
    }
    else {
        $tail = @($output | Select-Object -Last 15) -join [Environment]::NewLine
        $summary = "The quick check wrote no summary. The end of its output:" + [Environment]::NewLine + $tail
    }
    # The report is of the commit that was pinned, which the quick check was told to verify.
    $sha = $pinned
    $replacements = @{}
    foreach ($path in @($bundle, $repoRoot, $parent, (Get-EnvText "GITHUB_WORKSPACE"), (Get-EnvText "RUNNER_TEMP"))) {
        if (-not [string]::IsNullOrWhiteSpace($path)) {
            $replacements[$path] = "<runner>"
        }
    }
    if ($bundle) {
        $replacements[$bundle] = "<bundle>"
    }
    # The user's name where it is a folder of a path (C:\Users\name\...), never inside a word.
    $user = Get-EnvText "USERNAME"
    if ($user.Length -ge 2) {
        $replacements["\" + $user + "\"] = "\<user>\"
    }
    $clean = Hide-QuickPaths -Text $summary -Replacements $replacements

    $report = New-QuickReport -Summary $clean -Sha $sha -Passed ($code -eq 0) -RunUrl (Get-EnvText "RUN_URL") `
        -ArtifactName (Get-EnvText "ARTIFACT_NAME") -Mode $checkMode
    if ($bundle -and (Test-Path -LiteralPath $bundle)) {
        $copied = Copy-QuickArtifacts -Bundle $bundle -Destination $ArtifactDir -Replacements $replacements
        Write-Host "$copied file(s) copied to the artifact folder."
    }
    $stepSummary = Get-EnvText "GITHUB_STEP_SUMMARY"
    if ($stepSummary) {
        [System.IO.File]::AppendAllText($stepSummary, $report, (New-Object System.Text.UTF8Encoding($false)))
    }
    $token = Get-EnvText "GITHUB_TOKEN"
    $repository = Get-EnvText "REPOSITORY"
    $number = 0
    $pull = $null
    if ([int]::TryParse((Get-EnvText "PR_NUMBER"), [ref]$number) -and $number -gt 0) {
        $pull = $number
    }
    elseif ($token -and $repository) {
        $pull = Find-QuickPullNumber -Repository $repository -Branch $Branch -Token $token
    }
    if ($token -and $sha -and $pull) {
        $what = Publish-QuickReport -Repository $repository -PullNumber ([int]$pull) -Body $report -Sha $sha -Token $token
        Write-Host "Report $what on pull request #$pull."
    }
    else {
        Write-Host "No pull request comment: no token, no commit or no open pull request for $Branch."
    }
}
catch {
    Write-Host ("::warning::The report could not be completed: " + ($_.Exception.Message -replace '[\r\n]+', ' '))
}
exit $code
