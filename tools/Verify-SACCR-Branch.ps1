<#
.SYNOPSIS
Check out a branch of origin, certify it with tools\Certify-SACCR.ps1, or
run the quick pull-request check with -Quick, and print a short summary to
report.
.DESCRIPTION
For the branch of a pull request, with one command on the Windows host:
1. refuses a working tree with changes to tracked files, as
   Certify-SACCR.ps1 does, and never discards them;
2. fetches the branch from origin and checks it out: a local branch that
   tracks origin is created when there is none, and an existing one is
   fast-forwarded to origin. A local branch with commits origin does not
   have is refused;
3. without -Quick, runs tools\Certify-SACCR.ps1 of that branch, in its own
   PowerShell, with a new bundle folder next to the checkout:
   saccr-evidence-SHORTSHA, or saccr-evidence-SHORTSHA-yyyyMMdd-HHmmss when
   that folder exists already. Certification builds and compiles a
   development workbook, runs every suite on its own and Test_RunAll, runs
   the 14 protection checks and validates the bundle; it may ask the
   operator to enable content.
   With -Quick, runs the quick check of a pull request (#618) in the folder
   saccr-verify-SHORTSHA next to the checkout, unattended: it builds a
   development workbook, opens it with its macros enabled for that Excel
   session only (AutomationSecurity Low; the Trust Center is not changed),
   runs its Auto_Open as Excel does for a user, compiles it, runs
   Test_RunAll once (TEST_Suite.Verify, into summary.log)
   and runs the 14 protection checks with
   tools\Check-SACCR-Protection.ps1 -Unattended. No prompt is shown; a step
   that would need one stops the run. A compile error opens a dialog in the
   VBE, which the compile command waits on: a watchdog closes the Excel of
   this run when the compile has not returned after 60 seconds, and the run
   fails with that reason. The suites do not run on their own
   and no evidence record is written, so the folder is not an evidence
   bundle and tools\check_excel_evidence.py cannot validate it: a quick run
   is a pull-request check, never a release certification;
4. writes verify-summary.txt in the bundle and prints it: the branch, the
   commit, the mode, the duration, the Test_RunAll summary, the 14
   protection checks and the result.
Without -Quick the exit code is that of Certify-SACCR.ps1: 0 only when the
bundle validates and the 14 protection checks pass. With -Quick it is 0 only
when the project compiled, Test_RunAll printed RESULT=PASS and the 14
protection checks pass. The script never pushes and never changes the macro
setting or the access to the VBA project. Checks a person makes by eye, such
as the Ribbon's labels, stay with the person. A quick run needs the branch to
carry TEST_Suite.Verify and the -Unattended switch of
Check-SACCR-Protection.ps1; on an older branch it stops and says so.
.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Verify-SACCR-Branch.ps1 -Branch chore/218-tooling-cleanup
.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Verify-SACCR-Branch.ps1 -Branch fix/123-slug -Quick
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Branch,
    [Parameter(Mandatory = $false)]
    [switch]$Quick,
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

$msoAutomationSecurityLow = 1
$xlAutoOpen = 1
$compileControlId = 578
$compileSeconds = 60

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

function Write-Log {
    # Print one timed line of the quick run and append it to verify.log.
    param([string]$Text)
    $line = "{0}  {1}" -f (Get-Date).ToString("HH:mm:ss"), $Text
    Write-Host $line
    [System.IO.File]::AppendAllText($script:verifyLog, $line + "`r`n", $script:utf8NoBom)
}

function Start-CompileWatchdog {
    # A job that closes the Excel process of this run when the marker file
    # has not appeared after the given seconds: a compile error opens a modal
    # dialog in the VBE, which Debug > Compile VBAProject waits on, so the
    # compile call would otherwise wait for an operator. Returns the job, or
    # $null when the process could not be told apart from other Excels.
    param($Excel, [int]$Seconds, [string]$Marker)
    $handle = [int64]$Excel.Hwnd
    $process = @(Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue |
                 Where-Object { [int64]$_.MainWindowHandle -eq $handle })
    if ($process.Count -ne 1) {
        return $null
    }
    return Start-Job -ArgumentList $process[0].Id, $Seconds, $Marker -ScriptBlock {
        param([int]$ProcessId, [int]$Seconds, [string]$Marker)
        $deadline = (Get-Date).AddSeconds($Seconds)
        while ((Get-Date) -lt $deadline) {
            if (Test-Path -LiteralPath $Marker) {
                return "released"
            }
            Start-Sleep -Milliseconds 500
        }
        Stop-Process -Id $ProcessId -Force -ErrorAction SilentlyContinue
        return "closed"
    }
}

function Stop-CompileWatchdog {
    # Release the watchdog and return what it did: "released", "closed", or
    # "none" when there was no watchdog.
    param($Job, [string]$Marker)
    if ($null -eq $Job) {
        return "none"
    }
    [System.IO.File]::WriteAllText($Marker, "done", $script:utf8NoBom)
    $outcome = [string](Receive-Job -Job $Job -Wait -AutoRemoveJob)
    Remove-Item -LiteralPath $Marker -Force -ErrorAction SilentlyContinue
    return $outcome
}

function Invoke-QuickRun {
    # The quick check of a pull request, unattended; returns 0 when the
    # project compiled, Test_RunAll passed and the 14 protection checks pass.
    param([string]$Bundle, [string]$Sha)

    New-Item -ItemType Directory -Path $Bundle | Out-Null
    $script:verifyLog = Join-Path $Bundle "verify.log"
    $script:compileLine = "Compile: NOT RUN"
    Write-Log "K-SACCR quick verification of ${Sha}: a pull-request check, not a release certification"
    Write-Log "Started at $((Get-Date).ToString('yyyy-MM-ddTHH:mm:sszzz')); folder $Bundle"
    $workbookPath = Join-Path $Bundle "K-SACCR_Verification.xlsm"
    $excel = $null
    $workbook = $null
    $suitesRan = $false
    try {
        # Build the development workbook from the template and the source.
        Write-Log "Building the development workbook: tools\Sync-SACCR-VBA.ps1 -Mode Development -FromTemplate"
        $syncScript = Join-Path $RepoRoot "tools\Sync-SACCR-VBA.ps1"
        $syncOutput = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $syncScript `
            -WorkbookPath $workbookPath -Mode Development -FromTemplate -NoBackup -RepoRoot $RepoRoot 2>&1
        $syncCode = $LASTEXITCODE
        foreach ($line in $syncOutput) {
            Write-Log ("  sync: " + $line)
        }
        if ($syncCode -ne 0 -or -not (Test-Path -LiteralPath $workbookPath)) {
            throw "The workbook could not be built (Sync-SACCR-VBA exited with code $syncCode); see verify.log."
        }

        # Open it with its macros enabled for this Excel session only. The
        # Trust Center setting is read for the operator's information and
        # never changed.
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        $excel.AutomationSecurity = $msoAutomationSecurityLow
        $workbook = $excel.Workbooks.Open($workbookPath)
        $qualified = "'" + $workbook.Name + "'!"
        $ready = ""
        try {
            $ready = [string]$excel.Run($qualified + "TEST_Suite.CertificationReady")
        }
        catch {
            $ready = ""
        }
        if ($ready -ne "READY") {
            throw "The workbook's macros did not run although they were enabled for this Excel session."
        }
        Write-Log "Macros enabled for this Excel session; the workbook answered READY"
        # Excel runs Auto_Open when a user opens the workbook, or enables its
        # content, but not when automation opens it.
        [void]$workbook.RunAutoMacros($xlAutoOpen)
        Write-Log "Auto_Open run, as when a user opens the workbook"

        # Compile, as Certify-SACCR.ps1 does, but without asking the operator
        # when Debug > Compile VBAProject does not turn grey on its own, and
        # with a watchdog for the dialog a compile error opens.
        $marker = Join-Path $Bundle "compile.marker"
        $watchdog = Start-CompileWatchdog -Excel $excel -Seconds $compileSeconds -Marker $marker
        if ($null -eq $watchdog) {
            Write-Log "No compile watchdog: this run's Excel process was not found"
        }
        $vbe = $excel.VBE
        $vbe.MainWindow.Visible = $true
        try {
            $vbe.ActiveVBProject = $workbook.VBProject
        }
        catch {
            Write-Log ("The workbook's project could not be made the active one: " + $_.Exception.Message)
        }
        $compile = $vbe.CommandBars.FindControl(1, $compileControlId)
        if ($null -eq $compile) {
            $script:compileLine = "Compile: FAIL; Debug > Compile VBAProject was not found in the VBE"
            throw "Cannot find Debug > Compile VBAProject."
        }
        $compiled = $false
        $compileError = ""
        try {
            if ($compile.Enabled) {
                [void]$compile.Execute()
            }
            for ($attempt = 0; $attempt -lt 10 -and -not $compiled; $attempt++) {
                Start-Sleep -Milliseconds 500
                $compiled = -not $vbe.CommandBars.FindControl(1, $compileControlId).Enabled
            }
        }
        catch {
            # The call fails when the watchdog closed Excel under it.
            $compileError = $_.Exception.Message
        }
        finally {
            $outcome = Stop-CompileWatchdog -Job $watchdog -Marker $marker
        }
        if ($outcome -eq "closed") {
            $script:compileLine = ("Compile: FAIL; Debug > Compile VBAProject did not return in $compileSeconds s, " +
                                   "as when a compile error opens a dialog in the VBE; this run's Excel was closed")
            throw "The compile did not return; see the compile error in the VBE, without -Quick."
        }
        if ($compileError -ne "") {
            $script:compileLine = "Compile: FAIL; Debug > Compile VBAProject could not be run: " + $compileError
            throw "The compile could not be run: $compileError"
        }
        if (-not $compiled) {
            $script:compileLine = ("Compile: FAIL; Debug > Compile VBAProject did not complete on its own. " +
                                   "Look at the VBE, or run without -Quick, which asks")
            throw "The compile could not be confirmed without the operator."
        }
        $script:compileLine = "Compile: PASS; Debug > Compile VBAProject completed with no error"
        Write-Log "Compiled with no error"

        # One pass of the suites, into summary.log.
        Write-Log "Running TEST_Suite.Verify: Test_RunAll once, into summary.log"
        $verifyResult = ""
        try {
            $verifyResult = [string]$excel.Run($qualified + "TEST_Suite.Verify", $Bundle, $Sha)
        }
        catch {
            throw ("TEST_Suite.Verify could not be run; the branch may predate the quick mode (#618): " +
                   $_.Exception.Message)
        }
        Write-Log "Verify: $verifyResult"
        $suitesRan = $verifyResult.StartsWith("RESULT=DONE;")
        if (-not $suitesRan) {
            throw "The suites did not run: $verifyResult"
        }
    }
    catch {
        Write-Log ("Stopped: " + $_.Exception.Message)
    }
    finally {
        if ($null -ne $workbook) {
            try {
                [void]$workbook.Close($false)
                Write-Log "Development workbook closed without saving"
            }
            catch {
                Write-Log ("The development workbook could not be closed: " + $_.Exception.Message)
            }
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
        }
        if ($null -ne $excel) {
            try {
                [void]$excel.Quit()
            }
            catch {
            }
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }

    # The 14 protection checks on a production workbook, unattended.
    $protectionCode = 1
    if ($suitesRan) {
        Write-Log "Running the protection checks unattended: tools\Check-SACCR-Protection.ps1 -Unattended"
        try {
            & (Join-Path $RepoRoot "tools\Check-SACCR-Protection.ps1") -OutputDir $Bundle -Python $Python `
                -RepoRoot $RepoRoot -Unattended | Out-Host
            $protectionCode = $LASTEXITCODE
        }
        catch {
            Write-Log ("The protection checks stopped: " + $_.Exception.Message)
        }
        $protectionFile = Join-Path $Bundle "protection-checks.txt"
        if (Test-Path -LiteralPath $protectionFile) {
            Write-Log ("Protection checks: " + @(Get-Content -LiteralPath $protectionFile -Encoding UTF8)[-1])
        }
        else {
            Write-Log "Protection checks: no protection-checks.txt was written"
            $protectionCode = 1
        }
    }
    else {
        Write-Log "Protection checks not run: the suites did not run"
    }
    $summaryPassed = (Get-FileText (Join-Path $Bundle "summary.log")) -match '(?m)^RESULT=PASS; suites='
    Write-Log "Finished at $((Get-Date).ToString('yyyy-MM-ddTHH:mm:sszzz'))"
    if ($suitesRan -and $summaryPassed -and $protectionCode -eq 0) {
        return 0
    }
    return 1
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
# Certify, or the quick check
# ----------------------------------------------------------------------------
$parent = Split-Path -Parent $RepoRoot
$prefix = "saccr-evidence-"
if ($Quick) {
    $prefix = "saccr-verify-"
}
$bundle = Join-Path $parent ($prefix + $sha.Substring(0, 7))
if (Test-Path -LiteralPath $bundle) {
    $bundle = $bundle + "-" + (Get-Date).ToString("yyyyMMdd-HHmmss")
}
$started = Get-Date
Write-Host ""
if ($Quick) {
    $mode = "quick: a pull-request check, not a release certification"
    $quickResult = @(Invoke-QuickRun -Bundle $bundle -Sha $sha)
    $code = [int]$quickResult[-1]
}
else {
    $mode = "full: Certify-SACCR.ps1"
    $certify = Join-Path $RepoRoot "tools\Certify-SACCR.ps1"
    if (-not (Test-Path -LiteralPath $certify)) {
        throw "This branch has no tools\Certify-SACCR.ps1."
    }
    $arguments = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $certify, "-BundleDir", $bundle,
                   "-Python", $Python, "-RepoRoot", $RepoRoot)
    if (-not [string]::IsNullOrWhiteSpace($Operator)) {
        $arguments += @("-Operator", $Operator)
    }
    & powershell.exe @arguments
    $code = $LASTEXITCODE
}
$elapsed = (Get-Date) - $started

# ----------------------------------------------------------------------------
# Summary
# ----------------------------------------------------------------------------
$verdict = "FAIL"
if ($code -eq 0) {
    $verdict = "PASS"
}
$title = "K-SACCR branch verification"
$resultLine = "Result: " + $verdict + " (Certify-SACCR exit code " + $code + ")"
if ($Quick) {
    $title = "K-SACCR branch verification (quick): a pull-request check, not a release certification"
    $resultLine = "Result: " + $verdict + " (quick run; compile, Test_RunAll and the 14 protection checks)"
}
$lines = @(
    $title,
    ("Branch: " + $Branch),
    ("Commit: " + $sha),
    ("Mode: " + $mode),
    ("Bundle: " + $bundle),
    ("Duration: " + [int][Math]::Floor($elapsed.TotalMinutes) + " min " + $elapsed.Seconds + " s"),
    $resultLine
)
if ($Quick) {
    $lines += @($script:compileLine)
}
$lines += @(
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
