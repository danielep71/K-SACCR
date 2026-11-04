<#
.SYNOPSIS
Run the Excel certification of the checked-out commit with one command and
write a validated evidence bundle (#535).
.DESCRIPTION
For the commit checked out in this repository, with a clean working tree:
1. creates the evidence bundle folder, by default saccr-evidence-SHORTSHA next
   to the checkout, and starts session.log there;
2. records the environment as found: Windows, Excel version and build, Office
   bitness, the macro setting and "Trust access to the VBA project object
   model". It never changes any of them;
3. builds a development workbook in the bundle from the template and the
   source of the checkout (tools\Sync-SACCR-VBA.ps1 -Mode Development
   -FromTemplate);
4. opens it in a visible Excel that follows the Trust Center macro setting
   (AutomationSecurity "by UI"). When the macros are disabled, it asks the
   operator to click Enable Content and waits;
5. compiles the project (Debug > Compile VBAProject) and runs
   TEST_Suite.Certify, which runs the harness and every suite once each, on
   its own, writing each report to its own log with the commit's SHA as its
   first line, and then Test_RunAll into summary.log;
6. closes the workbook without saving and runs
   tools\Check-SACCR-Protection.ps1: the 14 manual protection checks on a
   production workbook built from the same checkout, written to
   protection-checks.txt in the bundle;
7. writes session.json and runs tools\write_excel_evidence.py, which writes
   host.json from the logs and validates the bundle with
   tools\check_excel_evidence.py.
The exit code is 0 only when the bundle validates as PASS and the 14
protection checks pass. The manual procedure of
docs\EXCEL_EVIDENCE.md stays available as a fallback.
.REQUIREMENTS
- Windows desktop Excel, and Python 3.10 or later on the PATH (or -Python).
- "Trust access to the VBA project object model" already on: the build needs
  it, and the script never turns it on.
- A clean checkout of the commit to certify.
.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Certify-SACCR.ps1
.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Certify-SACCR.ps1 -BundleDir "C:\Evidence\saccr-1.1.0" -Operator "Owner, release workstation"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$BundleDir,
    [Parameter(Mandatory = $false)]
    [string]$Operator = "Repository owner, local Windows workstation",
    [Parameter(Mandatory = $false)]
    [string]$Python = "python",
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

$msoAutomationSecurityByUI = 2
$compileControlId = 578
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

function Get-Timestamp {
    return (Get-Date).ToString("yyyy-MM-ddTHH:mm:sszzz")
}

function Write-Session {
    param([string]$Text)
    $line = "{0}  {1}" -f (Get-Date).ToString("HH:mm:ss"), $Text
    Write-Host $line
    [System.IO.File]::AppendAllText($script:sessionLog, $line + "`r`n", $script:utf8NoBom)
}

function Get-RegistryValue {
    param([string]$Path, [string]$Name)
    $item = Get-ItemProperty -LiteralPath $Path -Name $Name -ErrorAction SilentlyContinue
    if ($null -eq $item) {
        return $null
    }
    return $item.$Name
}

function Get-SecurityValue {
    # A value set by group policy takes precedence over the user's own setting.
    param([string]$OfficeVersion, [string]$Name)
    foreach ($root in @("HKCU:\Software\Policies\Microsoft\Office", "HKCU:\Software\Microsoft\Office")) {
        $value = Get-RegistryValue -Path "$root\$OfficeVersion\Excel\Security" -Name $Name
        if ($null -ne $value) {
            return $value
        }
    }
    return $null
}

function Get-MacroPolicy {
    param([string]$OfficeVersion)
    switch (Get-SecurityValue -OfficeVersion $OfficeVersion -Name "VBAWarnings") {
        1 { return "Enable VBA macros" }
        2 { return "Disable VBA macros with notification" }
        3 { return "Disable VBA macros except digitally signed macros" }
        4 { return "Disable VBA macros without notification" }
        default { return "Disable VBA macros with notification (default; no setting found)" }
    }
}

function Get-VbaProjectAccess {
    param([string]$OfficeVersion)
    if ((Get-SecurityValue -OfficeVersion $OfficeVersion -Name "AccessVBOM") -eq 1) {
        return "On"
    }
    return "Off"
}

# ----------------------------------------------------------------------------
# Candidate and bundle
# ----------------------------------------------------------------------------
$sha = (& git -C $RepoRoot rev-parse HEAD).Trim()
if ($LASTEXITCODE -ne 0 -or $sha -notmatch '^[0-9a-f]{40}$') {
    throw "Cannot read the checked-out commit of $RepoRoot."
}
$dirty = & git -C $RepoRoot status --porcelain --untracked-files=no
if ($LASTEXITCODE -ne 0 -or -not [string]::IsNullOrWhiteSpace(($dirty | Out-String))) {
    throw "The working tree has changes. Certify a clean checkout: git status must show none."
}
if ([string]::IsNullOrWhiteSpace($BundleDir)) {
    $BundleDir = Join-Path (Split-Path -Parent $RepoRoot) ("saccr-evidence-" + $sha.Substring(0, 7))
}
if (Test-Path -LiteralPath $BundleDir) {
    if (@(Get-ChildItem -LiteralPath $BundleDir -Force).Count -gt 0) {
        throw "The bundle folder is not empty: $BundleDir. A bundle is never overwritten."
    }
}
else {
    New-Item -ItemType Directory -Path $BundleDir | Out-Null
}
$BundleDir = (Resolve-Path -LiteralPath $BundleDir).Path
if ($BundleDir.StartsWith($RepoRoot + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "The bundle folder must be outside the checkout: $BundleDir"
}
$script:sessionLog = Join-Path $BundleDir "session.log"
$startedAt = Get-Timestamp
Write-Session "K-SACCR certification of candidate $sha"
Write-Session "Started at $startedAt; bundle $BundleDir"

# ----------------------------------------------------------------------------
# Environment, as found
# ----------------------------------------------------------------------------
$osInfo = Get-CimInstance -ClassName Win32_OperatingSystem
$osName = ($osInfo.Caption -replace '^Microsoft\s+', '').Trim()
$osText = "{0} (version {1}, build {2})" -f $osName, $osInfo.Version, $osInfo.BuildNumber
$osArchitecture = switch ($env:PROCESSOR_ARCHITECTURE) {
    "AMD64" { "x64" }
    "ARM64" { "ARM64" }
    "x86" { "x86" }
    default { $env:PROCESSOR_ARCHITECTURE }
}
$clickToRun = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
$platform = Get-RegistryValue -Path $clickToRun -Name "Platform"
$officeBitness = switch ($platform) {
    "x64" { "64-bit" }
    "x86" { "32-bit" }
    default { "unknown" }
}
$versionToReport = Get-RegistryValue -Path $clickToRun -Name "VersionToReport"
Write-Session "Windows: $osText; architecture $osArchitecture"

$pythonVersion = & $Python --version 2>&1
if ($LASTEXITCODE -ne 0) {
    throw "Python was not found as '$Python'. Install Python 3.10 or later, or pass -Python."
}
Write-Session "Python: $pythonVersion"

$session = [ordered]@{
    operator           = $Operator
    started_at         = $startedAt
    finished_at        = $null
    excel_version      = "unknown"
    excel_build        = "unknown"
    office_bitness     = $officeBitness
    os                 = $osText
    os_architecture    = $osArchitecture
    runtime            = "unknown"
    references         = @()
    macro_policy       = "unknown"
    vba_project_access = "unknown"
    import             = [ordered]@{ status = "NOT_RUN"; detail = "Not run" }
    compile            = [ordered]@{ status = "NOT_RUN"; detail = "Not run" }
    workbook_closed    = $false
}

$workbookPath = Join-Path $BundleDir "K-SACCR_Certification.xlsm"
$excel = $null
$workbook = $null
$certifyResult = ""
try {
    # ------------------------------------------------------------------------
    # Excel and its security settings, read and never changed
    # ------------------------------------------------------------------------
    $excel = New-Object -ComObject Excel.Application
    $officeVersion = [string]$excel.Version
    $session.excel_version = $officeVersion
    if ($null -ne $versionToReport) {
        $session.excel_build = "Version $versionToReport (build $($excel.Build))"
    }
    else {
        $session.excel_build = "Version $officeVersion (build $($excel.Build))"
    }
    $session.macro_policy = Get-MacroPolicy -OfficeVersion $officeVersion
    $session.vba_project_access = Get-VbaProjectAccess -OfficeVersion $officeVersion
    Write-Session "Excel: $($session.excel_build), $officeBitness"
    Write-Session "Macro setting: $($session.macro_policy); VBA project access: $($session.vba_project_access)"
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    $excel = $null
    if ($session.vba_project_access -ne "On") {
        throw ("'Trust access to the VBA project object model' is off. The build needs it; " +
               "turn it on yourself in Excel > Options > Trust Center if your policy allows. " +
               "This script never changes it.")
    }

    # ------------------------------------------------------------------------
    # Build the development workbook from the template and the source
    # ------------------------------------------------------------------------
    Write-Session "Building the development workbook: tools\Sync-SACCR-VBA.ps1 -Mode Development -FromTemplate"
    $syncScript = Join-Path $RepoRoot "tools\Sync-SACCR-VBA.ps1"
    $syncOutput = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $syncScript `
        -WorkbookPath $workbookPath -Mode Development -FromTemplate -NoBackup -RepoRoot $RepoRoot 2>&1
    $syncCode = $LASTEXITCODE
    foreach ($line in $syncOutput) {
        Write-Session ("  sync: " + $line)
    }
    if ($syncCode -ne 0 -or -not (Test-Path -LiteralPath $workbookPath)) {
        $session.import = [ordered]@{ status = "FAIL"; detail = "Sync-SACCR-VBA exited with code $syncCode" }
        throw "The workbook could not be built; see session.log."
    }
    $session.import = [ordered]@{
        status = "PASS"
        detail = "Sync-SACCR-VBA -Mode Development -FromTemplate from the checkout of $sha"
    }

    # ------------------------------------------------------------------------
    # Open with the Trust Center macro setting and wait for enabled content
    # ------------------------------------------------------------------------
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.AutomationSecurity = $msoAutomationSecurityByUI
    $workbook = $excel.Workbooks.Open($workbookPath)
    $qualified = "'" + $workbook.Name + "'!"
    $ready = $false
    while (-not $ready) {
        try {
            $ready = ([string]$excel.Run($qualified + "TEST_Suite.CertificationReady")) -eq "READY"
        }
        catch {
            $ready = $false
        }
        if (-not $ready) {
            Write-Host ""
            Write-Host "The workbook's macros are disabled. In the Excel window, click Enable Content in"
            Write-Host "the message bar (or File > Info > Enable Content), then press Enter here."
            $answer = Read-Host "Press Enter to try again, or type Q to stop"
            if ($answer -match '^\s*q') {
                throw "Stopped by the operator: the macros were not enabled."
            }
        }
    }
    Write-Session "Macros enabled; the workbook answered READY"
    # Read by the project itself: before the macros run, the VBProject of a
    # workbook opened "by UI" may not be readable from PowerShell.
    $referenceText = [string]$excel.Run($qualified + "TEST_Suite.CertificationReferences")
    $session.references = @($referenceText -split "`n" | Where-Object { $_ -ne "" })
    Write-Session ("References: " + ($session.references -join "; "))

    # ------------------------------------------------------------------------
    # Compile
    # ------------------------------------------------------------------------
    # Debug > Compile VBAProject acts on the VBE's active project and is
    # greyed out once that project is compiled. The VBE window is shown, so
    # that the command's state is current, and the state is read afresh.
    $vbe = $excel.VBE
    $vbe.MainWindow.Visible = $true
    try {
        $vbe.ActiveVBProject = $workbook.VBProject
    }
    catch {
        Write-Session ("The test workbook's project could not be made the active one: " + $_.Exception.Message)
    }
    $compile = $vbe.CommandBars.FindControl(1, $compileControlId)
    if ($null -eq $compile) {
        $session.compile = [ordered]@{ status = "FAIL"; detail = "Debug > Compile VBAProject was not found in the VBE" }
        throw "Cannot find Debug > Compile VBAProject."
    }
    if ($compile.Enabled) {
        $compile.Execute()
    }
    $compiled = $false
    for ($attempt = 0; $attempt -lt 10 -and -not $compiled; $attempt++) {
        Start-Sleep -Milliseconds 500
        $compiled = -not $vbe.CommandBars.FindControl(1, $compileControlId).Enabled
    }
    $compileDetail = "Debug > Compile VBAProject completed with no error"
    if (-not $compiled) {
        # The command can stay enabled without a compile error, so the
        # operator looks at the VBE instead of the run stopping on a guess.
        Write-Host ""
        Write-Host "Debug > Compile VBAProject still looks enabled. In the VBE window, choose"
        Write-Host "Debug > Compile VBAProject yourself. Type Y if it completes with no error and"
        Write-Host "the command is then greyed out; type N if it reports a compile error."
        $answer = Read-Host "Compiled with no error? (Y/N)"
        if ($answer -notmatch '^\s*y') {
            $session.compile = [ordered]@{ status = "FAIL"; detail = "Debug > Compile VBAProject reported an error, as the operator saw" }
            throw "The project does not compile; see the VBE."
        }
        $compileDetail = "Debug > Compile VBAProject completed with no error, as the operator saw in the VBE"
    }
    $session.compile = [ordered]@{ status = "PASS"; detail = $compileDetail }
    Write-Session ("Compiled with no error" + $(if ($compiled) { "" } else { ", as the operator saw in the VBE" }))

    # ------------------------------------------------------------------------
    # Run the certification
    # ------------------------------------------------------------------------
    Write-Session "Running TEST_Suite.Certify: the harness and the nine suites, then Test_RunAll"
    $certifyResult = [string]$excel.Run($qualified + "TEST_Suite.Certify", $BundleDir, $sha)
    Write-Session "Certify: $certifyResult"
}
catch {
    Write-Session ("Stopped: " + $_.Exception.Message)
}
finally {
    if ($null -ne $workbook) {
        try {
            $workbook.Close($false)
            $session.workbook_closed = $true
            Write-Session "Test workbook closed without saving"
        }
        catch {
            Write-Session ("The test workbook could not be closed: " + $_.Exception.Message)
        }
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }
    elseif ($session.import.status -ne "PASS" -or $session.compile.status -ne "PASS") {
        $session.workbook_closed = $true
    }
    if ($null -ne $excel) {
        try {
            $excel.Quit()
        }
        catch {
        }
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# ----------------------------------------------------------------------------
# Protection checks on a production workbook
# ----------------------------------------------------------------------------
# Run before session.log is recorded, as the record cites its digest.
$protectionCode = 1
if ($session.compile.status -eq "PASS") {
    Write-Session "Running the protection checks on a production workbook: tools\Check-SACCR-Protection.ps1"
    try {
        & (Join-Path $RepoRoot "tools\Check-SACCR-Protection.ps1") -OutputDir $BundleDir -Python $Python -RepoRoot $RepoRoot
        $protectionCode = $LASTEXITCODE
    }
    catch {
        Write-Session ("The protection checks stopped: " + $_.Exception.Message)
    }
    $protectionFile = Join-Path $BundleDir "protection-checks.txt"
    if (Test-Path -LiteralPath $protectionFile) {
        Write-Session ("Protection checks: " + @(Get-Content -LiteralPath $protectionFile -Encoding UTF8)[-1])
    }
    else {
        Write-Session "Protection checks: no protection-checks.txt was written"
        $protectionCode = 1
    }
}
else {
    Write-Session "Protection checks not run: the development workbook did not compile"
}

# ----------------------------------------------------------------------------
# Record and validate
# ----------------------------------------------------------------------------
$session.finished_at = Get-Timestamp
Write-Session "Finished at $($session.finished_at)"
[System.IO.File]::WriteAllText((Join-Path $BundleDir "session.json"),
    ($session | ConvertTo-Json -Depth 4), $utf8NoBom)
Write-Host ""
& $Python (Join-Path $RepoRoot "tools\write_excel_evidence.py") --root $RepoRoot --candidate-sha $sha --bundle $BundleDir
$code = $LASTEXITCODE
Write-Host ""
Write-Host "Bundle: $BundleDir"
if ($code -eq 0 -and $protectionCode -eq 0) {
    Write-Host "RESULT=PASS: the bundle validates for candidate $sha and the 14 protection checks pass."
    exit 0
}
if ($code -eq 0) {
    Write-Host "RESULT=FAIL: the bundle validates, but the protection checks did not pass; see protection-checks.txt."
    exit 1
}
Write-Host "RESULT=FAIL: see the outcomes above and session.log. Keep the bundle as it is."
exit $code
