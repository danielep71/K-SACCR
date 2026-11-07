<#
.SYNOPSIS
Run the manual protection checks of docs\workbook\WORKBOOK_PROTECTION.md on a
production workbook built from the checked-out commit, without a person at
the keyboard (#535).
.DESCRIPTION
tools\Certify-SACCR.ps1 runs this script after the Excel suites; it can also
be run on its own. It:
1. builds a production workbook from the template and the source of the
   checkout (tools\Sync-SACCR-VBA.ps1 -Mode Production -FromTemplate) in
   OutputDir\protection, and writes the import specification's example files
   next to it (tools\import_spec.py --examples);
2. opens the workbook with its macros disabled. Without the macros nothing
   protects the sheets for the user interface only, so Excel refuses an
   automation write to a locked cell as it refuses a user's typing. Checks 2,
   3, 4, 6 and 14 write, paste and add sheets this way; check 5 reads the
   protection settings Excel enforces for sorting, filtering, inserting rows,
   row groups and formula display;
3. opens it again following the Trust Center macro setting, as
   Certify-SACCR.ps1 does, runs its Auto_Open, and runs checks 1 and 7 to 13
   through the macros' silent entry points (M_Main.RunSACCR_Silent,
   ValidateInputs_Silent, ClearOutputs_Silent, ImportCsv_Silent,
   ValidateStaging_Silent, CommitStaging_Silent, ExportEvidence_Silent and
   ExportImportSpecification_Silent), saving, closing and reopening it where a
   check says so;
4. writes one line per check to OutputDir\protection-checks.txt and returns 0
   only when all 14 pass.
It never changes the macro setting or VBA-project access. Two steps of the
manual checks stay with a person if one wants them: the confirmation dialog of
Commit import, which the silent entry point does not show, and the sheet-tab
menu, whose effect, a protected workbook structure, check 6 tests.
.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Check-SACCR-Protection.ps1 -OutputDir ..\saccr-evidence-1234567
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$OutputDir,
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
$OutputDir = (Resolve-Path -LiteralPath $OutputDir).Path

$msoAutomationSecurityByUI = 2
$msoAutomationSecurityForceDisable = 3
$xlAutoOpen = 1
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

$script:results = New-Object System.Collections.Generic.List[string]
$script:failures = 0
$script:stops = New-Object System.Collections.Generic.List[string]
$script:lastWriteError = ""
$script:excel = $null
$script:workbook = $null

function Add-Result {
    param([int]$Number, [bool]$Passed, [string]$Detail)
    $verdict = "FAIL"
    if ($Passed) {
        $verdict = "PASS"
    }
    else {
        $script:failures++
    }
    $line = "{0} {1} - {2}" -f $Number, $verdict, $Detail
    $script:results.Add($line)
    Write-Host $line
}

function Add-Stop {
    # A session that stopped: its remaining checks are reported as not run.
    param([string]$Detail)
    $script:stops.Add($Detail)
    Write-Host ("STOPPED: " + $Detail)
}

function Get-SheetNames {
    $matrix = [System.IO.File]::ReadAllText((Join-Path $RepoRoot "docs\workbook\protection.json"), $utf8NoBom) |
        ConvertFrom-Json
    return @($matrix.sheets.PSObject.Properties | ForEach-Object { $_.Name })
}

function Get-BarLinks {
    # The hyperlinks of the shapes of a sheet's navigation bar (#615), read
    # from the groups NavBarTop and NavBar; a shape without a link is skipped.
    param($Sheet)
    $links = New-Object System.Collections.Generic.List[object]
    foreach ($shape in $Sheet.Shapes) {
        if ([string]$shape.Name -notlike "NavBar*") {
            continue
        }
        foreach ($item in $shape.GroupItems) {
            try {
                $link = $item.Hyperlink
                if ($null -ne $link -and [string]$link.SubAddress) {
                    $links.Add($link)
                }
            }
            catch {
            }
        }
    }
    return , $links
}

function Get-Unprotected {
    # The sheets not protected, and "workbook structure" when it is not.
    param($Workbook)
    $open = New-Object System.Collections.Generic.List[string]
    foreach ($name in $script:sheetNames) {
        if (-not $Workbook.Worksheets.Item($name).ProtectContents) {
            $open.Add($name)
        }
    }
    if (-not $Workbook.ProtectStructure) {
        $open.Add("workbook structure")
    }
    return , $open
}

function Test-Write {
    # True when Excel accepts the value; a refused write throws. The property
    # is set through reflection: PowerShell's own COM binding of one
    # assignment keeps the type of its first value, so a number or an array
    # written after a string would be refused as a failed cast to a string.
    param($Range, $Value)
    try {
        $arguments = New-Object 'object[]' 1
        $arguments[0] = $Value.PSObject.BaseObject
        [void][System.__ComObject].InvokeMember("Value2", [System.Reflection.BindingFlags]::SetProperty, $null,
                                                $Range.PSObject.BaseObject, $arguments)
        return $true
    }
    catch {
        $script:lastWriteError = $_.Exception.Message
        return $false
    }
}

function Get-StopDetail {
    # The message of an error record and the script line that raised it.
    param($Record)
    return ("{0} (line {1}: {2})" -f $Record.Exception.Message, $Record.InvocationInfo.ScriptLineNumber,
            $Record.InvocationInfo.Line.Trim())
}

function Invoke-Macro {
    param($Excel, $Workbook, [string]$Name, [object[]]$Arguments = @())
    $qualified = "'" + $Workbook.Name + "'!" + $Name
    switch ($Arguments.Count) {
        0 { return [string]$Excel.Run($qualified) }
        1 { return [string]$Excel.Run($qualified, $Arguments[0]) }
        2 { return [string]$Excel.Run($qualified, $Arguments[0], $Arguments[1]) }
        3 { return [string]$Excel.Run($qualified, $Arguments[0], $Arguments[1], $Arguments[2]) }
        default { throw "Invoke-Macro takes at most three arguments." }
    }
}

function Close-Excel {
    # Close the workbook without saving, quit Excel and release both; a
    # failure here never hides the outcome of a check.
    param($Excel, $Workbook)
    foreach ($item in @($Workbook, $Excel)) {
        if ($null -eq $item) {
            continue
        }
        try {
            if ($item -eq $Workbook) {
                [void]$item.Close($false)
            }
            else {
                [void]$item.Quit()
            }
        }
        catch {
        }
        try {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($item)
        }
        catch {
        }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

function Open-WithMacros {
    # Open the workbook in $script:excel and $script:workbook following the
    # Trust Center setting, wait until its macros run, then run Auto_Open as
    # Excel does when a user opens it.
    $script:excel = New-Object -ComObject Excel.Application
    $script:excel.Visible = $true
    $script:excel.DisplayAlerts = $false
    $script:excel.AutomationSecurity = $msoAutomationSecurityByUI
    $script:workbook = $script:excel.Workbooks.Open($script:workbookPath)
    $ready = $false
    while (-not $ready) {
        try {
            [void](Invoke-Macro $script:excel $script:workbook "M_Main.ResultsStatus")
            $ready = $true
        }
        catch {
            Write-Host ""
            Write-Host "The production workbook's macros are disabled. In the Excel window, click Enable"
            Write-Host "Content in the message bar (or File > Info > Enable Content), then press Enter here."
            $answer = Read-Host "Press Enter to try again, or type Q to stop"
            if ($answer -match '^\s*q') {
                throw "Stopped by the operator: the macros were not enabled."
            }
        }
    }
    [void]$script:workbook.RunAutoMacros($xlAutoOpen)
}

function Close-Workbook {
    # Close $script:workbook without saving and quit $script:excel.
    Close-Excel $script:excel $script:workbook
    $script:excel = $null
    $script:workbook = $null
}

function Get-HeaderColumn {
    # The column of Header in row Row of Sheet, or 0.
    param($Sheet, [int]$Row, [string]$Header)
    $last = $Sheet.Cells.Item($Row, $Sheet.Columns.Count).End(-4159).Column
    for ($column = 1; $column -le $last; $column++) {
        if ([string]$Sheet.Cells.Item($Row, $column).Value2 -eq $Header) {
            return $column
        }
    }
    return 0
}

function Test-Run {
    param([string]$Result)
    return $Result.StartsWith("RESULT=OK;") -and $Result.Contains("; cleanup=PASS")
}

# ----------------------------------------------------------------------------
# Build the production workbook and the example files
# ----------------------------------------------------------------------------
$script:sheetNames = Get-SheetNames
$work = Join-Path $OutputDir "protection"
if (Test-Path -LiteralPath $work) {
    throw "The folder exists already: $work. A check is never repeated into the same folder."
}
New-Item -ItemType Directory -Path $work | Out-Null
$script:workbookPath = Join-Path $work "K-SACCR_Production.xlsm"
Write-Host "Building the production workbook: tools\Sync-SACCR-VBA.ps1 -Mode Production -FromTemplate"
$syncOutput = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $RepoRoot "tools\Sync-SACCR-VBA.ps1") `
    -WorkbookPath $script:workbookPath -Mode Production -FromTemplate -NoBackup -RepoRoot $RepoRoot 2>&1
$syncCode = $LASTEXITCODE
foreach ($line in $syncOutput) {
    Write-Host ("  sync: " + $line)
}
if ($syncCode -ne 0 -or -not (Test-Path -LiteralPath $script:workbookPath)) {
    throw "The production workbook could not be built (Sync-SACCR-VBA exited with code $syncCode)."
}
& $Python (Join-Path $RepoRoot "tools\import_spec.py") --root $RepoRoot --examples $work
if ($LASTEXITCODE -ne 0) {
    throw "The example files could not be written."
}
$nettingSetsCsv = Join-Path $work "NettingSets.csv"
$tradesCsv = Join-Path $work "Trades.csv"

# ----------------------------------------------------------------------------
# Macros disabled: the protection as a user meets it (checks 14, 2 to 6)
# ----------------------------------------------------------------------------
$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = $msoAutomationSecurityForceDisable
    $workbook = $excel.Workbooks.Open($script:workbookPath)

    # 14: every sheet and the structure protected; no macro runs. The
    # workbook opens on Home, whose cell links reach every other sheet without
    # a macro (#330), and the navigation bar of every sheet links every other
    # sheet, also without a macro, on the protected sheets (#615). Only the
    # cell links (Type 0) of Home count as its links.
    $open = Get-Unprotected $workbook
    $macroRan = $true
    try {
        [void](Invoke-Macro $excel $workbook "M_Main.ResultsStatus")
    }
    catch {
        $macroRan = $false
    }
    $opensOn = [string]$workbook.ActiveSheet.Name
    $homeSheet = $workbook.Worksheets.Item("Home")
    $homeLinks = @($homeSheet.Hyperlinks | Where-Object { [int]$_.Type -eq 0 })
    $linkCount = $homeLinks.Count
    $broken = New-Object System.Collections.Generic.List[string]
    foreach ($link in $homeLinks) {
        $target = ([string]$link.SubAddress) -replace "^'?([^'!]+)'?!.*$", '$1'
        try {
            $link.Follow()
            if ([string]$workbook.ActiveSheet.Name -ne $target) {
                $broken.Add($target)
            }
        }
        catch {
            $broken.Add($target)
        }
    }
    $expectedLinks = [int]$workbook.Worksheets.Count - 1
    $sheetNames = @(foreach ($sheet in $workbook.Worksheets) { [string]$sheet.Name })
    $barWrong = New-Object System.Collections.Generic.List[string]
    $barBroken = New-Object System.Collections.Generic.List[string]
    foreach ($sheetName in $sheetNames) {
        $barLinks = Get-BarLinks $workbook.Worksheets.Item($sheetName)
        $targets = @(foreach ($link in $barLinks) {
            ([string]$link.SubAddress) -replace "^'?([^'!]+)'?!.*$", '$1'
        })
        $expected = @($sheetNames | Where-Object { $_ -ne $sheetName })
        if ((($targets | Sort-Object) -join "|") -ne (($expected | Sort-Object) -join "|")) {
            $barWrong.Add($sheetName)
        }
        foreach ($link in $barLinks) {
            $target = ([string]$link.SubAddress) -replace "^'?([^'!]+)'?!.*$", '$1'
            try {
                $link.Follow()
                if ([string]$workbook.ActiveSheet.Name -ne $target) {
                    $barBroken.Add($sheetName + ">" + $target)
                }
            }
            catch {
                $barBroken.Add($sheetName + ">" + $target)
            }
        }
    }
    Add-Result 14 ($open.Count -eq 0 -and -not $macroRan -and $opensOn -eq "Home" -and
                   $linkCount -eq $expectedLinks -and $broken.Count -eq 0 -and
                   $barWrong.Count -eq 0 -and $barBroken.Count -eq 0) `
        ("Opened with macros disabled: unprotected [" + ($open -join ", ") + "]; a macro ran: $macroRan; opens on " +
         $opensOn + "; Home links " + $linkCount + " of " + $expectedLinks + ", not reaching their sheet [" +
         ($broken -join ", ") + "]; navigation bar on " + ($sheetNames.Count - $barWrong.Count) + " of " +
         $sheetNames.Count + " sheets, wrong [" + ($barWrong -join ", ") + "], links not reaching their sheet [" +
         ($barBroken -join ", ") + "]")

    # 2: typing into a header, a grey Params value, an output cell, a
    # staging metadata cell and Home is refused.
    $targets = @(
        @("Trades", "C4"), @("Params", "E8"), @("Results", "K7"), @("Import_Trades", "D4"), @("Home", "D6")
    )
    $accepted = @($targets | Where-Object { Test-Write ($workbook.Worksheets.Item($_[0]).Range($_[1])) "x" } |
        ForEach-Object { $_[0] + "!" + $_[1] })
    Add-Result 2 ($accepted.Count -eq 0) `
        ("Writes to Trades!C4, Params!E8, Results!K7, Import_Trades!D4 and Home!D6; accepted [" +
         ($accepted -join ", ") + "]")

    # 3: inputs, Params configuration, an FX rate, a bank figure and a bank
    # tolerance are accepted, also thousands of rows down.
    $targets = @(
        @("Trades", "C5", "T-CHECK"), @("Trades", "C4000", "T-CHECK"), @("NettingSets", "C7", "NS-CHECK"),
        @("Params", "E6", 46022), @("Params", "E58", 1), @("BankNettingSets", "C9", "NS-CHECK"),
        @("BankNettingSets", "D5", 0.5)
    )
    $refused = @($targets | Where-Object { -not (Test-Write ($workbook.Worksheets.Item($_[0]).Range($_[1])) $_[2]) } |
        ForEach-Object { $_[0] + "!" + $_[1] })
    $reason = ""
    if ($refused.Count -gt 0) {
        $reason = "; last error: " + $script:lastWriteError
    }
    Add-Result 3 ($refused.Count -eq 0) `
        ("Writes to Trades!C5, Trades!C4000, NettingSets!C7, Params!E6, Params!E58, BankNettingSets!C9 and D5; refused [" +
         ($refused -join ", ") + "]" + $reason)

    # 4: several rows of values pasted into the input columns, then cleared.
    $block = New-Object 'object[,]' 3, 3
    for ($row = 0; $row -lt 3; $row++) {
        for ($column = 0; $column -lt 3; $column++) {
            $block[$row, $column] = "P" + $row + $column
        }
    }
    $pasteRange = $workbook.Worksheets.Item("Trades").Range("C80:E82")
    $pasted = Test-Write $pasteRange $block
    $reason = ""
    if (-not $pasted) {
        $reason = " (" + $script:lastWriteError + ")"
    }
    $cleared = $true
    try {
        [void]$pasteRange.ClearContents()
    }
    catch {
        $cleared = $false
    }
    Add-Result 4 ($pasted -and $cleared) "Three rows pasted into Trades!C80:E82: $pasted$reason; cleared: $cleared"

    # 5: what each sheet allows, as Excel enforces it, and formulas visible.
    $problems = New-Object System.Collections.Generic.List[string]
    $trades = $workbook.Worksheets.Item("Trades")
    if (-not ($trades.Protection.AllowFiltering -and $trades.AutoFilterMode)) {
        $problems.Add("Trades cannot be filtered from its header")
    }
    foreach ($name in @("Trades", "NettingSets")) {
        if (-not $workbook.Worksheets.Item($name).Protection.AllowInsertingRows) {
            $problems.Add("$name does not allow inserting rows")
        }
    }
    if ($workbook.Worksheets.Item("Results").Range("C5").FormulaHidden) {
        $problems.Add("the formula of Results!C5 is hidden")
    }
    # Copied into a new workbook rather than the clipboard, which a hidden
    # Excel may not reach.
    $scratch = $null
    try {
        $scratch = $excel.Workbooks.Add()
        [void]$workbook.Worksheets.Item("Results").Range("C5:E7").Copy($scratch.Worksheets.Item(1).Range("A1"))
    }
    catch {
        $problems.Add("output cells cannot be copied: " + $_.Exception.Message)
    }
    if ($null -ne $scratch) {
        try {
            [void]$scratch.Close($false)
        }
        catch {
        }
    }
    Add-Result 5 ($problems.Count -eq 0) `
        ("Filtering, inserting rows, copying, formulas visible; the README row groups are checked under 1 with the macros; problems [" +
         ($problems -join "; ") + "]")

    # 6: the structure: no sheet can be added, renamed or deleted.
    $changed = New-Object System.Collections.Generic.List[string]
    try {
        [void]$workbook.Worksheets.Add()
        $changed.Add("a sheet was added")
    }
    catch {
    }
    try {
        $workbook.Worksheets.Item("Checks").Name = "ChecksRenamed"
        $changed.Add("Checks was renamed")
    }
    catch {
    }
    try {
        $workbook.Worksheets.Item("Basel_vs_CRR").Delete()
        $changed.Add("Basel_vs_CRR was deleted")
    }
    catch {
    }
    Add-Result 6 ($changed.Count -eq 0) `
        ("Add, rename and delete a sheet (the tab menu's Delete, Rename, Move, Hide); done [" + ($changed -join "; ") + "]")
}
catch {
    Add-Stop ("The macros-disabled checks stopped: " + (Get-StopDetail $_))
}
finally {
    Close-Excel $excel $workbook
    $excel = $null
    $workbook = $null
}

# ----------------------------------------------------------------------------
# Macros enabled: the operations a user runs (checks 1, 7 to 13)
# ----------------------------------------------------------------------------
try {
    Open-WithMacros
    $excel = $script:excel
    $workbook = $script:workbook

    # 1: every sheet and the structure protected once Auto_Open has run, the
    # README row groups usable.
    $open = Get-Unprotected $workbook
    $outlining = [bool]$workbook.Worksheets.Item("README").EnableOutlining
    Add-Result 1 ($open.Count -eq 0 -and $outlining) `
        ("Opened with macros enabled: unprotected [" + ($open -join ", ") + "]; README row groups open and close: $outlining")

    # 7: Run SA-CCR, Validate inputs, Clear outputs, then Run SA-CCR again.
    $steps = @(
        @("Run SA-CCR", "M_Main.RunSACCR_Silent"), @("Validate inputs", "M_Main.ValidateInputs_Silent"),
        @("Clear outputs", "M_Main.ClearOutputs_Silent"), @("Run SA-CCR", "M_Main.RunSACCR_Silent")
    )
    $problems = New-Object System.Collections.Generic.List[string]
    foreach ($step in $steps) {
        $result = Invoke-Macro $excel $workbook $step[1]
        if (-not (Test-Run $result)) {
            $problems.Add($step[0] + ": " + $result)
        }
        $open = Get-Unprotected $workbook
        if ($open.Count -gt 0) {
            $problems.Add($step[0] + " left [" + ($open -join ", ") + "] unprotected")
        }
    }
    Add-Result 7 ($problems.Count -eq 0) ("Run, Validate inputs, Clear outputs, Run; problems [" + ($problems -join "; ") + "]")

    # 10: a changed Alpha after a maintenance unlock makes the run WHAT-IF;
    # the run protects Params again; the value is restored.
    $params = $workbook.Worksheets.Item("Params")
    $alphaFormula = [string]$params.Range("E8").Formula
    $params.Unprotect()
    $params.Range("E8").Value2 = 1.5
    $whatIf = Invoke-Macro $excel $workbook "M_Main.RunSACCR_Silent"
    $protectedAgain = [bool]$params.ProtectContents
    $params.Unprotect()
    $params.Range("E8").Formula = $alphaFormula
    $regular = Invoke-Macro $excel $workbook "M_Main.RunSACCR_Silent"
    Add-Result 10 ((Test-Run $whatIf) -and $whatIf.Contains("; run_class=WHAT-IF;") -and $protectedAgain -and
                   $regular.Contains("; run_class=REGULATORY;") -and [bool]$params.ProtectContents) `
        ("Alpha 1.5: $whatIf; Params protected again: $protectedAgain; restored: $regular")

    # 11: an output cell changed after a maintenance unlock refuses the
    # export as EDITED; a run rebuilds Results and protects it again.
    $resultsSheet = $workbook.Worksheets.Item("Results")
    $evidence = Join-Path $work "evidence-edited.csv"
    $resultsSheet.Unprotect()
    $resultsSheet.Range("K7").Value2 = 12345
    $refused = Invoke-Macro $excel $workbook "M_Main.ExportEvidence_Silent" @($evidence)
    $rebuilt = Invoke-Macro $excel $workbook "M_Main.RunSACCR_Silent"
    $status = Invoke-Macro $excel $workbook "M_Main.ResultsStatus"
    Add-Result 11 ($refused.StartsWith("RESULT=REFUSED;") -and $refused.Contains("changed after the run") -and
                   (Test-Run $rebuilt) -and [bool]$resultsSheet.ProtectContents -and $status -eq "CURRENT") `
        ("Export after an edit: $refused; run: $rebuilt; Results protected again: $($resultsSheet.ProtectContents); status $status")

    # 9: bank figures entered, a run reconciles them, the export succeeds.
    $bank = $workbook.Worksheets.Item("BankNettingSets")
    $eadColumn = Get-HeaderColumn $resultsSheet 6 "EAD"
    # Cast, so that Excel receives a plain string and number.
    $bank.Range("C9").Value2 = [string]$resultsSheet.Range("C7").Value2
    $bank.Range("D9").Value2 = [string]$resultsSheet.Range("E7").Value2
    if ($eadColumn -gt 0) {
        $bank.Range("E9").Value2 = [double]$resultsSheet.Cells.Item(7, $eadColumn).Value2
    }
    $reconciled = Invoke-Macro $excel $workbook "M_Main.RunSACCR_Silent"
    $exported = Invoke-Macro $excel $workbook "M_Main.ExportEvidence_Silent" @((Join-Path $work "evidence.csv"))
    Add-Result 9 ($eadColumn -gt 0 -and (Test-Run $reconciled) -and $reconciled.Contains("; reconciliation=DONE;") -and
                  $exported.StartsWith("RESULT=OK;")) ("Run with bank figures: $reconciled; export: $exported")
    [void]$bank.Range("C9:E9").ClearContents()

    # 8: Export import specifications, Import CSV on both staging sheets, a
    # staged value corrected, Validate staging, Commit import, Run SA-CCR.
    $specification = Invoke-Macro $excel $workbook "M_Main.ExportImportSpecification_Silent" @((Join-Path $work "specification.html"))
    $loadedNs = Invoke-Macro $excel $workbook "M_Main.ImportCsv_Silent" @("NettingSets", $nettingSetsCsv, $true)
    $loadedTr = Invoke-Macro $excel $workbook "M_Main.ImportCsv_Silent" @("Trades", $tradesCsv, $true)
    $staged = $workbook.Worksheets.Item("Import_Trades")
    $commentColumn = Get-HeaderColumn $staged 9 "Comment"
    $editable = $commentColumn -gt 0 -and (Test-Write ($staged.Cells.Item(10, $commentColumn)) "Corrected in the check")
    $lineLocked = [bool]$staged.Range("C10").Locked -and -not [bool]$staged.Range("D10").Locked
    $validated = Invoke-Macro $excel $workbook "M_Main.ValidateStaging_Silent"
    $committed = Invoke-Macro $excel $workbook "M_Main.CommitStaging_Silent"
    $ran = Invoke-Macro $excel $workbook "M_Main.RunSACCR_Silent"
    $open = Get-Unprotected $workbook
    Add-Result 8 ($specification.StartsWith("RESULT=OK;") -and $loadedNs.StartsWith("RESULT=LOADED;") -and
                  $loadedTr.StartsWith("RESULT=LOADED;") -and $editable -and $lineLocked -and
                  $validated.StartsWith("RESULT=VALID;") -and $committed.StartsWith("RESULT=COMMITTED;") -and
                  (Test-Run $ran) -and $ran.Contains("; errors=0;") -and $ran.Contains("; trades_used=5;") -and
                  $open.Count -eq 0) `
        ("Specification: $specification; NettingSets: $loadedNs; Trades: $loadedTr; staged value editable: $editable; " +
         "CSV line locked, values unlocked: $lineLocked; validation: $validated; commit: $committed; run: $ran; " +
         "unprotected [" + ($open -join ", ") + "]. The commit's confirmation dialog is not shown by the silent entry point")

    # 12: save, close and reopen, then Run SA-CCR.
    [void]$workbook.Save()
    Close-Workbook
    Open-WithMacros
    $excel = $script:excel
    $workbook = $script:workbook
    $open = Get-Unprotected $workbook
    $ran = Invoke-Macro $excel $workbook "M_Main.RunSACCR_Silent"
    Add-Result 12 ($open.Count -eq 0 -and (Test-Run $ran) -and $ran.Contains("; trades_used=5;")) `
        ("After save and reopen: unprotected [" + ($open -join ", ") + "]; run: $ran")

    # 13: a VALID staging validation stays current across save and reopen;
    # an edit saved and reopened refuses Commit import.
    $loadedNs = Invoke-Macro $excel $workbook "M_Main.ImportCsv_Silent" @("NettingSets", $nettingSetsCsv, $true)
    $loadedTr = Invoke-Macro $excel $workbook "M_Main.ImportCsv_Silent" @("Trades", $tradesCsv, $true)
    $validated = Invoke-Macro $excel $workbook "M_Main.ValidateStaging_Silent"
    [void]$workbook.Save()
    Close-Workbook
    Open-WithMacros
    $excel = $script:excel
    $workbook = $script:workbook
    $statusNs = [string]$workbook.Worksheets.Item("Import_NettingSets").Range("D7").Value2
    $statusTr = [string]$workbook.Worksheets.Item("Import_Trades").Range("D7").Value2
    $current = [bool]$excel.Run("'" + $workbook.Name + "'!M_Main.StagingValidationCurrent")
    $staged = $workbook.Worksheets.Item("Import_Trades")
    $staged.Cells.Item(10, $commentColumn).Value2 = "Changed before saving"
    [void]$workbook.Save()
    Close-Workbook
    Open-WithMacros
    $excel = $script:excel
    $workbook = $script:workbook
    $edited = [string]$workbook.Worksheets.Item("Import_Trades").Range("D7").Value2
    $refused = Invoke-Macro $excel $workbook "M_Main.CommitStaging_Silent"
    Add-Result 13 ($loadedNs.StartsWith("RESULT=LOADED;") -and $loadedTr.StartsWith("RESULT=LOADED;") -and
                   $validated.StartsWith("RESULT=VALID;") -and $statusNs -eq "VALID" -and $statusTr -eq "VALID" -and
                   $current -and $edited -eq "EDITED" -and $refused.StartsWith("RESULT=REFUSED;")) `
        ("Validation: $validated; after reopening: $statusNs and $statusTr, current: $current; after an edit, " +
         "save and reopen: $edited, commit: $refused")
}
catch {
    Add-Stop ("The macros-enabled checks stopped: " + (Get-StopDetail $_))
}
finally {
    Close-Workbook
}

# ----------------------------------------------------------------------------
# Record
# ----------------------------------------------------------------------------
$numbers = @($script:results | ForEach-Object { [int]($_.Split(" ")[0]) } | Sort-Object -Unique)
$missing = @(1..14 | Where-Object { $numbers -notcontains $_ })
foreach ($number in $missing) {
    Add-Result $number $false "Not run: an earlier step stopped"
}
$ordered = @($script:results | Sort-Object { [int]($_.Split(" ")[0]) })
$ordered += @($script:stops | ForEach-Object { "STOPPED - " + $_ })
$text = "Manual protection checks of docs/workbook/WORKBOOK_PROTECTION.md, run by tools/Check-SACCR-Protection.ps1`r`n" +
        "Workbook: production build of " + (& git -C $RepoRoot rev-parse HEAD).Trim() + "`r`n" +
        (($ordered | ForEach-Object { $_ }) -join "`r`n") + "`r`n"
$verdict = "RESULT=FAIL"
if ($script:failures -eq 0 -and $script:stops.Count -eq 0) {
    $verdict = "RESULT=PASS"
}
$text += "$verdict; checks=14; failed=$($script:failures)`r`n"
[System.IO.File]::WriteAllText((Join-Path $OutputDir "protection-checks.txt"), $text, $utf8NoBom)
Write-Host "$verdict; checks=14; failed=$($script:failures)"
if ($verdict -eq "RESULT=PASS") {
    exit 0
}
exit 1
