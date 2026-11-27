<#
.SYNOPSIS
The documentation screenshots (#616): dot-sourced by
tools\Verify-SACCR-Branch.ps1 -Screenshots, which calls
Invoke-DocumentationScreenshots in its own Excel session.
.DESCRIPTION
Invoke-DocumentationScreenshots takes the development workbook the run built
and opened, after the smoke suites, and pictures the sheets of the
documentation in four states, then two release checks, each prepared through
the workbook's own silent entry points:
1. template: Clear outputs, so that the input sheets show the template's
   content;
2. staging: the example files of the import specification loaded onto
   Import_NettingSets and Import_Trades (ImportCsv_Silent);
3. run: a run of the demonstration portfolio (RunSACCR_Silent);
4. bank: the bank figures of the reconciliation screenshot entered on
   BankNettingSets (BCBS-EX1 as on Results, MATCH; 5,300 for BCBS-EX3,
   DIFFERENCE; BANK-ONLY-1, MISSING IN CALCULATOR), bank figures per leg on
   BankTrades for the TradeReconciliation check, and a run;
5. tradefilter: TradeReconciliation filtered to the Notional measure, whose
   tiles then show the measure and its sums (#819, #821);
6. large: every notional and MtM on Trades times 70 and a run, so that the
   EAD is about a billion, for the Dashboard cards (#810).
The pictures of 5 and 6 are named check-*, not SACCR_*: they are release
checks, which tools\screenshots.py leaves out of docs\assets.
Each picture is Range.CopyPicture of rows and columns that the plan below
names, through the review's Save-SheetPicture (tools\Review-SACCR-Sheets.ps1),
into BUNDLE\screenshots\NAME.png. A picture that hides columns unprotects
its sheet, hides them, pictures and shows them again; the workbook is closed
without saving. The pictures are not processed here: tools\screenshots.py
reduces them to the 256-colour palette of docs\assets. The results go to
BUNDLE\screenshots.txt; the function returns its last line, RESULT=PASS or
RESULT=FAIL.
#>

Set-StrictMode -Version Latest

function Get-ScreenshotPlan {
    # The pictures of the documentation, by state: the file, the sheet, the
    # first and last rows, the last column, and the columns to hide (a range
    # such as "E:S", or "" for none). A last row of 0 is the sheet's content.
    return @(
        @{ State = "template"; File = "SACCR_Home"; Sheet = "Home"; First = 1; Last = 0; Column = "S"; Hide = "" },
        @{ State = "template"; File = "SACCR_NettingSets"; Sheet = "NettingSets"; First = 1; Last = 26; Column = "T"; Hide = "" },
        @{ State = "template"; File = "SACCR_Trades"; Sheet = "Trades"; First = 1; Last = 34; Column = "R"; Hide = "" },
        @{ State = "template"; File = "SACCR_Params"; Sheet = "Params"; First = 1; Last = 29; Column = "G"; Hide = "" },
        @{ State = "template"; File = "SACCR_Params_Supervisory"; Sheet = "Params"; First = 30; Last = 57; Column = "K"; Hide = "" },
        @{ State = "template"; File = "SACCR_README_sheet"; Sheet = "README"; First = 1; Last = 45; Column = "S"; Hide = "" },
        @{ State = "template"; File = "SACCR_TestCatalogue"; Sheet = "TestCatalogue"; First = 1; Last = 40; Column = "P"; Hide = "" },
        @{ State = "template"; File = "SACCR_FormulaTest"; Sheet = "FormulaTest"; First = 1; Last = 45; Column = "P"; Hide = "" },
        @{ State = "staging"; File = "SACCR_Import_NettingSets"; Sheet = "Import_NettingSets"; First = 1; Last = 0; Column = "U"; Hide = "" },
        @{ State = "staging"; File = "SACCR_Import_Trades"; Sheet = "Import_Trades"; First = 1; Last = 0; Column = "S"; Hide = "" },
        @{ State = "run"; File = "SACCR_Dashboard"; Sheet = "Dashboard"; First = 1; Last = 60; Column = "M"; Hide = "" },
        @{ State = "run"; File = "SACCR_Dashboard_Charts"; Sheet = "Dashboard"; First = 61; Last = 140; Column = "M"; Hide = "" },
        @{ State = "run"; File = "SACCR_Results_AddOns"; Sheet = "Results"; First = 1; Last = 0; Column = "S"; Hide = "" },
        @{ State = "run"; File = "SACCR_Results_EAD"; Sheet = "Results"; First = 1; Last = 0; Column = "AD"; Hide = "E:S" },
        @{ State = "run"; File = "SACCR_Checks"; Sheet = "Checks"; First = 1; Last = 0; Column = "J"; Hide = "" },
        @{ State = "run"; File = "SACCR_TradeCalc_Inputs"; Sheet = "TradeCalc"; First = 1; Last = 34; Column = "P"; Hide = "" },
        @{ State = "run"; File = "SACCR_TradeCalc_AddOns"; Sheet = "TradeCalc"; First = 1; Last = 34; Column = "AA"; Hide = "E:P" },
        @{ State = "run"; File = "SACCR_HedgingSets"; Sheet = "HedgingSets"; First = 1; Last = 34; Column = "P"; Hide = "" },
        @{ State = "run"; File = "SACCR_Buckets"; Sheet = "Buckets"; First = 1; Last = 34; Column = "M"; Hide = "" },
        @{ State = "run"; File = "SACCR_Basel_vs_CRR"; Sheet = "Basel_vs_CRR"; First = 1; Last = 16; Column = "K"; Hide = "" },
        @{ State = "bank"; File = "SACCR_BankNettingSets"; Sheet = "BankNettingSets"; First = 1; Last = 14; Column = "P"; Hide = "" },
        @{ State = "bank"; File = "SACCR_Reconciliation"; Sheet = "Reconciliation"; First = 1; Last = 0; Column = "N"; Hide = "" },
        @{ State = "bank"; File = "SACCR_ReconSummary"; Sheet = "ReconSummary"; First = 1; Last = 0; Column = "L"; Hide = "" },
        # Release checks, not documentation images (their names do not start with SACCR_, so
        # tools\screenshots.py leaves them out): the TradeReconciliation tiles (#819, #821) and the
        # Dashboard cards on a portfolio of about a billion (#810).
        @{ State = "bank"; File = "check-TradeReconciliation"; Sheet = "TradeReconciliation"; First = 1; Last = 30; Column = "O"; Hide = "" },
        @{ State = "tradefilter"; File = "check-TradeReconciliation-Notional"; Sheet = "TradeReconciliation"; First = 1; Last = 0; Column = "O"; Hide = "" },
        @{ State = "large"; File = "check-Dashboard-large"; Sheet = "Dashboard"; First = 1; Last = 30; Column = "M"; Hide = "" }
    )
}

function Get-ColumnNumber {
    # The number of a column from its letters: A is 1, AD is 30.
    param([string]$Letters)
    $number = 0
    foreach ($letter in $Letters.ToUpperInvariant().ToCharArray()) {
        $number = $number * 26 + ([int]$letter - 64)
    }
    return $number
}

function Get-NamedRow {
    # The row of the first cell of a workbook name.
    param($Workbook, [string]$Name)
    return [int]$Workbook.Names.Item($Name).RefersToRange.Row
}

function Save-Screenshot {
    # Picture one entry of the plan into Folder; returns a line for the report.
    param($Workbook, $Shot, $PictureSheet, [string]$Folder)
    $sheet = $Workbook.Worksheets.Item($Shot.Sheet)
    $last = [int]$Shot.Last
    if ($last -le 0) {
        $bound = Get-ContentBound $sheet
        $last = [Math]::Max([int]$bound[0], [int]$Shot.First)
    }
    $column = Get-ColumnNumber $Shot.Column
    $range = $sheet.Range($sheet.Cells.Item([int]$Shot.First, 1), $sheet.Cells.Item($last, $column))
    $path = Join-Path $Folder ($Shot.File + ".png")
    $hidden = $null
    try {
        if ($Shot.Hide -ne "") {
            if ($sheet.ProtectContents) {
                [void]$sheet.Unprotect()
            }
            $hidden = $sheet.Range($Shot.Hide).EntireColumn
            $hidden.Hidden = $true
        }
        $attempts = Save-SheetPicture $sheet $range $PictureSheet $path
    }
    finally {
        if ($null -ne $hidden) {
            $hidden.Hidden = $false
        }
    }
    if ($attempts -gt 0) {
        return ("{0} {1}: {2} rows {3} to {4}, columns A to {5}{6}" -f $Shot.State, $Shot.File, $Shot.Sheet,
                $Shot.First, $last, $Shot.Column, $(if ($Shot.Hide) { " without " + $Shot.Hide } else { "" }))
    }
    return ("{0} {1}: FAIL, no picture" -f $Shot.State, $Shot.File)
}

function Write-ScreenshotTrace {
    # Append a line to screenshots.txt as the work goes on, so that a run the
    # watchdog of Verify-SACCR-Branch.ps1 stops still shows its last step; the
    # finished report replaces the file.
    param([string]$Text)
    if (-not [string]::IsNullOrEmpty($script:screenshotTraceFile)) {
        [System.IO.File]::AppendAllText($script:screenshotTraceFile,
            ((Get-Date).ToString("HH:mm:ss") + "  " + $Text + [Environment]::NewLine), (New-Object System.Text.UTF8Encoding($false)))
    }
}

function Set-ScreenshotState {
    # Prepare a state through the workbook's silent entry points; returns the
    # lines for the report, the last of which starts with OK or FAIL.
    param($Excel, $Workbook, [string]$State, [string]$Bundle)
    $qualified = "'" + $Workbook.Name + "'!"
    $lines = New-Object System.Collections.Generic.List[string]
    switch ($State) {
        "template" {
            $outcome = [string]$Excel.Run($qualified + "M_Main.ClearOutputs_Silent")
            $lines.Add("template: ClearOutputs_Silent: " + $outcome)
            $ok = $outcome.StartsWith("RESULT=OK;")
        }
        "staging" {
            $ok = $true
            $examples = Join-Path $Bundle "examples"
            New-Item -ItemType Directory -Path $examples -Force | Out-Null
            $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
            foreach ($kind in @("NettingSets", "Trades")) {
                $text = [string]$Excel.Run($qualified + "CORE_ImportSpec.ImportSpecExample", $kind)
                $file = Join-Path $examples ("K-SACCR_" + $kind + "_example.csv")
                [System.IO.File]::WriteAllText($file, $text, $utf8NoBom)
                $outcome = [string]$Excel.Run($qualified + "M_Main.ImportCsv_Silent", $kind, $file, $true)
                $lines.Add(("staging: ImportCsv_Silent {0}: {1}" -f $kind, ($outcome -replace [regex]::Escape($examples), "<examples>")))
                $ok = $ok -and $outcome.StartsWith("RESULT=LOADED;")
            }
        }
        "run" {
            $outcome = [string]$Excel.Run($qualified + "M_Main.RunSACCR_Silent")
            $lines.Add("run: RunSACCR_Silent: " + $outcome)
            $ok = $outcome.StartsWith("RESULT=OK;")
        }
        "bank" {
            # The figures of the reconciliation screenshot: BCBS-EX1 as on Results, 5,300 for BCBS-EX3 (beyond the
            # relative tolerance of 1%), and a
            # netting set the calculator does not have.
            $results = $Workbook.Worksheets.Item("Results")
            $header = Get-NamedRow $Workbook "Results_Header"
            $eadColumn = 0
            for ($c = 3; $c -le 30; $c++) {
                if ([string]$results.Cells.Item($header, $c).Value2 -eq "EAD") {
                    $eadColumn = $c
                }
            }
            $ead = $null
            for ($r = $header + 1; $r -le $header + 60; $r++) {
                if ([string]$results.Cells.Item($r, 3).Value2 -eq "BCBS-EX1") {
                    $ead = $results.Cells.Item($r, $eadColumn).Value2
                }
            }
            Write-ScreenshotTrace "bank: ClearBank_Silent"
            $outcome = [string]$Excel.Run($qualified + "M_Main.ClearBank_Silent", "BankNettingSets")
            $lines.Add("bank: ClearBank_Silent: " + $outcome)
            $bank = $Workbook.Worksheets.Item("BankNettingSets")
            $first = (Get-NamedRow $Workbook "BankNettingSets_Header") + 1
            if ($null -eq $ead) {
                $lines.Add("bank: no EAD for BCBS-EX1 on Results")
                $lines.Add("FAIL " + $State)
                return $lines.ToArray()
            }
            $eadValue = [double]$ead
            $ids = [string[]]@("BCBS-EX1", "BCBS-EX3", "BANK-ONLY-1")
            $figures = [double[]]@($eadValue, 5300.0, 100.0)
            $comments = [string[]]@("As on Results: MATCH", "5,300 against 5,405: DIFFERENCE",
                                    "Not on NettingSets: MISSING IN CALCULATOR")
            Write-ScreenshotTrace "bank: writing the bank figures"
            for ($i = 0; $i -lt 3; $i++) {
                $row = [int]($first + $i)
                $bank.Cells.Item($row, 3).Value2 = [string]$ids[$i]
                $bank.Cells.Item($row, 4).Value2 = [string]"BCBS"
                $bank.Cells.Item($row, 5).Value2 = [double]$figures[$i]
                $bank.Cells.Item($row, 16).Value2 = [string]$comments[$i]
            }
            $lines.Add([string]("bank: BCBS-EX1 EAD " + $eadValue.ToString([System.Globalization.CultureInfo]::InvariantCulture) +
                                " from Results, BCBS-EX3 5300, BANK-ONLY-1 100"))
            # Bank figures per trade leg for the TradeReconciliation check: EX1-T1 and EX1-T2 as on Trades
            # (MATCH), 4,800 for EX1-T3 against 5,000 (DIFFERENCE), and a leg the calculator does not have.
            Write-ScreenshotTrace "bank: BankTrades"
            $outcome = [string]$Excel.Run($qualified + "M_Main.ClearBank_Silent", "BankTrades")
            $lines.Add("bank: ClearBank_Silent BankTrades: " + $outcome)
            $bankTrades = Join-Path $Bundle "bank-trades-check.csv"
            $csv = "trade_id;notional;mtm;maturity;supervisory_duration;adjusted_notional;supervisory_delta;" +
                   "mf_unmargined;mf_margined;effective_notional_unmargined;effective_notional_margined;" +
                   "standalone_addon;comment`r`n" +
                   "EX1-T1;10000;30;;;;;;;;;;As on Trades`r`n" +
                   "EX1-T2;10000;-20;;;;;;;;;;As on Trades`r`n" +
                   "EX1-T3;4800;;;;;;;;;;;4,800 against 5,000`r`n" +
                   "BANK-LEG-1;100;;;;;;;;;;;Not on Trades`r`n"
            [System.IO.File]::WriteAllText($bankTrades, $csv, (New-Object System.Text.UTF8Encoding($false)))
            $outcome = [string]$Excel.Run($qualified + "M_Main.ImportBankCsv_Silent", "BankTrades", $bankTrades, $true)
            $lines.Add("bank: ImportBankCsv_Silent BankTrades: " + ($outcome -replace [regex]::Escape($Bundle), "<bundle>"))
            Write-ScreenshotTrace "bank: RunSACCR_Silent"
            $outcome = [string]$Excel.Run($qualified + "M_Main.RunSACCR_Silent")
            $lines.Add("bank: RunSACCR_Silent: " + $outcome)
            $ok = ($null -ne $ead) -and $outcome.StartsWith("RESULT=OK;") -and $outcome.Contains("; reconciliation=DONE;")
        }
        "tradefilter" {
            # TradeReconciliation filtered to the Notional measure: its tiles then read the measure and its sums.
            $recon = $Workbook.Worksheets.Item("TradeReconciliation")
            $header = Get-NamedRow $Workbook "TradeReconciliation_Header"
            [void]$recon.Range($recon.Cells.Item($header, 3), $recon.Cells.Item($header, 15)).AutoFilter(6, "Notional")
            foreach ($address in @("R2", "X4", "X5", "Z1", "Z2", "Z3")) {
                $lines.Add(("tradefilter: TradeReconciliation {0} = {1}" -f $address, [string]$recon.Range($address).Text))
            }
            $ok = $true
        }
        "large" {
            # The demonstration portfolio with every notional and MtM on Trades times 70, so that the EAD is
            # about a billion, for the Dashboard cards (#810).
            $trades = $Workbook.Worksheets.Item("Trades")
            $header = Get-NamedRow $Workbook "Trades_Header"
            [void]$trades.Unprotect()
            $scaled = 0
            for ($r = $header + 1; $r -le $header + 200; $r++) {
                if ([string]$trades.Cells.Item($r, 3).Value2 -eq "") {
                    continue
                }
                foreach ($column in @(13, 15)) {
                    $value = $trades.Cells.Item($r, $column).Value2
                    if ($value -is [double]) {
                        $trades.Cells.Item($r, $column).Value2 = [double]$value * 70.0
                    }
                }
                $scaled++
            }
            $lines.Add(("large: notional and MtM of {0} legs times 70" -f $scaled))
            Write-ScreenshotTrace "large: RunSACCR_Silent"
            $outcome = [string]$Excel.Run($qualified + "M_Main.RunSACCR_Silent")
            $lines.Add("large: RunSACCR_Silent: " + $outcome)
            $ok = $outcome.StartsWith("RESULT=OK;")
        }
        default {
            $ok = $false
            $lines.Add("Unknown state: " + $State)
        }
    }
    if ($ok) {
        $lines.Add("OK " + $State)
    }
    else {
        $lines.Add("FAIL " + $State)
    }
    return $lines.ToArray()
}

function Invoke-DocumentationScreenshots {
    # The documentation screenshots of the open development workbook; see the
    # description above. Returns the last line of screenshots.txt.
    param($Excel, $Workbook, [string]$Bundle, [string]$RepoRoot)
    . (Join-Path $RepoRoot "tools\Review-SACCR-Sheets.ps1")
    $folder = Join-Path $Bundle "screenshots"
    New-Item -ItemType Directory -Path $folder -Force | Out-Null
    $report = New-Object System.Collections.Generic.List[string]
    $report.Add("Documentation screenshots (#616), tools\Take-SACCR-Screenshots.ps1")
    $script:screenshotTraceFile = Join-Path $Bundle "screenshots.txt"
    [System.IO.File]::WriteAllText($script:screenshotTraceFile, "", (New-Object System.Text.UTF8Encoding($false)))
    $plan = @(Get-ScreenshotPlan)
    $taken = 0
    $failed = 0
    $pictureBook = $null
    # Without events for the whole session: activating the workbook or a sheet, or editing a cell, makes
    # the workbook refresh its Ribbon (Workbook_Activate, Workbook_SheetActivate, Workbook_SheetChange);
    # Excel then asks for the callbacks once the picture workbook is the active one, does not find them
    # there and stops on "Cannot run the macro 'RibbonGetEnabled'" until someone closes the message.
    $events = [bool]$Excel.EnableEvents
    try {
        $Excel.EnableEvents = $false
        [void]$Workbook.Activate()
        $pictureBook = $Excel.Workbooks.Add()
        $pictureSheet = $pictureBook.Worksheets.Item(1)
        foreach ($state in @("template", "staging", "run", "bank", "tradefilter", "large")) {
            [void]$Workbook.Activate()
            Write-ScreenshotTrace ("preparing the state " + $state)
            try {
                $lines = @(Set-ScreenshotState $Excel $Workbook $state $Bundle)
            }
            catch {
                $lines = @(("{0}: stopped at line {1}: {2} ({3})" -f $state, $_.InvocationInfo.ScriptLineNumber,
                            $_.Exception.Message, $_.InvocationInfo.Line.Trim()), ("FAIL " + $state))
            }
            foreach ($line in $lines) {
                $report.Add($line)
                Write-ScreenshotTrace $line
            }
            if (-not $lines[-1].StartsWith("OK ")) {
                $failed++
                continue
            }
            foreach ($shot in @($plan | Where-Object { $_.State -eq $state })) {
                Write-ScreenshotTrace ("picturing " + $shot.File)
                try {
                    $line = Save-Screenshot $Workbook $shot $pictureSheet $folder
                }
                catch {
                    $line = ("{0} {1}: FAIL, {2}" -f $state, $shot.File, $_.Exception.Message)
                }
                $report.Add($line)
                if ($line.Contains(": FAIL")) {
                    $failed++
                }
                else {
                    $taken++
                }
            }
        }
    }
    catch {
        $report.Add(("Stopped at line {0}: {1}" -f $_.InvocationInfo.ScriptLineNumber, $_.Exception.Message))
        $failed++
    }
    finally {
        if ($null -ne $pictureBook) {
            try {
                [void]$pictureBook.Close($false)
            }
            catch {
            }
        }
        try {
            [void]$Workbook.Activate()
        }
        catch {
        }
        try {
            $Excel.EnableEvents = $events
        }
        catch {
        }
    }
    $verdict = "PASS"
    if ($failed -gt 0 -or $taken -ne $plan.Count) {
        $verdict = "FAIL"
    }
    $report.Add(("RESULT={0}; pictures={1}/{2}; failed={3}" -f $verdict, $taken, $plan.Count, $failed))
    $script:screenshotTraceFile = ""
    [System.IO.File]::WriteAllText((Join-Path $Bundle "screenshots.txt"),
        (($report -join [Environment]::NewLine) + [Environment]::NewLine), (New-Object System.Text.UTF8Encoding($false)))
    return $report[$report.Count - 1]
}
