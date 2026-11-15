<#
.SYNOPSIS
The sheet review of the quick run (#706): dot-sourced by
tools\Verify-SACCR-Branch.ps1, which calls Invoke-SheetReview in its own
Excel session.
.DESCRIPTION
Invoke-SheetReview takes the development workbook the quick run built and
opened, after Test_RunAll:
1. runs Clear outputs (M_Main.ClearOutputs_Silent), takes a picture of every
   sheet into BUNDLE\sheets\NN-Sheet-cleared.png and checks every sheet;
2. runs the demonstration portfolio (M_Main.RunSACCR_Silent), checks every
   sheet again and takes a second picture, NN-Sheet-run.png.
The checks, against docs\workbook\sheet-style.json and the roles of
docs\workbook\protection.json, are: every cell of the used range has the
style's font and vertical alignment; no cell displays ####; no cell holds a
date unless it has the style's date format; gridlines are off; the tab has the
color of the sheet's role. A picture is Range.CopyPicture pasted into a chart
of a separate temporary workbook, so never inside the copied range; the paste
is retried until the chart holds a shape, and a picture of one color only is
never counted. The results go to BUNDLE\sheet-review.txt; the function returns
its last line, RESULT=PASS or RESULT=FAIL.
#>

Set-StrictMode -Version Latest

$script:reviewPictureRows = 60
$script:reviewPictureColumns = 40
$script:reviewListed = 8
$script:xlCenter = -4108
$script:xlScreen = 1
$script:xlBitmap = 2
$script:xlPicture = -4147
$script:verticalNames = @{ -4108 = "center"; -4107 = "bottom"; -4160 = "top"; -4130 = "justify"; -4117 = "distributed" }

function ConvertTo-ExcelColor {
    # "RRGGBB" as Excel's color number, red + 256 * green + 65536 * blue.
    param([string]$Hex)
    $red = [Convert]::ToInt32($Hex.Substring(0, 2), 16)
    $green = [Convert]::ToInt32($Hex.Substring(2, 2), 16)
    $blue = [Convert]::ToInt32($Hex.Substring(4, 2), 16)
    return [double]($red + 256 * $green + 65536 * $blue)
}

function ConvertTo-HexColor {
    # Excel's color number as "RRGGBB".
    param([double]$Color)
    $value = [int64]$Color
    $red = [int]($value % 256)
    $green = [int](([int64][Math]::Floor($value / 256)) % 256)
    $blue = [int](([int64][Math]::Floor($value / 65536)) % 256)
    return ("{0:X2}{1:X2}{2:X2}" -f $red, $green, $blue)
}

function Get-CellArray {
    # The values of a range as a 1-based two-dimensional array, also for a
    # single cell, which Excel returns as a scalar.
    param($Range, [string]$Property)
    $value = $Range.$Property
    if ($value -is [System.Array]) {
        return , $value
    }
    $array = [System.Array]::CreateInstance([object], @(1, 1), @(1, 1))
    $array.SetValue($value, 1, 1)
    return , $array
}

function Test-Blank {
    # True when a value counts as an empty cell.
    param($Value)
    return ($null -eq $Value -or $Value -is [System.DBNull] -or ($Value -is [string] -and $Value -eq ""))
}

function Get-ContentBound {
    # The last row and column of the sheet holding a value or a formula, from
    # one read of the used range; @(0, 0) for an empty sheet.
    param($Sheet)
    $used = $Sheet.UsedRange
    $formulas = Get-CellArray $used "Formula"
    $lastRow = 0
    $lastColumn = 0
    for ($r = 1; $r -le $formulas.GetLength(0); $r++) {
        for ($c = 1; $c -le $formulas.GetLength(1); $c++) {
            if (-not (Test-Blank $formulas[$r, $c])) {
                $lastRow = [Math]::Max($lastRow, $used.Row + $r - 1)
                $lastColumn = [Math]::Max($lastColumn, $used.Column + $c - 1)
            }
        }
    }
    return @($lastRow, $lastColumn)
}

function Format-Listed {
    # At most the first few addresses, and how many more there are.
    param([System.Collections.Generic.List[string]]$Items)
    $shown = @($Items | Select-Object -First $script:reviewListed)
    $text = $shown -join ", "
    if ($Items.Count -gt $shown.Count) {
        $text += (" and {0} more" -f ($Items.Count - $shown.Count))
    }
    return $text
}

function Test-CellStyle {
    # Cells of the used range whose font or vertical alignment differ from the
    # style: the whole range first, then each mixed column, then its cells.
    param($Sheet, [string]$Font)
    $fonts = New-Object System.Collections.Generic.List[string]
    $aligns = New-Object System.Collections.Generic.List[string]
    $used = $Sheet.UsedRange
    $wholeFont = $used.Font.Name
    $wholeAlign = $used.VerticalAlignment
    if (($wholeFont -is [string] -and $wholeFont -eq $Font) -and
        (-not ($wholeAlign -is [System.DBNull]) -and [int]$wholeAlign -eq $script:xlCenter)) {
        return @{ Fonts = $fonts; Aligns = $aligns }
    }
    for ($c = 1; $c -le $used.Columns.Count; $c++) {
        $column = $used.Columns.Item($c)
        $columnFont = $column.Font.Name
        $columnAlign = $column.VerticalAlignment
        $fontMixed = -not ($columnFont -is [string] -and $columnFont -eq $Font)
        $alignMixed = ($columnAlign -is [System.DBNull]) -or [int]$columnAlign -ne $script:xlCenter
        if (-not ($fontMixed -or $alignMixed)) {
            continue
        }
        for ($r = 1; $r -le $column.Rows.Count; $r++) {
            $cell = $column.Cells.Item($r, 1)
            if ($fontMixed -and [string]$cell.Font.Name -ne $Font) {
                $fonts.Add(($cell.Address($false, $false) + " " + [string]$cell.Font.Name))
            }
            if ($alignMixed -and [int]$cell.VerticalAlignment -ne $script:xlCenter) {
                $name = $script:verticalNames[[int]$cell.VerticalAlignment]
                if ($null -eq $name) {
                    $name = [string]$cell.VerticalAlignment
                }
                $aligns.Add(($cell.Address($false, $false) + " " + $name))
            }
        }
    }
    return @{ Fonts = $fonts; Aligns = $aligns }
}

function Test-CellValue {
    # Cells showing #### (only a number or a date can) and cells holding a date
    # without the style's date format.
    param($Sheet, [string]$DateFormat)
    $hashes = New-Object System.Collections.Generic.List[string]
    $dates = New-Object System.Collections.Generic.List[string]
    $used = $Sheet.UsedRange
    $values = Get-CellArray $used "Value"
    for ($r = 1; $r -le $values.GetLength(0); $r++) {
        for ($c = 1; $c -le $values.GetLength(1); $c++) {
            $value = $values[$r, $c]
            if (-not ($value -is [double] -or $value -is [datetime] -or $value -is [decimal])) {
                continue
            }
            $cell = $used.Cells.Item($r, $c)
            if ([string]$cell.Text -match '^\s*#+\s*$') {
                $hashes.Add($cell.Address($false, $false))
            }
            if ($value -is [datetime] -and [string]$cell.NumberFormat -ne $DateFormat) {
                $dates.Add(($cell.Address($false, $false) + " " + [string]$cell.NumberFormat))
            }
        }
    }
    return @{ Hashes = $hashes; Dates = $dates }
}

function Test-PictureFile {
    # True when the PNG exists and shows more than one color.
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        return $false
    }
    Add-Type -AssemblyName System.Drawing
    $bitmap = [System.Drawing.Bitmap]::FromFile($Path)
    try {
        $first = $bitmap.GetPixel(0, 0).ToArgb()
        $stepX = [Math]::Max(1, [int]($bitmap.Width / 60))
        $stepY = [Math]::Max(1, [int]($bitmap.Height / 60))
        for ($y = 0; $y -lt $bitmap.Height; $y += $stepY) {
            for ($x = 0; $x -lt $bitmap.Width; $x += $stepX) {
                if ($bitmap.GetPixel($x, $y).ToArgb() -ne $first) {
                    return $true
                }
            }
        }
        return $false
    }
    finally {
        $bitmap.Dispose()
    }
}

function Save-SheetPicture {
    # Picture the range into Path through a chart of the temporary workbook;
    # returns the number of attempts it took, or 0 when every one failed.
    param($Sheet, $Range, $PictureSheet, [string]$Path)
    for ($attempt = 1; $attempt -le 6; $attempt++) {
        $holder = $null
        try {
            [void]$Sheet.Parent.Activate()
            [void]$Sheet.Activate()
            # A1 holds no list, so no dropdown arrow of a selected cell shows in the picture (#708).
            try {
                [void]$Sheet.Range("A1").Select()
            }
            catch {
            }
            $format = $script:xlBitmap
            if ($attempt -gt 3) {
                $format = $script:xlPicture
            }
            [void]$Range.CopyPicture($script:xlScreen, $format)
            [void]$PictureSheet.Parent.Activate()
            $holder = $PictureSheet.ChartObjects().Add(0, 0, [double]$Range.Width, [double]$Range.Height)
            [void]$holder.Activate()
            [void]$holder.Chart.Paste()
            if ($holder.Chart.Shapes.Count -gt 0) {
                if (Test-Path -LiteralPath $Path) {
                    Remove-Item -LiteralPath $Path -Force
                }
                [void]$holder.Chart.Export($Path, "PNG")
                if (Test-PictureFile $Path) {
                    return $attempt
                }
                Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
            }
        }
        catch {
            # CopyPicture or Paste failed, as they sometimes do: try again.
        }
        finally {
            if ($null -ne $holder) {
                try {
                    [void]$holder.Delete()
                }
                catch {
                }
            }
        }
        Start-Sleep -Milliseconds (300 * $attempt)
    }
    return 0
}

function Invoke-SheetState {
    # Check and picture every sheet in one state ("cleared" or "run"); returns
    # the report lines and the number of failed checks.
    param($Excel, $Workbook, [string]$State, $Style, $Roles, $PictureSheet, [string]$Folder)
    $lines = New-Object System.Collections.Generic.List[string]
    $failed = 0
    $pictures = 0
    $index = 0
    foreach ($sheet in @($Workbook.Worksheets)) {
        $index++
        $name = [string]$sheet.Name
        $results = New-Object System.Collections.Generic.List[string]
        $details = New-Object System.Collections.Generic.List[string]

        $cellStyle = Test-CellStyle $sheet $Style.font
        $cellValues = Test-CellValue $sheet $Style.date_format
        $checks = @(
            @("font", $cellStyle.Fonts, ("not " + $Style.font)),
            @("valign", $cellStyle.Aligns, "not vertically centered"),
            @("hashes", $cellValues.Hashes, "shows ####"),
            @("dates", $cellValues.Dates, ("a date not formatted " + $Style.date_format))
        )
        foreach ($check in $checks) {
            if ($check[1].Count -eq 0) {
                $results.Add($check[0] + " PASS")
            }
            else {
                $results.Add($check[0] + " FAIL")
                $details.Add(("  " + $check[0] + ": " + $check[1].Count + " cell(s) " + $check[2] + ": " +
                              (Format-Listed $check[1])))
                $failed++
            }
        }

        [void]$Workbook.Activate()
        [void]$sheet.Activate()
        if ($Excel.ActiveWindow.DisplayGridlines) {
            $results.Add("gridlines FAIL")
            $details.Add("  gridlines: shown")
            $failed++
        }
        else {
            $results.Add("gridlines PASS")
        }
        $role = $null
        $roleEntry = $Roles.PSObject.Properties[$name]
        if ($null -ne $roleEntry) {
            $role = [string]$roleEntry.Value.role
        }
        $expected = $null
        if ($null -ne $role -and $null -ne $Style.tab_colors.PSObject.Properties[$role]) {
            $expected = [string]$Style.tab_colors.PSObject.Properties[$role].Value
        }
        $tab = $sheet.Tab.Color
        $tabHex = "none"
        if ($tab -is [double] -or $tab -is [int]) {
            $tabHex = ConvertTo-HexColor ([double]$tab)
        }
        if ($null -ne $expected -and $tabHex -eq $expected) {
            $results.Add("tab PASS")
        }
        else {
            $results.Add("tab FAIL")
            $details.Add(("  tab: " + $tabHex + ", expected " + [string]$expected + " for role " + [string]$role))
            $failed++
        }

        $bounds = Get-ContentBound $sheet
        $lastRow = [Math]::Min([Math]::Max($bounds[0], 1), $script:reviewPictureRows)
        $lastColumn = [Math]::Min([Math]::Max($bounds[1], 1), $script:reviewPictureColumns)
        $range = $sheet.Range($sheet.Cells.Item(1, 1), $sheet.Cells.Item($lastRow, $lastColumn))
        $file = "{0:D2}-{1}-{2}.png" -f $index, ($name -replace '[^A-Za-z0-9_-]', '_'), $State
        $attempts = Save-SheetPicture $sheet $range $PictureSheet (Join-Path $Folder $file)
        if ($attempts -gt 0) {
            $pictures++
            $note = ""
            if ($bounds[0] -gt $lastRow -or $bounds[1] -gt $lastColumn) {
                $note = (" (first {0} rows and {1} columns of {2} x {3})" -f $lastRow, $lastColumn, $bounds[0], $bounds[1])
            }
            $results.Add("picture PASS")
            $details.Add(("  picture: sheets\" + $file + ", " + $range.Address($false, $false) + $note +
                          ", attempt " + $attempts))
        }
        else {
            $results.Add("picture FAIL")
            $details.Add(("  picture: sheets\" + $file + " stayed empty after every attempt"))
            $failed++
        }
        $lines.Add(("{0} {1}: {2}" -f $State, $name, ($results -join "; ")))
        foreach ($detail in $details) {
            $lines.Add($detail)
        }
    }
    return @{ Lines = $lines; Failed = $failed; Pictures = $pictures; Sheets = $index }
}

function Invoke-SheetReview {
    # The sheet review of the quick run; writes sheet-review.txt and the
    # pictures, and returns the last line of sheet-review.txt.
    param($Excel, $Workbook, [string]$Bundle, [string]$RepoRoot, [string]$Sha)
    $styleRules = Get-Content -LiteralPath (Join-Path $RepoRoot "docs\workbook\sheet-style.json") -Raw -Encoding UTF8 |
        ConvertFrom-Json
    $roles = (Get-Content -LiteralPath (Join-Path $RepoRoot "docs\workbook\protection.json") -Raw -Encoding UTF8 |
        ConvertFrom-Json).sheets
    $folder = Join-Path $Bundle "sheets"
    New-Item -ItemType Directory -Path $folder -Force | Out-Null
    $report = New-Object System.Collections.Generic.List[string]
    $report.Add("K-SACCR sheet review of $Sha (#706), against docs\workbook\sheet-style.json")
    $qualified = "'" + $Workbook.Name + "'!"
    $pictureBook = $null
    $failed = 0
    $pictures = 0
    $sheets = 0
    try {
        $pictureBook = $Excel.Workbooks.Add()
        $pictureSheet = $pictureBook.Worksheets.Item(1)
        foreach ($step in @(@("cleared", "M_Main.ClearOutputs_Silent"), @("run", "M_Main.RunSACCR_Silent"))) {
            # Only a completed operation gives a state worth reviewing: a run
            # that stopped (RESULT=STOPPED) leaves no results to picture.
            $outcome = [string]$Excel.Run($qualified + $step[1])
            $report.Add(($step[1] + ": " + $outcome))
            if (-not $outcome.StartsWith("RESULT=OK;")) {
                $report.Add(("Stopped: " + $step[1] + " did not complete; the " + $step[0] + " state is not reviewed"))
                $failed++
                break
            }
            $state = Invoke-SheetState $Excel $Workbook $step[0] $styleRules $roles $pictureSheet $folder
            foreach ($line in $state.Lines) {
                $report.Add($line)
            }
            $failed += $state.Failed
            $pictures += $state.Pictures
            $sheets = $state.Sheets
        }
    }
    catch {
        $report.Add(("Stopped: " + $_.Exception.Message))
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
    }
    $verdict = "PASS"
    if ($failed -gt 0 -or $pictures -ne 2 * $sheets -or $sheets -eq 0) {
        $verdict = "FAIL"
    }
    $report.Add(("RESULT={0}; sheets={1}; states=2; failed_checks={2}; pictures={3}/{4}" -f
                 $verdict, $sheets, $failed, $pictures, (2 * $sheets)))
    [System.IO.File]::WriteAllText((Join-Path $Bundle "sheet-review.txt"),
                                   (($report -join "`r`n") + "`r`n"), (New-Object System.Text.UTF8Encoding($false)))
    return $report[$report.Count - 1]
}
