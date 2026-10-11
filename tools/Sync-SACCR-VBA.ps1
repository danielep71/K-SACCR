<#
.SYNOPSIS
Synchronize repository VBA source into an existing local SACCR .xlsm workbook.

.DESCRIPTION
Updates VBA code only, from src\core, src\modules, src\workbook and
tests\modules, so the workbook holds every engine and test module. Standard
modules are replaced from source. ThisWorkbook and worksheet document modules
keep their workbook objects and only their code text is replaced. The script
does not write cells, formulas, names, formatting, tables, worksheets, or
workbook structure.

.REQUIREMENTS
- Windows desktop Excel.
- Excel > Options > Trust Center > Macro Settings >
  "Trust access to the VBA project object model" enabled.
- VBA project not password-locked.

.EXAMPLE
.\tools\Sync-SACCR-VBA.ps1 -WorkbookPath "C:\SACCR\SACCR.xlsm"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$WorkbookPath,

    [Parameter(Mandatory = $false)]
    [string]$RepoRoot,

    [Parameter(Mandatory = $false)]
    [switch]$NoBackup,

    [Parameter(Mandatory = $false)]
    [switch]$Visible
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Resolve the repository root after parameter binding. $PSScriptRoot is not
# reliable as a parameter default in every Windows PowerShell invocation mode.
if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $scriptPath = $MyInvocation.MyCommand.Path
    if ([string]::IsNullOrWhiteSpace($scriptPath)) {
        throw "Cannot determine the script path. Supply -RepoRoot explicitly."
    }

    $scriptDirectory = Split-Path -Parent $scriptPath
    $RepoRoot = Split-Path -Parent $scriptDirectory
}

$vbext_ct_StdModule = 1
$vbext_ct_ClassModule = 2
$vbext_ct_MSForm = 3
$vbext_ct_Document = 100
$xlCalculationManual = -4135
$xlCalculationAutomatic = -4105
$msoAutomationSecurityForceDisable = 3

function Resolve-FullPath {
    param([Parameter(Mandatory = $true)][string]$PathValue)
    return [System.IO.Path]::GetFullPath((Resolve-Path -LiteralPath $PathValue).Path)
}

function Get-VbaComponentName {
    param([Parameter(Mandatory = $true)][string]$SourcePath)

    foreach ($line in [System.IO.File]::ReadLines($SourcePath)) {
        if ($line -match '^\s*Attribute\s+VB_Name\s*=\s*"([^"]+)"\s*$') {
            return $Matches[1]
        }
    }

    return [System.IO.Path]::GetFileNameWithoutExtension($SourcePath)
}

function Get-VbaCodeBody {
    param([Parameter(Mandatory = $true)][string]$SourcePath)

    $lines = [System.IO.File]::ReadAllLines($SourcePath)
    if ($lines.Count -eq 0) { return "" }

    $lastAttribute = -1
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match '^\s*Attribute\s+VB_') {
            $lastAttribute = $i
        }
    }

    if ($lastAttribute -ge 0) {
        if (($lastAttribute + 1) -ge $lines.Count) { return "" }
        return [string]::Join(
            [Environment]::NewLine,
            $lines[($lastAttribute + 1)..($lines.Count - 1)]
        )
    }

    return [string]::Join([Environment]::NewLine, $lines)
}

function Find-VbaComponent {
    param(
        [Parameter(Mandatory = $true)]$VBProject,
        [Parameter(Mandatory = $true)][string]$Name
    )

    foreach ($component in $VBProject.VBComponents) {
        if ([string]::Equals(
            [string]$component.Name,
            $Name,
            [System.StringComparison]::OrdinalIgnoreCase
        )) {
            return $component
        }
    }

    return $null
}

function Get-SheetSnapshot {
    param([Parameter(Mandatory = $true)]$Workbook)

    $items = New-Object System.Collections.Generic.List[string]
    foreach ($sheet in $Workbook.Sheets) {
        $codeName = ""
        try { $codeName = [string]$sheet.CodeName }
        catch { $codeName = "<no-codename>" }

        $items.Add("$($sheet.Name)|$codeName")
    }

    return @($items)
}

function Assert-SheetSnapshotUnchanged {
    param(
        [Parameter(Mandatory = $true)][string[]]$Before,
        [Parameter(Mandatory = $true)][string[]]$After
    )

    if ($Before.Count -ne $After.Count) {
        throw "Workbook sheet count changed. Save aborted."
    }

    for ($i = 0; $i -lt $Before.Count; $i++) {
        if ($Before[$i] -cne $After[$i]) {
            throw "Workbook sheet structure changed. Save aborted."
        }
    }
}

function Replace-DocumentCode {
    param(
        [Parameter(Mandatory = $true)]$Component,
        [Parameter(Mandatory = $true)][string]$SourcePath
    )

    if ([int]$Component.Type -ne $vbext_ct_Document) {
        throw "Component '$($Component.Name)' is not a document module."
    }

    $module = $Component.CodeModule
    $count = [int]$module.CountOfLines
    if ($count -gt 0) {
        $module.DeleteLines(1, $count)
    }

    $code = Get-VbaCodeBody -SourcePath $SourcePath
    if (-not [string]::IsNullOrWhiteSpace($code)) {
        $module.AddFromString($code)
    }
}

function Sync-VbaSource {
    param(
        [Parameter(Mandatory = $true)]$VBProject,
        [Parameter(Mandatory = $true)][string]$SourcePath
    )

    $extension = [System.IO.Path]::GetExtension($SourcePath).ToLowerInvariant()
    $name = Get-VbaComponentName -SourcePath $SourcePath
    $existing = Find-VbaComponent -VBProject $VBProject -Name $name

    if ($null -ne $existing -and [int]$existing.Type -eq $vbext_ct_Document) {
        if ($extension -ne ".cls") {
            throw "Document module '$name' must come from a .cls source."
        }

        Replace-DocumentCode -Component $existing -SourcePath $SourcePath
        Write-Host ("{0,-30} document code replaced" -f $name)
        return
    }

    if ($extension -eq ".cls" -and $null -eq $existing) {
        if ($name -eq "ThisWorkbook" -or $name -like "sh*") {
            throw "Workbook document module '$name' was not found. Save aborted."
        }
    }

    switch ($extension) {
        ".bas" {
            if ($null -ne $existing) {
                if ([int]$existing.Type -ne $vbext_ct_StdModule) {
                    throw "Existing '$name' is not a standard module."
                }
                $VBProject.VBComponents.Remove($existing)
            }

            $imported = $VBProject.VBComponents.Import($SourcePath)
            Write-Host ("{0,-30} standard module imported" -f $imported.Name)
        }

        ".cls" {
            if ($null -ne $existing) {
                if ([int]$existing.Type -ne $vbext_ct_ClassModule) {
                    throw "Existing '$name' is not a class module."
                }
                $VBProject.VBComponents.Remove($existing)
            }

            $imported = $VBProject.VBComponents.Import($SourcePath)
            Write-Host ("{0,-30} class module imported" -f $imported.Name)
        }

        ".frm" {
            if ($null -ne $existing) {
                if ([int]$existing.Type -ne $vbext_ct_MSForm) {
                    throw "Existing '$name' is not a UserForm."
                }
                $VBProject.VBComponents.Remove($existing)
            }

            $imported = $VBProject.VBComponents.Import($SourcePath)
            Write-Host ("{0,-30} UserForm imported" -f $imported.Name)
        }

        default {
            throw "Unsupported VBA source extension '$extension'."
        }
    }
}

$excel = $null
$workbook = $null
$saved = $false
$backupPath = $null

try {
    $workbookFullPath = Resolve-FullPath -PathValue $WorkbookPath
    $repoFullPath = Resolve-FullPath -PathValue $RepoRoot

    if ([System.IO.Path]::GetExtension($workbookFullPath).ToLowerInvariant() -ne ".xlsm") {
        throw "WorkbookPath must point to an .xlsm file."
    }

    $sourceRoots = @(
        (Join-Path $repoFullPath "src\core"),
        (Join-Path $repoFullPath "src\modules"),
        (Join-Path $repoFullPath "src\workbook"),
        (Join-Path $repoFullPath "tests\modules")
    )

    foreach ($root in $sourceRoots) {
        if (-not (Test-Path -LiteralPath $root -PathType Container)) {
            throw "Required source directory not found: $root"
        }
    }

    $sourceFiles = @(
        foreach ($root in $sourceRoots) {
            Get-ChildItem -LiteralPath $root -File -Recurse |
                Where-Object {
                    $_.Extension.ToLowerInvariant() -in @(".bas", ".cls", ".frm")
                }
        }
    ) | Sort-Object FullName

    if ($sourceFiles.Count -eq 0) {
        throw "No VBA source files found."
    }

    $seen = @{}
    foreach ($file in $sourceFiles) {
        $name = Get-VbaComponentName -SourcePath $file.FullName
        $key = $name.ToLowerInvariant()

        if ($seen.ContainsKey($key)) {
            throw "Duplicate VBA component name '$name' in repository source."
        }

        $seen[$key] = $file.FullName
    }

    if (-not $NoBackup) {
        $directory = [System.IO.Path]::GetDirectoryName($workbookFullPath)
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($workbookFullPath)
        $extension = [System.IO.Path]::GetExtension($workbookFullPath)
        $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $backupPath = Join-Path $directory "$baseName.pre-vba-sync.$timestamp$extension"

        Copy-Item -LiteralPath $workbookFullPath -Destination $backupPath -Force
        Write-Host "Backup: $backupPath"
    }

    Write-Host "Workbook:   $workbookFullPath"
    Write-Host "Repository: $repoFullPath"
    Write-Host "Sources:    $($sourceFiles.Count)"
    Write-Host ""

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = [bool]$Visible
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = $msoAutomationSecurityForceDisable

    # Some Excel builds reject Application.Calculation changes until at least
    # one workbook is open (HRESULT 0x800A03EC). Open first, then disable
    # calculation-before-save for this isolated Excel instance.
    $workbook = $excel.Workbooks.Open($workbookFullPath, 0, $false)

    if ($workbook.ReadOnly) {
        throw "Workbook opened read-only. Close any other Excel instance using it."
    }

    $excel.Calculation = $xlCalculationManual
    $excel.CalculateBeforeSave = $false

    $before = Get-SheetSnapshot -Workbook $workbook

    try {
        $vbProject = $workbook.VBProject
        $null = $vbProject.VBComponents.Count
    }
    catch {
        throw (
            "Cannot access the VBA project. Enable Excel Trust Center > Macro Settings > " +
            "'Trust access to the VBA project object model' and ensure the VBA project " +
            "is not password-locked."
        )
    }

    # Refuse retired repository components before the first project mutation.
    # Do not delete arbitrary user modules or try to migrate workbook structure.
    $retiredNames = @(
        "CoreScaffold", "SaccrScaffold", "M_Config", "M_Engine", "M_Util", "M_Formulas",
        "TestHarness", "TestMainState", "CaseRunner", "TestCases", "TestInputValidation"
    )
    foreach ($retiredName in $retiredNames) {
        if ($null -ne (Find-VbaComponent -VBProject $vbProject -Name $retiredName)) {
            throw "Retired repository component '$retiredName' found. Rebuild from the current template and source; no VBA changes have been made."
        }
    }

    foreach ($file in $sourceFiles) {
        Sync-VbaSource -VBProject $vbProject -SourcePath $file.FullName
    }

    $after = Get-SheetSnapshot -Workbook $workbook
    Assert-SheetSnapshotUnchanged -Before $before -After $after

    # The synchronization uses manual calculation only to avoid recalculation
    # while modules are being replaced. The development workbook must leave
    # this isolated Excel instance in Automatic mode before it is saved.
    $excel.Calculation = $xlCalculationAutomatic
    $excel.CalculateBeforeSave = $true

    $workbook.Save()
    $saved = $true

    # The backup is only a transactional safety copy. Remove it after a
    # successful save so repeated synchronizations do not accumulate files.
    if (-not [string]::IsNullOrWhiteSpace($backupPath) -and
        (Test-Path -LiteralPath $backupPath -PathType Leaf)) {
        Remove-Item -LiteralPath $backupPath -Force
        Write-Host "Backup removed: $backupPath"
    }

    Write-Host ""
    Write-Host "SUCCESS: VBA synchronized. Workbook sheets were not removed or recreated."
}
catch {
    Write-Error $_
    if ($null -ne $workbook -and -not $saved) {
        Write-Warning "Synchronization failed. Workbook will close without saving."
    }

    if (-not [string]::IsNullOrWhiteSpace($backupPath) -and
        (Test-Path -LiteralPath $backupPath -PathType Leaf)) {
        Write-Warning "Safety backup retained: $backupPath"
    }

    exit 1
}
finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch {}
        try {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
        } catch {}
    }

    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        try {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        } catch {}
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
