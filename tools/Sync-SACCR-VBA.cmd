@echo off
setlocal EnableExtensions
set "REPO_ROOT=%~dp0.."
cd /d "%REPO_ROOT%"

set "SCRIPT=%~dp0Sync-SACCR-VBA.ps1"

if not exist "%SCRIPT%" (
    echo ERROR: PowerShell script not found:
    echo   %SCRIPT%
    echo.
    pause
    exit /b 1
)

rem Without a dropped workbook, a fresh development workbook is built from
rem the template in the ignored test-results folder, replaced on every run.
rem A dropped workbook is synchronized as it is and never rebuilt.
set "WORKBOOK=%~1"
set "BUILD="
if not defined WORKBOOK (
    set "WORKBOOK=%CD%\test-results\K-SACCR_Development.xlsm"
    set "BUILD=-Mode Development -FromTemplate"
)

echo SACCR VBA synchronization
echo.

if not defined BUILD if not exist "%WORKBOOK%" (
    echo.
    echo ERROR: Workbook not found:
    echo   %WORKBOOK%
    echo.
    pause
    exit /b 1
)

echo.
echo Repository:
echo   %CD%
echo Workbook:
echo   %WORKBOOK%
echo.
if defined BUILD (
    echo Building a development workbook from src\workbook\SACCR_Template.xlsx
    echo and synchronizing it, sheet code included...
) else (
    echo Using the dropped workbook as it is: its sheets are not rebuilt.
    echo Synchronizing VBA source...
)
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -WorkbookPath "%WORKBOOK%" -RepoRoot "%REPO_ROOT%" %BUILD%
set "RC=%ERRORLEVEL%"

echo.
if "%RC%"=="0" (
    echo SUCCESS: SACCR VBA synchronization completed.
    if defined BUILD (
        echo Open the workbook, run Debug - Compile VBAProject, then the tests:
        echo   %WORKBOOK%
    )
) else if "%RC%"=="2" (
    echo NOT CHANGED: nothing was imported; the workbook was not changed.
) else (
    echo FAILED: SACCR VBA synchronization returned exit code %RC%.
)
echo.
pause
exit /b %RC%
