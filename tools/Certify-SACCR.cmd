@echo off
setlocal EnableExtensions
set "REPO_ROOT=%~dp0.."
cd /d "%REPO_ROOT%"

set "SCRIPT=%~dp0Certify-SACCR.ps1"

if not exist "%SCRIPT%" (
    echo ERROR: PowerShell script not found:
    echo   %SCRIPT%
    echo.
    pause
    exit /b 1
)

rem Certifies the checked-out commit: builds a development workbook in a new
rem bundle folder next to the checkout, runs the harness and every suite,
rem writes host.json and validates the bundle. See docs\EXCEL_EVIDENCE.md.
echo K-SACCR certification
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" %*
set "CODE=%ERRORLEVEL%"
echo.
pause
exit /b %CODE%
