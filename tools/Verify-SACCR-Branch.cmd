@echo off
setlocal EnableExtensions
set "REPO_ROOT=%~dp0.."
cd /d "%REPO_ROOT%"

set "SCRIPT=%~dp0Verify-SACCR-Branch.ps1"

if not exist "%SCRIPT%" (
    echo ERROR: PowerShell script not found:
    echo   %SCRIPT%
    echo.
    pause
    exit /b 1
)

rem Checks out a branch of origin and certifies it with Certify-SACCR.ps1,
rem then prints a summary to report. The branch is the first argument; when
rem the file is double-clicked, it is asked for.
set "BRANCH=%~1"
if not defined BRANCH set /p "BRANCH=Branch to verify, as on GitHub: "
if not defined BRANCH (
    echo ERROR: no branch given.
    echo.
    pause
    exit /b 1
)

echo K-SACCR branch verification
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -Branch "%BRANCH%" -RepoRoot "%REPO_ROOT%"
set "CODE=%ERRORLEVEL%"
echo.
pause
exit /b %CODE%
