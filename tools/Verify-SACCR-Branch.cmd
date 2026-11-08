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
rem or runs the quick pull-request check when the second argument is quick,
rem then prints a summary to report. The branch is the first argument; when
rem the file is double-clicked, the branch and the mode are asked for.
set "BRANCH=%~1"
set "MODE=%~2"
set "QUICK="
set "KNOWN="
if not defined BRANCH (
    set /p "BRANCH=Branch to verify, as on GitHub: "
    set /p "MODE=Quick run for a pull request, without prompts (Y/N)? "
)
if not defined BRANCH (
    echo ERROR: no branch given.
    echo.
    pause
    exit /b 1
)
if not defined MODE set "MODE=full"
if /i "%MODE%"=="quick" set "QUICK=-Quick"
if /i "%MODE%"=="-quick" set "QUICK=-Quick"
if /i "%MODE%"=="/quick" set "QUICK=-Quick"
if /i "%MODE%"=="y" set "QUICK=-Quick"
if /i "%MODE%"=="yes" set "QUICK=-Quick"
if defined QUICK set "KNOWN=1"
if /i "%MODE%"=="full" set "KNOWN=1"
if /i "%MODE%"=="n" set "KNOWN=1"
if /i "%MODE%"=="no" set "KNOWN=1"
if not defined KNOWN (
    echo ERROR: the second argument must be quick, or left out for the full certification: %MODE%
    echo.
    pause
    exit /b 1
)

echo K-SACCR branch verification
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -Branch "%BRANCH%" -RepoRoot "%REPO_ROOT%" %QUICK%
set "CODE=%ERRORLEVEL%"
echo.
pause
exit /b %CODE%
