@echo off
setlocal EnableExtensions

cd /d "%~dp0"

where node >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Node.js is not installed or is not available in PATH.
    pause
    exit /b 1
)

where npm >nul 2>nul
if errorlevel 1 (
    echo [ERROR] npm is not available in PATH.
    pause
    exit /b 1
)

echo Running the Spreadsheet Phase 0 regression gate...
call npm run verify:spreadsheet:phase0
if errorlevel 1 (
    echo.
    echo [FAILED] Phase 0 verification found a regression.
    pause
    exit /b 1
)

echo.
echo [PASSED] All Spreadsheet Phase 0 checks passed.
pause
exit /b 0
