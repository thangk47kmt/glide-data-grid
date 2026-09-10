@echo off
setlocal EnableExtensions

cd /d "%~dp0"
set "TEST_URL=http://localhost:9009/iframe.html?id=extra-packages-spreadsheet-large-dataset--large-dataset-story"

where node >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Node.js is not installed or is not available in PATH.
    echo Install Node.js, then run this file again.
    pause
    exit /b 1
)

where npm >nul 2>nul
if errorlevel 1 (
    echo [ERROR] npm is not available in PATH.
    pause
    exit /b 1
)

if not exist "node_modules\.bin\storybook.cmd" (
    echo Installing project dependencies...
    call npm install
    if errorlevel 1 (
        echo [ERROR] npm install failed.
        pause
        exit /b 1
    )
)

call :check_server
if errorlevel 1 (
    echo Starting Glide Data Grid Storybook on port 9009...
    start "Glide Data Grid Storybook" /D "%~dp0" cmd.exe /k "call npm run debug:spreadsheet"
) else (
    echo Storybook is already running on port 9009.
)

echo Waiting for the 500k spreadsheet story...
for /L %%I in (1,1,120) do (
    call :check_server
    if not errorlevel 1 goto ready
    timeout /t 1 /nobreak >nul
)

echo [ERROR] Storybook did not become ready within 120 seconds.
echo Check the "Glide Data Grid Storybook" window for errors.
pause
exit /b 1

:ready
echo Opening:
echo %TEST_URL%
start "" "%TEST_URL%"
exit /b 0

:check_server
powershell.exe -NoLogo -NoProfile -NonInteractive -Command ^
    "try { $response = Invoke-WebRequest -UseBasicParsing -Uri '%TEST_URL%' -TimeoutSec 2; if ($response.StatusCode -ge 200 -and $response.StatusCode -lt 500) { exit 0 } } catch {}; exit 1" >nul 2>nul
exit /b %errorlevel%
