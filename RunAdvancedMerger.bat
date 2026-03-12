@echo off
SETLOCAL EnableDelayedExpansion

TITLE Advanced Merger Launcher


:: Check for Node.js
where node >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo [INFO] Node.js is not installed. Attempting to install via winget...
    where winget >nul 2>nul
    if %ERRORLEVEL% NEQ 0 (
        echo [ERROR] winget not found. Please install Node.js manually.
        pause
        exit /b 1
    )
    winget install OpenJS.NodeJS.LTS --silent --accept-package-agreements --accept-source-agreements
    echo [SUCCESS] Node.js installed! Please RESTART this script.
    pause
    exit /b 0
)


echo [INFO] Starting Advanced Merger...
node server.js
pause
