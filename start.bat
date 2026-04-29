@echo off
title Tool Allocation Dashboard
echo.
echo =========================================
echo   Tool Allocation Dashboard — Intel
echo =========================================
echo.

:: Check for Node.js
where node >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] Node.js is not installed or not on PATH.
    echo Please install Node.js from https://nodejs.org and re-run this script.
    pause
    exit /b 1
)

cd /d "%~dp0"

:: Install dependencies if node_modules is missing
if not exist "node_modules\" (
    echo [INFO] Installing dependencies (first run only)...
    npm install
    if %errorlevel% neq 0 (
        echo [ERROR] npm install failed. Check your internet connection.
        pause
        exit /b 1
    )
)

echo [INFO] Starting server on http://localhost:3000 ...
echo [INFO] Close this window to stop the dashboard.
echo.

:: Open browser after a short delay (start is non-blocking)
start "" /b cmd /c "timeout /t 2 >nul && start http://localhost:3000"

node server.js
pause
