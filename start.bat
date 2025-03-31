@echo off
setlocal enabledelayedexpansion
echo ===================================================
echo Starting Email Automation App Server...
echo ===================================================
echo.


:: Check if node.js is installed
where node >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo Node.js is not installed. Please install Node.js from https://nodejs.org/
    pause
    exit /b 1
)

:: Check if npm is installed
where npm >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo npm is not installed. Please install Node.js from https://nodejs.org/
    pause
    exit /b 1
)

:: Check if package.json exists
if not exist "package.json" (
    echo package.json not found. Please run install.bat first.
    pause
    exit /b 1
)

:: Check if node_modules exists
if not exist "node_modules" (
    echo node_modules not found. Running npm install...
    npm install
    if %ERRORLEVEL% neq 0 (
        echo Failed to install dependencies. Please run install.bat.
        pause
        exit /b 1
    )
)

:: Create data directory if it doesn't exist
if not exist "data" (
    mkdir data
    echo Created data directory
)

:: Kill any process using port 3000
echo Checking for processes using port 3000...
for /f "tokens=5" %%a in ('netstat -aon ^| find ":3000" ^| find "LISTENING"') do (
    echo Found process using port 3000. Attempting to kill process %%a...
    taskkill /F /PID %%a >nul 2>nul
    if !ERRORLEVEL! equ 0 (
        echo Successfully killed process %%a
    ) else (
        echo No process found using port 3000
    )
)
timeout /t 2 >nul

:: Start the server with debugging
echo Starting server with debugging enabled...
cd /d "%~dp0"
node --trace-warnings server.js

:: If we get here, the server has stopped
echo Server has stopped. Press any key to exit.
pause
