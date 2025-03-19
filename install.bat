@echo off
echo ===================================================
echo Installing dependencies for Email Automation App...
echo ===================================================
echo.

:: Check if npm is installed
where npm >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo ERROR: npm is not installed or not in your PATH.
    echo Please install Node.js from https://nodejs.org/
    echo.
    pause
    exit /b 1
)

:: Install dependencies
echo Installing npm packages...
call npm install

if %ERRORLEVEL% neq 0 (
    echo.
    echo ERROR: Failed to install dependencies.
    echo Please check your internet connection and try again.
    pause
    exit /b 1
) else (
    echo.
    echo ===================================================
    echo Dependencies installed successfully!
    echo.
    echo You can now run the application using:
    echo   start.bat
    echo ===================================================
    echo.
    pause
)
