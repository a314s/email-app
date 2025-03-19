@echo off
echo ===================================================
echo Starting Email Automation App Server...
echo ===================================================
echo.

:: Check if .env file exists
if not exist .env (
    echo WARNING: .env file not found.
    echo.
    echo You need to create a .env file with your API keys.
    echo Example:
    echo   PORT=3000
    echo   ANTHROPIC_API_KEY=your_api_key_here
    echo.
    echo Would you like to create a basic .env file now? (Y/N)
    set /p CREATE_ENV=
    
    if /i "%CREATE_ENV%"=="Y" (
        echo PORT=3000 > .env
        echo ANTHROPIC_API_KEY=your_api_key_here >> .env
        echo.
        echo Basic .env file created. Please edit it with your actual API keys.
        echo.
        pause
    )
)

:: Check if node_modules exists
if not exist node_modules (
    echo WARNING: node_modules directory not found.
    echo You need to install dependencies first.
    echo.
    echo Would you like to run the installation now? (Y/N)
    set /p INSTALL=
    
    if /i "%INSTALL%"=="Y" (
        call install.bat
    ) else (
        echo.
        echo Please run install.bat before starting the server.
        pause
        exit /b 1
    )
)

:: Start the server
echo Starting server...
echo.
echo Press Ctrl+C to stop the server.
echo.
call npm start

pause
