@echo off
setlocal enabledelayedexpansion

title DIGII Monday Bot - Administrative Automation Tool
color 0F

echo ===============================================================================
echo                DIGII ERP AUTOMATION - MONDAY BOT RUNNER
echo ===============================================================================
echo [INFO] Current Timestamp: %date% %time%
echo [INFO] Working Directory: %~dp0
echo.

cd /d "%~dp0"

:: Step 1: Verify Python Availability
echo [STEP 1/4] Checking Python environment...
python --version >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Python is not installed or not found in system PATH.
    echo Please install Python 3.10+ from https://python.org and ensure "Add Python to PATH" is checked.
    goto :FAIL
)
for /f "tokens=*" %%i in ('python --version') do set PYTHON_VER=%%i
echo [OK] Detected %PYTHON_VER%

:: Step 2: Check Configuration & Environment File
echo.
echo [STEP 2/4] Verifying configuration and credentials...
if not exist ".env" (
    if exist ".env.example" (
        echo [WARNING] .env file not found. Creating from .env.example template...
        copy ".env.example" ".env" >nul
        echo [NOTE] Created .env template. Please update with your DIGII ERP credentials.
    ) else (
        echo [ERROR] Neither .env nor .env.example found.
        goto :FAIL
    )
)
if not exist "config\urls_config.xlsx" (
    echo [ERROR] Configuration file 'config\urls_config.xlsx' is missing.
    goto :FAIL
)
echo [OK] Found .env and config\urls_config.xlsx

:: Step 3: Check and Install Required Dependencies
echo.
echo [STEP 3/4] Verifying required Python packages...
python -c "import playwright, pandas, openpyxl, dotenv" >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo [INFO] Missing required packages. Installing from requirements.txt...
    python -m pip install -r requirements.txt
    if %ERRORLEVEL% neq 0 (
        echo [ERROR] Failed to install dependencies.
        goto :FAIL
    )
    echo [INFO] Installing Playwright Chromium browser...
    python -m playwright install chromium
)
echo [OK] All dependencies satisfied.

:: Step 4: Execute Automation Bot
echo.
echo [STEP 4/4] Launching Digii_Monday_Bot...
echo -------------------------------------------------------------------------------
python digii_bot.py %*
set BOT_EXIT_CODE=%ERRORLEVEL%
echo -------------------------------------------------------------------------------

if %BOT_EXIT_CODE% equ 0 (
    echo.
    echo ===============================================================================
    echo [SUCCESS] DIGII Monday Bot completed run successfully!
    echo The Executive Audit Report has been generated and launched.
    echo ===============================================================================
) else (
    echo.
    echo ===============================================================================
    echo [WARNING] Bot finished with code %BOT_EXIT_CODE%. Check Audit Report for details.
    echo ===============================================================================
)

echo.
echo Press any key to close this console window...
pause >nul
exit /b %BOT_EXIT_CODE%

:FAIL
echo.
echo ===============================================================================
echo [FAILED] Execution aborted due to pre-flight setup errors.
echo ===============================================================================
echo.
echo Press any key to close this window...
pause >nul
exit /b 1
