@echo off
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"
title Email to Spreadsheet Wizard
color 0A

cls
echo ==============================================
echo     EMAIL TO SPREADSHEET FRAMEWORK
echo ==============================================
echo.
echo This wizard will help you:
echo  - Create reusable email processing profiles
echo  - Fetch emails from Outlook OR local files
echo  - Apply keyword/regex rules to extract data
echo  - Output to Excel/CSV (local or OneDrive)
echo.
echo ----------------------------------------------
echo 1) Installing dependencies
echo ----------------------------------------------
python -m pip install -r requirements.txt >nul 2>&1
if errorlevel 1 (
  echo.
  echo ERROR: Dependency installation failed.
  echo Please check your Python installation or network access.
  pause
  exit /b 1
)
echo ✓ Dependencies installed

echo.
echo ----------------------------------------------
echo 2) Checking Graph API configuration
echo ----------------------------------------------
if not defined CLIENT_ID (
  echo.
  echo ⚠ CLIENT_ID environment variable not set
  echo   Graph mode will be unavailable.
  echo   You can still use local .eml files as input.
  echo.
) else (
  echo ✓ CLIENT_ID configured
)

if not defined AUTHORITY (
  echo.
  echo ⚠ AUTHORITY environment variable not set
  echo   Graph mode will be unavailable.
  echo.
) else (
  echo ✓ AUTHORITY configured
)

echo.
echo ----------------------------------------------
echo 3) Launching Interactive Wizard
echo ----------------------------------------------
echo.
python run_wizard.py
if errorlevel 1 (
  echo.
  echo ERROR: Wizard exited with an error.
)

echo.
echo ----------------------------------------------
echo Done. Press any key to close.
echo ----------------------------------------------
pause >nul
endlocal
