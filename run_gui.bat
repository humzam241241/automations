@echo off
REM Quick launcher for GUI interface
cd /d "%~dp0"
title Email Automation Pro - GUI
python run_gui.py
if errorlevel 1 (
  echo.
  echo ERROR: Failed to launch GUI
  echo.
  echo Possible solutions:
  echo   1. Make sure Python is installed
  echo   2. Run: python -m pip install -r requirements.txt
  echo   3. Use run_wizard.py for command-line interface
  echo.
  pause
)
