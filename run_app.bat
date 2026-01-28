@echo off
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"
title Email to Spreadsheet Framework
color 0A

cls
echo ==============================================
echo     EMAIL TO SPREADSHEET FRAMEWORK
echo ==============================================
echo.
echo Select interface:
echo   1) GUI (Graphical Interface) - Recommended
echo   2) Command-line Wizard
echo   3) Exit
echo.
set /p choice="Your choice [1-3]: "

if "%choice%"=="1" goto gui_mode
if "%choice%"=="2" goto wizard_mode
if "%choice%"=="3" goto end

:gui_mode
echo.
echo ----------------------------------------------
echo Launching GUI Interface...
echo ----------------------------------------------
python run_gui.py
if errorlevel 1 (
  echo.
  echo ERROR: GUI failed to launch.
  echo.
  echo Possible causes:
  echo   - Python not installed correctly
  echo   - tkinter not available
  echo.
  echo Falling back to command-line wizard...
  timeout /t 3 >nul
  goto wizard_mode
)
goto end

:wizard_mode
echo.
echo ----------------------------------------------
echo Launching Command-line Wizard...
echo ----------------------------------------------
python run_wizard.py
if errorlevel 1 (
  echo.
  echo ERROR: Wizard exited with an error.
)
goto end

:end
echo.
echo ----------------------------------------------
echo Done. Press any key to close.
echo ----------------------------------------------
pause >nul
endlocal
