@echo off
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"
title Email Monitor - Outlook to OneDrive
color 0A

cls
echo ==============================================
echo        EMAIL MONITOR - OUTLOOK TO ONEDRIVE
echo ==============================================
echo.
echo This tool will:
echo  - Ask for folder/path/keywords
echo  - Scan the Outlook folder
echo  - Upload CSV reports to OneDrive
echo.
echo ----------------------------------------------
echo 1) Installing dependencies
echo ----------------------------------------------
python -m pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo ERROR: Dependency installation failed.
  echo Please check your Python installation or network access.
  pause
  exit /b 1
)

echo.
echo ----------------------------------------------
echo 2) Configuration (folders, keywords)
echo ----------------------------------------------
set "ONEDRIVE_FOLDER="
set "TEMPLATE_PATH="

:promptOnedrive
set "input="
set /p input=OneDrive folder path (e.g. /EmailReports) [leave blank to keep]: 
if "!input!"=="" goto promptTemplate
if not "!input:~0,1!"=="/" (
  echo Invalid OneDrive path. Use format like /EmailReports or /Folder/SubFolder.
  goto promptOnedrive
)
set "ONEDRIVE_FOLDER=!input!"

:promptTemplate
set "input="
set /p input=Template Excel path (optional .xlsx) [leave blank to skip]: 
if "!input!"=="" goto runSetup
if not exist "!input!" (
  echo Invalid path. File not found.
  goto promptTemplate
)
for %%F in ("!input!") do set "ext=%%~xF"
if /I not "!ext!"==".xlsx" (
  echo Invalid file type. Please select a .xlsx file.
  goto promptTemplate
)
set "TEMPLATE_PATH=!input!"

:runSetup
python setup_keywords.py
if errorlevel 1 (
  echo.
  echo ERROR: Configuration failed.
  pause
  exit /b 1
)

echo.
echo ----------------------------------------------
echo 3) Running email processor
echo ----------------------------------------------
python main.py
if errorlevel 1 (
  echo.
  echo ERROR: App exited with an error.
)

echo.
echo ----------------------------------------------
echo Done. Press any key to close.
echo ----------------------------------------------
pause >nul
endlocal
