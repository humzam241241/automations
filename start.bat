@echo off
cls
color 0B
echo.
echo ============================================================
echo.
echo          EMAIL AUTOMATION PRO - UNIFIED LAUNCHER
echo.
echo ============================================================
echo.
echo  Choose how you want to run the application:
echo.
echo  [1] Web Application (Recommended)
echo      - Modern browser interface
echo      - Runs on localhost:3000
echo      - Multi-user support
echo      - File browser with visual navigation
echo.
echo  [2] CLI Wizard
echo      - Command-line interface
echo      - Step-by-step prompts
echo.
echo  [3] Exit
echo.
echo ============================================================
echo.
set /p choice="Enter your choice (1-3): "

if "%choice%"=="1" goto web
if "%choice%"=="2" goto cli
if "%choice%"=="3" goto end

echo Invalid choice. Please try again.
timeout /t 2 /nobreak >nul
goto start

:web
cls
echo.
echo ============================================================
echo.
echo          Starting Web Application...
echo.
echo          Your browser will open to:
echo          http://localhost:3000
echo.
echo          Press Ctrl+C to stop the server
echo          The window will stay open if the server exits.
echo.
echo ============================================================
echo.
timeout /t 2 /nobreak >nul
start http://localhost:3000
cmd /k python app.py
goto end

:cli
cls
echo.
echo ============================================================
echo.
echo          Starting CLI Wizard...
echo.
echo ============================================================
echo.
python run_wizard.py
goto end

:end
echo.
echo Session ended. Close this window when ready.
pause
