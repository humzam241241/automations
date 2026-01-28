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
echo  [2] Desktop GUI
echo      - Native window interface
echo      - Traditional desktop app
echo.
echo  [3] CLI Wizard
echo      - Command-line interface
echo      - Step-by-step prompts
echo.
echo  [4] Exit
echo.
echo ============================================================
echo.
set /p choice="Enter your choice (1-4): "

if "%choice%"=="1" goto web
if "%choice%"=="2" goto desktop
if "%choice%"=="3" goto cli
if "%choice%"=="4" goto end

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
echo.
echo ============================================================
echo.
timeout /t 2 /nobreak >nul
start http://localhost:3000
python app.py
goto end

:desktop
cls
echo.
echo ============================================================
echo.
echo          Starting Desktop GUI...
echo.
echo ============================================================
echo.
python run_gui.py
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
echo Goodbye!
timeout /t 1 /nobreak >nul
