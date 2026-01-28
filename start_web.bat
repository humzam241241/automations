@echo off
echo.
echo ╔═══════════════════════════════════════════════════════╗
echo ║                                                       ║
echo ║        📧 Email Automation Pro - Web App 🚀          ║
echo ║                                                       ║
echo ║        Starting web server on localhost:3000...      ║
echo ║                                                       ║
echo ║        Your browser will open automatically          ║
echo ║        Or navigate to: http://localhost:3000         ║
echo ║                                                       ║
echo ║        Press Ctrl+C to stop the server              ║
echo ║                                                       ║
echo ╚═══════════════════════════════════════════════════════╝
echo.
timeout /t 2 /nobreak >nul
start http://localhost:3000
python app.py
pause
