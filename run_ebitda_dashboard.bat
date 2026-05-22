@echo off
setlocal
chcp 65001 >nul

cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo [ERROR] Environment not found. Run setup_friend_host.bat first.
    pause
    exit /b 1
)

if not exist "ebitda_dashboard.py" (
    echo [ERROR] File ebitda_dashboard.py not found in current folder.
    pause
    exit /b 1
)

set "PORT=8517"
if not "%~1"=="" set "PORT=%~1"

echo ================================================
echo   Starting EBITDA Dashboard
echo ================================================
echo Local   : http://localhost:%PORT%
echo Network : http://YOUR_IP:%PORT%
echo Press Ctrl+C to stop.
echo.

".venv\Scripts\python.exe" -m streamlit run "ebitda_dashboard.py" --server.port %PORT% --server.headless true

echo.
echo Dashboard stopped.
pause
exit /b 0
