@echo off
setlocal ENABLEDELAYEDEXPANSION
chcp 65001 >nul

cd /d "%~dp0"

echo ================================================
echo   Setup: EBITDA Dashboard Host (Windows)
echo ================================================
echo.

set "PY_CMD="
where py >nul 2>nul
if %errorlevel%==0 (
    set "PY_CMD=py -3.13"
) else (
    where python >nul 2>nul
    if %errorlevel%==0 set "PY_CMD=python"
)

if "%PY_CMD%"=="" (
    echo [ERROR] Python not found.
    echo Install Python 3.11+ from https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [1/4] Python check...
%PY_CMD% --version
if errorlevel 1 (
    echo [ERROR] Python launch failed.
    pause
    exit /b 1
)

echo [2/4] Create virtual environment (.venv)...
if not exist ".venv\Scripts\python.exe" (
    %PY_CMD% -m venv .venv
    if errorlevel 1 (
        echo [ERROR] venv creation failed.
        pause
        exit /b 1
    )
)

echo [3/4] Install dependencies...
".venv\Scripts\python.exe" -m pip install --upgrade pip
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
    echo [ERROR] Dependency installation failed.
    pause
    exit /b 1
)

echo [4/4] VS Code local settings...
if not exist ".vscode" mkdir ".vscode"
(
echo {
echo   "python.defaultInterpreterPath": ".venv\\Scripts\\python.exe",
echo   "python.terminal.activateEnvironment": true
echo }
) > ".vscode\settings.json"

echo.
echo [OK] Setup completed.
echo Next step: run_ebitda_dashboard.bat
echo.
pause
exit /b 0
