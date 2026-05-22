@echo off
chcp 65001 > nul
echo ============================================
echo   DOSCAR GROUP — EBITDA Dashboard
echo ============================================
echo.
echo Запуск дашборда...
echo.

set PYEXE=C:\Users\D.Muldabaev\AppData\Local\Programs\Python\Python313\python.exe
set SCRIPT=%~dp0ebitda_dashboard.py
set SHARE=%~dp0share_dashboard.py

REM Запускаем дашборд в фоне
start "" "%PYEXE%" -m streamlit run "%SCRIPT%" --server.port 8510 --server.headless true

REM Ждём 3 секунды пока поднимется
timeout /t 3 /nobreak > nul

REM Открываем браузер локально
start http://localhost:8510

REM Запускаем туннель для внешнего доступа
"%PYEXE%" "%SHARE%"
pause
