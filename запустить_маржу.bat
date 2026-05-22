@echo off
chcp 65001 > nul
cd /d "%~dp0"
echo.
echo ============================================================
echo   Ежедневный отчёт по марже продаж авто
echo ============================================================
echo.
echo  Сегодня: %date%
echo.
echo  Отчёт будет за СЕГОДНЯ.
echo  Чтобы задать другой период, закройте это окно и запустите:
echo     python daily_margin_report.py --from 2026-05-01 --to 2026-05-19
echo     python daily_margin_report.py --month 2026-05
set PYEXE=C:\Users\D.Muldabaev\AppData\Local\Programs\Python\Python313\python.exe
if not exist "%PYEXE%" set PYEXE=python
echo.
pause

"%PYEXE%" daily_margin_report.py

echo.
pause
