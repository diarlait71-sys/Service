@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ================================================
echo   EBITDA -> Cash | AKTAU  | port 8517
echo ================================================
echo http://localhost:8517
echo Ctrl+C - остановить
echo.

"C:\Users\D.Muldabaev\AppData\Local\Programs\Python\Python313\python.exe" ebitda_full_aktau.py
pause
