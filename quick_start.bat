@echo off
chcp 65001 >nul

echo.
echo ======================================================
echo   ЗАПУСК - Расчет бонусов отдела сервиса
echo ======================================================
echo.

cd /d "%~dp0"

echo [*] Установка зависимостей...
"C:\Users\D.Muldabaev\AppData\Local\Programs\Python\Python313\python.exe" -m pip install -q streamlit pandas openpyxl xlrd python-dateutil numpy sqlalchemy pyodbc

echo [*] Запуск приложения...
echo.

"C:\Users\D.Muldabaev\AppData\Local\Programs\Python\Python313\python.exe" -m streamlit run app.py

pause
