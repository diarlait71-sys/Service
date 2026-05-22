@echo off
chcp 65001 >nul
echo.
echo ======================================================
echo   ЗАПУСК - Расчет бонусов отдела сервиса (Обновлено)
echo ======================================================
echo.

cd /d "%~dp0"

echo [*] Установка зависимостей...
"C:\Users\D.Muldabaev\AppData\Local\Programs\Python\Python313\python.exe" -m pip install -q streamlit pandas openpyxl xlrd python-dateutil numpy sqlalchemy pyodbc plotly PyYAML pydantic pytest

echo [*] Запуск приложения (обновленная версия с поддержкой механиков)...
echo.

"C:\Users\D.Muldabaev\AppData\Local\Programs\Python\Python313\python.exe" -m streamlit run app_new.py --server.port=8506

pause
