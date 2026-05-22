@echo off
chcp 65001 >nul
echo.
echo ======================================================
echo   ЗАПУСК - Расчет бонусов отдела сервиса (v3.0)
echo ======================================================
echo.
echo   Режимы:
echo   - 🆕 Бонусы отдела сервиса
echo   - 👨‍🔧 Механики (с поддержкой маппинга типов нарядов)
echo   - 📊 Остальные сотрудники
echo.

cd /d "%~dp0"

echo [*] Установка зависимостей...
"C:\Users\D.Muldabaev\AppData\Local\Programs\Python\Python313\python.exe" -m pip install -q streamlit pandas openpyxl xlrd python-dateutil numpy sqlalchemy pyodbc plotly PyYAML pydantic pytest

echo [*] Запуск приложения (обновленная версия с улучшенной поддержкой механиков)...
echo   URL: http://localhost:8506
echo.

"C:\Users\D.Muldabaev\AppData\Local\Programs\Python\Python313\python.exe" -m streamlit run app_new.py --server.port=8506

pause
