# Скрипт запуска приложения (PowerShell)
# Правой клик -> Открыть с PowerShell

$PythonPath = "C:\Users\D.Muldabaev\AppData\Local\Programs\Python\Python313\python.exe"
$AppDir = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host ""
Write-Host "======================================================"
Write-Host "  ЗАПУСК - Расчет бонусов отдела сервиса"
Write-Host "======================================================"
Write-Host ""

Set-Location $AppDir

Write-Host "[*] Установка зависимостей..."
& $PythonPath -m pip install -q streamlit pandas openpyxl xlrd python-dateutil numpy sqlalchemy pyodbc 2>$null

Write-Host "[OK] Запуск приложения..."
Write-Host ""
Write-Host "    Откройте браузер: http://localhost:8501"
Write-Host "    Для выхода нажмите Ctrl+C"
Write-Host ""

& $PythonPath -m streamlit run app.py

Read-Host "Нажмите Enter для выхода"
