@echo off
chcp 65001 >nul
color 0A

echo ============================================================
echo  Загрузчик данных в Supabase
echo ============================================================
echo.

set PYTHON_EXE=C:\Users\D.Muldabaev\AppData\Local\Programs\Python\Python313\python.exe

if not exist "%PYTHON_EXE%" (
    echo ❌ Python не найден!
    pause
    exit /b 1
)

echo Запуск upload_to_supabase.py...
echo.

"%PYTHON_EXE%" upload_to_supabase.py

if %ERRORLEVEL% neq 0 (
    echo.
    echo ❌ Ошибка при выполнении скрипта!
    echo.
) else (
    echo.
    echo ✅ Готово!
    echo.
)

pause
