"""
Создаёт публичный туннель к дашборду через ngrok.
Нужна бесплатная регистрация на https://ngrok.com и токен.
"""
import sys
import time

try:
    from pyngrok import ngrok, conf
except ImportError:
    print("Устанавливаем pyngrok...")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pyngrok", "-q"])
    from pyngrok import ngrok, conf

# ── Токен ───────────────────────────────────────────────────────────────────
# Получите бесплатный токен на https://dashboard.ngrok.com/get-started/your-authtoken
# Вставьте его ниже:
NGROK_TOKEN = ""   # ← сюда вставить токен

PORT = 8510

if not NGROK_TOKEN:
    print()
    print("=" * 60)
    print("  Для публичного доступа нужен бесплатный токен ngrok")
    print()
    print("  1. Зайдите на https://ngrok.com/")
    print("  2. Зарегистрируйтесь (бесплатно, через Google/GitHub)")
    print("  3. Скопируйте токен с https://dashboard.ngrok.com/get-started/your-authtoken")
    print("  4. Вставьте токен в файл share_dashboard.py строку NGROK_TOKEN = '...'")
    print()
    print("  В офисной сети дашборд уже доступен по:")
    print(f"  http://10.2.5.141:{PORT}")
    print("=" * 60)
    input("\nНажмите Enter для выхода...")
    sys.exit(0)

try:
    conf.get_default().auth_token = NGROK_TOKEN
    tunnel = ngrok.connect(PORT, "http")
    public_url = tunnel.public_url

    print()
    print("=" * 60)
    print("  ✅ Дашборд доступен публично!")
    print()
    print(f"  Ссылка для отправки коллеге:")
    print(f"  {public_url}")
    print()
    print("  Ссылка работает пока это окно открыто.")
    print("  Закройте окно, чтобы остановить доступ.")
    print("=" * 60)
    print()

    # Держим туннель открытым
    try:
        while True:
            time.sleep(60)
    except KeyboardInterrupt:
        pass

    ngrok.disconnect(tunnel.public_url)
    print("Туннель закрыт.")

except Exception as e:
    print(f"Ошибка: {e}")
    print("Проверьте токен и подключение к интернету.")
    input("Нажмите Enter для выхода...")
