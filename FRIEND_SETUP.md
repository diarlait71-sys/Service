# Quick Setup For Your Friend (Windows + VS Code)

1. Install VS Code:
- https://code.visualstudio.com/

2. Install Python 3.13 (or 3.11+):
- https://www.python.org/downloads/
- During installation, enable: Add Python to PATH

3. Copy this full project folder to friend PC.
- Keep folder structure as-is, including:
  - ebitda_dashboard.py
  - requirements.txt
  - .streamlit/config.toml
  - folder "финанализ" with xls files

4. Open the folder in VS Code.

5. Run first-time setup:
- Double-click setup_friend_host.bat

6. Run dashboard:
- Double-click run_ebitda_dashboard.bat
- Optional custom port: run_ebitda_dashboard.bat 8518

7. Open in browser:
- Local: http://localhost:8517
- LAN: http://<friend_local_ip>:8517

## If friend sees package errors
Run setup_friend_host.bat again.

## If port is busy
Run: run_ebitda_dashboard.bat 8520

## Share to internet (optional)
Use ngrok/cloudflared separately if external public link is needed.
