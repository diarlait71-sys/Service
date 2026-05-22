@echo off
cd /d "%~dp0"
python -m streamlit run ebitda_aktau.py --server.port 8517 --server.headless true
pause
