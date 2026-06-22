@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo [*] Kiem tra thu vien...
python -m pip install --quiet --user streamlit pywin32 pandas openpyxl
echo [*] Dang khoi dong app... (trinh duyet se tu mo)
python -m streamlit run app.py
pause
