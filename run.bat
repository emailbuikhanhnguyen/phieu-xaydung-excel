@echo off
REM Mo cong cu phieu-xaydung-excel (giao dien Streamlit, chay cuc bo).
chcp 65001 >nul
cd /d "%~dp0"
if not exist .venv\Scripts\streamlit.exe (
  echo Chua cai dat. Chay setup.bat truoc.
  pause
  exit /b 1
)
call .venv\Scripts\activate.bat
set PYTHONPATH=%~dp0src
streamlit run app\streamlit_app.py
pause
