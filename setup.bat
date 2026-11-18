@echo off
REM Cai dat lan dau tren Windows: tao moi truong ao va cai thu vien.
REM Chay 1 lan sau khi clone repo. Can co Python 3.11+ va mang.
chcp 65001 >nul
cd /d "%~dp0"
where python >nul 2>nul || (echo Khong tim thay Python. Cai Python 3.11+ tai https://www.python.org/downloads/ roi chay lai. & pause & exit /b 1)
python -m venv .venv
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt
echo.
echo Cai dat xong. Chay run.bat de mo cong cu.
pause
