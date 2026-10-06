@echo off
REM Buoc 1: cai thu vien can thiet (chi can chay 1 lan). Can cai Python truoc: https://www.python.org/downloads/
cd /d "%~dp0"
python --version >nul 2>&1
if errorlevel 1 (
  echo Chua cai Python. Vao https://www.python.org/downloads/ tai ve, khi cai nho tick "Add Python to PATH".
  pause
  exit /b 1
)
python -m pip install -r requirements.txt
if not exist .env copy .env.example .env
echo.
echo Da cai xong. Mo file .env bang Notepad, dien GDT_USERNAME va GDT_PASSWORD, roi chay dang_nhap.bat de kiem tra.
pause
