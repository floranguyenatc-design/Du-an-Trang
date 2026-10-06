@echo off
REM Chay nhanh tren Windows, vi du:
REM   chay.bat --thang 12/2025
REM   chay.bat --chieu mua --tu-ngay 01/12/2025 --den-ngay 31/12/2025
cd /d "%~dp0"
if "%~1"=="" (
  python -m hddt pull --help
) else (
  python -m hddt pull %*
)
pause
