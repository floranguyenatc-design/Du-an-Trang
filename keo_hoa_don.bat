@echo off
REM Buoc 3: keo hoa don theo thang. Chay file nay, nhap thang dang mm/yyyy (vi du 09/2026)
cd /d "%~dp0"
set /p THANG=Nhap thang can keo (mm/yyyy): 
if "%THANG%"=="" (
  echo Chua nhap thang.
  pause
  exit /b 1
)
python -m hddt pull --thang %THANG% --log-file output\log_%THANG:/=_%.txt
echo.
echo Ket qua trong thu muc output. Neu co loi, gui file log trong output cho nguoi ho tro.
pause
