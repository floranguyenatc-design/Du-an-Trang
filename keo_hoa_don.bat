@echo off
setlocal
REM Buoc 3: keo hoa don. Nhap 1 thang (09/2026) hoac khoang ngay (01/01/2024-31/12/2024)
cd /d "%~dp0"
if not exist output mkdir output
set "KY="
set /p KY=Nhap thang (mm/yyyy) hoac khoang ngay (dd/mm/yyyy-dd/mm/yyyy): 
if "%KY%"=="" (
  echo Chua nhap ky can keo.
  pause
  exit /b 1
)
set "KY=%KY: =%"
set "TEN=%KY:/=_%"
set "TEN=%TEN:-=_den_%"
set "TU="
set "DEN="
for /f "tokens=1,2 delims=-" %%a in ("%KY%") do (
  set "TU=%%a"
  set "DEN=%%b"
)
if "%DEN%"=="" (
  echo Dang keo hoa don thang %KY% ...
  python -m hddt --log-file "output\log_%TEN%.txt" pull --thang %KY%
) else (
  echo Dang keo hoa don tu %TU% den %DEN% ...
  python -m hddt --log-file "output\log_%TEN%.txt" pull --tu-ngay %TU% --den-ngay %DEN%
)
echo.
echo Ket qua trong thu muc output. Neu co loi, gui file output\log_%TEN%.txt cho nguoi ho tro.
pause
