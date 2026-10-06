@echo off
cd /d "%~dp0"
where pythonw >nul 2>&1
if %errorlevel%==0 (
  start "" pythonw "%~dp0keo_hoa_don.py"
) else (
  where python >nul 2>&1
  if %errorlevel%==0 (
    start "" python "%~dp0keo_hoa_don.py"
  ) else (
    echo Chua cai Python. Tai tai https://www.python.org/downloads/ va tick "Add Python to PATH" khi cai.
    pause
  )
)
