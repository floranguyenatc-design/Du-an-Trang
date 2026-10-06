@echo off
REM Buoc 2: kiem tra dang nhap trang thue (tool tu giai CAPTCHA)
cd /d "%~dp0"
python -m hddt login
echo.
echo Neu bao loi CAPTCHA, chay them: python -m hddt captcha --so-lan 5 --luu output\captcha
pause
