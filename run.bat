@echo off
echo ========================================
echo   CHAY UNG DUNG PHAN TICH EXCEL
echo ========================================
echo.

REM Kich hoat moi truong ao
call .venv\Scripts\activate.bat

REM Chay ung dung Flask
echo Dang khoi dong server...
echo Mo trinh duyet va truy cap: http://localhost:5000
echo.
echo Nhan Ctrl+C de dung server
echo ========================================
echo.

python app.py


