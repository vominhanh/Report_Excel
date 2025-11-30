@echo off
echo ========================================
echo   CAI DAT DU AN PYTHON
echo ========================================
echo.

REM Kich hoat moi truong ao
echo [1/4] Kich hoat moi truong ao...
call .venv\Scripts\activate.bat

REM Dam bao pip co san
echo.
echo [2/4] Cai dat pip...
python -m ensurepip --upgrade

REM Nang cap pip
echo.
echo [3/4] Nang cap pip...
python -m pip install --upgrade pip

REM Cai dat cac thu vien
echo.
echo [4/4] Cai dat cac thu vien tu requirements.txt...
python -m pip install -r requirements.txt

echo.
echo ========================================
echo   HOAN THANH CAI DAT!
echo ========================================
echo.
echo Ban co the chay ung dung bang lenh:
echo   python app.py
echo.
pause


