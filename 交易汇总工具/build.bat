@echo off
rem ============================================================
rem  Build 交易汇总工具.exe (Windows one-click build)
rem  Requirements: Python 3.9+ installed and "python" in PATH
rem ============================================================
cd /d "%~dp0"

echo [1/3] Creating virtual environment...
if not exist ".venv" (
    python -m venv .venv
    if errorlevel 1 (
        echo.
        echo ERROR: Failed to create venv. Is Python installed and in PATH?
        echo Download: https://www.python.org/downloads/windows/
        echo Remember to tick "Add python.exe to PATH" during install.
        pause
        exit /b 1
    )
)

call .venv\Scripts\activate.bat

echo [2/3] Installing dependencies (pyinstaller, openpyxl)...
python -m pip install --upgrade pip
pip install pyinstaller openpyxl
if errorlevel 1 (
    echo ERROR: pip install failed. Check network connection.
    pause
    exit /b 1
)

echo [3/3] Building exe...
python make_icon.py
pyinstaller --noconfirm --clean 交易汇总工具.spec
if errorlevel 1 (
    echo ERROR: PyInstaller build failed.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  DONE! Program is at:  dist\交易汇总工具.exe
echo  Copy that single exe to any Windows PC and double-click it.
echo ============================================================
pause
