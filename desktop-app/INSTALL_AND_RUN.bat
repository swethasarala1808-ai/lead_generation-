@echo off
echo ================================================
echo  bizaxl LeadGen - Setup
echo ================================================
echo.

REM Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found!
    echo Please install Python from https://python.org
    echo Make sure to check "Add Python to PATH" during install
    pause
    exit /b 1
)

echo Python found. Installing dependencies...
pip install PyQt5 requests openpyxl --quiet

echo.
echo ================================================
echo  Starting bizaxl LeadGen...
echo ================================================
python app.py
pause
