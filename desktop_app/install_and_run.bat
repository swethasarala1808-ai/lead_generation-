@echo off
title bizaxl LeadGen - Setup
color 0A
echo.
echo  ============================================
echo   bizaxl LeadGen - Installing...
echo  ============================================
echo.

:: Check Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo  Python not found. Installing Python...
    echo  Please download Python from: https://www.python.org/downloads/
    echo  During install: CHECK "Add Python to PATH"
    pause
    start https://www.python.org/downloads/
    exit
)

echo  [1/3] Python found. Installing dependencies...
pip install PyQt5 requests openpyxl --quiet --upgrade

echo  [2/3] Dependencies installed.
echo  [3/3] Launching bizaxl LeadGen...
echo.
python app.py
pause
