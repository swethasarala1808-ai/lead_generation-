@echo off
title bizaxl LeadGen
python app.py
if %errorlevel% neq 0 (
    echo.
    echo Error starting app. Run install_and_run.bat first.
    pause
)
