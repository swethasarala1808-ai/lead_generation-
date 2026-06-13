@echo off
echo Building bizaxl LeadGen .exe...
pip install pyinstaller --quiet
pyinstaller --onefile --windowed --name="bizaxl_LeadGen" --icon=icon.ico app.py
echo.
echo Done! Find bizaxl_LeadGen.exe in the dist/ folder
pause
