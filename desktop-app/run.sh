#!/bin/bash
echo "================================================"
echo "  bizaxl LeadGen - Starting..."
echo "================================================"

# Install dependencies silently
pip3 install PyQt5 requests openpyxl --break-system-packages -q 2>/dev/null

# Try to run with display
if [ -z "$DISPLAY" ]; then
    export DISPLAY=:0
fi

python3 app.py
