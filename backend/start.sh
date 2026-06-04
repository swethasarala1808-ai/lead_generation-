#!/bin/bash
# LeadGen Pro — Backend Starter
# Run this from the lead_generation- folder

echo "================================================"
echo "  LeadGen Pro Backend"
echo "================================================"

# Install deps
pip3 install flask flask-cors openpyxl requests --break-system-packages -q 2>/dev/null || \
pip install flask flask-cors openpyxl requests -q 2>/dev/null

echo "Starting server on http://localhost:5000 ..."
echo "Keep this terminal open."
echo "================================================"

# Try python3 first, then python
if command -v python3 &>/dev/null; then
    python3 app.py
else
    python app.py
fi
