#!/bin/bash
echo "🎯 Starting LeadGen Pro backend..."
echo "  → IndiaMART Pull API: http://localhost:5000/api/indiamart/pull"
echo "  → Google Places:      http://localhost:5000/api/google/search"
echo "  → Health:             http://localhost:5000/api/health"
echo ""
pip install -r ../requirements.txt -q --break-system-packages
python app.py
