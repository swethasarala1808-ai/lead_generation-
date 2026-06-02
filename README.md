# 🎯 LeadGen Pro — B2B Lead Generation for Startups & MSMEs

A quality lead generation tool for India B2B businesses — works like IndiaMART for generating verified contact leads.

## Features
- 🏭 **10+ Industries**: Retail, Medical, Restaurant, IT, Education, Manufacturing, Real Estate, Logistics, Finance, Beauty
- 📍 **Location-based**: Target any Indian city
- 🏅 **Business Types**: Filter by MSME Registered, Startup India, SME, Proprietorship
- 📊 **Excel Export**: Professional .xlsx with 23 fields per lead + Summary sheet
- 📅 **Daily Schedule**: Set auto-generate time for fresh leads daily
- 🔍 **Search & Filter**: Real-time search through leads
- 📱 **WhatsApp Ready**: Flags which leads have WhatsApp

## Lead Fields Included
| Field | Description |
|-------|-------------|
| Company Name | Business name |
| Business Type | Retail/Medical/IT etc |
| Registration Type | MSME / Startup India / SME |
| Owner Name | Contact person |
| Mobile Number | Primary phone |
| Alternate Mobile | Secondary phone |
| Email ID | Business / personal email |
| Website | Company URL |
| Address | Full address |
| City / State / Pincode | Location details |
| Annual Turnover | Revenue bracket |
| No. of Employees | Company size |
| Year Established | Founded year |
| GSTIN | Tax registration |
| WhatsApp Available | WhatsApp contact |
| Lead Quality Score | AI-rated quality % |
| Source | Data source platform |
| Remarks | Business status notes |

## Quick Start

### Option 1: Frontend Only (No Setup Needed)
Just open `frontend/index.html` in a browser — works standalone!

### Option 2: Full Stack with Backend
```bash
# Install dependencies
pip install -r requirements.txt

# Start backend
cd backend
python app.py

# Open frontend
open frontend/index.html
```

Backend runs on `http://localhost:5000`

## Project Structure
```
lead_generation/
├── frontend/
│   └── index.html          # Main app (works standalone)
├── backend/
│   ├── app.py              # Flask API server
│   ├── lead_engine.py      # Lead generation logic
│   └── excel_exporter.py   # Professional Excel export
├── exports/                # Generated Excel files
├── requirements.txt
└── README.md
```

## Target Users
- 🚀 Startups looking for B2B customers
- 🏢 MSMEs wanting to expand their customer base
- 📈 Sales teams needing verified leads
- 🤝 Business development professionals

## Industries Covered
Retail | Medical | Restaurant | IT/Tech | Education | Manufacturing | Real Estate | Logistics | Finance/CA | Beauty/Fitness

---
Built for India B2B market 🇮🇳
