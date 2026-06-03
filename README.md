# 🎯 LeadGen Pro — IndiaMART + Google Places Lead Generator

Works exactly like IndiaMART Lead Manager — pull real buyer enquiries + search real businesses.

## Two Data Sources

### 🟠 IndiaMART Pull API (Real Buyer Enquiries)
Pulls actual enquiries sent to YOUR seller account on IndiaMART.

**What you get:**
- Buyer name, mobile, email, company
- City, state, pincode
- Product / requirement they asked for
- Buyer message
- Lead type: Web Enquiry, BuyLead, PNS Call
- Date/time received

**How to get your API key:**
1. Login at https://seller.indiamart.com
2. Lead Manager → ⋮ → Import/Export → Pull API
3. Click "Generate Key"
4. Requires: **Paid IndiaMART seller account** (Leader or Star plan)

**API Endpoint used:** `https://mapi.indiamart.com/wservce/crm/crmListing/v2/`

---

### 🔵 Google Places API (Real Business Search)
Finds real registered businesses in any Indian city.

**What you get:**
- Real business name, phone, address
- Website URL
- Google rating + review count
- Opening hours
- Direct Google Maps link

**How to get your API key:**
1. Go to https://console.cloud.google.com
2. Enable "Places API"
3. Credentials → Create API Key
4. $200 free credit/month (~28,500 searches)

---

## Quick Start

### Frontend Only (no backend needed for Google)
Open `frontend/index.html` in browser.
- Google tab works directly from browser
- IndiaMART tab needs the backend running

### With Backend (recommended for IndiaMART)
```bash
cd backend
pip install -r ../requirements.txt
python app.py
# Opens at http://localhost:5000
```
Then open `frontend/index.html`

---

## Project Structure
```
lead_generation/
├── frontend/
│   └── index.html              # Main app UI
├── backend/
│   ├── app.py                  # Flask API server
│   ├── indiamart_api.py        # IndiaMART Pull API v2 client
│   ├── google_places_api.py    # Google Places client
│   └── excel_exporter.py       # Professional Excel export
├── exports/                    # Generated Excel files
├── requirements.txt
└── README.md
```

## API Endpoints
| Method | Endpoint | Description |
|--------|----------|-------------|
| POST | `/api/indiamart/pull` | Pull leads (date range) |
| POST | `/api/indiamart/history` | Pull historical leads (up to 365 days) |
| POST | `/api/google/search` | Search businesses by industry + city |
| GET  | `/api/job/<id>` | Check job status + progress |
| GET  | `/api/job/<id>/leads` | Get all leads from job |
| GET  | `/api/job/<id>/export` | Download Excel file |
| GET  | `/api/health` | Health check |

---
Built for India B2B Sales Teams 🇮🇳
