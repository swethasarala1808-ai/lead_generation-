# 🎯 LeadGen Pro — Daily B2B Lead Generator

Auto-generates fresh B2B leads every day using GitHub Actions + OpenStreetMap.

## 🌐 Live Site
After setup: `https://swethasarala1808-ai.github.io/lead_generation-/`

## ⚡ Quick Setup (5 minutes)

### Step 1 — Clone and push
```bash
git clone https://github.com/swethasarala1808-ai/lead_generation-.git
cd lead_generation-
git pull origin main
```

### Step 2 — Enable GitHub Pages
1. Go to your repo on GitHub → **Settings**
2. Click **Pages** in the left sidebar
3. Under **Source**: select `Deploy from a branch`
4. Branch: `main` | Folder: `/docs`
5. Click **Save**
6. Wait 2 minutes → site is live at `https://swethasarala1808-ai.github.io/lead_generation-/`

### Step 3 — Run first lead generation
1. Go to **Actions** tab in your repo
2. Click **Daily Lead Generation**
3. Click **Run workflow** → choose industry & city → Run
4. Wait 2-3 minutes → refresh your site → leads appear!

### Step 4 — Daily auto-run (already configured)
GitHub Actions runs automatically every day at **7:30 AM IST**. No setup needed.

---

## 📊 What you get every day
- Real business names, phones, addresses from OpenStreetMap
- Excel + CSV download from the site
- 7+ days of history to browse
- Filters: by phone, website, area, rating
- Click-to-call, WhatsApp, Google Maps per lead

## 🔧 Customize

### Change industry / city
Edit `.github/workflows/daily_leads.yml`:
```yaml
env:
  INDUSTRY: medical      # retail/medical/it/restaurant/education/etc
  CITY: Mumbai           # any Indian city
  COUNT: 50
```

### Add Google Places for more leads
1. Get free key: https://console.cloud.google.com → Enable Places API
2. Go to repo → **Settings → Secrets → Actions → New secret**
3. Name: `GOOGLE_API_KEY`, Value: your key
4. Run workflow again — now gets Google data too

---

## 📁 Structure
```
lead_generation-/
├── .github/workflows/
│   └── daily_leads.yml     # Auto-runs every day at 7:30 AM IST
├── backend/
│   └── generate_daily.py   # Lead generation script
├── data/leads/
│   └── YYYY-MM-DD.json     # Daily lead files
├── docs/                   # GitHub Pages site
│   ├── index.html          # The web app
│   └── data/
│       ├── latest.json     # Today's leads
│       └── history.json    # All past runs index
└── README.md
```

---
Built for bizaxl sales team 🚀
