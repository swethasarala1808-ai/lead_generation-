"""
IndiaMART Pull API client (v2)
Docs: https://help.indiamart.com/knowledge-base/lms-crm-integration-v2

How it works:
  - Paid sellers on IndiaMART generate a Pull API key from:
      seller.indiamart.com -> Lead Manager -> ⋮ -> Import/Export -> Pull API
  - API URL:  https://mapi.indiamart.com/wservce/crm/crmListing/v2/
  - Method:   GET
  - Params:   glusr_crm_key, start_time, end_time
  - Window:   max 7 days per request
  - Output:   JSON list of leads
"""

import requests, time
from datetime import datetime, timedelta

INDIAMART_API_URL = "https://mapi.indiamart.com/wservce/crm/crmListing/v2/"

# Full field map from IndiaMART API v2 response
IM_FIELD_MAP = {
    "UNIQUE_QUERY_ID":      "Lead ID",
    "QUERY_TYPE":           "Lead Type",
    "QUERY_TIME":           "Received At",
    "SENDER_NAME":          "Buyer Name",
    "SENDER_MOBILE":        "Buyer Mobile",
    "SENDER_EMAIL":         "Buyer Email",
    "SENDER_COMPANY":       "Buyer Company",
    "SENDER_ADDRESS":       "Buyer Address",
    "SENDER_CITY":          "Buyer City",
    "SENDER_STATE":         "Buyer State",
    "SENDER_PINCODE":       "Buyer Pincode",
    "SENDER_COUNTRY_ISO":   "Country",
    "SENDER_MOBILE_ALT":    "Alternate Mobile",
    "SENDER_EMAIL_ALT":     "Alternate Email",
    "SUBJECT":              "Product / Requirement",
    "QUERY_MESSAGE":        "Buyer Message",
    "QUERY_PRODUCT_NAME":   "Product Name",
    "RECEIVER_GLID":        "Seller GLID",
    "RECEIVER_MOBILE":      "Seller Mobile",
    "CALL_DURATION":        "Call Duration",
    "RECEIVER_MOBILE2":     "Seller Mobile 2",
}

LEAD_TYPE_MAP = {
    "W": "Web Enquiry (Direct)",
    "B": "BuyLead (Consumed)",
    "P": "PNS Call",
    "C": "Catalog View Lead",
}

class IndiaMARTClient:
    def __init__(self, api_key: str):
        self.api_key = api_key

    def _fetch_window(self, start: str, end: str) -> list:
        """Fetch one window (max 7 days). start/end: 'YYYYMMDD HH:MM:SS'"""
        params = {
            "glusr_crm_key": self.api_key,
            "start_time":    start,
            "end_time":      end,
        }
        resp = requests.get(INDIAMART_API_URL, params=params, timeout=30)
        resp.raise_for_status()
        data = resp.json()

        # IndiaMART returns {"STATUS":200,"RESPONSE":[...]} or error codes
        if isinstance(data, dict):
            code = data.get("STATUS") or data.get("CODE")
            if code and int(code) != 200:
                msg = data.get("MESSAGE") or data.get("message") or str(data)
                raise ValueError(f"IndiaMART API error {code}: {msg}")
            rows = data.get("RESPONSE") or data.get("response") or []
        elif isinstance(data, list):
            rows = data
        else:
            rows = []

        return [self._map_lead(r) for r in rows]

    def _map_lead(self, row: dict) -> dict:
        lead = {}
        for src_key, label in IM_FIELD_MAP.items():
            lead[label] = row.get(src_key, "")
        # Friendly lead type
        lt = lead.get("Lead Type", "")
        lead["Lead Type"] = LEAD_TYPE_MAP.get(lt, lt)
        # Source
        lead["Source"] = "IndiaMART (Live)"
        return lead

    def pull_leads(self, start: str = "", end: str = "") -> list:
        """Pull leads between two timestamps (max 7-day window)."""
        if not start:
            start = (datetime.now() - timedelta(days=7)).strftime("%Y%m%d 00:00:00")
        if not end:
            end = datetime.now().strftime("%Y%m%d %H:%M:%S")
        return self._fetch_window(start, end)

    def pull_history(self, days: int = 30) -> list:
        """Pull up to 365 days by batching 7-day windows."""
        days  = min(days, 365)
        leads = []
        end   = datetime.now()
        # Walk backwards in 7-day chunks
        while days > 0:
            chunk = min(days, 7)
            start = end - timedelta(days=chunk)
            try:
                batch = self._fetch_window(
                    start.strftime("%Y%m%d 00:00:00"),
                    end.strftime("%Y%m%d 23:59:59"),
                )
                leads.extend(batch)
            except Exception:
                pass   # skip failed windows, continue
            end   = start
            days -= chunk
            time.sleep(0.3)   # polite rate limiting

        # De-duplicate by Lead ID
        seen = set()
        unique = []
        for l in leads:
            lid = l.get("Lead ID", "")
            if lid and lid not in seen:
                seen.add(lid)
                unique.append(l)
        return unique
