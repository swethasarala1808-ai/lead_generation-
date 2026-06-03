from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import os, time, random, threading, json
from indiamart_api import IndiaMARTClient
from google_places_api import GooglePlacesClient
from excel_exporter import ExcelExporter

app = Flask(__name__)
CORS(app)

exporter = ExcelExporter()
jobs = {}   # in-memory job store

# ── helpers ──────────────────────────────────────────────────────────────────

def make_job_id():
    return f"job_{int(time.time())}_{random.randint(1000, 9999)}"

def run_in_thread(fn, *args):
    t = threading.Thread(target=fn, args=args, daemon=True)
    t.start()

# ── IndiaMART routes ──────────────────────────────────────────────────────────

@app.route("/api/indiamart/pull", methods=["POST"])
def indiamart_pull():
    """Pull fresh leads from IndiaMART using seller's API key."""
    body = request.json or {}
    api_key  = body.get("api_key", "").strip()
    start_dt = body.get("start", "")   # YYYYMMDD HH:MM:SS
    end_dt   = body.get("end", "")

    if not api_key:
        return jsonify({"error": "IndiaMART API key required"}), 400

    job_id = make_job_id()
    jobs[job_id] = {"status": "running", "progress": 0, "leads": [], "source": "indiamart"}

    def work():
        try:
            client = IndiaMARTClient(api_key)
            jobs[job_id]["progress"] = 20
            leads = client.pull_leads(start_dt, end_dt)
            jobs[job_id]["leads"]    = leads
            jobs[job_id]["progress"] = 100
            jobs[job_id]["status"]   = "done"
        except Exception as e:
            jobs[job_id]["status"]  = "error"
            jobs[job_id]["error"]   = str(e)

    run_in_thread(work)
    return jsonify({"job_id": job_id})


@app.route("/api/indiamart/history", methods=["POST"])
def indiamart_history():
    """Fetch up to 365 days of historical leads."""
    body    = request.json or {}
    api_key = body.get("api_key", "").strip()
    days    = min(int(body.get("days", 7)), 365)

    if not api_key:
        return jsonify({"error": "IndiaMART API key required"}), 400

    job_id = make_job_id()
    jobs[job_id] = {"status": "running", "progress": 0, "leads": [], "source": "indiamart"}

    def work():
        try:
            client = IndiaMARTClient(api_key)
            jobs[job_id]["progress"] = 20
            leads = client.pull_history(days)
            jobs[job_id]["leads"]    = leads
            jobs[job_id]["progress"] = 100
            jobs[job_id]["status"]   = "done"
        except Exception as e:
            jobs[job_id]["status"]  = "error"
            jobs[job_id]["error"]   = str(e)

    run_in_thread(work)
    return jsonify({"job_id": job_id})


# ── Google Places routes ──────────────────────────────────────────────────────

INDUSTRY_QUERIES = {
    "retail":       ["retail shop", "grocery store", "supermarket", "general store", "clothing store"],
    "medical":      ["clinic", "hospital", "pharmacy", "diagnostic centre", "nursing home"],
    "restaurant":   ["restaurant", "cafe", "dhaba", "catering", "cloud kitchen"],
    "it":           ["software company", "IT company", "digital agency", "web development"],
    "education":    ["coaching centre", "training institute", "school", "college"],
    "manufacturing":["manufacturer", "factory", "fabrication", "industries"],
    "real_estate":  ["real estate", "property dealer", "builder", "construction company"],
    "logistics":    ["logistics", "courier", "transport", "cargo"],
    "finance":      ["CA firm", "chartered accountant", "tax consultant", "insurance agent"],
    "beauty":       ["salon", "spa", "beauty parlour", "gym", "fitness centre"],
}

CITY_COORDS = {
    "bangalore":      "12.9716,77.5946", "mumbai":       "19.0760,72.8777",
    "delhi":          "28.6139,77.2090", "chennai":      "13.0827,80.2707",
    "hyderabad":      "17.3850,78.4867", "pune":         "18.5204,73.8567",
    "kolkata":        "22.5726,88.3639", "ahmedabad":    "23.0225,72.5714",
    "surat":          "21.1702,72.8311", "jaipur":       "26.9124,75.7873",
    "lucknow":        "26.8467,80.9462", "coimbatore":   "11.0168,76.9558",
    "kochi":           "9.9312,76.2673", "chandigarh":   "30.7333,76.7794",
    "indore":         "22.7196,75.8577", "nagpur":       "21.1458,79.0882",
    "visakhapatnam":  "17.6868,83.2185", "patna":        "25.5941,85.1376",
    "bhopal":         "23.2599,77.4126", "agra":         "27.1767,78.0081",
}

@app.route("/api/google/search", methods=["POST"])
def google_search():
    body     = request.json or {}
    api_key  = body.get("api_key", "").strip()
    industry = body.get("industry", "retail")
    city     = body.get("city", "bangalore").lower()
    count    = min(int(body.get("count", 20)), 50)

    if not api_key:
        return jsonify({"error": "Google Places API key required"}), 400

    latlng   = CITY_COORDS.get(city, "12.9716,77.5946")
    queries  = INDUSTRY_QUERIES.get(industry, [industry])
    city_cap = city.capitalize()

    job_id = make_job_id()
    jobs[job_id] = {"status": "running", "progress": 0, "leads": [], "source": "google"}

    def work():
        try:
            client = GooglePlacesClient(api_key)
            jobs[job_id]["progress"] = 10
            leads = client.search_leads(queries, city_cap, latlng, count, job_id, jobs)
            jobs[job_id]["leads"]    = leads
            jobs[job_id]["progress"] = 100
            jobs[job_id]["status"]   = "done"
        except Exception as e:
            jobs[job_id]["status"]  = "error"
            jobs[job_id]["error"]   = str(e)

    run_in_thread(work)
    return jsonify({"job_id": job_id})


# ── Status / Export ───────────────────────────────────────────────────────────

@app.route("/api/job/<job_id>", methods=["GET"])
def job_status(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404
    return jsonify({
        "status":   job["status"],
        "progress": job["progress"],
        "count":    len(job.get("leads", [])),
        "source":   job.get("source", ""),
        "error":    job.get("error", ""),
        "preview":  job.get("leads", [])[:3],
    })

@app.route("/api/job/<job_id>/leads", methods=["GET"])
def get_leads(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Not found"}), 404
    return jsonify({"leads": job.get("leads", []), "status": job["status"], "source": job.get("source","")})

@app.route("/api/job/<job_id>/export", methods=["GET"])
def export_leads(job_id):
    job = jobs.get(job_id)
    if not job or job["status"] != "done":
        return jsonify({"error": "Job not ready"}), 404
    path = exporter.export(job["leads"], job_id, job.get("source",""))
    return send_file(path, as_attachment=True, download_name=f"leads_{job_id}.xlsx")

@app.route("/api/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "time": time.strftime("%Y-%m-%d %H:%M:%S")})

if __name__ == "__main__":
    app.run(debug=True, port=5000, threaded=True)
