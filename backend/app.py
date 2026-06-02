from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import json, os, time, random, threading
from lead_engine import LeadEngine
from excel_exporter import ExcelExporter

app = Flask(__name__)
CORS(app)

engine = LeadEngine()
exporter = ExcelExporter()

# In-memory job store
jobs = {}

INDUSTRIES = {
    "retail": ["retail shop", "clothing store", "grocery store", "supermarket", "boutique", "general store"],
    "medical": ["clinic", "hospital", "pharmacy", "diagnostic centre", "doctor", "medical store", "nursing home"],
    "restaurant": ["restaurant", "hotel", "cafe", "dhaba", "food court", "catering", "cloud kitchen"],
    "it": ["software company", "IT company", "tech startup", "web development", "digital agency"],
    "education": ["school", "college", "coaching centre", "training institute", "tuition centre"],
    "manufacturing": ["manufacturer", "factory", "production unit", "industrial", "fabrication"],
    "real_estate": ["real estate", "property dealer", "builder", "construction company", "interior designer"],
    "logistics": ["logistics", "courier", "transport", "freight", "warehouse", "cargo"],
    "finance": ["CA firm", "chartered accountant", "tax consultant", "financial advisor", "insurance agent"],
    "beauty": ["salon", "spa", "beauty parlour", "wellness center", "gym", "fitness center"],
}

@app.route("/api/industries", methods=["GET"])
def get_industries():
    return jsonify({"industries": list(INDUSTRIES.keys())})

@app.route("/api/generate", methods=["POST"])
def generate_leads():
    data = request.json
    industry = data.get("industry", "retail")
    location = data.get("location", "India")
    count = min(int(data.get("count", 30)), 50)
    business_type = data.get("business_type", "all")  # startup, msme, all

    job_id = f"job_{int(time.time())}_{random.randint(1000,9999)}"
    jobs[job_id] = {"status": "running", "progress": 0, "leads": [], "total": count}

    def run_job():
        keywords = INDUSTRIES.get(industry, [industry])
        leads = engine.generate_leads(industry, keywords, location, count, business_type, job_id, jobs)
        jobs[job_id]["leads"] = leads
        jobs[job_id]["status"] = "done"
        jobs[job_id]["progress"] = 100

    t = threading.Thread(target=run_job)
    t.start()
    return jsonify({"job_id": job_id})

@app.route("/api/status/<job_id>", methods=["GET"])
def job_status(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404
    return jsonify({
        "status": job["status"],
        "progress": job["progress"],
        "count": len(job.get("leads", [])),
        "leads": job.get("leads", [])[:5]  # preview first 5
    })

@app.route("/api/export/<job_id>", methods=["GET"])
def export_leads(job_id):
    job = jobs.get(job_id)
    if not job or job["status"] != "done":
        return jsonify({"error": "Job not ready"}), 404
    leads = job["leads"]
    filepath = exporter.export(leads, job_id)
    return send_file(filepath, as_attachment=True, download_name=f"leads_{job_id}.xlsx")

@app.route("/api/leads/<job_id>", methods=["GET"])
def get_leads(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Not found"}), 404
    return jsonify({"leads": job.get("leads", []), "status": job["status"]})

if __name__ == "__main__":
    app.run(debug=True, port=5000)
