"""
COMP5 Reports Web App — Standalone Flask application
Serves:
  /            → reports home page
  /api/tq-sdr  → generate TQ & SDR Excel reports
  /api/comp5   → generate Weekly COMP5 Issued Documents report
  /download/<id>/<name> → download a generated file
"""

import os
import uuid
import threading
from flask import Flask, render_template, request, jsonify, send_file
from io import BytesIO

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB

# In-memory store for generated files (keyed by job_id)
_store: dict = {}
_lock = threading.Lock()


def _save(data: bytes, filename: str) -> str:
    job_id = str(uuid.uuid4())
    with _lock:
        _store[job_id] = {"data": data, "filename": filename}
    return job_id


# ── Pages ──────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("reports.html")


# ── API: TQ & SDR Report ───────────────────────────────────────────────────

@app.route("/api/tq-sdr", methods=["POST"])
def api_tq_sdr():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    from report_core import generate_tq_sdr
    try:
        raw = request.files["file"].read()
        result = generate_tq_sdr(raw)
        tqy_id = _save(result["tqy_bytes"], result["tqy_name"])
        sdr_id = _save(result["sdr_bytes"], result["sdr_name"])
        return jsonify({
            "tqy": {"job_id": tqy_id, "filename": result["tqy_name"]},
            "sdr": {"job_id": sdr_id, "filename": result["sdr_name"]},
            "summary": result["summary"],
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ── API: Weekly COMP5 Issued Documents Report ──────────────────────────────

@app.route("/api/comp5", methods=["POST"])
def api_comp5():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    from report_core import generate_comp5
    try:
        raw = request.files["file"].read()
        result = generate_comp5(raw)
        job_id = _save(result["bytes"], result["filename"])
        return jsonify({
            "job_id": job_id,
            "filename": result["filename"],
            "summary": result["summary"],
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ── Download ───────────────────────────────────────────────────────────────

@app.route("/download/<job_id>")
def download(job_id):
    with _lock:
        item = _store.get(job_id)
    if not item:
        return "File not found or expired", 404
    return send_file(
        BytesIO(item["data"]),
        download_name=item["filename"],
        as_attachment=True,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ── Run ────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
