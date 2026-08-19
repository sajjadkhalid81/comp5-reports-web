"""
COMP5 Reports Web App — app.py
Existing routes (DO NOT TOUCH):
  GET  /                        → reports page
  POST /api/tq-sdr/summary      → TQ & SDR summary JSON
  POST /api/tq-sdr/tqy          → download TQY Excel
  POST /api/tq-sdr/sdr          → download SDR Excel
  POST /api/comp5/summary       → COMP5 Issued Docs summary JSON
  POST /api/comp5/download      → download COMP5 Issued Docs Excel

New routes added at bottom (Report 3 — MR & TBE):
  POST /api/mr/summary          → MR summary JSON
  POST /api/mr/download         → download MR Excel
  POST /api/tbe/summary         → TBE summary JSON
  POST /api/tbe/download        → download TBE Excel
"""

import os
from io import BytesIO
from flask import Flask, render_template, request, jsonify, send_file

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


# ── Page ───────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("reports.html")


# ══════════════════════════════════════════════════════════════════════════
# EXISTING — Report 1: TQ & SDR  (DO NOT TOUCH)
# ══════════════════════════════════════════════════════════════════════════

@app.route("/api/tq-sdr/summary", methods=["POST"])
def api_tq_sdr_summary():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    try:
        from report_core import generate_tq_sdr
        result = generate_tq_sdr(request.files["file"].read())
        return jsonify({"summary": result["summary"]})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/tq-sdr/tqy", methods=["POST"])
def api_download_tqy():
    if "file" not in request.files:
        return "No file uploaded", 400
    try:
        from report_core import generate_tq_sdr
        result = generate_tq_sdr(request.files["file"].read())
        return send_file(BytesIO(result["tqy_bytes"]),
                         download_name=result["tqy_name"],
                         as_attachment=True, mimetype=XLSX_MIME)
    except Exception as e:
        return str(e), 500


@app.route("/api/tq-sdr/sdr", methods=["POST"])
def api_download_sdr():
    if "file" not in request.files:
        return "No file uploaded", 400
    try:
        from report_core import generate_tq_sdr
        result = generate_tq_sdr(request.files["file"].read())
        return send_file(BytesIO(result["sdr_bytes"]),
                         download_name=result["sdr_name"],
                         as_attachment=True, mimetype=XLSX_MIME)
    except Exception as e:
        return str(e), 500


# ══════════════════════════════════════════════════════════════════════════
# EXISTING — Report 2: COMP5 Issued Documents  (DO NOT TOUCH)
# ══════════════════════════════════════════════════════════════════════════

@app.route("/api/comp5/summary", methods=["POST"])
def api_comp5_summary():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    try:
        from report_core import generate_comp5
        result = generate_comp5(request.files["file"].read())
        return jsonify({"summary": result["summary"], "filename": result["filename"]})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/comp5/download", methods=["POST"])
def api_download_comp5():
    if "file" not in request.files:
        return "No file uploaded", 400
    try:
        from report_core import generate_comp5
        result = generate_comp5(request.files["file"].read())
        return send_file(BytesIO(result["bytes"]),
                         download_name=result["filename"],
                         as_attachment=True, mimetype=XLSX_MIME)
    except Exception as e:
        return str(e), 500


# ══════════════════════════════════════════════════════════════════════════
# NEW — Report 3: MR & TBE  (added below, nothing above touched)
# ══════════════════════════════════════════════════════════════════════════

@app.route("/api/mr/summary", methods=["POST"])
def api_mr_summary():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    try:
        from mr_tbe_report import generate_mr
        result = generate_mr(request.files["file"].read())
        return jsonify({"summary": result["summary"], "filename": result["filename"]})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/mr/download", methods=["POST"])
def api_download_mr():
    if "file" not in request.files:
        return "No file uploaded", 400
    try:
        from mr_tbe_report import generate_mr
        result = generate_mr(request.files["file"].read())
        return send_file(BytesIO(result["bytes"]),
                         download_name=result["filename"],
                         as_attachment=True, mimetype=XLSX_MIME)
    except Exception as e:
        return str(e), 500


@app.route("/api/tbe/summary", methods=["POST"])
def api_tbe_summary():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    try:
        from mr_tbe_report import generate_tbe
        result = generate_tbe(request.files["file"].read())
        return jsonify({"summary": result["summary"], "filename": result["filename"]})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/tbe/download", methods=["POST"])
def api_download_tbe():
    if "file" not in request.files:
        return "No file uploaded", 400
    try:
        from mr_tbe_report import generate_tbe
        result = generate_tbe(request.files["file"].read())
        return send_file(BytesIO(result["bytes"]),
                         download_name=result["filename"],
                         as_attachment=True, mimetype=XLSX_MIME)
    except Exception as e:
        return str(e), 500


@app.route("/api/mvdr/summary", methods=["POST"])
def api_mvdr_summary():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    try:
        from mvdr_report import summarise_mvdr
        from datetime import datetime
        summary = summarise_mvdr(request.files["file"].read())
        fname = f"COMP5_MVDR_Report_{datetime.today().strftime('%d%b%Y').upper()}.xlsx"
        return jsonify({"summary": summary, "filename": fname})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/mvdr/download", methods=["POST"])
def api_download_mvdr():
    if "file" not in request.files:
        return "No file uploaded", 400
    try:
        from mvdr_report import generate_mvdr
        result = generate_mvdr(request.files["file"].read())
        return send_file(BytesIO(result["bytes"]),
                         download_name=result["filename"],
                         as_attachment=True, mimetype=XLSX_MIME)
    except Exception as e:
        return str(e), 500


# ── Run ────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
