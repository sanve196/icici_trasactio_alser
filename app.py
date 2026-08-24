import io
import os

from flask import Flask, jsonify, request, send_from_directory

from parser import StatementFormatError, build_summary, parse_workbook

app = Flask(__name__, static_folder="static", static_url_path="")

MAX_UPLOAD_MB = 20
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024


@app.get("/")
def index():
    return send_from_directory(app.static_folder, "index.html")


@app.get("/health")
def health():
    return jsonify({"status": "ok"})


@app.post("/api/analyze")
def analyze():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded. Attach an .xlsx statement as 'file'."}), 400

    upload = request.files["file"]
    if not upload.filename:
        return jsonify({"error": "No file selected."}), 400
    if not upload.filename.lower().endswith(".xlsx"):
        return jsonify({"error": "Please upload an .xlsx file."}), 400

    try:
        data = io.BytesIO(upload.read())
        transactions = parse_workbook(data)
        summary = build_summary(transactions)
    except StatementFormatError as e:
        return jsonify({"error": str(e)}), 422
    except Exception as e:
        return jsonify({"error": f"Could not parse this file: {e}"}), 422

    return jsonify({
        "summary": summary,
        "transactions": transactions,
    })


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
