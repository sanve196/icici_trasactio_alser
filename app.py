import io
import os

from flask import Flask, jsonify, request, send_from_directory

from parser import StatementFormatError, build_summary, parse_workbook
from db import check_connection, init_db, save_statement, list_statements, get_statement

app = Flask(__name__, static_folder="static", static_url_path="")

MAX_UPLOAD_MB = 20
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024

try:
    init_db()
except Exception as e:
    # Don't crash app startup if the DB isn't reachable yet (or isn't configured
    # locally) — uploads still work, they just won't be saved for later.
    print(f"DB init skipped: {e}")


@app.get("/")
def index():
    return send_from_directory(app.static_folder, "index.html")


@app.get("/health")
def health():
    return jsonify({"status": "ok"})


@app.get("/health/db")
def health_db():
    ok, detail = check_connection()
    return jsonify({"db_connected": ok, "detail": detail}), (200 if ok else 503)


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
        transactions, account_info = parse_workbook(data)
        summary = build_summary(transactions)
    except StatementFormatError as e:
        return jsonify({"error": str(e)}), 422
    except Exception as e:
        return jsonify({"error": f"Could not parse this file: {e}"}), 422

    statement_id = None
    try:
        statement_id = save_statement(upload.filename, account_info, summary, transactions)
    except Exception as e:
        # Analysis still succeeds even if we couldn't save it for later.
        print(f"Could not save statement to DB: {e}")

    return jsonify({
        "statement_id": statement_id,
        "filename": upload.filename,
        "account_no": account_info.get("account_no"),
        "account_name": account_info.get("account_name"),
        "summary": summary,
        "transactions": transactions,
    })


@app.get("/api/statements")
def api_list_statements():
    try:
        rows = list_statements()
    except Exception as e:
        return jsonify({"error": f"Statement history isn't available right now: {e}"}), 503

    for r in rows:
        if r.get("period_start"):
            r["period_start"] = r["period_start"].isoformat()
        if r.get("period_end"):
            r["period_end"] = r["period_end"].isoformat()
        if r.get("uploaded_at"):
            r["uploaded_at"] = r["uploaded_at"].isoformat()
        if r.get("total_in") is not None:
            r["total_in"] = float(r["total_in"])
        if r.get("total_out") is not None:
            r["total_out"] = float(r["total_out"])

    return jsonify({"statements": rows})


@app.get("/api/statements/<int:statement_id>")
def api_get_statement(statement_id):
    try:
        meta, transactions = get_statement(statement_id)
    except Exception as e:
        return jsonify({"error": f"Statement history isn't available right now: {e}"}), 503

    if meta is None:
        return jsonify({"error": "That statement couldn't be found."}), 404

    summary = build_summary(transactions)
    return jsonify({
        "statement_id": meta["id"],
        "filename": meta["filename"],
        "account_no": meta["account_no"],
        "account_name": meta["account_name"],
        "summary": summary,
        "transactions": transactions,
    })


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
