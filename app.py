import io
import os

from flask import Flask, jsonify, request, send_from_directory

from parser import StatementFormatError, build_summary, parse_workbook
from db import (
    check_connection, init_db, save_statement, list_statements, get_statement,
    delete_statement, save_batch, list_batches, get_batch, delete_batch,
)

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


def _analyze_one(upload):
    """
    Parse + save a single uploaded file. Returns a dict describing the outcome —
    never raises; failures come back as {"ok": False, "error": ...} so a batch
    of files can process independently.
    """
    filename = upload.filename or "(unnamed file)"

    if not filename.lower().endswith((".xlsx", ".xls")):
        return {"ok": False, "filename": filename, "error": "Please upload an .xlsx or .xls file."}

    try:
        data = io.BytesIO(upload.read())
        transactions, account_info = parse_workbook(data)
        summary = build_summary(transactions)
    except StatementFormatError as e:
        return {"ok": False, "filename": filename, "error": str(e)}
    except Exception as e:
        return {"ok": False, "filename": filename, "error": f"Could not parse this file: {e}"}

    statement_id = None
    try:
        statement_id = save_statement(filename, account_info, summary, transactions)
    except Exception as e:
        print(f"Could not save statement to DB: {e}")

    return {
        "ok": True,
        "filename": filename,
        "statement_id": statement_id,
        "account_no": account_info.get("account_no"),
        "account_name": account_info.get("account_name"),
        "summary": summary,
        "transactions": transactions,
    }


@app.post("/api/analyze")
def analyze():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded. Attach an .xlsx statement as 'file'."}), 400

    upload = request.files["file"]
    if not upload.filename:
        return jsonify({"error": "No file selected."}), 400

    result = _analyze_one(upload)
    if not result["ok"]:
        status = 422
        return jsonify({"error": result["error"]}), status

    return jsonify({
        "statement_id": result["statement_id"],
        "filename": result["filename"],
        "account_no": result["account_no"],
        "account_name": result["account_name"],
        "summary": result["summary"],
        "transactions": result["transactions"],
    })


def _ranges_overlap(a_start, a_end, b_start, b_end):
    if not (a_start and a_end and b_start and b_end):
        return False
    return a_start <= b_end and b_start <= a_end


@app.post("/api/analyze/bulk")
def analyze_bulk():
    uploads = request.files.getlist("files")
    if not uploads:
        return jsonify({"error": "No files uploaded. Attach one or more files as 'files'."}), 400

    files_result = []
    combined_transactions = []

    for upload in uploads:
        if not upload.filename:
            continue
        result = _analyze_one(upload)

        if not result["ok"]:
            files_result.append({
                "filename": result["filename"],
                "status": "error",
                "error": result["error"],
            })
            continue

        s = result["summary"]
        files_result.append({
            "filename": result["filename"],
            "status": "ok",
            "statement_id": result["statement_id"],
            "account_no": result["account_no"],
            "account_name": result["account_name"],
            "period_start": s["period_start"],
            "period_end": s["period_end"],
            "total_in": s["total_in"],
            "total_out": s["total_out"],
            "transaction_count": s["total_transactions"],
        })

        for t in result["transactions"]:
            tagged = dict(t)
            tagged["source_filename"] = result["filename"]
            tagged["source_account_no"] = result["account_no"]
            tagged["source_account_name"] = result["account_name"]
            combined_transactions.append(tagged)

    # Flag overlapping statement periods for the same account — a likely
    # accidental duplicate upload, not auto-deduplicated.
    warnings = []
    ok_files = [f for f in files_result if f["status"] == "ok"]
    for i in range(len(ok_files)):
        for j in range(i + 1, len(ok_files)):
            a, b = ok_files[i], ok_files[j]
            if a["account_no"] and a["account_no"] == b["account_no"] and _ranges_overlap(
                a["period_start"], a["period_end"], b["period_start"], b["period_end"]
            ):
                warnings.append(
                    f"\"{a['filename']}\" and \"{b['filename']}\" cover overlapping dates "
                    f"for the same account — check they aren't duplicates."
                )

    combined_summary = build_summary(combined_transactions) if combined_transactions else None

    batch_id = None
    try:
        statement_ids = [f.get("statement_id") for f in files_result if f["status"] == "ok"]
        batch_id = save_batch(statement_ids, combined_summary, warnings)
    except Exception as e:
        print(f"Could not save batch to DB: {e}")

    return jsonify({
        "batch_id": batch_id,
        "files": files_result,
        "warnings": warnings,
        "combined_summary": combined_summary,
        "transactions": combined_transactions,
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


@app.delete("/api/statements/<int:statement_id>")
def api_delete_statement(statement_id):
    try:
        deleted = delete_statement(statement_id)
    except Exception as e:
        return jsonify({"error": f"Could not delete this statement: {e}"}), 503

    if not deleted:
        return jsonify({"error": "That statement couldn't be found."}), 404

    return jsonify({"deleted": True})


@app.get("/api/batches")
def api_list_batches():
    try:
        rows = list_batches()
    except Exception as e:
        return jsonify({"error": f"Batch history isn't available right now: {e}"}), 503

    for r in rows:
        if r.get("created_at"):
            r["created_at"] = r["created_at"].isoformat()
        if r.get("total_in") is not None:
            r["total_in"] = float(r["total_in"])
        if r.get("total_out") is not None:
            r["total_out"] = float(r["total_out"])
        r["warnings"] = r["warnings"].split("\n") if r.get("warnings") else []

    return jsonify({"batches": rows})


@app.get("/api/batches/<int:batch_id>")
def api_get_batch(batch_id):
    try:
        batch, statements = get_batch(batch_id)
    except Exception as e:
        return jsonify({"error": f"Batch history isn't available right now: {e}"}), 503

    if batch is None:
        return jsonify({"error": "That bulk upload couldn't be found."}), 404

    files_result = []
    combined_transactions = []
    for s in statements:
        try:
            _, transactions = get_statement(s["id"])
        except Exception as e:
            return jsonify({"error": f"Could not reload one of this batch's statements: {e}"}), 503

        files_result.append({
            "filename": s["filename"],
            "status": "ok",
            "statement_id": s["id"],
            "account_no": s["account_no"],
            "account_name": s["account_name"],
            "period_start": s["period_start"].isoformat() if s["period_start"] else None,
            "period_end": s["period_end"].isoformat() if s["period_end"] else None,
            "total_in": float(s["total_in"]) if s["total_in"] is not None else 0.0,
            "total_out": float(s["total_out"]) if s["total_out"] is not None else 0.0,
            "transaction_count": s["transaction_count"],
        })
        for t in transactions:
            tagged = dict(t)
            tagged["source_filename"] = s["filename"]
            tagged["source_account_no"] = s["account_no"]
            tagged["source_account_name"] = s["account_name"]
            combined_transactions.append(tagged)

    combined_summary = build_summary(combined_transactions) if combined_transactions else None

    return jsonify({
        "batch_id": batch["id"],
        "files": files_result,
        "warnings": batch["warnings"].split("\n") if batch.get("warnings") else [],
        "combined_summary": combined_summary,
        "transactions": combined_transactions,
    })


@app.delete("/api/batches/<int:batch_id>")
def api_delete_batch(batch_id):
    delete_statements_too = request.args.get("delete_statements", "false").lower() == "true"
    try:
        deleted = delete_batch(batch_id, delete_statements_too=delete_statements_too)
    except Exception as e:
        return jsonify({"error": f"Could not delete this bulk upload: {e}"}), 503

    if not deleted:
        return jsonify({"error": "That bulk upload couldn't be found."}), 404

    return jsonify({"deleted": True})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
