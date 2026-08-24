"""
Parses an ICICI Statement of Account (SOA) export and extracts, per transaction,
a best-effort counterparty name/UPI-id and a transaction category, by reading the
bank's free-text Narration field.

Expected input: an .xlsx file with a header row containing at least these columns
(case-sensitive, matching ICICI's export format):
    Ac_No, AC_Name, Tran_ID, Tran_Date, Inst_Type, Inst_Num,
    Dr_Amt, Cr_Amt, Balance, Narration, pstd_dt
"""
from datetime import datetime
import openpyxl

REQUIRED_COLUMNS = [
    "Ac_No", "AC_Name", "Tran_ID", "Tran_Date", "Inst_Type", "Inst_Num",
    "Dr_Amt", "Cr_Amt", "Balance", "Narration", "pstd_dt",
]

BANK_FEE_PREFIXES = (
    "IMPS Chg", "Chq rtn Chg", "Int on delayed", "REJECT", "EZY",
    "Debit Card", "SMS Charge", "AMC", "ATM", "GST", "CHEQUE BOOK",
)

EXCLUDE_CATEGORIES_FROM_COUNTERPARTY_VIEWS = {
    "Bank Charges / Fees / Returns",
    "Cash Deposit",
}


class StatementFormatError(ValueError):
    """Raised when the uploaded file doesn't look like a supported ICICI SOA export."""


def _clean_name(s):
    if not s:
        return ""
    return s.strip(" -/.")


def parse_narration(narration, dr, cr):
    """Return (category, counterparty, direction) for one transaction row."""
    n = narration or ""
    direction = "IN" if cr and cr > 0 else ("OUT" if dr and dr > 0 else "NONE")

    for pfx in BANK_FEE_PREFIXES:
        if n.startswith(pfx):
            return ("Bank Charges / Fees / Returns", _clean_name(n), direction)

    parts = n.split("/")

    if n.startswith("UPI/"):
        vpa = parts[3] if len(parts) > 3 else ""
        remark = parts[2] if len(parts) > 2 else ""
        vpa_id = vpa.split("@")[0] if "@" in vpa else vpa
        label = remark if remark and not remark.lower().startswith("payment from ph") else vpa_id
        cp = _clean_name(label) or _clean_name(vpa_id) or "Unknown UPI counterparty"
        return ("UPI", cp, direction)

    if n.startswith("MMT/IMPS/") or n.startswith("MMT/"):
        name = parts[4] if len(parts) > 4 else (parts[3] if len(parts) > 3 else "")
        return ("IMPS", _clean_name(name) or "Unknown IMPS counterparty", direction)

    if n.startswith("INF/INFT") or n.startswith("INF/"):
        name = parts[4] if len(parts) > 4 else (parts[3] if len(parts) > 3 else "")
        return ("Internal Fund Transfer (INFT)", _clean_name(name) or "Unknown Internal-Transfer counterparty", direction)

    if n.startswith("CLG/"):
        name = parts[1] if len(parts) > 1 else ""
        return ("Cheque Clearing (CLG)", _clean_name(name) or "Unknown", direction)

    if n.startswith("TRF/"):
        name = parts[1] if len(parts) > 1 else ""
        return ("Fund Transfer (TRF)", _clean_name(name) or "Unknown", direction)

    if n.startswith("TRFR TO"):
        name = n.split(":", 1)[1] if ":" in n else ""
        return ("Internal Transfer OUT", _clean_name(name) or "Unknown", direction)

    if n.startswith("TRFR FROM"):
        name = n.split(":", 1)[1] if ":" in n else ""
        return ("Internal Transfer IN", _clean_name(name) or "Unknown", direction)

    if n.startswith("NEFT"):
        p = n.replace("NEFT-", "").replace("NEFT/", "").split("-")
        name = p[1] if len(p) > 1 else ""
        return ("NEFT", _clean_name(name) or "Unknown", direction)

    if n.startswith("RTGS"):
        if "-" in n:
            p = n.split("-")
            name = p[2] if len(p) > 2 else ""
        else:
            p = n.split("/")
            name = p[3] if len(p) > 3 else ""
        return ("RTGS", _clean_name(name) or "Unknown", direction)

    if n.startswith("BIL/"):
        name = parts[3] if len(parts) > 3 else ""
        return ("Bill Payment", _clean_name(name) or "Unknown biller", direction)

    if n.startswith("CASH PAID"):
        name = n.split(":", 1)[1] if ":" in n else ""
        return ("Cash Withdrawal/Payment", _clean_name(name) or "Cash", direction)

    if n.startswith("CAM/"):
        return ("Cash Deposit", "Cash Deposit (depositor not named)", direction)

    return ("Other/Unclassified", _clean_name(n) or "Unknown", direction)


def parse_workbook(file_stream):
    """
    Parse an uploaded ICICI SOA .xlsx file (a file-like object) into a list of
    transaction dicts. Raises StatementFormatError if the expected columns
    aren't found.
    """
    wb = openpyxl.load_workbook(file_stream, data_only=True, read_only=True)
    ws = wb.active

    header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
    header = [str(h).strip() if h else "" for h in header_row]
    col_index = {name: idx for idx, name in enumerate(header)}

    missing = [c for c in REQUIRED_COLUMNS if c not in col_index]
    if missing:
        raise StatementFormatError(
            "This doesn't look like a supported ICICI statement export. "
            f"Missing expected column(s): {', '.join(missing)}."
        )

    results = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row is None or all(v is None for v in row):
            continue
        tran_date = row[col_index["Tran_Date"]]
        dr = row[col_index["Dr_Amt"]] or 0
        cr = row[col_index["Cr_Amt"]] or 0
        balance = row[col_index["Balance"]]
        narration = row[col_index["Narration"]]

        category, counterparty, direction = parse_narration(narration, dr, cr)
        amount = cr if cr > 0 else dr

        date_iso = None
        if isinstance(tran_date, datetime):
            date_iso = tran_date.strftime("%Y-%m-%d")
        elif isinstance(tran_date, str):
            for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d"):
                try:
                    date_iso = datetime.strptime(tran_date, fmt).strftime("%Y-%m-%d")
                    break
                except ValueError:
                    continue

        results.append({
            "date": date_iso,
            "category": category,
            "counterparty": counterparty,
            "direction": direction,
            "dr": float(dr) if dr else 0.0,
            "cr": float(cr) if cr else 0.0,
            "balance": float(balance) if balance is not None else None,
            "narration": narration,
        })

    if not results:
        raise StatementFormatError("No transaction rows were found in this file.")

    return results


def build_summary(transactions):
    """Compute overall totals and per-counterparty aggregates for both directions."""
    def agg_for(direction):
        agg = {}
        for t in transactions:
            if t["direction"] != direction:
                continue
            if t["category"] in EXCLUDE_CATEGORIES_FROM_COUNTERPARTY_VIEWS:
                continue
            key = t["counterparty"]
            entry = agg.setdefault(key, {"name": key, "total": 0.0, "count": 0, "first": t["date"], "last": t["date"]})
            entry["total"] += t["dr"] if direction == "OUT" else t["cr"]
            entry["count"] += 1
            if t["date"] and (entry["first"] is None or t["date"] < entry["first"]):
                entry["first"] = t["date"]
            if t["date"] and (entry["last"] is None or t["date"] > entry["last"]):
                entry["last"] = t["date"]
        return sorted(agg.values(), key=lambda x: -x["total"])

    total_in = sum(t["cr"] for t in transactions if t["direction"] == "IN")
    total_out = sum(t["dr"] for t in transactions if t["direction"] == "OUT")
    dates = sorted(t["date"] for t in transactions if t["date"])

    return {
        "total_transactions": len(transactions),
        "period_start": dates[0] if dates else None,
        "period_end": dates[-1] if dates else None,
        "total_in": total_in,
        "total_out": total_out,
        "net": total_in - total_out,
        "payers": agg_for("IN"),
        "payees": agg_for("OUT"),
    }
