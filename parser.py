"""
Parses an ICICI Statement of Account (SOA) export and extracts, per transaction,
a best-effort counterparty name/UPI-id and a transaction category, by reading the
bank's free-text Narration field.

Accepts both modern (.xlsx) and legacy (.xls) Excel exports — the actual format
is detected from the file's contents, not its filename extension.

Expected input: a header row containing at least these columns (case-sensitive,
matching ICICI's export format):
    Ac_No, AC_Name, Tran_ID, Tran_Date, Inst_Type, Inst_Num,
    Dr_Amt, Cr_Amt, Balance, Narration, pstd_dt
"""
from datetime import datetime
import openpyxl
import xlrd

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

XLSX_MAGIC = b"PK\x03\x04"      # zip-based: .xlsx / .xlsm
XLS_MAGIC = b"\xd0\xcf\x11\xe0"  # OLE2-based: legacy .xls


class StatementFormatError(ValueError):
    """Raised when the uploaded file doesn't look like a supported ICICI SOA export."""


def _clean_name(s):
    if not s:
        return ""
    return s.strip(" -/.")


def _detect_excel_format(file_stream):
    """Sniff the first bytes to tell modern (.xlsx) from legacy (.xls) Excel files."""
    header = file_stream.read(8)
    file_stream.seek(0)
    if header.startswith(XLSX_MAGIC):
        return "xlsx"
    if header.startswith(XLS_MAGIC):
        return "xls"
    return None


def _rows_from_xlsx(file_stream):
    """Yield (header_row, data_row_iterator) for a modern .xlsx/.xlsm file."""
    wb = openpyxl.load_workbook(file_stream, data_only=True, read_only=True)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)
    header_row = next(rows_iter)
    return header_row, rows_iter


def _rows_from_xls(file_stream):
    """Yield (header_row, data_row_iterator) for a legacy .xls file, via xlrd."""
    wb = xlrd.open_workbook(file_contents=file_stream.read())
    ws = wb.sheet_by_index(0)
    datemode = wb.datemode

    def cell_value(r, c):
        cell = ws.cell(r, c)
        if cell.ctype == xlrd.XL_CELL_DATE:
            try:
                return xlrd.xldate_as_datetime(cell.value, datemode)
            except (xlrd.XLDateError, ValueError):
                return cell.value
        if cell.ctype == xlrd.XL_CELL_EMPTY:
            return None
        return cell.value

    header_row = [cell_value(0, c) for c in range(ws.ncols)]

    def data_rows():
        for r in range(1, ws.nrows):
            yield [cell_value(r, c) for c in range(ws.ncols)]

    return header_row, data_rows()


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
    Parse an uploaded ICICI SOA Excel file (a file-like object; .xlsx or legacy .xls,
    detected automatically) into a list of transaction dicts, plus a small dict of
    account info (account_no, account_name) read off the data rows. Raises
    StatementFormatError if the file isn't a recognized Excel format, or the
    expected columns aren't found.
    """
    fmt = _detect_excel_format(file_stream)
    if fmt == "xlsx":
        header_row, data_rows = _rows_from_xlsx(file_stream)
    elif fmt == "xls":
        header_row, data_rows = _rows_from_xls(file_stream)
    else:
        raise StatementFormatError(
            "This file doesn't look like a valid Excel file (.xlsx or .xls). "
            "Please upload the original statement export from ICICI, unmodified."
        )

    header = [str(h).strip() if h else "" for h in header_row]
    col_index = {name: idx for idx, name in enumerate(header)}

    missing = [c for c in REQUIRED_COLUMNS if c not in col_index]
    if missing:
        raise StatementFormatError(
            "This doesn't look like a supported ICICI statement export. "
            f"Missing expected column(s): {', '.join(missing)}."
        )

    account_info = {"account_no": None, "account_name": None}
    results = []
    for row in data_rows:
        if row is None or all(v is None for v in row):
            continue
        if account_info["account_no"] is None:
            ac_no = row[col_index["Ac_No"]]
            ac_name = row[col_index["AC_Name"]]
            if ac_no is not None:
                if isinstance(ac_no, float) and ac_no.is_integer():
                    account_info["account_no"] = str(int(ac_no))
                else:
                    account_info["account_no"] = str(ac_no)
            if ac_name is not None:
                account_info["account_name"] = str(ac_name)
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
            for fmt_str in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d"):
                try:
                    date_iso = datetime.strptime(tran_date, fmt_str).strftime("%Y-%m-%d")
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

    return results, account_info


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
