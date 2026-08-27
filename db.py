"""
Thin Postgres connection helper.

Provides a connection, a health check, and persistence for analyzed
statements — so a statement only needs to be uploaded and parsed once,
then can be reopened later without re-uploading.
"""
import os
import psycopg2
from psycopg2.extras import execute_values, RealDictCursor

DATABASE_URL = os.environ.get("DATABASE_URL")


def get_connection():
    """Return a new psycopg2 connection. Raises if DATABASE_URL isn't set."""
    if not DATABASE_URL:
        raise RuntimeError("DATABASE_URL is not set")
    # Render's connection strings work as-is with psycopg2.
    return psycopg2.connect(DATABASE_URL)


def check_connection():
    """Return (ok: bool, detail: str) — used by the /health/db route."""
    if not DATABASE_URL:
        return False, "DATABASE_URL is not set"
    try:
        conn = get_connection()
        with conn.cursor() as cur:
            cur.execute("SELECT 1;")
            cur.fetchone()
        conn.close()
        return True, "connected"
    except Exception as e:
        return False, str(e)


def init_db():
    """Create the statements/transactions tables if they don't exist yet."""
    if not DATABASE_URL:
        return
    conn = get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS statements (
                    id SERIAL PRIMARY KEY,
                    filename TEXT NOT NULL,
                    account_no TEXT,
                    account_name TEXT,
                    period_start DATE,
                    period_end DATE,
                    total_in NUMERIC,
                    total_out NUMERIC,
                    transaction_count INTEGER,
                    uploaded_at TIMESTAMPTZ DEFAULT now()
                );
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS transactions (
                    id SERIAL PRIMARY KEY,
                    statement_id INTEGER REFERENCES statements(id) ON DELETE CASCADE,
                    tran_date DATE,
                    category TEXT,
                    counterparty TEXT,
                    direction TEXT,
                    dr NUMERIC,
                    cr NUMERIC,
                    balance NUMERIC,
                    narration TEXT
                );
            """)
            cur.execute(
                "CREATE INDEX IF NOT EXISTS idx_transactions_statement_id "
                "ON transactions(statement_id);"
            )
            cur.execute("""
                CREATE TABLE IF NOT EXISTS batches (
                    id SERIAL PRIMARY KEY,
                    file_count INTEGER,
                    total_in NUMERIC,
                    total_out NUMERIC,
                    transaction_count INTEGER,
                    warnings TEXT,
                    created_at TIMESTAMPTZ DEFAULT now()
                );
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS batch_statements (
                    batch_id INTEGER REFERENCES batches(id) ON DELETE CASCADE,
                    statement_id INTEGER REFERENCES statements(id) ON DELETE CASCADE,
                    PRIMARY KEY (batch_id, statement_id)
                );
            """)
        conn.commit()
    finally:
        conn.close()


def save_statement(filename, account_info, summary, transactions):
    """Persist one analyzed statement and all its transactions. Returns the new statement id."""
    conn = get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute(
                """
                INSERT INTO statements
                    (filename, account_no, account_name, period_start, period_end,
                     total_in, total_out, transaction_count)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s)
                RETURNING id;
                """,
                (
                    filename,
                    account_info.get("account_no"),
                    account_info.get("account_name"),
                    summary.get("period_start"),
                    summary.get("period_end"),
                    summary.get("total_in"),
                    summary.get("total_out"),
                    summary.get("total_transactions"),
                ),
            )
            statement_id = cur.fetchone()[0]

            rows = [
                (
                    statement_id,
                    t.get("date"),
                    t.get("category"),
                    t.get("counterparty"),
                    t.get("direction"),
                    t.get("dr"),
                    t.get("cr"),
                    t.get("balance"),
                    t.get("narration"),
                )
                for t in transactions
            ]
            execute_values(
                cur,
                """
                INSERT INTO transactions
                    (statement_id, tran_date, category, counterparty, direction, dr, cr, balance, narration)
                VALUES %s
                """,
                rows,
            )
        conn.commit()
        return statement_id
    finally:
        conn.close()


def list_statements():
    """Return metadata for every saved statement, most recent first."""
    conn = get_connection()
    try:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute(
                """
                SELECT id, filename, account_no, account_name, period_start, period_end,
                       total_in, total_out, transaction_count, uploaded_at
                FROM statements
                ORDER BY uploaded_at DESC;
                """
            )
            return [dict(r) for r in cur.fetchall()]
    finally:
        conn.close()


def get_statement(statement_id):
    """Return (metadata dict, transactions list) for one saved statement, or (None, None)."""
    conn = get_connection()
    try:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("SELECT * FROM statements WHERE id=%s;", (statement_id,))
            meta = cur.fetchone()
            if not meta:
                return None, None
            cur.execute(
                """
                SELECT tran_date, category, counterparty, direction, dr, cr, balance, narration
                FROM transactions
                WHERE statement_id=%s
                ORDER BY tran_date;
                """,
                (statement_id,),
            )
            txn_rows = cur.fetchall()
            transactions = [
                {
                    "date": r["tran_date"].strftime("%Y-%m-%d") if r["tran_date"] else None,
                    "category": r["category"],
                    "counterparty": r["counterparty"],
                    "direction": r["direction"],
                    "dr": float(r["dr"]) if r["dr"] is not None else 0.0,
                    "cr": float(r["cr"]) if r["cr"] is not None else 0.0,
                    "balance": float(r["balance"]) if r["balance"] is not None else None,
                    "narration": r["narration"],
                }
                for r in txn_rows
            ]
            return dict(meta), transactions
    finally:
        conn.close()


def save_batch(statement_ids, combined_summary, warnings):
    """
    Persist a bulk-upload batch as a list of statement ids that belong
    together, plus the combined totals at the time of upload. Returns the
    new batch id, or None if there were no successfully saved statements
    to link (e.g. every file in the batch failed to parse).
    """
    statement_ids = [sid for sid in statement_ids if sid is not None]
    if not statement_ids:
        return None

    conn = get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute(
                """
                INSERT INTO batches (file_count, total_in, total_out, transaction_count, warnings)
                VALUES (%s, %s, %s, %s, %s)
                RETURNING id;
                """,
                (
                    len(statement_ids),
                    combined_summary.get("total_in") if combined_summary else None,
                    combined_summary.get("total_out") if combined_summary else None,
                    combined_summary.get("total_transactions") if combined_summary else None,
                    "\n".join(warnings) if warnings else None,
                ),
            )
            batch_id = cur.fetchone()[0]
            execute_values(
                cur,
                "INSERT INTO batch_statements (batch_id, statement_id) VALUES %s",
                [(batch_id, sid) for sid in statement_ids],
            )
        conn.commit()
        return batch_id
    finally:
        conn.close()


def list_batches():
    """Return metadata for every saved bulk-upload batch, most recent first."""
    conn = get_connection()
    try:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute(
                """
                SELECT id, file_count, total_in, total_out, transaction_count, warnings, created_at
                FROM batches
                ORDER BY created_at DESC;
                """
            )
            return [dict(r) for r in cur.fetchall()]
    finally:
        conn.close()


def get_batch(batch_id):
    """
    Return (batch metadata dict, list of statement metadata dicts belonging
    to it) or (None, None) if the batch doesn't exist. Each statement dict
    includes its own id/filename/account info, but not its transactions —
    fetch those per-statement with get_statement().
    """
    conn = get_connection()
    try:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("SELECT * FROM batches WHERE id=%s;", (batch_id,))
            batch = cur.fetchone()
            if not batch:
                return None, None
            cur.execute(
                """
                SELECT s.id, s.filename, s.account_no, s.account_name,
                       s.period_start, s.period_end, s.total_in, s.total_out, s.transaction_count
                FROM batch_statements bs
                JOIN statements s ON s.id = bs.statement_id
                WHERE bs.batch_id = %s
                ORDER BY s.id;
                """,
                (batch_id,),
            )
            statements = [dict(r) for r in cur.fetchall()]
            return dict(batch), statements
    finally:
        conn.close()
