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
