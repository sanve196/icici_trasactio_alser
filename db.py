"""
Thin Postgres connection helper.

For now this only provides a connection + a health check, so the deployment
step can prove the database is reachable. Actual tables/storage for
uploaded statements will be added in the next phase (functional changes).
"""
import os
import psycopg2

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
