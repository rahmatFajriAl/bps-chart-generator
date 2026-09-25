"""Penyimpanan SQLite. SEMUA query memakai parameter (?) -> aman dari SQL injection."""
import os
import sqlite3
import time
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path

DB_PATH = Path(os.environ.get("CHARTGEN_DB", Path(__file__).parent / "data" / "chartgen.db"))
MAX_NAME_LEN = 120
MAX_SAVED = 200


@contextmanager
def _conn():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(DB_PATH, timeout=10)
    con.row_factory = sqlite3.Row
    try:
        yield con
        con.commit()
    except Exception:
        con.rollback()
        raise
    finally:
        con.close()


def _now():
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def init_db():
    with _conn() as c:
        c.execute("""CREATE TABLE IF NOT EXISTS saved_tables (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            name       TEXT    NOT NULL,
            ext        TEXT    NOT NULL CHECK (ext IN ('xlsx','xlsm','csv')),
            content    BLOB    NOT NULL,
            size       INTEGER NOT NULL,
            created_at TEXT    NOT NULL,
            updated_at TEXT    NOT NULL)""")
        c.execute("""CREATE TABLE IF NOT EXISTS login_guard (
            id           INTEGER PRIMARY KEY CHECK (id = 1),
            fail_count   INTEGER NOT NULL DEFAULT 0,
            locked_until REAL    NOT NULL DEFAULT 0)""")
        c.execute("INSERT OR IGNORE INTO login_guard (id) VALUES (1)")
    try:
        os.chmod(DB_PATH, 0o600)
    except OSError:
        pass


# ------------------------------------------------------------- tabel ----
def _unique_name(con, name, exclude_id=None):
    base, n, candidate = name[:MAX_NAME_LEN - 6], 1, name[:MAX_NAME_LEN]
    while True:
        row = con.execute("SELECT id FROM saved_tables WHERE name = ? COLLATE NOCASE AND id IS NOT ?",
                          (candidate, exclude_id)).fetchone()
        if row is None:
            return candidate
        n += 1
        candidate = f"{base} ({n})"


def count_tables():
    with _conn() as c:
        return c.execute("SELECT COUNT(*) FROM saved_tables").fetchone()[0]


def list_tables():
    with _conn() as c:
        rows = c.execute("SELECT id, name, ext, size, created_at, updated_at FROM saved_tables "
                         "ORDER BY updated_at DESC, id DESC").fetchall()
    return [dict(r) for r in rows]


def get_table(table_id):
    with _conn() as c:
        row = c.execute("SELECT * FROM saved_tables WHERE id = ?", (int(table_id),)).fetchone()
    return dict(row) if row else None


def create_table(name, ext, content: bytes):
    with _conn() as c:
        name = _unique_name(c, name)
        now = _now()
        cur = c.execute("INSERT INTO saved_tables (name, ext, content, size, created_at, updated_at) "
                        "VALUES (?, ?, ?, ?, ?, ?)", (name, ext, content, len(content), now, now))
        return cur.lastrowid


def update_table(table_id, name, ext, content: bytes):
    with _conn() as c:
        name = _unique_name(c, name, exclude_id=int(table_id))
        c.execute("UPDATE saved_tables SET name = ?, ext = ?, content = ?, size = ?, updated_at = ? "
                  "WHERE id = ?", (name, ext, content, len(content), _now(), int(table_id)))
    return name


def delete_table(table_id):
    with _conn() as c:
        c.execute("DELETE FROM saved_tables WHERE id = ?", (int(table_id),))


# ------------------------------------------------- pembatas login ----
def lock_remaining():
    """Sisa detik terkunci (0 = tidak terkunci)."""
    with _conn() as c:
        row = c.execute("SELECT locked_until FROM login_guard WHERE id = 1").fetchone()
    return max(0, int(row["locked_until"] - time.time()))


def register_failure(max_fails, lock_seconds):
    with _conn() as c:
        c.execute("UPDATE login_guard SET fail_count = fail_count + 1 WHERE id = 1")
        row = c.execute("SELECT fail_count FROM login_guard WHERE id = 1").fetchone()
        if row["fail_count"] >= max_fails:
            c.execute("UPDATE login_guard SET fail_count = 0, locked_until = ? WHERE id = 1",
                      (time.time() + lock_seconds,))


def reset_failures():
    with _conn() as c:
        c.execute("UPDATE login_guard SET fail_count = 0, locked_until = 0 WHERE id = 1")