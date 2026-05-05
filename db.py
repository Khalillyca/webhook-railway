"""
db.py — SQLite persistence layer

Tables:
  thread_state            — tracks msgCount per thread for delta-based trigger logic
  processed_notifications — deduplication of Microsoft Graph notification message IDs
  runtime_config          — stores targets, thread filter, overdue hours, scan date
"""

import sqlite3
import threading
import logging
from datetime import datetime, timezone
from typing import Optional

log = logging.getLogger(__name__)

DB_FILE = "tracker.db"
_local = threading.local()


# ------------------------------------------------------------------ #
#  Connection (thread-safe)                                           #
# ------------------------------------------------------------------ #

def _conn() -> sqlite3.Connection:
    """Get a thread-local SQLite connection."""
    if not hasattr(_local, "conn") or _local.conn is None:
        _local.conn = sqlite3.connect(DB_FILE, check_same_thread=False)
        _local.conn.row_factory = sqlite3.Row
        _local.conn.execute("PRAGMA journal_mode=WAL")
    return _local.conn


def init_db():
    """Create all tables if they don't exist."""
    conn = _conn()
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS thread_state (
            thread_id       TEXT PRIMARY KEY,
            subject         TEXT DEFAULT '',
            msg_count       INTEGER DEFAULT 0,
            last_status     TEXT DEFAULT 'Open',
            last_summary    TEXT DEFAULT '',
            last_department TEXT DEFAULT '',
            last_priority   TEXT DEFAULT 'Medium',
            last_activity   TEXT DEFAULT '',
            last_triggered  TEXT DEFAULT '',
            created_at      TEXT DEFAULT ''
        );

        CREATE TABLE IF NOT EXISTS processed_notifications (
            message_id  TEXT PRIMARY KEY,
            processed_at TEXT DEFAULT ''
        );

        CREATE TABLE IF NOT EXISTS runtime_config (
            key   TEXT PRIMARY KEY,
            value TEXT DEFAULT ''
        );
    """)
    conn.commit()
    log.info("[DB] Database initialized — all tables ready.")


# ------------------------------------------------------------------ #
#  THREAD STATE                                                        #
# ------------------------------------------------------------------ #

def get_thread_state(thread_id: str) -> Optional[dict]:
    """Get the stored state for a thread. Returns None if not tracked."""
    row = _conn().execute(
        "SELECT * FROM thread_state WHERE thread_id = ?", (thread_id,)
    ).fetchone()
    return dict(row) if row else None


def get_all_thread_states() -> list:
    """Get all stored thread states."""
    rows = _conn().execute("SELECT * FROM thread_state ORDER BY last_activity DESC").fetchall()
    return [dict(r) for r in rows]


def upsert_thread_state(
    thread_id: str,
    msg_count: int,
    subject: str = "",
    last_status: str = "Open",
    last_summary: str = "",
    last_department: str = "",
    last_priority: str = "Medium",
    last_activity: str = "",
):
    """Insert or update thread state. Returns (old_count, new_count)."""
    now = datetime.now(timezone.utc).isoformat()
    existing = get_thread_state(thread_id)
    old_count = existing["msg_count"] if existing else 0

    _conn().execute(
        """
        INSERT INTO thread_state
            (thread_id, subject, msg_count, last_status, last_summary,
             last_department, last_priority, last_activity, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(thread_id) DO UPDATE SET
            msg_count       = excluded.msg_count,
            subject         = CASE WHEN excluded.subject != '' THEN excluded.subject ELSE thread_state.subject END,
            last_status     = excluded.last_status,
            last_summary    = excluded.last_summary,
            last_department = excluded.last_department,
            last_priority   = excluded.last_priority,
            last_activity   = excluded.last_activity
        """,
        (thread_id, subject, msg_count, last_status, last_summary,
         last_department, last_priority, last_activity, now),
    )
    _conn().commit()
    return old_count, msg_count


def set_thread_triggered(thread_id: str):
    """Mark a thread as triggered right now."""
    now = datetime.now(timezone.utc).isoformat()
    _conn().execute(
        "UPDATE thread_state SET last_triggered = ? WHERE thread_id = ?",
        (now, thread_id),
    )
    _conn().commit()


def get_pending_threads(overdue_hours: float) -> list:
    """
    Get threads with status 'Pending Reply' whose last_activity
    is older than overdue_hours ago.
    """
    rows = _conn().execute(
        """
        SELECT * FROM thread_state
        WHERE last_status = 'Pending Reply'
        ORDER BY last_activity ASC
        """
    ).fetchall()

    results = []
    now = datetime.now(timezone.utc)
    for row in rows:
        row_dict = dict(row)
        last_act = row_dict.get("last_activity", "")
        if not last_act:
            continue
        try:
            # Parse ISO datetime
            act_dt = datetime.fromisoformat(last_act.replace("Z", "+00:00"))
            if act_dt.tzinfo is None:
                act_dt = act_dt.replace(tzinfo=timezone.utc)
            hours_elapsed = (now - act_dt).total_seconds() / 3600
            if hours_elapsed >= overdue_hours:
                row_dict["hours_elapsed"] = round(hours_elapsed, 1)
                results.append(row_dict)
        except (ValueError, TypeError):
            continue

    return results


# ------------------------------------------------------------------ #
#  PROCESSED NOTIFICATIONS (dedup)                                     #
# ------------------------------------------------------------------ #

def is_message_processed(message_id: str) -> bool:
    """Check if a message ID has already been processed."""
    row = _conn().execute(
        "SELECT 1 FROM processed_notifications WHERE message_id = ?",
        (message_id,),
    ).fetchone()
    return row is not None


def mark_message_processed(message_id: str):
    """Mark a message ID as processed."""
    now = datetime.now(timezone.utc).isoformat()
    _conn().execute(
        "INSERT OR IGNORE INTO processed_notifications (message_id, processed_at) VALUES (?, ?)",
        (message_id, now),
    )
    _conn().commit()


def cleanup_old_processed(keep_last_n: int = 5000):
    """Keep only the most recent N processed message IDs."""
    _conn().execute(
        f"""
        DELETE FROM processed_notifications
        WHERE message_id NOT IN (
            SELECT message_id FROM processed_notifications
            ORDER BY processed_at DESC
            LIMIT ?
        )
        """,
        (keep_last_n,),
    )
    _conn().commit()


# ------------------------------------------------------------------ #
#  RUNTIME CONFIG                                                      #
# ------------------------------------------------------------------ #

def set_config(key: str, value: str):
    """Set a runtime config key."""
    _conn().execute(
        "INSERT OR REPLACE INTO runtime_config (key, value) VALUES (?, ?)",
        (key, value),
    )
    _conn().commit()


def get_config(key: str, default: str = "") -> str:
    """Get a runtime config value."""
    row = _conn().execute(
        "SELECT value FROM runtime_config WHERE key = ?", (key,)
    ).fetchone()
    return row["value"] if row else default


def get_config_targets() -> list:
    """Get target emails/domains as a list."""
    raw = get_config("targets", "")
    if not raw:
        return []
    return [t.strip() for t in raw.split(",") if t.strip()]


def get_config_thread_filter() -> str:
    """Get optional thread ID filter (empty = all)."""
    return get_config("thread_filter", "")


def get_config_overdue_hours() -> float:
    """Get overdue threshold in hours."""
    val = get_config("overdue_hours", "6")
    try:
        return float(val)
    except ValueError:
        return 6.0


def get_config_scan_date() -> str:
    """Get the bootstrap scan start date."""
    return get_config("scan_date", "")
