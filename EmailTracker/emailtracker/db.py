from __future__ import annotations

import json
import sqlite3
from contextlib import contextmanager
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Iterable, Iterator, Sequence

SCHEMA = """
CREATE TABLE IF NOT EXISTS messages (
    id               TEXT PRIMARY KEY,
    conversation_id  TEXT,
    direction        TEXT NOT NULL CHECK(direction IN ('in','out')),
    folder           TEXT NOT NULL,
    sender_name      TEXT,
    sender_address   TEXT,
    to_recipients    TEXT NOT NULL DEFAULT '[]',
    cc_recipients    TEXT NOT NULL DEFAULT '[]',
    bcc_recipients   TEXT NOT NULL DEFAULT '[]',
    subject          TEXT,
    received_at      TEXT,
    is_reply         INTEGER NOT NULL DEFAULT 0,
    is_forward       INTEGER NOT NULL DEFAULT 0,
    has_attachments  INTEGER NOT NULL DEFAULT 0,
    requires_reply   INTEGER NOT NULL DEFAULT 0,
    body_preview     TEXT NOT NULL DEFAULT '',
    ingested_at      TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_messages_conversation ON messages(conversation_id);
CREATE INDEX IF NOT EXISTS idx_messages_received_at ON messages(received_at DESC);
CREATE INDEX IF NOT EXISTS idx_messages_sender ON messages(sender_address);

CREATE TABLE IF NOT EXISTS sync_state (
    folder           TEXT PRIMARY KEY,
    watermark        TEXT,
    last_success_at  TEXT,
    last_error       TEXT
);

CREATE TABLE IF NOT EXISTS filter_rules (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    name             TEXT NOT NULL,
    sender_pattern   TEXT,
    subject_pattern  TEXT,
    enabled          INTEGER NOT NULL DEFAULT 1,
    created_at       TEXT NOT NULL,
    CHECK (sender_pattern IS NOT NULL OR subject_pattern IS NOT NULL)
);
"""


def now_utc_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def connect(db_path: Path) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path, isolation_level=None, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def init_schema(conn: sqlite3.Connection) -> None:
    conn.executescript(SCHEMA)
    # Migrations: add columns that may not exist in older databases
    for col in [
        "has_attachments INTEGER NOT NULL DEFAULT 0",
        "requires_reply INTEGER NOT NULL DEFAULT 0",
        "body_preview TEXT NOT NULL DEFAULT ''",
    ]:
        try:
            conn.execute(f"ALTER TABLE messages ADD COLUMN {col}")
        except sqlite3.OperationalError:
            pass  # column already exists


@contextmanager
def transaction(conn: sqlite3.Connection) -> Iterator[sqlite3.Connection]:
    conn.execute("BEGIN")
    try:
        yield conn
    except Exception:
        conn.execute("ROLLBACK")
        raise
    else:
        conn.execute("COMMIT")


def upsert_message(conn: sqlite3.Connection, row: dict[str, Any]) -> None:
    conn.execute(
        """
        INSERT INTO messages (
            id, conversation_id, direction, folder,
            sender_name, sender_address,
            to_recipients, cc_recipients, bcc_recipients,
            subject, received_at, is_reply, is_forward, has_attachments, requires_reply, body_preview, ingested_at
        ) VALUES (
            :id, :conversation_id, :direction, :folder,
            :sender_name, :sender_address,
            :to_recipients, :cc_recipients, :bcc_recipients,
            :subject, :received_at, :is_reply, :is_forward, :has_attachments, :requires_reply, :body_preview, :ingested_at
        )
        ON CONFLICT(id) DO UPDATE SET
            conversation_id = excluded.conversation_id,
            direction       = excluded.direction,
            folder          = excluded.folder,
            sender_name     = excluded.sender_name,
            sender_address  = excluded.sender_address,
            to_recipients   = excluded.to_recipients,
            cc_recipients   = excluded.cc_recipients,
            bcc_recipients  = excluded.bcc_recipients,
            subject         = excluded.subject,
            received_at     = excluded.received_at,
            is_reply        = excluded.is_reply,
            is_forward      = excluded.is_forward,
            has_attachments = excluded.has_attachments,
            requires_reply  = excluded.requires_reply,
            body_preview    = excluded.body_preview
        """,
        row,
    )


def get_sync_state(conn: sqlite3.Connection, folder: str) -> sqlite3.Row | None:
    return conn.execute(
        "SELECT folder, watermark, last_success_at, last_error FROM sync_state WHERE folder = ?",
        (folder,),
    ).fetchone()


def save_sync_success(conn: sqlite3.Connection, folder: str, watermark: str | None) -> None:
    conn.execute(
        """
        INSERT INTO sync_state (folder, watermark, last_success_at, last_error)
        VALUES (?, ?, ?, NULL)
        ON CONFLICT(folder) DO UPDATE SET
            watermark       = excluded.watermark,
            last_success_at = excluded.last_success_at,
            last_error      = NULL
        """,
        (folder, watermark, now_utc_iso()),
    )


def save_sync_error(conn: sqlite3.Connection, folder: str, error: str) -> None:
    conn.execute(
        """
        INSERT INTO sync_state (folder, watermark, last_success_at, last_error)
        VALUES (?, NULL, NULL, ?)
        ON CONFLICT(folder) DO UPDATE SET last_error = excluded.last_error
        """,
        (folder, error),
    )


def list_filter_rules(conn: sqlite3.Connection, only_enabled: bool = False) -> list[sqlite3.Row]:
    sql = "SELECT id, name, sender_pattern, subject_pattern, enabled, created_at FROM filter_rules"
    if only_enabled:
        sql += " WHERE enabled = 1"
    sql += " ORDER BY created_at DESC"
    return list(conn.execute(sql))


def create_filter_rule(
    conn: sqlite3.Connection,
    name: str,
    sender_pattern: str | None,
    subject_pattern: str | None,
) -> int:
    cur = conn.execute(
        """
        INSERT INTO filter_rules (name, sender_pattern, subject_pattern, enabled, created_at)
        VALUES (?, ?, ?, 1, ?)
        """,
        (name, sender_pattern or None, subject_pattern or None, now_utc_iso()),
    )
    return int(cur.lastrowid)


def toggle_filter_rule(conn: sqlite3.Connection, rule_id: int) -> None:
    conn.execute("UPDATE filter_rules SET enabled = 1 - enabled WHERE id = ?", (rule_id,))


def delete_filter_rule(conn: sqlite3.Connection, rule_id: int) -> None:
    conn.execute("DELETE FROM filter_rules WHERE id = ?", (rule_id,))


def count_messages(conn: sqlite3.Connection) -> int:
    row = conn.execute("SELECT COUNT(*) AS c FROM messages").fetchone()
    return int(row["c"]) if row else 0


def get_status(conn: sqlite3.Connection) -> dict[str, Any]:
    rows = list(conn.execute("SELECT folder, last_success_at, last_error FROM sync_state"))
    per_folder = {
        r["folder"]: {"last_success_at": r["last_success_at"], "last_error": r["last_error"]}
        for r in rows
    }
    last_sync = max(
        (r["last_success_at"] for r in rows if r["last_success_at"]),
        default=None,
    )
    any_error = any(r["last_error"] for r in rows)
    return {
        "last_sync": last_sync,
        "tracked_count": count_messages(conn),
        "poller_healthy": bool(rows) and not any_error,
        "per_folder": per_folder,
    }


def get_metrics(
    conn: sqlite3.Connection,
    days: int | None = None,
    rules: Sequence[sqlite3.Row] | None = None,
    date: str | None = None,
) -> dict[str, Any]:
    """Compute dashboard metrics, optionally filtered to the last N days, a specific date, and active rules."""
    conditions: list[str] = []
    params: list[Any] = []

    if date:
        # Specific date filter: e.g. "2026-04-15" → that whole day
        conditions.append("received_at >= ?")
        params.append(f"{date}T00:00:00+00:00")
        conditions.append("received_at < ?")
        next_day = (datetime.strptime(date, "%Y-%m-%d") + timedelta(days=1)).strftime("%Y-%m-%d")
        params.append(f"{next_day}T00:00:00+00:00")
    elif days is not None and days > 0:
        # Start of today (midnight) minus (days-1) to get the period start.
        # DB timestamps are stored as local time with +00:00 suffix, so compare
        # against local midnight directly (no UTC conversion).
        today_start = datetime.now(timezone.utc).replace(hour=0, minute=0, second=0, microsecond=0)
        period_start = today_start - timedelta(days=days - 1)
        cutoff = period_start.isoformat()
        conditions.append("received_at > ?")
        params.append(cutoff)

    # Apply filter rules (same logic as feed)
    if rules is not None:
        filter_sql, filter_params = build_filter_clause(rules)
        if filter_sql != "1=1":
            conditions.append(filter_sql)
            params.extend(filter_params)

    where = (" WHERE " + " AND ".join(conditions)) if conditions else ""

    # Total inbound / outbound
    row = conn.execute(
        f"SELECT "
        f"  COUNT(*) AS total, "
        f"  SUM(CASE WHEN direction='in' THEN 1 ELSE 0 END) AS inbound, "
        f"  SUM(CASE WHEN direction='out' THEN 1 ELSE 0 END) AS outbound "
        f"FROM messages{where}",
        params,
    ).fetchone()
    total = int(row["total"]) if row else 0
    inbound = int(row["inbound"] or 0) if row else 0
    outbound = int(row["outbound"] or 0) if row else 0

    # Pending replies: inbound messages (matching filters + time) with no outbound reply
    pending_conditions = ["m.direction = 'in'"]
    pending_params: list[Any] = []

    if date:
        pending_conditions.append("m.received_at >= ?")
        pending_params.append(f"{date}T00:00:00+00:00")
        pending_conditions.append("m.received_at < ?")
        next_day = (datetime.strptime(date, "%Y-%m-%d") + timedelta(days=1)).strftime("%Y-%m-%d")
        pending_params.append(f"{next_day}T00:00:00+00:00")
    elif days is not None and days > 0:
        pending_conditions.append("m.received_at > ?")
        pending_params.append(cutoff)

    if rules is not None:
        filter_sql, filter_params = build_filter_clause(rules)
        if filter_sql != "1=1":
            # Qualify column references for the outer table alias 'm'
            qualified = filter_sql.replace("sender_address", "m.sender_address").replace("subject", "m.subject")
            pending_conditions.append(qualified)
            pending_params.extend(filter_params)

    pending_sql = (
        "SELECT COUNT(*) AS c FROM messages m WHERE "
        + " AND ".join(pending_conditions)
        + " AND NOT EXISTS ("
        "    SELECT 1 FROM messages o "
        "    WHERE o.conversation_id = m.conversation_id "
        "    AND o.direction = 'out' "
        "    AND o.is_reply = 1 "
        "    AND o.received_at > m.received_at"
        ")"
    )
    pending_row = conn.execute(pending_sql, pending_params).fetchone()
    pending = int(pending_row["c"]) if pending_row else 0

    # Replied count
    replied = inbound - pending

    return {
        "total": total,
        "inbound": inbound,
        "outbound": outbound,
        "pending_replies": pending,
        "replied": replied,
    }


def _pattern_to_like(pattern: str) -> str:
    if "*" in pattern:
        return pattern.replace("*", "%")
    return f"%{pattern}%"


def build_filter_clause(rules: Sequence[sqlite3.Row]) -> tuple[str, list[Any]]:
    """Return (sql_fragment, params) for the filter WHERE clause.

    - No enabled rules → ``1=1`` (match everything).
    - Multiple rules → OR'd together.
    - Within a rule, sender + subject patterns are AND'd.
    - Matching is case-insensitive via LOWER(...).
    - ``*`` in a pattern is treated as SQL ``%``; otherwise the pattern is
      wrapped as a substring (``%pattern%``).
    """
    enabled = [r for r in rules if int(r["enabled"]) == 1]
    if not enabled:
        return "1=1", []

    branches: list[str] = []
    params: list[Any] = []
    for r in enabled:
        conds: list[str] = []
        if r["sender_pattern"]:
            conds.append("LOWER(IFNULL(sender_address,'')) LIKE LOWER(?)")
            params.append(_pattern_to_like(r["sender_pattern"]))
        if r["subject_pattern"]:
            conds.append("LOWER(IFNULL(subject,'')) LIKE LOWER(?)")
            params.append(_pattern_to_like(r["subject_pattern"]))
        if conds:
            branches.append("(" + " AND ".join(conds) + ")")

    if not branches:
        return "1=1", []
    return "(" + " OR ".join(branches) + ")", params


def search_messages(
    conn: sqlite3.Connection,
    rules: Sequence[sqlite3.Row],
    query: str | None,
    limit: int = 100,
    days: int | None = None,
    date: str | None = None,
) -> list[sqlite3.Row]:
    filter_sql, filter_params = build_filter_clause(rules)
    where_parts = [filter_sql]
    params: list[Any] = list(filter_params)

    # Date/period filtering
    if date:
        where_parts.append("received_at >= ?")
        params.append(f"{date}T00:00:00+00:00")
        next_day = (datetime.strptime(date, "%Y-%m-%d") + timedelta(days=1)).strftime("%Y-%m-%d")
        where_parts.append("received_at < ?")
        params.append(f"{next_day}T00:00:00+00:00")
    elif days is not None and days > 0:
        today_start = datetime.now(timezone.utc).replace(hour=0, minute=0, second=0, microsecond=0)
        period_start = today_start - timedelta(days=days - 1)
        where_parts.append("received_at > ?")
        params.append(period_start.isoformat())

    if query:
        where_parts.append(
            "("
            "LOWER(IFNULL(sender_address,'')) LIKE LOWER(?) "
            "OR LOWER(IFNULL(subject,'')) LIKE LOWER(?) "
            "OR LOWER(IFNULL(to_recipients,'')) LIKE LOWER(?)"
            ")"
        )
        like = f"%{query}%"
        params.extend([like, like, like])

    sql = (
        "SELECT id, conversation_id, direction, folder, sender_name, sender_address, "
        "to_recipients, cc_recipients, bcc_recipients, subject, received_at, "
        "is_reply, is_forward, has_attachments, requires_reply, body_preview, ingested_at "
        "FROM messages WHERE "
        + " AND ".join(where_parts)
        + " ORDER BY received_at DESC LIMIT ?"
    )
    params.append(limit)
    return list(conn.execute(sql, params))


def get_conversation(conn: sqlite3.Connection, conversation_id: str) -> list[sqlite3.Row]:
    return list(
        conn.execute(
            "SELECT id, conversation_id, direction, folder, sender_name, sender_address, "
            "to_recipients, cc_recipients, bcc_recipients, subject, received_at, "
            "is_reply, is_forward, has_attachments, requires_reply, body_preview, ingested_at "
            "FROM messages WHERE conversation_id = ? "
            "ORDER BY received_at DESC",
            (conversation_id,),
        )
    )


def serialize_recipients(recipients: Iterable[dict[str, Any]] | None) -> str:
    if not recipients:
        return "[]"
    out: list[dict[str, str]] = []
    for r in recipients:
        addr = (r or {}).get("emailAddress") or {}
        out.append({"name": addr.get("name", ""), "address": addr.get("address", "")})
    return json.dumps(out, ensure_ascii=False)


def deserialize_recipients(s: str | None) -> list[dict[str, str]]:
    if not s:
        return []
    try:
        return json.loads(s)
    except json.JSONDecodeError:
        return []
