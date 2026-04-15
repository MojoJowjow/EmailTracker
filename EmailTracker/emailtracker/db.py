from __future__ import annotations

import json
import sqlite3
from contextlib import contextmanager
from datetime import datetime, timezone
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
    ingested_at      TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_messages_conversation ON messages(conversation_id);
CREATE INDEX IF NOT EXISTS idx_messages_received_at ON messages(received_at DESC);
CREATE INDEX IF NOT EXISTS idx_messages_sender ON messages(sender_address);

CREATE TABLE IF NOT EXISTS sync_state (
    folder           TEXT PRIMARY KEY,
    delta_link       TEXT,
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
            subject, received_at, is_reply, is_forward, ingested_at
        ) VALUES (
            :id, :conversation_id, :direction, :folder,
            :sender_name, :sender_address,
            :to_recipients, :cc_recipients, :bcc_recipients,
            :subject, :received_at, :is_reply, :is_forward, :ingested_at
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
            is_forward      = excluded.is_forward
        """,
        row,
    )


def get_sync_state(conn: sqlite3.Connection, folder: str) -> sqlite3.Row | None:
    return conn.execute(
        "SELECT folder, delta_link, last_success_at, last_error FROM sync_state WHERE folder = ?",
        (folder,),
    ).fetchone()


def save_sync_success(conn: sqlite3.Connection, folder: str, delta_link: str | None) -> None:
    conn.execute(
        """
        INSERT INTO sync_state (folder, delta_link, last_success_at, last_error)
        VALUES (?, ?, ?, NULL)
        ON CONFLICT(folder) DO UPDATE SET
            delta_link      = excluded.delta_link,
            last_success_at = excluded.last_success_at,
            last_error      = NULL
        """,
        (folder, delta_link, now_utc_iso()),
    )


def save_sync_error(conn: sqlite3.Connection, folder: str, error: str) -> None:
    conn.execute(
        """
        INSERT INTO sync_state (folder, delta_link, last_success_at, last_error)
        VALUES (?, NULL, NULL, ?)
        ON CONFLICT(folder) DO UPDATE SET last_error = excluded.last_error
        """,
        (folder, error),
    )


def clear_delta_link(conn: sqlite3.Connection, folder: str) -> None:
    conn.execute("UPDATE sync_state SET delta_link = NULL WHERE folder = ?", (folder,))


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
) -> list[sqlite3.Row]:
    filter_sql, filter_params = build_filter_clause(rules)
    where_parts = [filter_sql]
    params: list[Any] = list(filter_params)

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
        "is_reply, is_forward, ingested_at "
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
            "is_reply, is_forward, ingested_at "
            "FROM messages WHERE conversation_id = ? "
            "ORDER BY received_at ASC",
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
