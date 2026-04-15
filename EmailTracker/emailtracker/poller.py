from __future__ import annotations

import asyncio
import logging
import sqlite3
from typing import Any

from . import db
from .config import Settings
from .graph_client import GraphClient, GraphDeltaGone

log = logging.getLogger(__name__)

# (folder_key, graph_well_known_folder_name, direction)
FOLDERS: list[tuple[str, str, str]] = [
    ("inbox", "inbox", "in"),
    ("sentitems", "sentitems", "out"),
]

_INITIAL_BACKOFF = 5.0
_MAX_BACKOFF = 300.0


def _map_message(raw: dict[str, Any], folder_key: str, direction: str) -> dict[str, Any]:
    from_field = raw.get("from") or raw.get("sender") or {}
    email = (from_field or {}).get("emailAddress") or {}
    subject = raw.get("subject") or ""
    received_at = raw.get("receivedDateTime") or raw.get("sentDateTime")
    subject_lower = subject.lower()
    return {
        "id": raw["id"],
        "conversation_id": raw.get("conversationId"),
        "direction": direction,
        "folder": folder_key,
        "sender_name": email.get("name") or "",
        "sender_address": email.get("address") or "",
        "to_recipients": db.serialize_recipients(raw.get("toRecipients")),
        "cc_recipients": db.serialize_recipients(raw.get("ccRecipients")),
        "bcc_recipients": db.serialize_recipients(raw.get("bccRecipients")),
        "subject": subject,
        "received_at": received_at,
        "is_reply": 1 if subject_lower.startswith("re:") else 0,
        "is_forward": 1 if subject_lower.startswith(("fw:", "fwd:")) else 0,
        "ingested_at": db.now_utc_iso(),
    }


class Poller:
    def __init__(
        self,
        settings: Settings,
        conn: sqlite3.Connection,
        graph: GraphClient,
    ) -> None:
        self._settings = settings
        self._conn = conn
        self._graph = graph
        self._stop = asyncio.Event()
        self._backoff = _INITIAL_BACKOFF

    async def run_forever(self) -> None:
        log.info("Poller starting (interval=%ss)", self._settings.poll_interval_seconds)
        while not self._stop.is_set():
            try:
                await self._tick()
                self._backoff = _INITIAL_BACKOFF
                await self._wait(self._settings.poll_interval_seconds)
            except asyncio.CancelledError:
                break
            except Exception as exc:  # noqa: BLE001
                log.exception("Poller tick failed")
                for folder_key, _, _ in FOLDERS:
                    try:
                        db.save_sync_error(self._conn, folder_key, f"{type(exc).__name__}: {exc}")
                    except Exception:  # noqa: BLE001
                        log.exception("Failed to write sync error for %s", folder_key)
                await self._wait(self._backoff)
                self._backoff = min(self._backoff * 2, _MAX_BACKOFF)

    async def _wait(self, seconds: float) -> None:
        try:
            await asyncio.wait_for(self._stop.wait(), timeout=seconds)
        except asyncio.TimeoutError:
            pass

    def stop(self) -> None:
        self._stop.set()

    async def _tick(self) -> None:
        for folder_key, graph_folder, direction in FOLDERS:
            await self._sync_folder(folder_key, graph_folder, direction)

    async def _sync_folder(self, folder_key: str, graph_folder: str, direction: str) -> None:
        state = db.get_sync_state(self._conn, folder_key)
        delta_link = state["delta_link"] if state else None

        try:
            final_delta: str | None = None
            total = 0
            async for page, page_delta in self._graph.list_messages_delta(graph_folder, delta_link):
                if page:
                    with db.transaction(self._conn):
                        for raw in page:
                            if raw.get("@removed"):
                                continue
                            row = _map_message(raw, folder_key, direction)
                            db.upsert_message(self._conn, row)
                            total += 1
                if page_delta:
                    final_delta = page_delta
            db.save_sync_success(self._conn, folder_key, final_delta)
            if total:
                log.info("[%s] upserted %d messages", folder_key, total)
        except GraphDeltaGone:
            log.warning("[%s] delta link expired; forcing full resync next tick", folder_key)
            db.clear_delta_link(self._conn, folder_key)
            db.save_sync_error(self._conn, folder_key, "delta link expired; will full-resync")
            raise
