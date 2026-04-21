from __future__ import annotations

import logging
import sqlite3
import threading
from datetime import datetime, timedelta, timezone

from . import db
from .config import Settings
from .outlook_reader import OutlookError, OutlookReader

log = logging.getLogger(__name__)

FOLDERS: list[tuple[str, str]] = [
    ("inbox", "in"),
    ("sentitems", "out"),
]

_INITIAL_BACKOFF = 5.0
_MAX_BACKOFF = 300.0


def _parse_iso(s: str | None) -> datetime | None:
    if not s:
        return None
    try:
        return datetime.fromisoformat(s)
    except ValueError:
        return None


class Poller:
    """Threaded poller that reads from Outlook via COM and writes to SQLite.

    COM must be initialized on the same thread that calls Outlook, so the
    poller runs in a dedicated daemon thread rather than an asyncio task.
    """

    def __init__(self, settings: Settings, conn: sqlite3.Connection) -> None:
        self._settings = settings
        self._conn = conn
        self._stop = threading.Event()
        self._thread: threading.Thread | None = None

    def start(self) -> None:
        if self._thread is not None:
            return
        self._thread = threading.Thread(
            target=self._run, name="emailtracker-poller", daemon=True
        )
        self._thread.start()

    def stop(self, timeout: float = 10.0) -> None:
        self._stop.set()
        t = self._thread
        if t is not None:
            t.join(timeout=timeout)

    def _run(self) -> None:
        log.info(
            "Poller thread starting (interval=%ss, mailbox=%s)",
            self._settings.poll_interval_seconds,
            self._settings.shared_mailbox,
        )
        reader = OutlookReader(self._settings.shared_mailbox)
        try:
            reader.connect()
        except Exception as exc:  # noqa: BLE001
            log.exception("Failed to connect to Outlook")
            msg = f"{type(exc).__name__}: {exc}"
            for folder_key, _ in FOLDERS:
                try:
                    db.save_sync_error(self._conn, folder_key, msg)
                except Exception:  # noqa: BLE001
                    log.exception("Failed to persist connect error")
            reader.disconnect()
            return

        backoff = _INITIAL_BACKOFF
        try:
            while not self._stop.is_set():
                try:
                    self._tick(reader)
                    backoff = _INITIAL_BACKOFF
                    if self._stop.wait(timeout=self._settings.poll_interval_seconds):
                        break
                except OutlookError as exc:
                    log.exception("Outlook error during tick")
                    self._record_error(f"{type(exc).__name__}: {exc}")
                    if self._stop.wait(timeout=backoff):
                        break
                    backoff = min(backoff * 2, _MAX_BACKOFF)
                except Exception as exc:  # noqa: BLE001
                    log.exception("Unexpected error during tick")
                    self._record_error(f"{type(exc).__name__}: {exc}")
                    if self._stop.wait(timeout=backoff):
                        break
                    backoff = min(backoff * 2, _MAX_BACKOFF)
        finally:
            reader.disconnect()
            log.info("Poller thread exiting")

    def _record_error(self, msg: str) -> None:
        for folder_key, _ in FOLDERS:
            try:
                db.save_sync_error(self._conn, folder_key, msg)
            except Exception:  # noqa: BLE001
                log.exception("Failed to persist error for %s", folder_key)

    def _tick(self, reader: OutlookReader) -> None:
        initial_cutoff = datetime.now(timezone.utc) - timedelta(
            days=self._settings.initial_sync_days
        )
        for folder_key, direction in FOLDERS:
            state = db.get_sync_state(self._conn, folder_key)
            since = _parse_iso(state["watermark"] if state else None) or initial_cutoff

            max_seen = since
            count = 0
            with db.transaction(self._conn):
                for row in reader.iter_since(folder_key, since, direction):
                    db.upsert_message(self._conn, row)
                    count += 1
                    r_dt = _parse_iso(row.get("received_at"))
                    if r_dt and r_dt > max_seen:
                        max_seen = r_dt
            db.save_sync_success(self._conn, folder_key, max_seen.isoformat())
            if count:
                log.info("[%s] upserted %d messages (new watermark=%s)", folder_key, count, max_seen.isoformat())
