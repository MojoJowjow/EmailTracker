from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv

_PROJECT_ROOT = Path(__file__).resolve().parent.parent


@dataclass(frozen=True)
class Settings:
    shared_mailbox: str
    poll_interval_seconds: int
    initial_sync_days: int
    db_path: Path
    web_host: str
    web_port: int


def _require(name: str) -> str:
    value = os.environ.get(name, "").strip()
    if not value:
        raise RuntimeError(
            f"Missing required environment variable: {name}. "
            f"Copy .env.example to .env and fill it in."
        )
    return value


def _resolve(path_str: str) -> Path:
    p = Path(path_str)
    return p if p.is_absolute() else (_PROJECT_ROOT / p).resolve()


def load_settings() -> Settings:
    load_dotenv(_PROJECT_ROOT / ".env")
    return Settings(
        shared_mailbox=_require("SHARED_MAILBOX"),
        poll_interval_seconds=int(os.environ.get("POLL_INTERVAL_SECONDS", "60")),
        initial_sync_days=int(os.environ.get("INITIAL_SYNC_DAYS", "30")),
        db_path=_resolve(os.environ.get("DB_PATH", "./emailtracker.db")),
        web_host=os.environ.get("WEB_HOST", "127.0.0.1"),
        web_port=int(os.environ.get("WEB_PORT", "8000")),
    )
