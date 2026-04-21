from __future__ import annotations

import logging
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from .. import db
from ..config import load_settings
from ..poller import Poller
from . import routes

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(name)s: %(message)s",
)
log = logging.getLogger("emailtracker")


_WEB_DIR = Path(__file__).resolve().parent
TEMPLATES = Jinja2Templates(directory=str(_WEB_DIR / "templates"))


@asynccontextmanager
async def lifespan(app: FastAPI):
    settings = load_settings()
    conn = db.connect(settings.db_path)
    db.init_schema(conn)

    poller = Poller(settings, conn)

    app.state.settings = settings
    app.state.db = conn
    app.state.poller = poller
    app.state.templates = TEMPLATES

    poller.start()
    log.info("Poller thread started (mailbox=%s)", settings.shared_mailbox)
    try:
        yield
    finally:
        log.info("Shutting down poller...")
        poller.stop(timeout=15)
        conn.close()
        log.info("Shutdown complete")


def create_app() -> FastAPI:
    app = FastAPI(title="EmailTracker", lifespan=lifespan)
    app.mount("/static", StaticFiles(directory=str(_WEB_DIR / "static")), name="static")
    app.include_router(routes.router)
    return app


app = create_app()
