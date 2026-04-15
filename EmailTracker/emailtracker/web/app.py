from __future__ import annotations

import asyncio
import logging
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from .. import db
from ..auth import GraphAuth
from ..config import load_settings
from ..graph_client import GraphClient
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

    auth = GraphAuth(settings)
    graph = GraphClient(auth, settings.shared_mailbox)
    poller = Poller(settings, conn, graph)

    app.state.settings = settings
    app.state.db = conn
    app.state.graph = graph
    app.state.poller = poller
    app.state.templates = TEMPLATES

    task = asyncio.create_task(poller.run_forever(), name="emailtracker-poller")
    log.info("Poller task started")
    try:
        yield
    finally:
        log.info("Shutting down poller...")
        poller.stop()
        try:
            await asyncio.wait_for(task, timeout=10)
        except asyncio.TimeoutError:
            task.cancel()
        await graph.aclose()
        conn.close()
        log.info("Shutdown complete")


def create_app() -> FastAPI:
    app = FastAPI(title="EmailTracker", lifespan=lifespan)
    app.mount("/static", StaticFiles(directory=str(_WEB_DIR / "static")), name="static")
    app.include_router(routes.router)
    return app


app = create_app()
