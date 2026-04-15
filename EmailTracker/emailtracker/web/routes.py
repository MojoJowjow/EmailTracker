from __future__ import annotations

import json
import sqlite3
from typing import Any

from fastapi import APIRouter, Form, HTTPException, Request
from fastapi.responses import HTMLResponse, JSONResponse, Response

from .. import db

router = APIRouter()


def _conn(request: Request) -> sqlite3.Connection:
    return request.app.state.db


def _templates(request: Request):
    return request.app.state.templates


def _row_to_message(row: sqlite3.Row) -> dict[str, Any]:
    return {
        "id": row["id"],
        "conversation_id": row["conversation_id"],
        "direction": row["direction"],
        "folder": row["folder"],
        "sender_name": row["sender_name"] or "",
        "sender_address": row["sender_address"] or "",
        "to_recipients": db.deserialize_recipients(row["to_recipients"]),
        "cc_recipients": db.deserialize_recipients(row["cc_recipients"]),
        "bcc_recipients": db.deserialize_recipients(row["bcc_recipients"]),
        "subject": row["subject"] or "(no subject)",
        "received_at": row["received_at"],
        "is_reply": bool(row["is_reply"]),
        "is_forward": bool(row["is_forward"]),
    }


def _format_recipients(recipients: list[dict[str, str]], max_shown: int = 3) -> str:
    if not recipients:
        return ""
    labels = [(r.get("address") or r.get("name") or "").strip() for r in recipients]
    labels = [l for l in labels if l]
    if len(labels) <= max_shown:
        return ", ".join(labels)
    return ", ".join(labels[:max_shown]) + f" +{len(labels) - max_shown}"


@router.get("/", response_class=HTMLResponse)
async def dashboard(request: Request) -> Response:
    conn = _conn(request)
    rules = db.list_filter_rules(conn)
    messages = [_row_to_message(r) for r in db.search_messages(conn, rules, None, limit=100)]
    status = db.get_status(conn)
    return _templates(request).TemplateResponse(
        request,
        "dashboard.html",
        {
            "messages": messages,
            "rules": rules,
            "status": status,
            "query": "",
            "format_recipients": _format_recipients,
        },
    )


@router.get("/feed", response_class=HTMLResponse)
async def feed(request: Request, q: str = "", limit: int = 100) -> Response:
    conn = _conn(request)
    rules = db.list_filter_rules(conn)
    messages = [_row_to_message(r) for r in db.search_messages(conn, rules, q or None, limit=limit)]
    return _templates(request).TemplateResponse(
        request,
        "_feed_rows.html",
        {"messages": messages, "format_recipients": _format_recipients},
    )


@router.get("/thread/{conversation_id}", response_class=HTMLResponse)
async def thread(request: Request, conversation_id: str) -> Response:
    conn = _conn(request)
    rows = db.get_conversation(conn, conversation_id)
    if not rows:
        raise HTTPException(status_code=404, detail="Conversation not found")
    messages = [_row_to_message(r) for r in rows]
    return _templates(request).TemplateResponse(
        request,
        "thread.html",
        {
            "conversation_id": conversation_id,
            "messages": messages,
            "format_recipients": _format_recipients,
        },
    )


@router.get("/status")
async def status(request: Request) -> JSONResponse:
    return JSONResponse(db.get_status(_conn(request)))


@router.get("/rules", response_class=HTMLResponse)
async def rules_panel(request: Request) -> Response:
    rules = db.list_filter_rules(_conn(request))
    return _templates(request).TemplateResponse(
        request, "_rules_panel.html", {"rules": rules}
    )


@router.post("/rules", response_class=HTMLResponse)
async def create_rule(
    request: Request,
    name: str = Form(...),
    sender_pattern: str = Form(""),
    subject_pattern: str = Form(""),
) -> Response:
    conn = _conn(request)
    name = name.strip()
    sender = sender_pattern.strip() or None
    subject = subject_pattern.strip() or None
    if not name:
        raise HTTPException(status_code=400, detail="Rule name is required")
    if not sender and not subject:
        raise HTTPException(
            status_code=400,
            detail="At least one of sender_pattern or subject_pattern is required",
        )
    db.create_filter_rule(conn, name, sender, subject)
    rules = db.list_filter_rules(conn)
    return _templates(request).TemplateResponse(
        request, "_rules_panel.html", {"rules": rules}
    )


@router.post("/rules/{rule_id}/toggle", response_class=HTMLResponse)
async def toggle_rule(request: Request, rule_id: int) -> Response:
    conn = _conn(request)
    db.toggle_filter_rule(conn, rule_id)
    rules = db.list_filter_rules(conn)
    return _templates(request).TemplateResponse(
        request, "_rules_panel.html", {"rules": rules}
    )


@router.delete("/rules/{rule_id}", response_class=HTMLResponse)
async def delete_rule(request: Request, rule_id: int) -> Response:
    conn = _conn(request)
    db.delete_filter_rule(conn, rule_id)
    rules = db.list_filter_rules(conn)
    return _templates(request).TemplateResponse(
        request, "_rules_panel.html", {"rules": rules}
    )
