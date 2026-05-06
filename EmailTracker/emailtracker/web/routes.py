from __future__ import annotations

import json
import sqlite3
from datetime import datetime
from typing import Any

from fastapi import APIRouter, Form, HTTPException, Request
from fastapi.responses import HTMLResponse, JSONResponse, Response

from .. import db

router = APIRouter()


def _conn(request: Request) -> sqlite3.Connection:
    return request.app.state.db


def _templates(request: Request):
    return request.app.state.templates


def _format_dt(iso_str: str | None) -> str:
    """Convert ISO-8601 timestamp to dd/mm/yy HH:MM format."""
    if not iso_str:
        return ""
    try:
        dt = datetime.fromisoformat(iso_str)
        return dt.strftime("%d/%m/%y %H:%M")
    except (ValueError, TypeError):
        return iso_str


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
        "received_at": _format_dt(row["received_at"]),
        "received_at_raw": row["received_at"] or "",
        "is_reply": bool(row["is_reply"]),
        "is_forward": bool(row["is_forward"]),
        "has_attachments": bool(row["has_attachments"]),
        "requires_reply": bool(row["requires_reply"]),
        "body_preview": row["body_preview"] or "",
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
async def dashboard(request: Request, days: int = 0, date: str = "") -> Response:
    conn = _conn(request)
    rules = db.list_filter_rules(conn)
    messages = [
        _row_to_message(r)
        for r in db.search_messages(
            conn, rules, None, limit=100,
            days=days if days > 0 else None,
            date=date or None,
        )
    ]
    _mark_awaiting_reply(conn, messages)
    status = db.get_status(conn)
    status["last_sync"] = _format_dt(status.get("last_sync"))
    metrics = db.get_metrics(conn, days if days > 0 else None, rules=rules, date=date or None)
    return _templates(request).TemplateResponse(
        request,
        "dashboard.html",
        {
            "messages": messages,
            "rules": rules,
            "status": status,
            "metrics": metrics,
            "days": days,
            "date": date,
            "query": "",
            "format_recipients": _format_recipients,
        },
    )


@router.get("/metrics", response_class=HTMLResponse)
async def metrics_panel(request: Request, days: int = 0, date: str = "") -> Response:
    conn = _conn(request)
    rules = db.list_filter_rules(conn)
    metrics = db.get_metrics(conn, days if days > 0 else None, rules=rules, date=date or None)
    return _templates(request).TemplateResponse(
        request,
        "_metrics.html",
        {"metrics": metrics, "days": days, "date": date},
    )



def _mark_awaiting_reply(conn: sqlite3.Connection, messages: list[dict[str, Any]]) -> None:
    """Set 'awaiting_reply' on messages that require a reply from FSJ but haven't been replied to yet."""
    needs = [m for m in messages if m["requires_reply"] and m["direction"] == "in"]
    if not needs:
        for m in messages:
            m["awaiting_reply"] = False
        return
    conv_ids = list({m["conversation_id"] for m in needs})
    placeholders = ",".join("?" * len(conv_ids))
    rows = conn.execute(
        f"SELECT DISTINCT conversation_id FROM messages "
        f"WHERE conversation_id IN ({placeholders}) AND direction = 'out' AND is_reply = 1",
        conv_ids,
    ).fetchall()
    replied_convs = {r["conversation_id"] for r in rows}
    for m in messages:
        if m["requires_reply"] and m["direction"] == "in":
            m["awaiting_reply"] = m["conversation_id"] not in replied_convs
        else:
            m["awaiting_reply"] = False


@router.get("/feed", response_class=HTMLResponse)
async def feed(
    request: Request, q: str = "", limit: int = 100,
    filter_days: int = 0, filter_date: str = "",
) -> Response:
    conn = _conn(request)
    rules = db.list_filter_rules(conn)
    messages = [
        _row_to_message(r)
        for r in db.search_messages(
            conn, rules, q or None, limit=limit,
            days=filter_days if filter_days > 0 else None,
            date=filter_date or None,
        )
    ]
    _mark_awaiting_reply(conn, messages)
    return _templates(request).TemplateResponse(
        request,
        "_feed_rows.html",
        {"messages": messages, "format_recipients": _format_recipients},
    )


def _annotate_reply_status(messages: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Add a 'status' field to each message in a thread.

    For inbound messages:
      - 'Replied'        if a later outbound RE: exists in the thread
      - 'Forwarded'      if a later outbound FW:/FWD: exists (and no reply)
      - 'Awaiting Reply' if no outbound follows it
    For outbound messages:
      - 'Reply Sent'     if it's a reply (RE:)
      - 'Forwarded'      if it's a forward (FW:)
      - 'Sent'           otherwise
    """
    # Collect outbound messages sorted by time
    outbound = [m for m in messages if m["direction"] == "out"]

    for m in messages:
        if m["direction"] == "out":
            if m["is_reply"]:
                m["status"] = "Reply Sent"
                m["status_class"] = "status-replied"
            elif m["is_forward"]:
                m["status"] = "Forwarded"
                m["status_class"] = "status-forwarded"
            else:
                m["status"] = "Sent"
                m["status_class"] = "status-sent"
        else:
            # Inbound: check if any later outbound exists
            has_reply = any(
                o for o in outbound
                if o["is_reply"] and (o["received_at_raw"] or "") > (m["received_at_raw"] or "")
            )
            has_forward = any(
                o for o in outbound
                if o["is_forward"] and (o["received_at_raw"] or "") > (m["received_at_raw"] or "")
            )
            if has_reply:
                m["status"] = "Replied"
                m["status_class"] = "status-replied"
            elif has_forward:
                m["status"] = "Forwarded"
                m["status_class"] = "status-forwarded"
            else:
                m["status"] = "Awaiting Reply"
                m["status_class"] = "status-awaiting"
    return messages


@router.get("/thread/{conversation_id}", response_class=HTMLResponse)
async def thread(request: Request, conversation_id: str) -> Response:
    conn = _conn(request)
    rows = db.get_conversation(conn, conversation_id)
    if not rows:
        raise HTTPException(status_code=404, detail="Conversation not found")
    messages = _annotate_reply_status([_row_to_message(r) for r in rows])
    _mark_awaiting_reply(conn, messages)
    # Thread subject from the first message
    thread_subject = messages[0]["subject"] if messages else ""
    return _templates(request).TemplateResponse(
        request,
        "thread.html",
        {
            "conversation_id": conversation_id,
            "thread_subject": thread_subject,
            "messages": messages,
            "format_recipients": _format_recipients,
        },
    )


@router.get("/status")
async def status(request: Request) -> JSONResponse:
    s = db.get_status(_conn(request))
    s["last_sync"] = _format_dt(s.get("last_sync"))
    return JSONResponse(s)


@router.get("/rules", response_class=HTMLResponse)
async def rules_panel(request: Request) -> Response:
    rules = db.list_filter_rules(_conn(request))
    return _templates(request).TemplateResponse(
        request, "_rules_panel.html", {"rules": rules, "error": None}
    )


@router.post("/rules", response_class=HTMLResponse)
async def create_rule(
    request: Request,
    name: str = Form(""),
    sender_pattern: str = Form(""),
    subject_pattern: str = Form(""),
) -> Response:
    conn = _conn(request)
    name = name.strip()
    sender = sender_pattern.strip() or None
    subject = subject_pattern.strip() or None
    error = None
    if not name:
        error = "Rule name is required."
    elif not sender and not subject:
        error = "Fill in at least one of sender or subject pattern."
    if error:
        rules = db.list_filter_rules(conn)
        return _templates(request).TemplateResponse(
            request, "_rules_panel.html", {"rules": rules, "error": error}
        )
    db.create_filter_rule(conn, name, sender, subject)
    rules = db.list_filter_rules(conn)
    return _templates(request).TemplateResponse(
        request, "_rules_panel.html", {"rules": rules, "error": None}
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
