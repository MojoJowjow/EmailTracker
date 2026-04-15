from __future__ import annotations

import logging
from typing import Any, AsyncIterator
from urllib.parse import quote

import httpx

from .auth import GraphAuth

log = logging.getLogger(__name__)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# Fields we need for metadata tracking. Keep this tight; bodies/attachments excluded.
_SELECT = ",".join(
    [
        "id",
        "conversationId",
        "from",
        "sender",
        "toRecipients",
        "ccRecipients",
        "bccRecipients",
        "subject",
        "receivedDateTime",
        "sentDateTime",
    ]
)


class GraphDeltaGone(Exception):
    """Raised when the stored delta link has expired (Graph returns 410)."""


class GraphClient:
    def __init__(self, auth: GraphAuth, shared_mailbox: str) -> None:
        self._auth = auth
        self._mailbox = shared_mailbox
        self._client = httpx.AsyncClient(timeout=60.0)

    async def aclose(self) -> None:
        await self._client.aclose()

    def _headers(self) -> dict[str, str]:
        return {
            "Authorization": f"Bearer {self._auth.get_access_token()}",
            "Accept": "application/json",
            "Prefer": 'odata.maxpagesize=100, outlook.body-content-type="text"',
        }

    def _initial_delta_url(self, folder: str) -> str:
        mbox = quote(self._mailbox, safe="@")
        return (
            f"{GRAPH_BASE}/users/{mbox}/mailFolders/{folder}/messages/delta"
            f"?$select={_SELECT}"
        )

    async def list_messages_delta(
        self,
        folder: str,
        delta_link: str | None,
    ) -> AsyncIterator[tuple[list[dict[str, Any]], str | None]]:
        """Yield (page_of_messages, next_delta_link_or_None).

        Only the final page's tuple has a non-None delta link. Callers should
        persist that delta link for the next sync run.
        """
        url = delta_link or self._initial_delta_url(folder)
        while True:
            resp = await self._client.get(url, headers=self._headers())
            if resp.status_code == 410:
                raise GraphDeltaGone(f"Delta link expired for folder {folder!r}")
            resp.raise_for_status()
            payload = resp.json()
            messages: list[dict[str, Any]] = payload.get("value", [])
            next_link = payload.get("@odata.nextLink")
            final_delta = payload.get("@odata.deltaLink")
            if next_link:
                yield messages, None
                url = next_link
                continue
            yield messages, final_delta
            return
