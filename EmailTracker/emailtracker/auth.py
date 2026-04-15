from __future__ import annotations

import atexit
import logging
import threading
from pathlib import Path

import msal

from .config import Settings

log = logging.getLogger(__name__)

GRAPH_SCOPES = ["Mail.Read", "Mail.Read.Shared"]


class GraphAuth:
    """MSAL device-code flow with a persisted token cache.

    First call triggers device-code sign-in (prints URL + code). After that,
    MSAL refreshes silently as long as the refresh token remains valid.
    """

    def __init__(self, settings: Settings) -> None:
        self._settings = settings
        self._lock = threading.Lock()
        self._cache = msal.SerializableTokenCache()
        self._cache_path: Path = settings.token_cache_path
        if self._cache_path.exists():
            self._cache.deserialize(self._cache_path.read_text(encoding="utf-8"))
        atexit.register(self._persist_cache)

        authority = f"https://login.microsoftonline.com/{settings.tenant_id}"
        self._app = msal.PublicClientApplication(
            client_id=settings.client_id,
            authority=authority,
            token_cache=self._cache,
        )

    def _persist_cache(self) -> None:
        if self._cache.has_state_changed:
            self._cache_path.parent.mkdir(parents=True, exist_ok=True)
            self._cache_path.write_text(self._cache.serialize(), encoding="utf-8")

    def get_access_token(self) -> str:
        with self._lock:
            result = None
            accounts = self._app.get_accounts()
            if accounts:
                result = self._app.acquire_token_silent(GRAPH_SCOPES, account=accounts[0])

            if not result:
                flow = self._app.initiate_device_flow(scopes=GRAPH_SCOPES)
                if "user_code" not in flow:
                    raise RuntimeError(
                        f"Failed to start device-code flow: {flow.get('error_description') or flow}"
                    )
                print("\n" + "=" * 70)
                print("EmailTracker: sign-in required")
                print(flow["message"])
                print("=" * 70 + "\n", flush=True)
                result = self._app.acquire_token_by_device_flow(flow)

            if not result or "access_token" not in result:
                err = (result or {}).get("error_description") or str(result)
                raise RuntimeError(f"Failed to acquire Graph access token: {err}")

            self._persist_cache()
            return result["access_token"]
