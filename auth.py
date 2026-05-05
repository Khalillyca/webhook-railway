"""
auth.py — MSAL authentication for personal @outlook.com
On Railway: reads token cache from TOKEN_CACHE_JSON env var (base64 encoded)
Locally: reads from .token_cache.bin file
"""

import os
import base64
import logging
import threading
import msal

log = logging.getLogger(__name__)

CLIENT_ID = "c2e7c427-f51e-4f54-bcde-392d7b34cdde"
TOKEN_CACHE_FILE = ".token_cache.bin"

SCOPES = [
    "Mail.ReadWrite",
    "Mail.Send",
    "Mail.Read",
    "User.Read",
]

_token_lock = threading.Lock()


def _get_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()

    # First try env var (Railway deployment)
    token_cache_b64 = os.getenv("TOKEN_CACHE_B64")
    if token_cache_b64:
        try:
            cache_json = base64.b64decode(token_cache_b64).decode("utf-8")
            cache.deserialize(cache_json)
            log.info("[AUTH] Token cache loaded from TOKEN_CACHE_B64 env var.")
            return cache
        except Exception as e:
            log.warning(f"[AUTH] Failed to load TOKEN_CACHE_B64: {e}")

    # Fall back to local file
    if os.path.exists(TOKEN_CACHE_FILE):
        with open(TOKEN_CACHE_FILE, "r") as f:
            cache.deserialize(f.read())
        log.debug("[AUTH] Token cache loaded from local file.")

    return cache


def _save_cache(cache: msal.SerializableTokenCache):
    if cache.has_state_changed:
        # Only save to file locally (not on Railway)
        if not os.getenv("TOKEN_CACHE_B64"):
            with open(TOKEN_CACHE_FILE, "w") as f:
                f.write(cache.serialize())
            log.debug("[AUTH] Token cache saved to disk.")


def get_token() -> str:
    with _token_lock:
        cache = _get_cache()
        app = msal.PublicClientApplication(
            client_id=CLIENT_ID,
            authority="https://login.microsoftonline.com/consumers",
            token_cache=cache,
        )

        accounts = app.get_accounts()
        result = None

        if accounts:
            result = app.acquire_token_silent(SCOPES, account=accounts[0])
            if result and "access_token" in result:
                log.debug("[AUTH] Silent token refresh successful.")

        if not result or "access_token" not in result:
            raise RuntimeError(
                "[AUTH] No cached token available. "
                "Run locally first to generate token cache, "
                "then set TOKEN_CACHE_B64 env var on Railway."
            )

        _save_cache(cache)
        log.info("[AUTH] Access token acquired successfully.")
        return result["access_token"]


def get_my_email(token: str) -> str:
    import requests
    r = requests.get(
        "https://graph.microsoft.com/v1.0/me",
        headers={"Authorization": f"Bearer {token}"},
    )
    r.raise_for_status()
    data = r.json()
    email = data.get("mail") or data.get("userPrincipalName", "")
    log.info(f"[AUTH] Logged in as: {email}")
    return email
