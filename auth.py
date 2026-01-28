import json
import os
from pathlib import Path
from typing import List, Optional

import msal


CLIENT_ID = os.getenv("CLIENT_ID", "d3590ed6-52b3-4102-aeff-aad2292ab01c")
AUTHORITY = os.getenv("AUTHORITY", "https://login.microsoftonline.com/aca3c8d6-aa71-4e1a-a10e-03572fc58c0b")
CACHE_PATH = Path("token_cache.bin")


class AuthError(RuntimeError):
    pass


def _load_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    if CACHE_PATH.exists():
        cache.deserialize(CACHE_PATH.read_text(encoding="utf-8"))
    return cache


def _save_cache(cache: msal.SerializableTokenCache) -> None:
    if cache.has_state_changed:
        CACHE_PATH.write_text(cache.serialize(), encoding="utf-8")


def acquire_token(scopes: List[str]) -> str:
    cache = _load_cache()
    app = msal.PublicClientApplication(
        client_id=CLIENT_ID,
        authority=AUTHORITY,
        token_cache=cache,
    )

    accounts = app.get_accounts()
    result: Optional[dict] = None
    if accounts:
        result = app.acquire_token_silent(scopes=scopes, account=accounts[0])

    if not result:
        flow = app.initiate_device_flow(scopes=scopes)
        if "user_code" not in flow:
            raise AuthError(f"Failed to start device flow: {flow}")
        print(flow["message"])
        result = app.acquire_token_by_device_flow(flow)

    _save_cache(cache)

    if not result or "access_token" not in result:
        raise AuthError(f"Authentication failed: {result.get('error_description') if result else 'Unknown'}")
    return result["access_token"]
