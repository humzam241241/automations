from typing import Any, Dict, Optional

import msal
import requests


class GraphClient:
    def __init__(self, tenant_id: str, client_id: str, client_secret: str) -> None:
        self._tenant_id = tenant_id
        self._client_id = client_id
        self._client_secret = client_secret
        self._authority = f"https://login.microsoftonline.com/{tenant_id}"
        self._app = msal.ConfidentialClientApplication(
            client_id=client_id,
            client_credential=client_secret,
            authority=self._authority,
        )

    def get_access_token(self) -> str:
        result = self._app.acquire_token_silent(
            scopes=["https://graph.microsoft.com/.default"],
            account=None,
        )
        if not result:
            result = self._app.acquire_token_for_client(
                scopes=["https://graph.microsoft.com/.default"]
            )
        if "access_token" not in result:
            raise RuntimeError(
                f"Failed to acquire token: {result.get('error_description')}"
            )
        return result["access_token"]

    def _headers(self) -> Dict[str, str]:
        return {"Authorization": f"Bearer {self.get_access_token()}"}

    def get(self, url: str, params: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        response = requests.get(url, headers=self._headers(), params=params, timeout=30)
        response.raise_for_status()
        return response.json()

    def post(self, url: str, payload: Dict[str, Any]) -> Dict[str, Any]:
        response = requests.post(
            url,
            headers={**self._headers(), "Content-Type": "application/json"},
            json=payload,
            timeout=30,
        )
        response.raise_for_status()
        if response.text:
            return response.json()
        return {}

    def put(self, url: str, data: bytes, content_type: str) -> None:
        response = requests.put(
            url,
            headers={**self._headers(), "Content-Type": content_type},
            data=data,
            timeout=60,
        )
        response.raise_for_status()
