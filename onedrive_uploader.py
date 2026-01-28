import logging
import time

import requests


def _ensure_folder(token: str, folder_path: str) -> None:
    base_path = folder_path.strip("/")
    if not base_path:
        return

    headers = {"Authorization": f"Bearer {token}"}
    parts = base_path.split("/")
    current_path = ""
    for part in parts:
        current_path = f"{current_path}/{part}" if current_path else part
        check_url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{current_path}"
        response = requests.get(check_url, headers=headers, timeout=30)
        if response.status_code == 200:
            continue
        if response.status_code != 404:
            response.raise_for_status()

        parent_path = "/".join(current_path.split("/")[:-1])
        parent_url = "https://graph.microsoft.com/v1.0/me/drive/root"
        if parent_path:
            parent_url = f"{parent_url}:/{parent_path}"
        create_url = f"{parent_url}:/children"
        payload = {
            "name": part,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "replace",
        }
        create_response = requests.post(
            create_url,
            headers={**headers, "Content-Type": "application/json"},
            json=payload,
            timeout=30,
        )
        create_response.raise_for_status()
        logging.info("Created OneDrive folder: %s", current_path)


def upload_file(token: str, folder_path: str, filename: str, content: bytes) -> dict:
    headers = {"Authorization": f"Bearer {token}"}
    base_path = folder_path.strip("/")
    _ensure_folder(token, base_path)
    if base_path:
        url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{base_path}/{filename}:/content"
    else:
        url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{filename}:/content"

    attempts = 3
    delay = 2
    for attempt in range(1, attempts + 1):
        try:
            response = requests.put(
                url,
                headers={**headers, "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
                data=content,
                timeout=60,
            )
            response.raise_for_status()
            if response.text:
                return response.json()
            return {}
        except requests.RequestException as exc:
            if attempt == attempts:
                raise
            logging.warning("Upload failed (%s). Retrying in %ss...", exc, delay)
            time.sleep(delay)
            delay *= 2
