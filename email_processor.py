import json
import logging
import time
from pathlib import Path
from typing import Dict, List, Tuple

import re
from dateutil import parser as date_parser

import requests

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def _request_with_retry(method: str, url: str, headers: Dict[str, str], params: Dict = None, json_body=None) -> Dict:
    attempts = 3
    delay = 2
    for attempt in range(1, attempts + 1):
        try:
            response = requests.request(
                method,
                url,
                headers=headers,
                params=params,
                json=json_body,
                timeout=30,
            )
            response.raise_for_status()
            if response.text:
                return response.json()
            return {}
        except requests.RequestException as exc:
            if attempt == attempts:
                raise
            logging.warning("Network error (%s). Retrying in %ss...", exc, delay)
            time.sleep(delay)
            delay *= 2
    return {}


def load_config(path: Path = Path("config.json")) -> Dict:
    return json.loads(path.read_text(encoding="utf-8"))


def load_processed_ids(path: Path = Path("processed_emails.json")) -> List[str]:
    if not path.exists():
        return []
    data = json.loads(path.read_text(encoding="utf-8"))
    return data.get("processed_ids", [])


def save_processed_ids(ids: List[str], path: Path = Path("processed_emails.json")) -> None:
    path.write_text(json.dumps({"processed_ids": sorted(set(ids))}, indent=2), encoding="utf-8")


def get_folder_id(token: str, folder_name: str) -> str:
    headers = {"Authorization": f"Bearer {token}"}
    url = f"{GRAPH_BASE}/me/mailFolders"
    data = _request_with_retry("GET", url, headers)
    for folder in data.get("value", []):
        if folder.get("displayName", "").lower() == folder_name.lower():
            return folder.get("id")
    raise RuntimeError(f"Folder '{folder_name}' not found.")


def fetch_messages(token: str, folder_id: str, top: int = 25) -> List[Dict]:
    headers = {"Authorization": f"Bearer {token}"}
    url = f"{GRAPH_BASE}/me/mailFolders/{folder_id}/messages"
    params = {
        "$select": "id,subject,from,receivedDateTime,body,hasAttachments,conversationId",
        "$orderby": "receivedDateTime desc",
        "$top": top,
    }
    data = _request_with_retry("GET", url, headers, params=params)
    return data.get("value", [])


def fetch_attachments(token: str, message_id: str) -> List[str]:
    headers = {"Authorization": f"Bearer {token}"}
    url = f"{GRAPH_BASE}/me/messages/{message_id}/attachments"
    data = _request_with_retry("GET", url, headers)
    return [item.get("name", "") for item in data.get("value", []) if item.get("name")]


def normalize_keywords(keywords: List[str]) -> List[str]:
    normalized: List[str] = []
    for keyword in keywords:
        if not isinstance(keyword, str):
            continue
        parts = [part.strip() for part in keyword.split(",") if part.strip()]
        cleaned_parts = []
        for part in parts or [keyword.strip()]:
            cleaned = part.strip().strip("\"'").strip()
            if cleaned:
                cleaned_parts.append(cleaned)
        normalized.extend(cleaned_parts)
    return [kw for kw in normalized if kw]


def extract_keywords(content: str, keywords: List[str]) -> List[str]:
    content_lower = content.lower()
    return [keyword for keyword in keywords if keyword.lower() in content_lower]


def _safe_parse_date(value: str):
    try:
        return date_parser.parse(value)
    except (ValueError, TypeError):
        return None


def _detect_miscommunication(content: str, participants: List[str]) -> str:
    if len(participants) < 2:
        return "No"
    signals = [
        "misunderstand",
        "confusing",
        "clarify",
        "not clear",
        "incorrect",
        "no update",
        "didn't receive",
        "did not receive",
        "missing",
        "issue",
    ]
    content_lower = content.lower()
    return "Yes" if any(signal in content_lower for signal in signals) else "No"


def _strip_html(text: str) -> str:
    return re.sub(r"<[^>]+>", " ", text or "")


def _build_thread_text(messages: List[Dict]) -> str:
    parts = []
    for msg in messages:
        sender = (
            msg.get("from", {})
            .get("emailAddress", {})
            .get("address", "")
        )
        subject = msg.get("subject", "")
        body = _strip_html(msg.get("body", {}).get("content", ""))
        parts.append(f"From: {sender} | Subject: {subject} | Body: {body}")
    return "\n".join(parts).strip()


def build_thread_row(messages: List[Dict], attachments: List[str], keywords: List[str]) -> Tuple[List[str], List[str], str]:
    sorted_messages = sorted(
        messages,
        key=lambda msg: _safe_parse_date(msg.get("receivedDateTime")) or 0,
        reverse=True,
    )
    latest = sorted_messages[0]
    subject = latest.get("subject", "")
    sender = latest.get("from", {}).get("emailAddress", {}).get("address", "")
    received = latest.get("receivedDateTime", "")

    participants = sorted(
        {
            msg.get("from", {})
            .get("emailAddress", {})
            .get("address", "")
            for msg in sorted_messages
            if msg.get("from")
        }
    )

    content_parts = []
    for msg in sorted_messages:
        content_parts.append(msg.get("subject", ""))
        content_parts.append(msg.get("body", {}).get("content", ""))
    combined_content = " ".join(content_parts)

    normalized_keywords = normalize_keywords(keywords)
    keyword_hits = extract_keywords(combined_content, normalized_keywords)
    keyword_hit_set = set(keyword_hits)

    miscommunication = _detect_miscommunication(combined_content, participants)
    summary = _build_thread_text(sorted_messages)

    dynamic_headers = normalized_keywords
    headers = [
        "Subject",
        "From",
        "Received",
        "Attachments",
        "Participants",
        "Message_Count",
    ] + dynamic_headers + ["Miscommunication", "Summary"]

    row = [
        subject,
        sender,
        received,
        "; ".join(attachments),
        ", ".join(participants),
        str(len(sorted_messages)),
    ]
    for header in dynamic_headers:
        row.append("Yes" if header in keyword_hit_set else "")
    row.append(miscommunication)
    row.append(summary)
    return headers, row, summary
