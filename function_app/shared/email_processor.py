import base64
import re
from typing import Dict, List, Tuple


def strip_html(html: str) -> str:
    return re.sub(r"<[^>]+>", " ", html or "").replace("&nbsp;", " ")


def normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip().lower()


def extract_message_id(resource: str) -> str:
    if not resource:
        return ""
    match = re.search(r"messages\\('([^']+)'\\)", resource)
    if match:
        return match.group(1)
    if "/messages/" in resource:
        return resource.rsplit("/messages/", 1)[-1].split("?", 1)[0]
    return resource.rsplit("/", 1)[-1]


def attachment_text(attachments: List[Dict]) -> Tuple[List[str], str]:
    names = []
    text_blobs = []
    for attachment in attachments:
        name = attachment.get("name")
        if name:
            names.append(name)
        content_type = (attachment.get("contentType") or "").lower()
        content_bytes = attachment.get("contentBytes")
        if not content_bytes:
            continue
        if content_type.startswith("text/") or name.lower().endswith(".txt"):
            try:
                decoded = base64.b64decode(content_bytes).decode("utf-8", errors="ignore")
                text_blobs.append(decoded)
            except Exception:
                continue
    return names, " ".join(text_blobs)


def apply_keyword_map(
    content: str,
    keyword_map: Dict,
    attachment_names: List[str],
    notes_text: str,
) -> Tuple[List[str], List[str]]:
    columns = keyword_map.get("columns", [])
    headers = [col.get("header") for col in columns]
    row = []
    normalized = normalize_text(content)
    for column in columns:
        keywords = [kw.lower() for kw in column.get("keywords", [])]
        matched = any(kw in normalized for kw in keywords)
        value = column.get("value", "TRUE") if matched else column.get("unmatched_value", "")
        row.append(value)

    attachment_column = keyword_map.get("attachment_column")
    if attachment_column and attachment_column not in headers:
        headers.append(attachment_column)
        row.append("; ".join(attachment_names))

    notes_column = keyword_map.get("notes_column")
    if notes_column and notes_column not in headers:
        headers.append(notes_column)
        row.append(notes_text)

    return headers, row
