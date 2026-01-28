import json
import logging
from datetime import datetime
from typing import Dict, List

import azure.functions as func

from shared.config_loader import load_app_settings, load_keyword_map
from shared.email_processor import (
    apply_keyword_map,
    attachment_text,
    extract_message_id,
    strip_html,
)
from shared.excel_builder import build_workbook
from shared.graph_client import GraphClient


def _build_notes(message: Dict, attachment_names: List[str], max_len: int) -> str:
    subject = message.get("subject", "")
    sender = (
        message.get("from", {})
        .get("emailAddress", {})
        .get("address", "")
    )
    body = message.get("body", {}).get("content", "")
    body_text = strip_html(body)
    notes = f"Subject: {subject}\nFrom: {sender}\nAttachments: {', '.join(attachment_names)}\nBody: {body_text}"
    return notes[:max_len]


def _upload_excel(graph: GraphClient, settings: Dict, file_bytes: bytes, filename: str) -> None:
    onedrive_path = settings["onedrive_path"].strip("/")
    mailbox_user = settings["mailbox_user"]
    upload_url = (
        f"https://graph.microsoft.com/v1.0/users/{mailbox_user}/drive/root:"
        f"/{onedrive_path}/{filename}:/content"
    )
    graph.put(upload_url, file_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def _process_notification(notification: Dict, settings: Dict, keyword_map: Dict) -> None:
    resource = notification.get("resource", "")
    message_id = extract_message_id(resource)
    if not message_id:
        logging.warning("No message id found in resource: %s", resource)
        return

    graph = GraphClient(
        tenant_id=settings["tenant_id"],
        client_id=settings["client_id"],
        client_secret=settings["client_secret"],
    )

    mailbox_user = settings["mailbox_user"]
    message_url = (
        f"https://graph.microsoft.com/v1.0/users/{mailbox_user}/messages/{message_id}"
        "?$select=subject,body,from,receivedDateTime"
    )
    message = graph.get(message_url)

    attachments_url = (
        f"https://graph.microsoft.com/v1.0/users/{mailbox_user}/messages/{message_id}/attachments"
    )
    attachments = graph.get(attachments_url).get("value", [])
    attachment_names, attachment_content = attachment_text(attachments)

    subject = message.get("subject", "")
    body = strip_html(message.get("body", {}).get("content", ""))
    combined_content = f"{subject} {body} {attachment_content}"

    notes = _build_notes(
        message,
        attachment_names,
        settings.get("max_notes_length", 500),
    )

    headers, row = apply_keyword_map(
        combined_content, keyword_map, attachment_names, notes
    )
    file_bytes = build_workbook(headers, row)

    timestamp = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
    safe_id = message_id.replace("/", "_")
    filename = f"email_{timestamp}_{safe_id}.xlsx"
    _upload_excel(graph, settings, file_bytes, filename)


def main(req: func.HttpRequest) -> func.HttpResponse:
    logging.info("Notifications endpoint invoked.")

    validation_token = req.params.get("validationToken")
    if validation_token:
        return func.HttpResponse(validation_token, status_code=200)

    try:
        payload = req.get_json()
    except ValueError:
        return func.HttpResponse("Invalid JSON", status_code=400)

    settings = load_app_settings()
    keyword_map = load_keyword_map()

    client_state_expected = settings.get("client_state")
    notifications = payload.get("value", [])

    for notification in notifications:
        client_state = notification.get("clientState")
        if client_state_expected and client_state != client_state_expected:
            logging.warning("Skipping notification due to clientState mismatch.")
            continue
        try:
            _process_notification(notification, settings, keyword_map)
        except Exception as exc:
            logging.exception("Failed to process notification: %s", exc)

    return func.HttpResponse(json.dumps({"status": "accepted"}), status_code=202)
