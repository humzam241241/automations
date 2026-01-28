import logging
import sys
from datetime import datetime

from auth import acquire_token
from email_processor import (
    build_thread_row,
    fetch_attachments,
    fetch_messages,
    get_folder_id,
    load_config,
    load_processed_ids,
    save_processed_ids,
)
from excel_generator import fill_template, generate_excel
from onedrive_uploader import upload_file


SCOPES = ["Mail.Read", "Files.ReadWrite"]


def setup_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )


def main() -> None:
    setup_logging()
    config = load_config()
    folder_name = config.get("outlook_folder", "auto")
    onedrive_folder = config.get("onedrive_folder", "/EmailReports")
    keyword_map = config.get("keywords", [])
    template_path = config.get("template_path", "").strip()

    logging.info("Starting Outlook monitor")
    logging.info("Folder: %s | OneDrive: %s", folder_name, onedrive_folder)
    logging.info("Keywords: %s", ", ".join(keyword_map))

    processed_ids = set(load_processed_ids())
    processed_count = 0

    try:
        token = acquire_token(SCOPES)
    except Exception as exc:
        logging.error("Authentication failed: %s", exc)
        sys.exit(1)

    try:
        folder_id = get_folder_id(token, folder_name)
        messages = fetch_messages(token, folder_id)
    except Exception as exc:
        logging.error("Failed to fetch emails: %s", exc)
        sys.exit(1)

    try:
        logging.info("Fetched %d messages from folder '%s'.", len(messages), folder_name)
        threads = {}
        for message in messages:
            conversation_id = message.get("conversationId") or message.get("id")
            if conversation_id:
                threads.setdefault(conversation_id, []).append(message)

        for conversation_id, thread_messages in threads.items():
            if conversation_id in processed_ids:
                continue

            latest_message = sorted(
                thread_messages,
                key=lambda msg: msg.get("receivedDateTime", ""),
                reverse=True,
            )[0]
            attachments = []
            for msg in thread_messages:
                if msg.get("hasAttachments") and msg.get("id"):
                    attachments.extend(fetch_attachments(token, msg["id"]))

            headers, row, summary = build_thread_row(thread_messages, attachments, keyword_map)
            if template_path:
                file_bytes = fill_template(template_path, headers, row, summary)
                extension = "xlsx"
            else:
                file_bytes = generate_excel(headers, row)
                extension = "xlsx"

            timestamp = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
            filename = f"EmailThread_{timestamp}.{extension}"
            upload_result = upload_file(token, onedrive_folder, filename, file_bytes)
            if upload_result:
                logging.info(
                    "Uploaded to OneDrive: %s",
                    upload_result.get("webUrl", "upload complete"),
                )
            else:
                logging.info("Uploaded to OneDrive: %s", filename)

            processed_ids.add(conversation_id)
            processed_count += 1
            logging.info("Processed thread %s | total=%d", conversation_id, processed_count)
    except KeyboardInterrupt:
        logging.info("Shutting down gracefully.")
    finally:
        save_processed_ids(list(processed_ids))


if __name__ == "__main__":
    main()
