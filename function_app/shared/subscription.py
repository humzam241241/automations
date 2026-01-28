from datetime import datetime, timedelta, timezone
from typing import Dict

from .graph_client import GraphClient


def create_folder_subscription(
    graph: GraphClient,
    mailbox_user: str,
    folder_id: str,
    notification_url: str,
    client_state: str,
    minutes_valid: int = 60 * 24,
) -> Dict:
    expiration = datetime.now(timezone.utc) + timedelta(minutes=minutes_valid)
    payload = {
        "changeType": "created",
        "notificationUrl": notification_url,
        "resource": f"/users/{mailbox_user}/mailFolders/{folder_id}/messages",
        "expirationDateTime": expiration.isoformat(),
        "clientState": client_state,
    }
    return graph.post("https://graph.microsoft.com/v1.0/subscriptions", payload)
