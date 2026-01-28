import uuid
from pathlib import Path
import sys

sys.path.append(str(Path(__file__).resolve().parents[1]))

from shared.config_loader import load_app_settings
from shared.graph_client import GraphClient
from shared.subscription import create_folder_subscription


def main() -> None:
    settings = load_app_settings()
    graph = GraphClient(
        tenant_id=settings["tenant_id"],
        client_id=settings["client_id"],
        client_secret=settings["client_secret"],
    )
    notification_url = settings["notification_url"]
    client_state = settings.get("client_state") or str(uuid.uuid4())
    subscription = create_folder_subscription(
        graph=graph,
        mailbox_user=settings["mailbox_user"],
        folder_id=settings["target_folder_id"],
        notification_url=notification_url,
        client_state=client_state,
    )
    print(subscription)


if __name__ == "__main__":
    main()
