import json
import os
from pathlib import Path


def _repo_root() -> Path:
    return Path(__file__).resolve().parents[2]


def load_json(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def load_app_settings() -> dict:
    """Load app settings, preferring environment variables over file config."""
    config_path = Path(
        (Path.cwd() / "config" / "app_settings.json").resolve()
    )
    if not config_path.exists():
        config_path = _repo_root() / "config" / "app_settings.json"
    
    # Load from file if it exists
    settings = {}
    if config_path.exists():
        settings = load_json(config_path)
    
    # Override with environment variables (production-safe)
    if os.getenv("TENANT_ID"):
        settings["tenant_id"] = os.getenv("TENANT_ID")
    if os.getenv("CLIENT_ID"):
        settings["client_id"] = os.getenv("CLIENT_ID")
    if os.getenv("CLIENT_SECRET"):
        settings["client_secret"] = os.getenv("CLIENT_SECRET")
    if os.getenv("MAILBOX_USER"):
        settings["mailbox_user"] = os.getenv("MAILBOX_USER")
    if os.getenv("TARGET_FOLDER_ID"):
        settings["target_folder_id"] = os.getenv("TARGET_FOLDER_ID")
    if os.getenv("ONEDRIVE_PATH"):
        settings["onedrive_path"] = os.getenv("ONEDRIVE_PATH")
    if os.getenv("NOTIFICATION_URL"):
        settings["notification_url"] = os.getenv("NOTIFICATION_URL")
    if os.getenv("CLIENT_STATE"):
        settings["client_state"] = os.getenv("CLIENT_STATE")
    
    return settings


def load_keyword_map() -> dict:
    map_path = Path((Path.cwd() / "config" / "keyword_map.json").resolve())
    if not map_path.exists():
        map_path = _repo_root() / "config" / "keyword_map.json"
    return load_json(map_path)
