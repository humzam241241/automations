import json
import os
from pathlib import Path


CONFIG_PATH = Path("config.json")


def load_config() -> dict:
    if CONFIG_PATH.exists():
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    return {
        "keywords": {},
        "outlook_folder": "auto",
        "onedrive_folder": "/EmailReports",
        "check_interval": 120,
    }


def main() -> None:
    config = load_config()
    print("Configure folders and interval (press Enter to keep current value).")
    current_folder = config.get("outlook_folder", "auto")
    folder_input = input(f"Outlook folder [{current_folder}]: ").strip()
    if folder_input:
        config["outlook_folder"] = folder_input

    onedrive_env = os.getenv("ONEDRIVE_FOLDER", "").strip()
    if onedrive_env:
        config["onedrive_folder"] = onedrive_env
    else:
        current_onedrive = config.get("onedrive_folder", "/EmailReports")
        while True:
            onedrive_input = input(f"OneDrive folder [{current_onedrive}]: ").strip()
            if not onedrive_input:
                break
            if not onedrive_input.startswith("/"):
                print("Invalid OneDrive path. Use format like /EmailReports or /Folder/SubFolder.")
                continue
            config["onedrive_folder"] = onedrive_input
            break

    template_env = os.getenv("TEMPLATE_PATH", "").strip()
    if template_env:
        config["template_path"] = template_env
    else:
        current_template = config.get("template_path", "")
        while True:
            template_input = input(f"Template Excel path [{current_template}]: ").strip()
            if not template_input:
                break
            if not Path(template_input).exists():
                print("Invalid path. File not found. Please enter a valid .xlsx path.")
                continue
            config["template_path"] = template_input
            break

    current_interval = str(config.get("check_interval", 120))
    interval_input = input(f"Check interval seconds [{current_interval}]: ").strip()
    if interval_input:
        try:
            config["check_interval"] = int(interval_input)
        except ValueError:
            print("Invalid interval. Keeping existing value.")

    print("Enter keywords. You can input multiple separated by commas.")
    print("Leave blank to finish.")
    keywords = []
    while True:
        keyword_input = input("Keyword(s): ").strip()
        if not keyword_input:
            break
        parts = [part.strip() for part in keyword_input.split(",") if part.strip()]
        keywords.extend(parts or [keyword_input])

    if not keywords:
        print("No keywords entered. Keeping existing config.")
        return

    config["keywords"] = keywords
    CONFIG_PATH.write_text(json.dumps(config, indent=2), encoding="utf-8")
    print(f"Saved {len(keywords)} keywords to {CONFIG_PATH}.")


if __name__ == "__main__":
    main()
