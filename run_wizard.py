"""
Offline wizard: Interactive profile creation and execution.
Detects Graph availability and offers local fallback.
"""
import json
import logging
import sys
from pathlib import Path
from typing import Optional

import msal

from core.engine import ExecutionEngine
from core.profile_loader import ProfileLoader


def setup_logging():
    """Configure logging."""
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )


def test_graph_access(token: str) -> bool:
    """Test if Graph API is accessible."""
    import requests
    
    try:
        response = requests.get(
            "https://graph.microsoft.com/v1.0/me",
            headers={"Authorization": f"Bearer {token}"},
            timeout=10,
        )
        return response.status_code == 200
    except Exception:
        return False


def acquire_graph_token() -> Optional[str]:
    """Attempt to acquire Graph token via device code flow."""
    import os
    
    client_id = os.getenv("CLIENT_ID", "d3590ed6-52b3-4102-aeff-aad2292ab01c")
    authority = os.getenv("AUTHORITY", "https://login.microsoftonline.com/aca3c8d6-aa71-4e1a-a10e-03572fc58c0b")
    
    cache_path = Path("token_cache.bin")
    cache = msal.SerializableTokenCache()
    if cache_path.exists():
        cache.deserialize(cache_path.read_text(encoding="utf-8"))
    
    app = msal.PublicClientApplication(
        client_id=client_id,
        authority=authority,
        token_cache=cache,
    )
    
    scopes = ["Mail.Read", "Files.ReadWrite"]
    
    # Try silent first
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes=scopes, account=accounts[0])
        if result and "access_token" in result:
            if cache.has_state_changed:
                cache_path.write_text(cache.serialize(), encoding="utf-8")
            return result["access_token"]
    
    # Device code flow
    print("\n" + "="*60)
    print("ATTEMPTING MICROSOFT GRAPH AUTHENTICATION")
    print("="*60)
    
    try:
        flow = app.initiate_device_flow(scopes=scopes)
        if "user_code" not in flow:
            print("Failed to start device flow.")
            return None
        
        print(flow["message"])
        result = app.acquire_token_by_device_flow(flow)
        
        if result and "access_token" in result:
            if cache.has_state_changed:
                cache_path.write_text(cache.serialize(), encoding="utf-8")
            return result["access_token"]
        else:
            print(f"Authentication failed: {result.get('error_description') if result else 'Unknown'}")
            return None
    except Exception as e:
        print(f"Authentication error: {e}")
        return None


def wizard_create_profile() -> dict:
    """Interactive wizard to create a profile."""
    print("\n" + "="*60)
    print("CREATE NEW PROFILE")
    print("="*60)
    
    profile = {}
    
    # Profile name
    profile["name"] = input("Profile name: ").strip()
    profile["description"] = input("Description (optional): ").strip()
    
    # Input source
    print("\nSelect input source:")
    print("  1. Microsoft Graph (requires Graph permissions)")
    print("  2. Local .eml files (no Graph required)")
    print("  3. Local CSV file (no Graph required)")
    
    source_choice = input("Choice [1-3]: ").strip()
    
    if source_choice == "1":
        profile["input_source"] = "graph"
        folder_name = input("Outlook folder name (e.g., Inbox): ").strip()
        top = input("Number of newest messages [25]: ").strip() or "25"
        profile["email_selection"] = {
            "folder_name": folder_name,
            "newest_n": int(top),
            "search_query": None,
        }
    elif source_choice == "2":
        profile["input_source"] = "local_eml"
        directory = input("Directory path [./input_emails]: ").strip() or "./input_emails"
        pattern = input("File pattern [*.eml]: ").strip() or "*.eml"
        profile["email_selection"] = {
            "directory": directory,
            "pattern": pattern,
        }
    elif source_choice == "3":
        profile["input_source"] = "local_csv"
        csv_path = input("CSV file path [./emails.csv]: ").strip() or "./emails.csv"
        profile["email_selection"] = {
            "csv_path": csv_path,
        }
    else:
        print("Invalid choice. Defaulting to local .eml.")
        profile["input_source"] = "local_eml"
        profile["email_selection"] = {"directory": "./input_emails", "pattern": "*.eml"}
    
    # Schema
    print("\nDefine schema columns (comma-separated, e.g., Subject,From,Date,Billing,IT):")
    columns_input = input("Columns: ").strip()
    columns = [c.strip() for c in columns_input.split(",") if c.strip()]
    
    profile["schema"] = {
        "columns": [{"name": col, "type": "text"} for col in columns]
    }
    
    # Rules
    print("\nDefine keyword rules (leave blank to skip):")
    rules = []
    while True:
        col_name = input("Column name (or blank to finish): ").strip()
        if not col_name:
            break
        
        keywords_input = input(f"  Keywords for '{col_name}' (comma-separated): ").strip()
        keywords = [kw.strip() for kw in keywords_input.split(",") if kw.strip()]
        
        value = input(f"  Value if matched [Yes]: ").strip() or "Yes"
        
        rules.append({
            "column": col_name,
            "keywords": keywords,
            "value": value,
            "unmatched_value": "",
        })
    
    profile["rules"] = rules
    
    # Output
    print("\nSelect output format:")
    print("  1. Excel")
    print("  2. CSV")
    format_choice = input("Choice [1-2]: ").strip()
    output_format = "excel" if format_choice == "1" else "csv"
    
    print("\nSelect output destination:")
    print("  1. OneDrive (requires Graph permissions)")
    print("  2. Local file")
    dest_choice = input("Choice [1-2]: ").strip()
    
    if dest_choice == "1":
        destination = "onedrive"
        onedrive_path = input("OneDrive folder path [/EmailReports]: ").strip() or "/EmailReports"
        profile["output"] = {
            "format": output_format,
            "destination": destination,
            "onedrive_path": onedrive_path,
            "filename_template": f"{profile['name']}_{{timestamp}}.{output_format if output_format == 'csv' else 'xlsx'}",
        }
    else:
        destination = "local"
        local_path = input("Local output directory [./output]: ").strip() or "./output"
        profile["output"] = {
            "format": output_format,
            "destination": destination,
            "local_path": local_path,
            "filename_template": f"{profile['name']}_{{timestamp}}.{output_format if output_format == 'csv' else 'xlsx'}",
        }
    
    return profile


def main():
    """Main wizard entry point."""
    setup_logging()
    
    print("="*60)
    print("EMAIL TO SPREADSHEET WIZARD")
    print("="*60)
    
    # Try to get Graph token
    print("\nChecking Microsoft Graph access...")
    token = acquire_graph_token()
    
    graph_available = False
    if token:
        graph_available = test_graph_access(token)
        if graph_available:
            print("✓ Microsoft Graph access confirmed")
        else:
            print("✗ Microsoft Graph access denied (permissions issue)")
            token = None
    else:
        print("✗ Microsoft Graph authentication failed")
    
    if not graph_available:
        print("\n" + "="*60)
        print("NO-GRAPH MODE")
        print("You can still use local .eml files or CSV inputs.")
        print("="*60)
    
    # Load existing profiles
    loader = ProfileLoader()
    existing_profiles = loader.list_profiles()
    
    if existing_profiles:
        print(f"\nExisting profiles: {', '.join(existing_profiles)}")
    
    print("\nOptions:")
    print("  1. Run existing profile")
    print("  2. Create new profile")
    print("  3. Exit")
    
    choice = input("\nChoice [1-3]: ").strip()
    
    if choice == "1":
        profile_name = input("Profile name: ").strip()
        profile = loader.load_profile(profile_name)
        
        if not profile:
            print(f"Profile '{profile_name}' not found.")
            sys.exit(1)
    
    elif choice == "2":
        profile = wizard_create_profile()
        
        # Validate
        errors = loader.validate_profile(profile)
        if errors:
            print("\nProfile validation errors:")
            for error in errors:
                print(f"  - {error}")
            sys.exit(1)
        
        # Save
        loader.save_profile(profile)
        print(f"\n✓ Profile '{profile['name']}' saved to profiles/{profile['name']}.json")
    
    else:
        print("Exiting.")
        sys.exit(0)
    
    # Check if profile requires Graph
    requires_graph = (
        profile.get("input_source") == "graph" or
        profile.get("output", {}).get("destination") == "onedrive"
    )
    
    if requires_graph and not token:
        print("\nERROR: This profile requires Microsoft Graph access, but authentication failed.")
        print("Please use a profile with local input/output, or resolve Graph permissions.")
        sys.exit(1)
    
    # Execute profile
    print(f"\nExecuting profile: {profile.get('name')}")
    engine = ExecutionEngine(access_token=token)
    
    try:
        result = engine.run_profile(profile, explain=True)
        print("\n" + "="*60)
        print("EXECUTION COMPLETE")
        print("="*60)
        print(json.dumps(result, indent=2))
    except Exception as e:
        logging.exception("Execution failed")
        print(f"\nERROR: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
