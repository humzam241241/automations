"""
Permissions diagnostic tool - test Microsoft Graph permissions.
"""
import os
import sys
from pathlib import Path

import msal
import requests


def acquire_token() -> str:
    """Acquire token using device code flow."""
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
    
    scopes = ["https://graph.microsoft.com/.default"]
    
    # Try silent
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes=scopes, account=accounts[0])
        if result and "access_token" in result:
            if cache.has_state_changed:
                cache_path.write_text(cache.serialize(), encoding="utf-8")
            return result["access_token"]
    
    # Device code
    flow = app.initiate_device_flow(scopes=scopes)
    if "user_code" not in flow:
        raise RuntimeError("Failed to start device flow")
    
    print(flow["message"])
    result = app.acquire_token_by_device_flow(flow)
    
    if not result or "access_token" not in result:
        raise RuntimeError(f"Auth failed: {result.get('error_description') if result else 'Unknown'}")
    
    if cache.has_state_changed:
        cache_path.write_text(cache.serialize(), encoding="utf-8")
    
    return result["access_token"]


def test_permission(token: str, url: str, name: str, required_permission: str) -> dict:
    """Test a specific Graph API endpoint."""
    try:
        response = requests.get(
            url,
            headers={"Authorization": f"Bearer {token}"},
            timeout=15,
        )
        
        if response.status_code == 200:
            return {
                "name": name,
                "status": "✓ PASSED",
                "permission": required_permission,
                "message": "Access granted",
            }
        elif response.status_code == 403:
            return {
                "name": name,
                "status": "✗ DENIED",
                "permission": required_permission,
                "message": "Permission denied. Add this permission in Azure AD app registration.",
            }
        elif response.status_code == 401:
            return {
                "name": name,
                "status": "✗ UNAUTHORIZED",
                "permission": required_permission,
                "message": "Token invalid or expired.",
            }
        else:
            return {
                "name": name,
                "status": f"✗ ERROR ({response.status_code})",
                "permission": required_permission,
                "message": response.text[:200],
            }
    except Exception as e:
        return {
            "name": name,
            "status": "✗ EXCEPTION",
            "permission": required_permission,
            "message": str(e),
        }


def main():
    """Run diagnostics."""
    print("="*70)
    print("MICROSOFT GRAPH PERMISSIONS DIAGNOSTIC")
    print("="*70)
    
    print("\n1. Acquiring access token...")
    try:
        token = acquire_token()
        print("   ✓ Token acquired")
    except Exception as e:
        print(f"   ✗ Failed to acquire token: {e}")
        sys.exit(1)
    
    print("\n2. Testing Graph API permissions...\n")
    
    # Define tests
    tests = [
        {
            "url": "https://graph.microsoft.com/v1.0/me",
            "name": "Read user profile (/me)",
            "permission": "User.Read",
        },
        {
            "url": "https://graph.microsoft.com/v1.0/me/mailFolders",
            "name": "List mail folders",
            "permission": "Mail.Read",
        },
        {
            "url": "https://graph.microsoft.com/v1.0/me/messages?$top=1",
            "name": "Read messages",
            "permission": "Mail.Read",
        },
        {
            "url": "https://graph.microsoft.com/v1.0/me/drive",
            "name": "Access OneDrive",
            "permission": "Files.ReadWrite",
        },
        {
            "url": "https://graph.microsoft.com/v1.0/me/drive/root/children",
            "name": "List OneDrive files",
            "permission": "Files.ReadWrite",
        },
    ]
    
    results = []
    for test in tests:
        result = test_permission(token, test["url"], test["name"], test["permission"])
        results.append(result)
        
        status_color = result["status"]
        print(f"{status_color:20} {result['name']:30} ({result['permission']})")
        if "DENIED" in result["status"] or "ERROR" in result["status"] or "EXCEPTION" in result["status"]:
            print(f"{'':20} → {result['message']}")
        print()
    
    # Summary
    print("="*70)
    print("SUMMARY")
    print("="*70)
    
    passed = sum(1 for r in results if "PASSED" in r["status"])
    total = len(results)
    
    print(f"\nPassed: {passed}/{total}")
    
    failed = [r for r in results if "PASSED" not in r["status"]]
    if failed:
        print("\nMissing or denied permissions:")
        for r in failed:
            print(f"  - {r['permission']:20} (required for: {r['name']})")
        
        print("\nTO FIX:")
        print("  1. Go to Azure Portal → App Registrations → Your App")
        print("  2. Navigate to 'API permissions'")
        print("  3. Add the missing permissions listed above")
        print("  4. Click 'Grant admin consent' if required by your organization")
        print("  5. Re-run this diagnostic")
    else:
        print("\n✓ All permissions are correctly configured!")
    
    print("\n" + "="*70)


if __name__ == "__main__":
    main()
