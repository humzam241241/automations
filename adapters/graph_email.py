"""
Adapter: Fetch emails via Microsoft Graph API.
"""
import base64
import re
from typing import Dict, List, Optional

import requests


class GraphEmailAdapter:
    """Fetch emails using Microsoft Graph."""
    
    def __init__(self, access_token: str):
        self.access_token = access_token
        self.base_url = "https://graph.microsoft.com/v1.0"
    
    def _headers(self) -> Dict[str, str]:
        return {"Authorization": f"Bearer {self.access_token}"}
    
    def list_folders(self) -> List[Dict]:
        """List all mail folders."""
        url = f"{self.base_url}/me/mailFolders"
        response = requests.get(url, headers=self._headers(), timeout=30)
        response.raise_for_status()
        return response.json().get("value", [])
    
    def get_folder_id(self, folder_name: str) -> Optional[str]:
        """Get folder ID by display name."""
        folders = self.list_folders()
        for folder in folders:
            if folder.get("displayName", "").lower() == folder_name.lower():
                return folder.get("id")
        return None
    
    def fetch_messages(
        self,
        folder_id: str = None,
        search_query: str = None,
        top: int = 25
    ) -> List[Dict]:
        """
        Fetch messages from a folder or via search.
        
        Args:
            folder_id: Mail folder ID (uses inbox if None)
            search_query: Optional search query
            top: Max messages to fetch
        
        Returns:
            List of message dicts with standard fields
        """
        if folder_id:
            url = f"{self.base_url}/me/mailFolders/{folder_id}/messages"
        else:
            url = f"{self.base_url}/me/messages"
        
        params = {
            "$select": "id,subject,from,toRecipients,receivedDateTime,body,hasAttachments",
            "$orderby": "receivedDateTime desc",
            "$top": top,
        }
        
        if search_query:
            params["$search"] = search_query
        
        response = requests.get(url, headers=self._headers(), params=params, timeout=30)
        response.raise_for_status()
        messages = response.json().get("value", [])
        
        # Normalize to standard format
        return [self._normalize_message(msg) for msg in messages]
    
    def _normalize_message(self, msg: Dict) -> Dict:
        """Convert Graph message to standard format."""
        from_addr = msg.get("from", {}).get("emailAddress", {}).get("address", "")
        to_addrs = [
            r.get("emailAddress", {}).get("address", "")
            for r in msg.get("toRecipients", [])
        ]
        
        body_html = msg.get("body", {}).get("content", "")
        body_text = self._strip_html(body_html)
        
        attachments = []
        attachment_text = []
        if msg.get("hasAttachments"):
            attachments, attachment_text = self._fetch_attachments(msg["id"])
        
        return {
            "id": msg.get("id", ""),
            "subject": msg.get("subject", ""),
            "from": from_addr,
            "to": ", ".join(to_addrs),
            "date": msg.get("receivedDateTime", ""),
            "body": body_text,
            "attachments": attachments,
            "attachment_text": attachment_text,
        }
    
    def _fetch_attachments(self, message_id: str) -> tuple[List[str], List[str]]:
        """Fetch attachment names and extract text from text attachments."""
        url = f"{self.base_url}/me/messages/{message_id}/attachments"
        response = requests.get(url, headers=self._headers(), timeout=30)
        response.raise_for_status()
        
        attachments_data = response.json().get("value", [])
        names = []
        text_blobs = []
        
        for att in attachments_data:
            name = att.get("name", "")
            if name:
                names.append(name)
            
            content_type = (att.get("contentType") or "").lower()
            content_bytes = att.get("contentBytes")
            
            if content_bytes and (content_type.startswith("text/") or name.lower().endswith(".txt")):
                try:
                    decoded = base64.b64decode(content_bytes).decode("utf-8", errors="ignore")
                    text_blobs.append(decoded)
                except Exception:
                    pass
        
        return names, text_blobs
    
    def _strip_html(self, html: str) -> str:
        """Strip HTML tags."""
        return re.sub(r"<[^>]+>", " ", html or "").replace("&nbsp;", " ").strip()
