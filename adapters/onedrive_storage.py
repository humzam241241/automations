"""
Adapter: Upload/download files to OneDrive via Microsoft Graph.
"""
import logging
from typing import Optional

import requests


class OneDriveAdapter:
    """Upload and download files to OneDrive."""
    
    def __init__(self, access_token: str):
        self.access_token = access_token
        self.base_url = "https://graph.microsoft.com/v1.0/me/drive"
    
    def _headers(self) -> dict:
        return {"Authorization": f"Bearer {self.access_token}"}
    
    def ensure_folder(self, folder_path: str) -> None:
        """
        Ensure a folder exists, creating it if necessary.
        
        Args:
            folder_path: OneDrive path like "/EmailReports" or "/Folder/SubFolder"
        """
        base_path = folder_path.strip("/")
        if not base_path:
            return
        
        parts = base_path.split("/")
        current_path = ""
        
        for part in parts:
            current_path = f"{current_path}/{part}" if current_path else part
            check_url = f"{self.base_url}/root:/{current_path}"
            
            response = requests.get(check_url, headers=self._headers(), timeout=30)
            if response.status_code == 200:
                continue
            
            if response.status_code != 404:
                response.raise_for_status()
            
            # Create folder
            parent_path = "/".join(current_path.split("/")[:-1])
            parent_url = f"{self.base_url}/root"
            if parent_path:
                parent_url = f"{parent_url}:/{parent_path}"
            
            create_url = f"{parent_url}:/children"
            payload = {
                "name": part,
                "folder": {},
                "@microsoft.graph.conflictBehavior": "replace",
            }
            
            create_response = requests.post(
                create_url,
                headers={**self._headers(), "Content-Type": "application/json"},
                json=payload,
                timeout=30,
            )
            create_response.raise_for_status()
            logging.info(f"Created OneDrive folder: {current_path}")
    
    def upload_file(
        self,
        folder_path: str,
        filename: str,
        content: bytes,
        content_type: str = "application/octet-stream"
    ) -> Optional[dict]:
        """
        Upload a file to OneDrive.
        
        Args:
            folder_path: Target folder (e.g., "/EmailReports")
            filename: File name
            content: File bytes
            content_type: MIME type
        
        Returns:
            Upload response dict with webUrl, or None on failure
        """
        self.ensure_folder(folder_path)
        
        base_path = folder_path.strip("/")
        if base_path:
            url = f"{self.base_url}/root:/{base_path}/{filename}:/content"
        else:
            url = f"{self.base_url}/root:/{filename}:/content"
        
        response = requests.put(
            url,
            headers={**self._headers(), "Content-Type": content_type},
            data=content,
            timeout=60,
        )
        response.raise_for_status()
        
        if response.text:
            return response.json()
        return {}
    
    def download_file(self, file_path: str) -> bytes:
        """
        Download a file from OneDrive.
        
        Args:
            file_path: Full path like "/EmailReports/file.xlsx"
        
        Returns:
            File bytes
        """
        url = f"{self.base_url}/root:/{file_path.strip('/')}:/content"
        response = requests.get(url, headers=self._headers(), timeout=60)
        response.raise_for_status()
        return response.content
