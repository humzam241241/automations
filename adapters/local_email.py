"""
Adapter: Fetch emails from local files (.eml, .msg, or CSV) - NO GRAPH REQUIRED.
"""
import csv
import email
from email import policy
from email.parser import BytesParser
from pathlib import Path
from typing import Dict, List


class LocalEmailAdapter:
    """Load emails from local sources without Graph API."""
    
    def __init__(self):
        pass
    
    def load_from_eml_files(self, file_paths: List[str]) -> List[Dict]:
        """
        Load emails from .eml files.
        
        Args:
            file_paths: List of .eml file paths
        
        Returns:
            List of normalized email dicts
        """
        emails = []
        for path_str in file_paths:
            path = Path(path_str)
            if not path.exists():
                continue
            
            try:
                with path.open("rb") as f:
                    msg = BytesParser(policy=policy.default).parse(f)
                    emails.append(self._normalize_eml(msg))
            except Exception as e:
                print(f"Failed to parse {path}: {e}")
        
        return emails
    
    def load_from_csv(self, csv_path: str) -> List[Dict]:
        """
        Load emails from CSV export.
        Expected columns: subject, from, to, date, body
        
        Args:
            csv_path: Path to CSV file
        
        Returns:
            List of normalized email dicts
        """
        emails = []
        path = Path(csv_path)
        
        if not path.exists():
            return emails
        
        with path.open("r", encoding="utf-8", errors="ignore") as f:
            reader = csv.DictReader(f)
            for row in reader:
                emails.append({
                    "id": row.get("id", ""),
                    "subject": row.get("subject", row.get("Subject", "")),
                    "from": row.get("from", row.get("From", "")),
                    "to": row.get("to", row.get("To", "")),
                    "date": row.get("date", row.get("Date", "")),
                    "body": row.get("body", row.get("Body", "")),
                    "attachments": self._parse_attachment_list(row.get("attachments", "")),
                    "attachment_text": [],
                })
        
        return emails
    
    def scan_directory(self, directory: str, pattern: str = "*.eml") -> List[Dict]:
        """
        Scan a directory for .eml files and load them.
        
        Args:
            directory: Directory path
            pattern: Glob pattern (default *.eml)
        
        Returns:
            List of normalized email dicts
        """
        dir_path = Path(directory)
        if not dir_path.exists():
            return []
        
        file_paths = [str(p) for p in dir_path.glob(pattern)]
        return self.load_from_eml_files(file_paths)
    
    def _normalize_eml(self, msg) -> Dict:
        """Convert email.message.Message to standard format."""
        subject = msg.get("Subject", "")
        from_addr = msg.get("From", "")
        to_addr = msg.get("To", "")
        date_str = msg.get("Date", "")
        
        # Extract body
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    body += part.get_content()
        else:
            body = msg.get_content()
        
        # Extract attachment names (basic)
        attachments = []
        for part in msg.walk():
            filename = part.get_filename()
            if filename:
                attachments.append(filename)
        
        return {
            "id": msg.get("Message-ID", ""),
            "subject": subject,
            "from": from_addr,
            "to": to_addr,
            "date": date_str,
            "body": body,
            "attachments": attachments,
            "attachment_text": [],
        }
    
    def _parse_attachment_list(self, attachments_str: str) -> List[str]:
        """Parse semicolon-separated attachment list."""
        if not attachments_str:
            return []
        return [a.strip() for a in attachments_str.split(";") if a.strip()]
