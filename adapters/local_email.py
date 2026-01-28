"""
Adapter: Fetch emails from local files (.eml, .msg, PDF, DOCX, CSV) - NO GRAPH REQUIRED.
Supports extracting text from email attachments!
"""
import csv
import email
from email import policy
from email.parser import BytesParser
from pathlib import Path
from typing import Dict, List

# Import attachment reader
try:
    from .attachment_reader import (
        extract_text_from_file, 
        extract_attachment_from_bytes,
        extract_excel_rows_from_bytes
    )
except ImportError:
    from adapters.attachment_reader import (
        extract_text_from_file, 
        extract_attachment_from_bytes,
        extract_excel_rows_from_bytes
    )


class LocalEmailAdapter:
    """Load emails from local sources without Graph API."""
    
    def __init__(self, extract_attachments: bool = True):
        """
        Initialize adapter.
        
        Args:
            extract_attachments: If True, extract text from PDF/DOCX attachments
        """
        self.extract_attachments = extract_attachments
    
    def load_from_files(self, file_paths: List[str]) -> List[Dict]:
        """
        Load data from multiple files of mixed types.
        
        Supports: .eml, .pdf, .docx, .txt, .csv
        
        Args:
            file_paths: List of file paths
        
        Returns:
            List of normalized data dicts
        """
        results = []
        
        for path_str in file_paths:
            path = Path(path_str)
            if not path.exists():
                print(f"File not found: {path_str}")
                continue
            
            ext = path.suffix.lower()
            
            try:
                if ext == '.eml':
                    # Parse as email
                    with path.open("rb") as f:
                        msg = BytesParser(policy=policy.default).parse(f)
                        results.append(self._normalize_eml(msg, path_str))
                
                elif ext == '.pdf':
                    # Extract PDF text as a "document"
                    text = extract_text_from_file(path_str)
                    results.append(self._create_document_record(path_str, text, 'PDF'))
                
                elif ext == '.docx':
                    # Extract Word doc text
                    text = extract_text_from_file(path_str)
                    results.append(self._create_document_record(path_str, text, 'Word'))
                
                elif ext in ['.txt', '.log', '.md']:
                    # Plain text file
                    text = extract_text_from_file(path_str)
                    results.append(self._create_document_record(path_str, text, 'Text'))
                
                elif ext == '.csv':
                    # CSV file - load all rows
                    results.extend(self.load_from_csv(path_str))
                
                else:
                    print(f"Unsupported file type: {ext}")
            
            except Exception as e:
                print(f"Failed to process {path}: {e}")
        
        return results
    
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
            
            ext = path.suffix.lower()
            
            try:
                if ext == '.eml':
                    with path.open("rb") as f:
                        msg = BytesParser(policy=policy.default).parse(f)
                        emails.append(self._normalize_eml(msg, path_str))
                elif ext in ['.pdf', '.docx', '.txt']:
                    # Also accept document files
                    text = extract_text_from_file(path_str)
                    emails.append(self._create_document_record(path_str, text, ext.upper()))
            except Exception as e:
                print(f"Failed to parse {path}: {e}")
        
        return emails
    
    def _create_document_record(self, file_path: str, text: str, doc_type: str) -> Dict:
        """Create a record from a document file (PDF, DOCX, etc.)."""
        path = Path(file_path)
        return {
            "id": file_path,
            "subject": f"[{doc_type}] {path.name}",
            "from": "Document",
            "to": "",
            "date": "",
            "body": text,
            "attachments": [path.name],
            "attachment_text": [text],
            "source_file": file_path,
            "source_type": doc_type,
        }
    
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
    
    def _normalize_eml(self, msg, source_path: str = "") -> Dict:
        """Convert email.message.Message to standard format with attachment text extraction."""
        subject = msg.get("Subject", "")
        from_addr = msg.get("From", "")
        to_addr = msg.get("To", "")
        date_str = msg.get("Date", "")
        
        # Extract body
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    try:
                        body += part.get_content()
                    except:
                        pass
                elif part.get_content_type() == "text/html" and not body:
                    # Fallback to HTML if no plain text
                    try:
                        import re
                        html = part.get_content()
                        # Strip HTML tags
                        body = re.sub(r'<[^>]+>', ' ', html)
                        body = re.sub(r'\s+', ' ', body).strip()
                    except:
                        pass
        else:
            try:
                body = msg.get_content()
            except:
                body = ""
        
        # Extract attachments and their text content
        attachments = []
        attachment_texts = []
        excel_rows = []  # NEW: Store Excel data as structured rows
        
        for part in msg.walk():
            filename = part.get_filename()
            if filename:
                attachments.append(filename)
                
                # Extract text from attachment if enabled
                if self.extract_attachments:
                    ext = Path(filename).suffix.lower()
                    
                    # For Excel files, extract both text AND structured rows
                    if ext in ['.xlsx', '.xls']:
                        try:
                            payload = part.get_payload(decode=True)
                            if payload:
                                # Get text version
                                text = extract_attachment_from_bytes(payload, filename)
                                if text:
                                    attachment_texts.append(f"[{filename}]:\n{text}")
                                
                                # Also get structured row data
                                rows = extract_excel_rows_from_bytes(payload)
                                if rows:
                                    for row in rows:
                                        row["_source_file"] = filename
                                    excel_rows.extend(rows)
                        except Exception as e:
                            print(f"Could not extract from Excel {filename}: {e}")
                    
                    elif ext in ['.pdf', '.docx', '.txt', '.csv']:
                        try:
                            payload = part.get_payload(decode=True)
                            if payload:
                                text = extract_attachment_from_bytes(payload, filename)
                                if text:
                                    attachment_texts.append(f"[{filename}]:\n{text}")
                        except Exception as e:
                            print(f"Could not extract from attachment {filename}: {e}")
        
        return {
            "id": msg.get("Message-ID", ""),
            "subject": subject,
            "from": from_addr,
            "to": to_addr,
            "date": date_str,
            "body": body,
            "attachments": attachments,
            "attachment_text": attachment_texts,
            "excel_rows": excel_rows,  # NEW: Structured Excel data
            "source_file": source_path,
        }
    
    def _parse_attachment_list(self, attachments_str: str) -> List[str]:
        """Parse semicolon-separated attachment list."""
        if not attachments_str:
            return []
        return [a.strip() for a in attachments_str.split(";") if a.strip()]
