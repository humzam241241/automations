"""
Core execution engine - runs profiles with selected adapters.
"""
import logging
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

from adapters.csv_writer import CSVWriter
from adapters.excel_writer import ExcelWriter
from adapters.graph_email import GraphEmailAdapter
from adapters.local_email import LocalEmailAdapter
from adapters.onedrive_storage import OneDriveAdapter
from jobs.email_to_table import email_to_row


class ExecutionEngine:
    """Execute profiles using configured adapters."""
    
    def __init__(self, access_token: Optional[str] = None):
        """
        Initialize engine.
        
        Args:
            access_token: Microsoft Graph access token (required for Graph/OneDrive)
        """
        self.access_token = access_token
    
    def run_profile(self, profile: Dict, explain: bool = False) -> Dict:
        """
        Execute a profile configuration.
        
        Args:
            profile: Profile dict
            explain: If True, include rule match explanations
        
        Returns:
            Result dict with status, output_path, and optional explanation
        """
        logging.info(f"Running profile: {profile.get('name')}")
        
        # 1. Load emails based on input_source
        emails = self._load_emails(profile)
        logging.info(f"Loaded {len(emails)} emails")
        
        if not emails:
            return {"status": "no_emails", "message": "No emails found"}
        
        # 2. Convert emails to rows
        schema = profile.get("schema", {})
        rules = profile.get("rules", [])
        
        all_headers = None
        all_rows = []
        explanations = []
        
        for email_data in emails:
            headers, row, explanation = email_to_row(email_data, schema, rules, explain)
            if all_headers is None:
                all_headers = headers
            all_rows.append(row)
            if explain:
                explanations.append(explanation)
        
        # 3. Write output
        output_result = self._write_output(profile, all_headers, all_rows)
        
        result = {
            "status": "success",
            "emails_processed": len(emails),
            "output": output_result,
        }
        
        if explain:
            result["explanations"] = explanations
        
        logging.info(f"Profile execution complete: {result}")
        return result
    
    def _load_emails(self, profile: Dict) -> List[Dict]:
        """Load emails based on profile input_source."""
        input_source = profile.get("input_source")
        email_selection = profile.get("email_selection", {})
        
        if input_source == "graph":
            if not self.access_token:
                raise RuntimeError("Graph access token required for input_source='graph'")
            
            adapter = GraphEmailAdapter(self.access_token)
            folder_name = email_selection.get("folder_name")
            search_query = email_selection.get("search_query")
            top = email_selection.get("newest_n", 25)
            
            folder_id = None
            if folder_name:
                folder_id = adapter.get_folder_id(folder_name)
            
            return adapter.fetch_messages(folder_id, search_query, top)
        
        elif input_source == "local_eml":
            adapter = LocalEmailAdapter()
            directory = email_selection.get("directory", "./input_emails")
            pattern = email_selection.get("pattern", "*.eml")
            file_paths = email_selection.get("file_paths")
            
            if file_paths:
                return adapter.load_from_eml_files(file_paths)
            else:
                return adapter.scan_directory(directory, pattern)
        
        elif input_source == "local_csv":
            adapter = LocalEmailAdapter()
            csv_path = email_selection.get("csv_path", "./emails.csv")
            return adapter.load_from_csv(csv_path)
        
        else:
            raise ValueError(f"Unknown input_source: {input_source}")
    
    def _write_output(self, profile: Dict, headers: List[str], rows: List[List[str]]) -> Dict:
        """Write output based on profile configuration."""
        output = profile.get("output", {})
        output_format = output.get("format", "excel")
        destination = output.get("destination", "local")
        
        # Generate filename
        filename_template = output.get("filename_template", "output_{timestamp}.xlsx")
        timestamp = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
        filename = filename_template.replace("{timestamp}", timestamp)
        
        # Generate file content
        if output_format == "excel":
            writer = ExcelWriter()
            template_path = output.get("template_path")
            if template_path:
                content = writer.fill_template(template_path, headers, rows)
            else:
                content = writer.create_workbook(headers, rows)
        elif output_format == "csv":
            writer = CSVWriter()
            content = writer.create_csv(headers, rows)
        else:
            raise ValueError(f"Unknown output format: {output_format}")
        
        # Write to destination
        if destination == "onedrive":
            if not self.access_token:
                raise RuntimeError("Graph access token required for destination='onedrive'")
            
            onedrive_path = output.get("onedrive_path", "/EmailReports")
            adapter = OneDriveAdapter(self.access_token)
            
            content_type = (
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                if output_format == "excel"
                else "text/csv"
            )
            
            result = adapter.upload_file(onedrive_path, filename, content, content_type)
            return {
                "destination": "onedrive",
                "path": f"{onedrive_path}/{filename}",
                "web_url": result.get("webUrl") if result else None,
            }
        
        elif destination == "local":
            local_path = output.get("local_path", "./output")
            Path(local_path).mkdir(parents=True, exist_ok=True)
            
            output_file = Path(local_path) / filename
            output_file.write_bytes(content)
            
            return {
                "destination": "local",
                "path": str(output_file.absolute()),
            }
        
        else:
            raise ValueError(f"Unknown destination: {destination}")
