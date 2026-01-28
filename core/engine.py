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
from jobs.email_to_table import email_to_record


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
        Execute a profile configuration with pipeline support.
        
        Args:
            profile: Profile dict
            explain: If True, include rule match explanations
        
        Returns:
            Result dict with status, output_path, and optional explanation
        """
        logging.info(f"Running profile: {profile.get('name')}")
        
        # Get pipeline (default to email_to_table for backward compat)
        pipeline = profile.get("pipeline", ["email_to_table"])
        
        # 1. Load emails based on input_source
        emails = self._load_emails(profile)
        logging.info(f"Loaded {len(emails)} emails")
        
        if not emails:
            return {"status": "no_emails", "message": "No emails found"}
        
        # 2. Execute pipeline
        records = []
        explanations = []
        
        for job_name in pipeline:
            if job_name == "email_to_table":
                records, explanations = self._job_email_to_table(
                    emails, profile, explain
                )
            elif job_name == "append_to_master":
                records = self._job_append_to_master(records, profile)
            elif job_name == "excel_to_biready":
                records = self._job_excel_to_biready(records, profile)
            else:
                logging.warning(f"Unknown job: {job_name}")
        
        if not records:
            return {"status": "no_data", "message": "Pipeline produced no data"}
        
        # 3. Convert records to headers + rows for output
        headers, rows = self._records_to_table(records, profile)
        
        # 4. Check if should export to BI
        output_config = profile.get("output", {})
        should_export_bi = (
            output_config.get("destination") == "bi_dashboard" or
            output_config.get("also_export_bi", False) or
            profile.get("input_source") in ["excel_file", "csv_file"]
        )
        
        # 5. Write output
        output_result = self._write_output(profile, headers, rows)
        
        # 6. Generate BI dashboard if requested
        if should_export_bi:
            bi_result = self._export_bi_dashboard(profile, headers, rows)
            output_result["bi_dashboard"] = bi_result
        
        result = {
            "status": "success",
            "emails_processed": len(emails),
            "records_produced": len(records),
            "output": output_result,
            "headers": headers,
            "rows": rows,
            "profile_name": profile.get("name", "Unknown"),
        }
        
        if explain:
            result["explanations"] = explanations
        
        logging.info(f"Profile execution complete: {result}")
        return result
    
    def _job_email_to_table(
        self,
        emails: List[Dict],
        profile: Dict,
        explain: bool
    ) -> tuple[List[Dict], List[Dict]]:
        """Execute email_to_table job."""
        schema = profile.get("schema", {})
        rules = profile.get("rules", [])
        
        records = []
        explanations = []
        
        for email_data in emails:
            record, explanation = email_to_record(email_data, schema, rules, explain)
            records.append(record)
            if explain:
                explanations.append(explanation)
        
        return records, explanations
    
    def _job_append_to_master(
        self,
        records: List[Dict],
        profile: Dict
    ) -> List[Dict]:
        """Execute append_to_master job."""
        try:
            from jobs.append_to_master import append_rows
            
            master_config = profile.get("master_dataset", {})
            master_path = master_config.get("path")
            dedupe_column = master_config.get("deduplicate_column")
            
            if not master_path:
                logging.warning("No master_dataset.path in profile, skipping append")
                return records
            
            # Load existing master if exists
            existing_records = self._load_master_records(master_path)
            
            # Convert to headers+rows for append_rows function
            if existing_records:
                existing_headers = list(existing_records[0].keys())
                existing_rows = [list(r.values()) for r in existing_records]
            else:
                existing_headers = []
                existing_rows = []
            
            new_headers = list(records[0].keys()) if records else []
            new_rows = [list(r.values()) for r in records]
            
            merged_headers, merged_rows = append_rows(
                existing_headers,
                existing_rows,
                new_headers,
                new_rows,
                dedupe_column
            )
            
            # Convert back to records
            merged_records = []
            for row in merged_rows:
                record = {h: v for h, v in zip(merged_headers, row)}
                merged_records.append(record)
            
            # Save master
            self._save_master_records(master_path, merged_records)
            
            return merged_records
        
        except ImportError:
            logging.warning("append_to_master job not available")
            return records
    
    def _job_excel_to_biready(
        self,
        records: List[Dict],
        profile: Dict
    ) -> List[Dict]:
        """Execute excel_to_biready job."""
        try:
            from jobs.excel_to_biready import apply_bi_transformations
            
            transformations = profile.get("bi_transformations", [])
            if not transformations:
                return records
            
            # Convert to headers+rows
            headers = list(records[0].keys()) if records else []
            rows = [list(r.values()) for r in records]
            
            # Apply transformations
            new_headers, new_rows = apply_bi_transformations(
                headers, rows, transformations
            )
            
            # Convert back to records
            new_records = []
            for row in new_rows:
                record = {h: v for h, v in zip(new_headers, row)}
                new_records.append(record)
            
            return new_records
        
        except ImportError:
            logging.warning("excel_to_biready job not available")
            return records
    
    def _records_to_table(
        self,
        records: List[Dict],
        profile: Dict
    ) -> tuple[List[str], List[List[str]]]:
        """Convert records to headers + rows, preserving schema order."""
        if not records:
            return [], []
        
        # Get headers from profile schema to preserve order
        schema = profile.get("schema", {})
        if schema and schema.get("columns"):
            headers = [
                col.get("name", col.get("header", ""))
                for col in schema.get("columns", [])
            ]
        else:
            # Fallback: use keys from first record
            headers = list(records[0].keys())
        
        # Build rows aligned to headers
        rows = []
        for record in records:
            row = [record.get(h, "") for h in headers]
            rows.append(row)
        
        return headers, rows
    
    def _load_master_records(self, path: str) -> List[Dict]:
        """Load existing master dataset."""
        from pathlib import Path
        import csv
        
        file_path = Path(path)
        if not file_path.exists():
            return []
        
        records = []
        if path.endswith(".csv"):
            with file_path.open("r", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                records = list(reader)
        elif path.endswith(".xlsx"):
            from openpyxl import load_workbook
            wb = load_workbook(path)
            ws = wb.active
            headers = [cell.value for cell in ws[1]]
            for row in ws.iter_rows(min_row=2, values_only=True):
                record = {h: v for h, v in zip(headers, row)}
                records.append(record)
        
        return records
    
    def _save_master_records(self, path: str, records: List[Dict]):
        """Save master dataset."""
        if not records:
            return
        
        headers = list(records[0].keys())
        rows = [list(r.values()) for r in records]
        
        if path.endswith(".csv"):
            writer = CSVWriter()
            content = writer.create_csv(headers, rows)
            Path(path).write_bytes(content)
        elif path.endswith(".xlsx"):
            writer = ExcelWriter()
            content = writer.create_workbook(headers, rows)
            Path(path).write_bytes(content)
    
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
        
        elif input_source == "excel_file":
            from adapters.excel_csv_email import ExcelCSVEmailAdapter
            adapter = ExcelCSVEmailAdapter()
            file_path = email_selection.get("file_path", "")
            
            if not file_path:
                raise ValueError("Excel file path is required")
            
            headers, emails = adapter.load_from_excel(file_path)
            
            # If auto-detect enabled, update schema
            if profile.get("auto_detect_columns") and headers:
                profile["schema"] = {
                    "columns": [{"name": h, "type": "text"} for h in headers]
                }
                logging.info(f"Auto-detected {len(headers)} columns from Excel")
            
            return emails
        
        elif input_source == "csv_file":
            from adapters.excel_csv_email import ExcelCSVEmailAdapter
            adapter = ExcelCSVEmailAdapter()
            file_path = email_selection.get("file_path", "")
            
            if not file_path:
                raise ValueError("CSV file path is required")
            
            headers, emails = adapter.load_from_csv(file_path)
            
            # If auto-detect enabled, update schema
            if profile.get("auto_detect_columns") and headers:
                profile["schema"] = {
                    "columns": [{"name": h, "type": "text"} for h in headers]
                }
                logging.info(f"Auto-detected {len(headers)} columns from CSV")
            
            return emails
        
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
    
    def _export_bi_dashboard(self, profile: Dict, headers: List[str], rows: List[List[str]]) -> Dict:
        """Export to BI dashboard."""
        try:
            from bi_dashboard_export import generate_dashboard, save_dashboard
            import webbrowser
            
            profile_name = profile.get("name", "Dashboard")
            html = generate_dashboard(headers, rows, profile_name)
            
            # Save dashboard
            output_dir = Path("./output/dashboards")
            output_dir.mkdir(parents=True, exist_ok=True)
            
            path = save_dashboard(html, output_dir)
            
            # Open in browser
            webbrowser.open(f"file:///{path}")
            
            logging.info(f"BI Dashboard saved: {path}")
            
            return {
                "enabled": True,
                "path": path,
                "opened_in_browser": True
            }
        
        except Exception as e:
            logging.error(f"Failed to export BI dashboard: {e}")
            return {
                "enabled": False,
                "error": str(e)
            }