"""
Adapter: Write data to CSV files.
"""
import csv
from pathlib import Path
from typing import List


class CSVWriter:
    """Write tabular data to CSV."""
    
    def create_csv(self, headers: List[str], rows: List[List[str]]) -> bytes:
        """
        Create CSV content.
        
        Args:
            headers: Column headers
            rows: Data rows
        
        Returns:
            CSV bytes (UTF-8)
        """
        from io import StringIO
        
        buffer = StringIO()
        writer = csv.writer(buffer)
        writer.writerow(headers)
        writer.writerows(rows)
        return buffer.getvalue().encode("utf-8")
    
    def append_to_file(
        self,
        file_path: str,
        headers: List[str],
        rows: List[List[str]]
    ) -> None:
        """
        Append rows to an existing CSV file.
        
        Args:
            file_path: Path to CSV
            headers: New data headers
            rows: New data rows
        """
        path = Path(file_path)
        
        if not path.exists():
            # Create new
            content = self.create_csv(headers, rows)
            path.write_bytes(content)
            return
        
        # Read existing
        with path.open("r", encoding="utf-8", newline="") as f:
            reader = csv.reader(f)
            existing_headers = next(reader, [])
            existing_rows = list(reader)
        
        # Align and append
        for row in rows:
            aligned_row = self._align_row(row, headers, existing_headers)
            existing_rows.append(aligned_row)
        
        # Write back
        with path.open("w", encoding="utf-8", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(existing_headers)
            writer.writerows(existing_rows)
    
    def _align_row(
        self,
        row: List[str],
        source_headers: List[str],
        target_headers: List[str]
    ) -> List[str]:
        """Align a row from source to target schema."""
        result = [""] * len(target_headers)
        for i, header in enumerate(source_headers):
            if header in target_headers and i < len(row):
                target_idx = target_headers.index(header)
                result[target_idx] = row[i]
        return result
