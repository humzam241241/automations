"""
Table Extractor - Extract tables from email bodies (HTML and plain text).
Finds the largest table and converts it to structured rows.
"""
import re
from typing import Dict, List, Tuple, Optional
from html.parser import HTMLParser


class HTMLTableParser(HTMLParser):
    """Parse HTML and extract all tables."""
    
    def __init__(self):
        super().__init__()
        self.tables = []  # List of tables, each table is list of rows
        self.current_table = []
        self.current_row = []
        self.current_cell = ""
        self.in_table = False
        self.in_row = False
        self.in_cell = False
    
    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        if tag == 'table':
            self.in_table = True
            self.current_table = []
        elif tag == 'tr' and self.in_table:
            self.in_row = True
            self.current_row = []
        elif tag in ('td', 'th') and self.in_row:
            self.in_cell = True
            self.current_cell = ""
    
    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag == 'table' and self.in_table:
            if self.current_table:
                self.tables.append(self.current_table)
            self.current_table = []
            self.in_table = False
        elif tag == 'tr' and self.in_row:
            if self.current_row:
                self.current_table.append(self.current_row)
            self.current_row = []
            self.in_row = False
        elif tag in ('td', 'th') and self.in_cell:
            self.current_row.append(self.current_cell.strip())
            self.current_cell = ""
            self.in_cell = False
    
    def handle_data(self, data):
        if self.in_cell:
            self.current_cell += data


def extract_tables_from_html(html: str) -> List[List[List[str]]]:
    """
    Extract all tables from HTML content.
    
    Returns:
        List of tables, where each table is a list of rows,
        and each row is a list of cell values.
    """
    if not html:
        return []
    
    try:
        parser = HTMLTableParser()
        parser.feed(html)
        return parser.tables
    except Exception as e:
        print(f"HTML table parsing error: {e}")
        return []


def extract_tables_from_text(text: str) -> List[List[List[str]]]:
    """
    Extract tables from plain text using common delimiters.
    Supports:
    - Pipe-delimited: | col1 | col2 | col3 |
    - Tab-delimited: col1\tcol2\tcol3
    - Comma-delimited (CSV-like): col1, col2, col3
    - Fixed-width columns (detected by alignment)
    
    Returns:
        List of detected tables.
    """
    if not text:
        return []
    
    tables = []
    lines = text.splitlines()
    
    # Strategy 1: Pipe-delimited tables
    pipe_table = extract_pipe_table(lines)
    if pipe_table and len(pipe_table) >= 2:
        tables.append(pipe_table)
    
    # Strategy 2: Tab-delimited tables
    tab_table = extract_tab_table(lines)
    if tab_table and len(tab_table) >= 2:
        tables.append(tab_table)
    
    # Strategy 3: Detect consistent column patterns (multiple spaces)
    space_table = extract_space_aligned_table(lines)
    if space_table and len(space_table) >= 2:
        tables.append(space_table)
    
    return tables


def extract_pipe_table(lines: List[str]) -> List[List[str]]:
    """Extract pipe-delimited table from lines."""
    table_rows = []
    in_table = False
    
    for line in lines:
        # Check if line has pipe separators (at least 2 pipes)
        if line.count('|') >= 2:
            # Skip separator lines (like |---|---|)
            if re.match(r'^[\s|:\-]+$', line):
                continue
            
            # Parse cells
            cells = [c.strip() for c in line.split('|')]
            # Remove empty first/last if line starts/ends with |
            if cells and not cells[0]:
                cells = cells[1:]
            if cells and not cells[-1]:
                cells = cells[:-1]
            
            if cells:
                table_rows.append(cells)
                in_table = True
        elif in_table and line.strip():
            # End of table (non-pipe line after pipe lines)
            break
    
    return table_rows


def extract_tab_table(lines: List[str]) -> List[List[str]]:
    """Extract tab-delimited table from lines."""
    table_rows = []
    
    for line in lines:
        if '\t' in line:
            cells = [c.strip() for c in line.split('\t')]
            if len(cells) >= 2 and any(c for c in cells):
                table_rows.append(cells)
    
    # Only return if we have consistent column counts
    if table_rows:
        col_counts = [len(row) for row in table_rows]
        most_common = max(set(col_counts), key=col_counts.count)
        # Keep only rows with the most common column count
        table_rows = [row for row in table_rows if len(row) == most_common]
    
    return table_rows


def extract_space_aligned_table(lines: List[str]) -> List[List[str]]:
    """
    Extract table from space-aligned columns.
    Looks for lines with multiple values separated by 2+ spaces.
    """
    potential_rows = []
    
    for line in lines:
        # Split on 2+ spaces
        parts = re.split(r'\s{2,}', line.strip())
        if len(parts) >= 3:  # At least 3 columns
            potential_rows.append(parts)
    
    if not potential_rows:
        return []
    
    # Find the most common column count
    col_counts = [len(row) for row in potential_rows]
    if not col_counts:
        return []
    
    most_common = max(set(col_counts), key=col_counts.count)
    
    # Only keep rows with that column count
    table_rows = [row for row in potential_rows if len(row) == most_common]
    
    return table_rows


def find_largest_table(tables: List[List[List[str]]]) -> Optional[List[List[str]]]:
    """
    Find the largest table (most cells = rows * columns).
    
    Returns:
        The largest table, or None if no tables.
    """
    if not tables:
        return None
    
    def table_size(table):
        if not table:
            return 0
        rows = len(table)
        cols = max(len(row) for row in table) if table else 0
        return rows * cols
    
    return max(tables, key=table_size)


def table_to_dicts(table: List[List[str]]) -> List[Dict[str, str]]:
    """
    Convert a table (list of rows) to list of dicts.
    First row is used as headers.
    
    Returns:
        List of dicts, one per data row.
    """
    if not table or len(table) < 2:
        return []
    
    # First row as headers
    headers = table[0]
    # Clean headers
    headers = [h.strip() if h else f"Column_{i}" for i, h in enumerate(headers)]
    
    rows = []
    for row in table[1:]:
        row_dict = {"_source": "email_body_table"}
        has_data = False
        
        for i, cell in enumerate(row):
            header = headers[i] if i < len(headers) else f"Column_{i}"
            value = str(cell).strip() if cell else ""
            row_dict[header] = value
            if value:
                has_data = True
        
        if has_data:
            rows.append(row_dict)
    
    return rows


def extract_body_table_rows(email_body: str, html_body: str = "") -> List[Dict[str, str]]:
    """
    Main function: Extract the largest table from email body.
    
    Args:
        email_body: Plain text body
        html_body: Raw HTML body (if available)
    
    Returns:
        List of dicts representing table rows, or empty list if no table found.
    """
    all_tables = []
    
    # Try HTML first (more reliable structure)
    if html_body:
        html_tables = extract_tables_from_html(html_body)
        all_tables.extend(html_tables)
    
    # Also try plain text
    if email_body:
        text_tables = extract_tables_from_text(email_body)
        all_tables.extend(text_tables)
    
    # Find largest table
    largest = find_largest_table(all_tables)
    
    if not largest:
        return []
    
    # Convert to dicts
    return table_to_dicts(largest)


def get_table_preview(table: List[List[str]], max_rows: int = 5) -> str:
    """Get a preview string of a table for debugging."""
    if not table:
        return "Empty table"
    
    lines = []
    for i, row in enumerate(table[:max_rows]):
        lines.append(f"Row {i}: {' | '.join(str(c)[:20] for c in row)}")
    
    if len(table) > max_rows:
        lines.append(f"... and {len(table) - max_rows} more rows")
    
    return "\n".join(lines)
