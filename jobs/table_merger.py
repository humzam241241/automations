"""
Table Merger - Intelligently merge tables from multiple emails with different structures.
Handles:
- Different column sets across tables
- Missing columns in some tables
- Merging by index column (Work Order number)
- Filling gaps from email context when columns aren't in tables
"""
import logging
from typing import Dict, List, Set, Optional


def merge_tables(
    all_records: List[Dict],
    index_column: Optional[str] = None,
    schema_columns: Optional[List[str]] = None
) -> List[Dict]:
    """
    Merge multiple records (from different emails/tables) into a unified master table.
    
    Args:
        all_records: List of record dicts from multiple emails
        index_column: Column name to use as index (e.g., "order", "wo number")
        schema_columns: Expected column names from profile schema (for filling gaps)
    
    Returns:
        Merged list of records with all columns present
    """
    if not all_records:
        return []
    
    # Step 1: Collect all unique column names across all records
    all_columns: Set[str] = set()
    for record in all_records:
        all_columns.update(record.keys())
    
    # Add schema columns if provided (ensures expected columns exist)
    if schema_columns:
        all_columns.update(schema_columns)
    
    # Remove internal metadata columns
    all_columns.discard("_source")
    all_columns.discard("_source_file")
    all_columns.discard("_sheet")
    all_columns.discard("_focus_order")
    
    all_columns = sorted(list(all_columns))
    
    # Step 2: If index column specified, group by index value
    if index_column:
        # Find actual index column name (handle synonyms/case)
        index_col_name = _find_index_column(index_column, all_columns)
        
        if index_col_name:
            return _merge_by_index(all_records, index_col_name, all_columns)
    
    # Step 3: No index column - just normalize structure (add missing columns)
    normalized = []
    for record in all_records:
        normalized_record = {}
        for col in all_columns:
            normalized_record[col] = record.get(col, "")
        normalized.append(normalized_record)
    
    return normalized


def _find_index_column(index_column: str, available_columns: List[str]) -> Optional[str]:
    """Find the actual index column name (case-insensitive, handles synonyms)."""
    index_lower = index_column.lower()
    
    # Direct match
    for col in available_columns:
        if col.lower() == index_lower:
            return col
    
    # Partial match (e.g., "order" matches "Order Number")
    for col in available_columns:
        if index_lower in col.lower() or col.lower() in index_lower:
            return col
    
    # Try synonyms
    try:
        from jobs.synonym_resolver import get_synonyms
        
        for col in available_columns:
            synonyms = get_synonyms(col)
            if index_column.lower() in [s.lower() for s in synonyms]:
                return col
    except ImportError:
        pass
    
    return None


def _merge_by_index(
    records: List[Dict],
    index_col_name: str,
    all_columns: List[str]
) -> List[Dict]:
    """
    Merge records that have the same index value.
    When multiple records share the same index, combine their columns intelligently.
    """
    # Group records by index value
    indexed: Dict[str, List[Dict]] = {}
    
    for record in records:
        index_val = str(record.get(index_col_name, "")).strip()
        if not index_val or index_val.lower() in ["", "none", "n/a", "null"]:
            # No index value - keep as separate record
            indexed.setdefault("_no_index", []).append(record)
        else:
            indexed.setdefault(index_val, []).append(record)
    
    merged = []
    
    for index_val, group in indexed.items():
        if index_val == "_no_index":
            # Records without index - normalize structure but keep separate
            for record in group:
                normalized = {}
                for col in all_columns:
                    normalized[col] = record.get(col, "")
                merged.append(normalized)
        elif len(group) == 1:
            # Single record with this index - just normalize structure
            normalized = {}
            for col in all_columns:
                normalized[col] = group[0].get(col, "")
            merged.append(normalized)
        else:
            # Multiple records with same index - merge them
            merged_record = _merge_record_group(group, all_columns, index_col_name)
            merged.append(merged_record)
    
    logging.info(f"Merged {len(records)} records into {len(merged)} by index '{index_col_name}'")
    return merged


def _merge_record_group(group: List[Dict], all_columns: List[str], index_col: str) -> Dict:
    """
    Merge a group of records that share the same index value.
    Strategy:
    1. Index column value is preserved (should be same anyway)
    2. For other columns, prefer non-empty values
    3. If multiple non-empty values, prefer longer/more complete value
    4. If still tied, prefer value from record with more filled columns
    """
    merged = {}
    
    # Score each record by how complete it is
    record_scores = []
    for record in group:
        filled_count = sum(1 for k, v in record.items() if v and str(v).strip())
        record_scores.append((filled_count, record))
    
    # Sort by completeness (most complete first)
    record_scores.sort(reverse=True, key=lambda x: x[0])
    
    # Merge columns
    for col in all_columns:
        best_value = ""
        best_length = 0
        
        for _, record in record_scores:
            value = record.get(col, "")
            value_str = str(value).strip()
            
            if value_str:
                # Prefer longer, more complete values
                if len(value_str) > best_length:
                    best_value = value_str
                    best_length = len(value_str)
        
        merged[col] = best_value
    
    # Ensure index column is set (use first non-empty value)
    if not merged.get(index_col):
        for _, record in record_scores:
            index_val = record.get(index_col, "")
            if index_val and str(index_val).strip():
                merged[index_col] = str(index_val).strip()
                break
    
    return merged


def fill_missing_columns_from_context(
    record: Dict,
    email_data: Dict,
    schema_columns: List[Dict],
    use_synonyms: bool = True
) -> Dict:
    """
    Fill missing columns in a record by extracting from email context.
    Used when tables don't have all expected columns.
    
    Args:
        record: Record dict (may have missing columns)
        email_data: Original email data for context extraction
        schema_columns: Schema column definitions [{name, extract_type}, ...]
        use_synonyms: Whether to use synonym matching
    
    Returns:
        Record with missing columns filled from email context
    """
    from jobs.email_to_table import smart_search_column
    
    filled_record = dict(record)
    
    for col_def in schema_columns:
        col_name = col_def.get("name", "")
        extract_type = col_def.get("extract_type", col_def.get("type", "auto"))
        
        # Skip if column already has a value
        if col_name in filled_record and filled_record[col_name] and str(filled_record[col_name]).strip():
            continue
        
        # Try to extract from email context
        value, meta = smart_search_column(email_data, col_name, extract_type, use_synonyms)
        
        if value and str(value).strip():
            filled_record[col_name] = value
    
    return filled_record
