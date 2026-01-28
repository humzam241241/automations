"""
Core job: Append new rows to a master dataset (Excel or CSV).
"""
from typing import List


def append_rows(
    existing_headers: List[str],
    existing_rows: List[List[str]],
    new_headers: List[str],
    new_rows: List[List[str]],
    deduplicate_column: str = None
) -> tuple[List[str], List[List[str]]]:
    """
    Append new rows to existing dataset.
    
    Args:
        existing_headers: Current headers
        existing_rows: Current rows
        new_headers: Headers of new data
        new_rows: New data rows
        deduplicate_column: Optional column name to check for duplicates
    
    Returns:
        (merged_headers, merged_rows)
    """
    # Align headers
    merged_headers = existing_headers.copy()
    for header in new_headers:
        if header not in merged_headers:
            merged_headers.append(header)
    
    # Convert existing rows to match merged headers
    result_rows = []
    for row in existing_rows:
        aligned_row = _align_row(row, existing_headers, merged_headers)
        result_rows.append(aligned_row)
    
    # Deduplicate if requested
    if deduplicate_column and deduplicate_column in merged_headers:
        existing_values = set()
        col_idx = merged_headers.index(deduplicate_column)
        for row in result_rows:
            if col_idx < len(row):
                existing_values.add(row[col_idx])
        
        new_col_idx = new_headers.index(deduplicate_column) if deduplicate_column in new_headers else -1
        for row in new_rows:
            aligned_row = _align_row(row, new_headers, merged_headers)
            if new_col_idx >= 0 and new_col_idx < len(row):
                if row[new_col_idx] not in existing_values:
                    result_rows.append(aligned_row)
            else:
                result_rows.append(aligned_row)
    else:
        # Just append all
        for row in new_rows:
            aligned_row = _align_row(row, new_headers, merged_headers)
            result_rows.append(aligned_row)
    
    return merged_headers, result_rows


def _align_row(row: List[str], source_headers: List[str], target_headers: List[str]) -> List[str]:
    """Align a row from source schema to target schema."""
    result = [""] * len(target_headers)
    for i, header in enumerate(source_headers):
        if header in target_headers and i < len(row):
            target_idx = target_headers.index(header)
            result[target_idx] = row[i]
    return result
