"""
Core job: Transform raw Excel data into BI-ready format.
Apply data quality checks, pivot, denormalize, etc.
"""
from typing import Dict, List


def clean_value(value: str, rules: Dict) -> str:
    """Apply cleaning rules to a value."""
    if not value:
        return ""
    
    # Trim whitespace
    if rules.get("trim", True):
        value = value.strip()
    
    # Lowercase
    if rules.get("lowercase"):
        value = value.lower()
    
    # Remove special chars
    if rules.get("remove_special_chars"):
        import re
        value = re.sub(r"[^\w\s-]", "", value)
    
    return value


def apply_bi_transformations(
    headers: List[str],
    rows: List[List[str]],
    transformations: List[Dict]
) -> tuple[List[str], List[List[str]]]:
    """
    Apply BI transformations to tabular data.
    
    Args:
        headers: Column headers
        rows: Data rows
        transformations: List of transformation dicts:
            - {type: "clean", column: "...", rules: {...}}
            - {type: "add_column", name: "...", formula: "..."}
            - {type: "filter", column: "...", condition: "..."}
    
    Returns:
        (transformed_headers, transformed_rows)
    """
    # Create a working copy
    result_headers = headers.copy()
    result_rows = [row.copy() for row in rows]
    
    for transform in transformations:
        transform_type = transform.get("type")
        
        if transform_type == "clean":
            col_name = transform.get("column")
            rules = transform.get("rules", {})
            if col_name in result_headers:
                col_idx = result_headers.index(col_name)
                for row in result_rows:
                    if col_idx < len(row):
                        row[col_idx] = clean_value(row[col_idx], rules)
        
        elif transform_type == "add_column":
            col_name = transform.get("name")
            default_value = transform.get("default", "")
            result_headers.append(col_name)
            for row in result_rows:
                row.append(default_value)
        
        elif transform_type == "filter":
            col_name = transform.get("column")
            condition = transform.get("condition", "")
            if col_name in result_headers:
                col_idx = result_headers.index(col_name)
                result_rows = [
                    row for row in result_rows
                    if col_idx < len(row) and _eval_condition(row[col_idx], condition)
                ]
    
    return result_headers, result_rows


def _eval_condition(value: str, condition: str) -> bool:
    """Safely evaluate a simple condition."""
    if not condition:
        return True
    
    # Support simple conditions like "not_empty", "contains:text", "equals:value"
    if condition == "not_empty":
        return bool(value.strip())
    
    if condition.startswith("contains:"):
        search_text = condition[9:]
        return search_text.lower() in value.lower()
    
    if condition.startswith("equals:"):
        expected = condition[7:]
        return value.strip().lower() == expected.lower()
    
    return True
