"""
Core job: Extract emails into tabular data based on schema and rules.
"""
import re
from typing import Dict, List, Tuple


def normalize_text(text: str) -> str:
    """Normalize text for keyword matching."""
    return re.sub(r"\s+", " ", text or "").strip().lower()


def apply_keyword_rules(content: str, rules: List[Dict]) -> Dict[str, str]:
    """
    Apply keyword/regex rules to content.
    
    Args:
        content: Combined email text (subject + body + attachments)
        rules: List of rule dicts with {column, keywords, regex, value, unmatched_value}
    
    Returns:
        Dict mapping column name to matched value
    """
    normalized = normalize_text(content)
    result = {}
    
    for rule in rules:
        column = rule.get("column", rule.get("header", ""))
        keywords = rule.get("keywords", [])
        regex_pattern = rule.get("regex")
        value = rule.get("value", "Yes")
        unmatched_value = rule.get("unmatched_value", "")
        
        matched = False
        
        # Keyword matching
        if keywords:
            keywords_lower = [kw.lower() for kw in keywords]
            matched = any(kw in normalized for kw in keywords_lower)
        
        # Regex matching
        if regex_pattern and not matched:
            try:
                if re.search(regex_pattern, content, re.IGNORECASE):
                    matched = True
            except re.error:
                pass
        
        result[column] = value if matched else unmatched_value
    
    return result


def extract_email_fields(email_data: Dict) -> Dict[str, str]:
    """
    Extract standard fields from email data.
    
    Args:
        email_data: Dict with keys like subject, from, to, date, body, attachments
    
    Returns:
        Dict of standard fields
    """
    return {
        "Subject": email_data.get("subject", ""),
        "From": email_data.get("from", ""),
        "To": email_data.get("to", ""),
        "Date": email_data.get("date", ""),
        "Body": email_data.get("body", ""),
        "Attachments": "; ".join(email_data.get("attachments", [])),
    }


def email_to_row(
    email_data: Dict,
    schema: Dict,
    rules: List[Dict],
    explain: bool = False
) -> Tuple[List[str], List[str], Dict]:
    """
    Convert email to table row.
    
    Args:
        email_data: Email dict
        schema: Schema dict with {columns: [{name, type}, ...]}
        rules: List of keyword/regex rules
        explain: If True, return explanation of which rules matched
    
    Returns:
        (headers, row_values, explanation_dict)
    """
    standard_fields = extract_email_fields(email_data)
    
    # Build combined content for rule matching
    combined_content = " ".join([
        email_data.get("subject", ""),
        email_data.get("body", ""),
        " ".join(email_data.get("attachment_text", [])),
    ])
    
    rule_results = apply_keyword_rules(combined_content, rules)
    
    headers = []
    row = []
    explanation = {}
    
    # Add schema columns
    for col in schema.get("columns", []):
        col_name = col.get("name", col.get("header", ""))
        headers.append(col_name)
        
        # Check if it's a standard field
        if col_name in standard_fields:
            row.append(standard_fields[col_name])
        # Check if it's a rule-based column
        elif col_name in rule_results:
            row.append(rule_results[col_name])
            if explain and rule_results[col_name]:
                explanation[col_name] = f"Matched in content"
        else:
            row.append("")
    
    return headers, row, explanation
