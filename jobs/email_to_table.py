"""
Core job: Extract emails into tabular data based on schema and rules.
"""
import re
from typing import Dict, List, Tuple


def normalize_text(text: str) -> str:
    """Normalize text for keyword matching."""
    return re.sub(r"\s+", " ", text or "").strip().lower()


def extract_snippet(text: str, match_pos: int, context: int = 30) -> str:
    """Extract a snippet around a match position."""
    start = max(0, match_pos - context)
    end = min(len(text), match_pos + context)
    snippet = text[start:end]
    if start > 0:
        snippet = "..." + snippet
    if end < len(text):
        snippet = snippet + "..."
    return snippet.strip()


def get_search_content(email_data: Dict, search_in: List[str]) -> str:
    """
    Build search content based on search_in specification.
    
    Args:
        email_data: Email dict
        search_in: List of fields to search in
    
    Returns:
        Combined text from specified fields
    """
    parts = []
    
    if "all" in search_in:
        search_in = ["subject", "body", "attachments_text", "attachments_names", "from", "to"]
    
    if "subject" in search_in:
        parts.append(email_data.get("subject", ""))
    
    if "body" in search_in:
        parts.append(email_data.get("body", ""))
    
    if "attachments_text" in search_in:
        parts.extend(email_data.get("attachment_text", []))
    
    if "attachments_names" in search_in:
        parts.extend(email_data.get("attachments", []))
    
    if "from" in search_in:
        parts.append(email_data.get("from", ""))
    
    if "to" in search_in:
        parts.append(email_data.get("to", ""))
    
    return " ".join(parts)


def apply_keyword_rules_enhanced(
    email_data: Dict,
    rules: List[Dict],
    explain: bool = False
) -> Tuple[Dict[str, str], Dict[str, Dict]]:
    """
    Apply keyword/regex rules with enhanced matching and explanation.
    
    Args:
        email_data: Email dict
        rules: List of rule dicts
        explain: If True, collect detailed match info
    
    Returns:
        (result_dict, explanation_dict)
        result_dict: {column: value}
        explanation_dict: {column: {matched, match_type, matched_term, search_in, snippet}}
    """
    result = {}
    explanations = {}
    
    # Sort rules by priority (higher first), then by order
    sorted_rules = sorted(
        rules,
        key=lambda r: (r.get("priority", 0), -rules.index(r)),
        reverse=True
    )
    
    for rule in sorted_rules:
        column = rule.get("column", rule.get("header", ""))
        
        # First match wins - skip if column already has a value
        if column in result:
            continue
        
        keywords = rule.get("keywords", [])
        regex_pattern = rule.get("regex")
        value = rule.get("value", "Yes")
        unmatched_value = rule.get("unmatched_value", "")
        search_in = rule.get("search_in", ["all"])
        word_boundary = rule.get("word_boundary", True)
        
        # Build search content
        content = get_search_content(email_data, search_in)
        normalized = normalize_text(content)
        
        matched = False
        match_type = None
        matched_term = None
        snippet = None
        
        # Keyword matching
        if keywords:
            for keyword in keywords:
                keyword_lower = keyword.lower()
                
                if word_boundary:
                    # Word-aware matching
                    pattern = r'\b' + re.escape(keyword_lower) + r'\b'
                    match = re.search(pattern, normalized)
                    if match:
                        matched = True
                        match_type = "keyword_word"
                        matched_term = keyword
                        if explain:
                            snippet = extract_snippet(content, match.start())
                        break
                else:
                    # Substring matching
                    if keyword_lower in normalized:
                        matched = True
                        match_type = "keyword_substring"
                        matched_term = keyword
                        if explain:
                            pos = normalized.find(keyword_lower)
                            snippet = extract_snippet(content, pos)
                        break
        
        # Regex matching
        if regex_pattern and not matched:
            try:
                match = re.search(regex_pattern, content, re.IGNORECASE)
                if match:
                    matched = True
                    match_type = "regex"
                    matched_term = regex_pattern
                    if explain:
                        snippet = extract_snippet(content, match.start())
            except re.error:
                pass
        
        result[column] = value if matched else unmatched_value
        
        if explain:
            explanations[column] = {
                "matched": matched,
                "match_type": match_type,
                "matched_term": matched_term,
                "search_in": search_in,
                "snippet": snippet,
            }
    
    return result, explanations


def extract_email_fields(email_data: Dict) -> Dict[str, str]:
    """
    Extract standard fields from email data.
    
    Args:
        email_data: Dict with keys like subject, from, to, date, body, attachments
    
    Returns:
        Dict of standard fields (case-insensitive mapping)
    """
    # Build both exact and lowercase versions for flexible matching
    standard = {
        # Standard casing
        "Subject": email_data.get("subject", ""),
        "From": email_data.get("from", ""),
        "To": email_data.get("to", ""),
        "Date": email_data.get("date", ""),
        "Body": email_data.get("body", ""),
        "Attachments": "; ".join(email_data.get("attachments", [])),
        # Lowercase versions
        "subject": email_data.get("subject", ""),
        "from": email_data.get("from", ""),
        "to": email_data.get("to", ""),
        "date": email_data.get("date", ""),
        "body": email_data.get("body", ""),
        "attachments": "; ".join(email_data.get("attachments", [])),
        # Common variations
        "Sender": email_data.get("from", ""),
        "sender": email_data.get("from", ""),
        "From Address": email_data.get("from", ""),
        "Recipient": email_data.get("to", ""),
        "recipient": email_data.get("to", ""),
        "To Address": email_data.get("to", ""),
        "Received": email_data.get("date", ""),
        "received": email_data.get("date", ""),
        "Sent": email_data.get("date", ""),
        "sent": email_data.get("date", ""),
        "Message": email_data.get("body", ""),
        "message": email_data.get("body", ""),
        "Content": email_data.get("body", ""),
        "content": email_data.get("body", ""),
        "Email Body": email_data.get("body", ""),
    }
    return standard


def smart_search_column(email_data: Dict, column_name: str) -> str:
    """
    Smart search for a column value by searching the email content.
    Uses the column name as a keyword to find related content.
    
    Args:
        email_data: Email dict
        column_name: Column name to search for
    
    Returns:
        Extracted value or 'Yes'/'No' depending on match
    """
    # Get individual content parts
    subject = str(email_data.get("subject", ""))
    body = str(email_data.get("body", ""))
    from_addr = str(email_data.get("from", ""))
    to_addr = str(email_data.get("to", ""))
    attachments = " ".join(email_data.get("attachments", []))
    
    # Combine all content
    all_content = f"{subject} {body} {from_addr} {to_addr} {attachments}"
    all_content_lower = all_content.lower()
    
    column_lower = column_name.lower().strip()
    
    # SPECIAL CASES: Extract actual data for common column types
    
    # Date columns - extract dates from content
    if any(kw in column_lower for kw in ['date', 'time', 'when', 'received', 'sent']):
        # Return the email date
        email_date = email_data.get("date", "")
        if email_date:
            return str(email_date)
        # Try to find dates in body
        date_pattern = r'\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4}|\d{4}[/\-]\d{1,2}[/\-]\d{1,2}'
        match = re.search(date_pattern, all_content)
        if match:
            return match.group()
    
    # Number/Amount columns - extract numbers
    if any(kw in column_lower for kw in ['amount', 'total', 'price', 'cost', 'number', 'qty', 'quantity', '#']):
        # Look for currency amounts
        amount_pattern = r'[\$€£]\s*[\d,]+\.?\d*|\d{1,3}(?:,\d{3})*(?:\.\d{2})?'
        match = re.search(amount_pattern, all_content)
        if match:
            return match.group()
    
    # ID/Reference columns - extract IDs
    if any(kw in column_lower for kw in ['id', 'ref', 'reference', 'ticket', 'case', 'order', 'po', 'invoice']):
        # Look for reference numbers
        ref_pattern = rf'{re.escape(column_lower)}\s*[:#]?\s*([A-Z0-9\-]+)'
        match = re.search(ref_pattern, all_content, re.IGNORECASE)
        if match:
            return match.group(1)
        # General alphanumeric ID pattern
        id_pattern = r'[A-Z]{2,5}[-#]?\d{4,10}'
        match = re.search(id_pattern, all_content)
        if match:
            return match.group()
    
    # Priority/Status columns - look for keywords
    if any(kw in column_lower for kw in ['priority', 'urgent', 'importance']):
        if any(kw in all_content_lower for kw in ['urgent', 'asap', 'critical', 'high priority', 'important']):
            return "High"
        elif any(kw in all_content_lower for kw in ['low priority', 'fyi', 'when possible']):
            return "Low"
        else:
            return "Normal"
    
    if any(kw in column_lower for kw in ['status']):
        if any(kw in all_content_lower for kw in ['complete', 'done', 'finished', 'resolved', 'closed']):
            return "Complete"
        elif any(kw in all_content_lower for kw in ['pending', 'waiting', 'in progress', 'ongoing']):
            return "Pending"
        elif any(kw in all_content_lower for kw in ['open', 'new', 'unread']):
            return "Open"
    
    # YES/NO detection - search for the column keyword
    # Look for column name followed by colon and value
    extract_pattern = rf'{re.escape(column_lower)}\s*[:\-=]\s*([^\n,;]+)'
    match = re.search(extract_pattern, all_content_lower, re.IGNORECASE)
    if match:
        value = match.group(1).strip()
        if len(value) < 100:
            return value.title()  # Return extracted value
    
    # Word boundary search - just check if keyword exists
    pattern = r'\b' + re.escape(column_lower) + r'\b'
    if re.search(pattern, all_content_lower):
        return "Yes"  # Found the keyword
    
    # Check for partial matches (column might be abbreviation)
    if len(column_lower) >= 3:
        if column_lower in all_content_lower:
            return "Yes"
    
    return ""  # Not found


def email_to_record(
    email_data: Dict,
    schema: Dict,
    rules: List[Dict],
    explain: bool = False
) -> Tuple[Dict[str, str], Dict]:
    """
    Convert email to record dict (NEW canonical format).
    
    Args:
        email_data: Email dict
        schema: Schema dict with {columns: [{name, type}, ...]}
        rules: List of keyword/regex rules
        explain: If True, return explanation of which rules matched
    
    Returns:
        (record_dict, explanation_dict)
        record_dict: {column_name: value}
        explanation_dict: {column_name: match_details}
    """
    standard_fields = extract_email_fields(email_data)
    
    # Apply rules with enhanced matching
    rule_results, rule_explanations = apply_keyword_rules_enhanced(
        email_data, rules, explain
    )
    
    # Build record dict based on schema order
    record = {}
    explanations = dict(rule_explanations) if rule_explanations else {}
    
    for col in schema.get("columns", []):
        col_name = col.get("name", col.get("header", ""))
        
        # Priority 1: Standard email fields (Subject, From, To, Date, Body, etc.)
        if col_name in standard_fields and standard_fields[col_name]:
            record[col_name] = standard_fields[col_name]
            if explain:
                explanations[col_name] = {
                    "matched": True,
                    "match_type": "standard_field",
                    "matched_term": col_name,
                    "search_in": ["email_headers"],
                    "snippet": str(standard_fields[col_name])[:50],
                }
        
        # Priority 2: Explicit rule results
        elif col_name in rule_results and rule_results[col_name]:
            record[col_name] = rule_results[col_name]
        
        # Priority 3: Smart search - use column name as keyword
        else:
            smart_value = smart_search_column(email_data, col_name)
            record[col_name] = smart_value
            
            if explain and smart_value:
                explanations[col_name] = {
                    "matched": True,
                    "match_type": "smart_search",
                    "matched_term": col_name,
                    "search_in": ["all"],
                    "snippet": f"Found '{col_name}' in email content",
                }
            elif explain:
                explanations[col_name] = {
                    "matched": False,
                    "match_type": None,
                    "matched_term": None,
                    "search_in": ["all"],
                    "snippet": None,
                }
    
    return record, explanations


def email_to_row(
    email_data: Dict,
    schema: Dict,
    rules: List[Dict],
    explain: bool = False
) -> Tuple[List[str], List[str], Dict]:
    """
    Convert email to table row (BACKWARD COMPATIBLE).
    
    Args:
        email_data: Email dict
        schema: Schema dict with {columns: [{name, type}, ...]}
        rules: List of keyword/regex rules
        explain: If True, return explanation of which rules matched
    
    Returns:
        (headers, row_values, explanation_dict)
    """
    record, explanation = email_to_record(email_data, schema, rules, explain)
    
    # Convert record to headers + row
    headers = list(record.keys())
    row = list(record.values())
    
    return headers, row, explanation
