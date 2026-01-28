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
    
    INTELLIGENT EXTRACTION:
    - "MI #" → Extracts MI-12345, MI #12345, MI: 12345
    - "CAPA #" → Extracts CAPA-001, CAPA #001
    - "QE" → Extracts QE numbers
    - "Work Order" → Extracts WO-xxx, Work Order: xxx
    
    Args:
        email_data: Email dict
        column_name: Column name to search for
    
    Returns:
        Extracted value(s)
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
    column_clean = re.sub(r'[#\s]+', '', column_lower)  # Remove # and spaces
    
    # ============================================================
    # INTELLIGENT PREFIX + NUMBER EXTRACTION
    # Handles: MI #, MM #, CAPA #, QE, EQQ, PO, WO, etc.
    # ============================================================
    
    # Common prefixes that are followed by numbers
    known_prefixes = {
        'mi': ['MI', 'mi'],
        'mm': ['MM', 'mm'],
        'capa': ['CAPA', 'capa', 'Capa'],
        'qe': ['QE', 'qe', 'Qe'],
        'eqq': ['EQQ', 'eqq', 'Eqq'],
        'po': ['PO', 'po', 'P.O.', 'P.O'],
        'wo': ['WO', 'wo', 'Work Order', 'work order', 'WorkOrder'],
        'inv': ['INV', 'inv', 'Invoice', 'invoice'],
        'ref': ['REF', 'ref', 'Reference', 'reference', 'Ref'],
        'ticket': ['Ticket', 'ticket', 'TICKET', 'TKT'],
        'case': ['Case', 'case', 'CASE'],
        'cr': ['CR', 'cr', 'Change Request'],
        'pr': ['PR', 'pr', 'Purchase Request'],
        'sr': ['SR', 'sr', 'Service Request'],
        'prt': ['PRT', 'prt'],
        'ncr': ['NCR', 'ncr', 'Non-Conformance'],
        'car': ['CAR', 'car', 'Corrective Action'],
        'deviation': ['DEV', 'dev', 'Deviation', 'deviation'],
        'lot': ['LOT', 'lot', 'Lot', 'Batch', 'batch'],
        'batch': ['BATCH', 'batch', 'Batch'],
        'order': ['Order', 'order', 'ORDER'],
    }
    
    # Check if column matches any known prefix
    prefix_match = None
    for prefix_key, variations in known_prefixes.items():
        if column_clean == prefix_key or column_lower.replace('#', '').replace(' ', '') == prefix_key:
            prefix_match = variations
            break
        # Also check if column starts with the prefix
        if column_clean.startswith(prefix_key):
            prefix_match = variations
            break
    
    # If we found a prefix, extract all matching numbers
    if prefix_match:
        extracted_numbers = []
        
        for prefix in prefix_match:
            # Pattern: PREFIX followed by optional separator then number
            # Matches: MI-12345, MI #12345, MI: 12345, MI12345, MI 12345
            patterns = [
                rf'{re.escape(prefix)}\s*[#:\-_]?\s*(\d+[\-\d]*)',  # MI #12345, MI-12345
                rf'{re.escape(prefix)}(\d+)',  # MI12345
            ]
            
            for pattern in patterns:
                matches = re.findall(pattern, all_content, re.IGNORECASE)
                for match in matches:
                    full_ref = f"{prefix.upper()}-{match}" if not match.startswith('-') else f"{prefix.upper()}{match}"
                    # Clean up the reference
                    full_ref = re.sub(r'-+', '-', full_ref)  # Remove double dashes
                    if full_ref not in extracted_numbers:
                        extracted_numbers.append(full_ref)
        
        if extracted_numbers:
            # Return all found numbers, separated by semicolons
            return "; ".join(extracted_numbers[:5])  # Limit to first 5
    
    # ============================================================
    # GENERIC "PREFIX #" PATTERN
    # If column contains # sign, extract PREFIX + NUMBER
    # ============================================================
    
    if '#' in column_name:
        # Extract the prefix part (before #)
        prefix = column_name.split('#')[0].strip()
        if prefix:
            # Find PREFIX followed by number
            pattern = rf'{re.escape(prefix)}\s*[#:\-]?\s*(\d+[\-\d]*)'
            matches = re.findall(pattern, all_content, re.IGNORECASE)
            if matches:
                results = [f"{prefix.upper()}-{m}" for m in matches[:5]]
                return "; ".join(results)
    
    # ============================================================
    # STANDARD FIELD EXTRACTION (Date, Amount, Status, etc.)
    # ============================================================
    
    # Date columns - extract dates from content
    if any(kw in column_lower for kw in ['date', 'time', 'when', 'received', 'sent']):
        email_date = email_data.get("date", "")
        if email_date:
            return str(email_date)
        date_pattern = r'\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4}|\d{4}[/\-]\d{1,2}[/\-]\d{1,2}'
        match = re.search(date_pattern, all_content)
        if match:
            return match.group()
    
    # Amount columns - extract currency
    if any(kw in column_lower for kw in ['amount', 'total', 'price', 'cost', 'qty', 'quantity']):
        amount_pattern = r'[\$€£]\s*[\d,]+\.?\d*|\d{1,3}(?:,\d{3})*(?:\.\d{2})?'
        match = re.search(amount_pattern, all_content)
        if match:
            return match.group()
    
    # Priority columns
    if any(kw in column_lower for kw in ['priority', 'urgent', 'importance']):
        if any(kw in all_content_lower for kw in ['urgent', 'asap', 'critical', 'high priority', 'important']):
            return "High"
        elif any(kw in all_content_lower for kw in ['low priority', 'fyi', 'when possible']):
            return "Low"
        return "Normal"
    
    # Status columns
    if 'status' in column_lower:
        if any(kw in all_content_lower for kw in ['complete', 'done', 'finished', 'resolved', 'closed']):
            return "Complete"
        elif any(kw in all_content_lower for kw in ['pending', 'waiting', 'in progress', 'ongoing']):
            return "Pending"
        elif any(kw in all_content_lower for kw in ['open', 'new', 'unread']):
            return "Open"
    
    # ============================================================
    # GENERIC EXTRACTION: Column Name followed by value
    # ============================================================
    
    # Look for "ColumnName: value" or "ColumnName - value" patterns
    extract_pattern = rf'{re.escape(column_lower)}\s*[:\-=]\s*([^\n,;]+)'
    match = re.search(extract_pattern, all_content, re.IGNORECASE)
    if match:
        value = match.group(1).strip()
        if len(value) < 100:
            return value
    
    # ============================================================
    # KEYWORD EXISTENCE CHECK
    # ============================================================
    
    # Word boundary search
    pattern = r'\b' + re.escape(column_lower) + r'\b'
    if re.search(pattern, all_content_lower):
        return "Yes"
    
    # Partial match for abbreviations
    if len(column_lower) >= 2 and column_lower in all_content_lower:
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
