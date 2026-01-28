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


def _extract_number_only(column_name: str, content: str) -> str:
    """
    Extract ONLY numbers/IDs - NEVER returns "Yes".
    
    For columns like "MI #", "Building Number", "Work Order ID".
    """
    column_lower = column_name.lower().strip()
    
    # Get the prefix (part before number/id/#/etc.)
    prefix = column_name
    for suffix in ['number', 'num', '#', 'no.', 'no', 'id', 'code', 'ref', 'reference']:
        prefix = re.sub(rf'\s*{re.escape(suffix)}\s*$', '', prefix, flags=re.IGNORECASE)
    prefix = prefix.strip()
    
    # If we have a prefix, search for PREFIX + NUMBER
    if prefix:
        extracted = extract_prefix_numbers(prefix, content)
        if extracted:
            return extracted
    
    # If column name is just a short code (MI, QE, MM), search directly
    column_stripped = re.sub(r'[#\s\.\-_]+', '', column_name)
    if len(column_stripped) <= 5 and column_stripped.isupper():
        extracted = extract_prefix_numbers(column_stripped, content)
        if extracted:
            return extracted
    
    # Last resort: look for "ColumnName: 12345" pattern
    pattern = rf'{re.escape(column_name)}\s*[:\-=]\s*(\d[\d\-]*)'
    match = re.search(pattern, content, re.IGNORECASE)
    if match:
        return match.group(1)
    
    # Return empty string - NOT "Yes"!
    return ""


def smart_search_column(email_data: Dict, column_name: str, extract_type: str = "auto") -> str:
    """
    FULLY INTELLIGENT column extraction - works with ANY column name!
    
    Now supports explicit TYPE to prevent wrong extractions:
    - "number" → ONLY extract numbers, never return "Yes"
    - "text" → Extract text values
    - "date" → Extract dates
    - "amount" → Extract currency
    - "yesno" → Check keyword presence (Yes/No)
    - "email_field" → Standard email field
    - "auto" → Auto-detect from column name
    
    Args:
        email_data: Email dict
        column_name: ANY column name the user types
        extract_type: The expected value type (number, text, date, amount, yesno, auto)
    
    Returns:
        Extracted value based on type
    """
    # Get all content from email
    subject = str(email_data.get("subject", ""))
    body = str(email_data.get("body", ""))
    from_addr = str(email_data.get("from", ""))
    to_addr = str(email_data.get("to", ""))
    attachments = " ".join(email_data.get("attachments", []))
    attachment_text = str(email_data.get("attachment_text", ""))
    
    all_content = f"{subject}\n{body}\n{from_addr}\n{to_addr}\n{attachments}\n{attachment_text}"
    all_content_lower = all_content.lower()
    
    column_lower = column_name.lower().strip()
    
    # ============================================================
    # EXPLICIT TYPE HANDLING - User selected a specific type
    # This prevents "Yes" being returned when expecting a number
    # ============================================================
    
    if extract_type == "number":
        # USER WANTS NUMBERS ONLY - never return "Yes"
        return _extract_number_only(column_name, all_content)
    
    elif extract_type == "date":
        # USER WANTS DATES ONLY
        email_date = email_data.get("date", "")
        if email_date:
            return str(email_date)
        date_pattern = r'\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4}|\d{4}[/\-]\d{1,2}[/\-]\d{1,2}'
        match = re.search(date_pattern, all_content)
        return match.group() if match else ""
    
    elif extract_type == "amount":
        # USER WANTS CURRENCY/AMOUNTS ONLY
        amount_pattern = r'[\$€£]\s*[\d,]+\.?\d*|\d{1,3}(?:,\d{3})*(?:\.\d{2})?'
        match = re.search(amount_pattern, all_content)
        return match.group() if match else ""
    
    elif extract_type == "yesno":
        # USER WANTS YES/NO - check keyword presence
        pattern = r'\b' + re.escape(column_lower) + r'\b'
        return "Yes" if re.search(pattern, all_content_lower) else "No"
    
    elif extract_type == "text":
        # USER WANTS TEXT - extract value after "Column: value"
        extract_pattern = rf'{re.escape(column_name)}\s*[:\-=]\s*([^\n,;]+)'
        match = re.search(extract_pattern, all_content, re.IGNORECASE)
        if match:
            value = match.group(1).strip()
            if 0 < len(value) < 200:
                return value
        return ""
    
    elif extract_type == "email_field":
        # USER WANTS STANDARD EMAIL FIELD
        field_map = {
            'subject': subject, 'from': from_addr, 'sender': from_addr,
            'to': to_addr, 'recipient': to_addr, 'date': email_data.get("date", ""),
            'body': body, 'content': body, 'message': body,
            'attachments': ", ".join(email_data.get("attachments", []))
        }
        for key, value in field_map.items():
            if key in column_lower:
                return str(value) if value else ""
        return ""
    
    # ============================================================
    # AUTO-DETECT MODE (default) - Smart detection from column name
    # ============================================================
    
    number_indicators = ['#', 'number', 'no.', 'no', 'id', 'num', 'code', 'ref', 'reference']
    is_number_column = any(ind in column_lower for ind in number_indicators)
    
    # Also check if column is short uppercase (likely abbreviation like MI, QE, MM)
    column_stripped = re.sub(r'[#\s\.\-_]+', '', column_name)
    is_abbreviation = len(column_stripped) <= 5 and column_stripped.isupper()
    
    # If it looks like a number column, extract numbers (no "Yes" fallback)
    if is_number_column or is_abbreviation:
        result = _extract_number_only(column_name, all_content)
        if result:
            return result
        # Don't fall back to "Yes" for number columns!
        return ""
    
    # Date columns
    if any(kw in column_lower for kw in ['date', 'time', 'when', 'received', 'sent', 'due']):
        email_date = email_data.get("date", "")
        if email_date:
            return str(email_date)
        date_pattern = r'\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4}|\d{4}[/\-]\d{1,2}[/\-]\d{1,2}'
        match = re.search(date_pattern, all_content)
        if match:
            return match.group()
    
    # Amount/Money columns
    if any(kw in column_lower for kw in ['amount', 'total', 'price', 'cost', 'qty', 'quantity', 'dollar', 'payment']):
        amount_pattern = r'[\$€£]\s*[\d,]+\.?\d*|\d{1,3}(?:,\d{3})*(?:\.\d{2})?'
        match = re.search(amount_pattern, all_content)
        if match:
            return match.group()
    
    # Priority columns
    if any(kw in column_lower for kw in ['priority', 'urgent', 'importance']):
        if any(kw in all_content_lower for kw in ['urgent', 'asap', 'critical', 'high priority', 'important', 'immediately']):
            return "High"
        elif any(kw in all_content_lower for kw in ['low priority', 'fyi', 'when possible', 'no rush']):
            return "Low"
        return "Normal"
    
    # Status columns
    if any(kw in column_lower for kw in ['status', 'state']):
        if any(kw in all_content_lower for kw in ['complete', 'done', 'finished', 'resolved', 'closed', 'approved']):
            return "Complete"
        elif any(kw in all_content_lower for kw in ['pending', 'waiting', 'in progress', 'ongoing', 'review']):
            return "Pending"
        elif any(kw in all_content_lower for kw in ['open', 'new', 'unread', 'submitted']):
            return "Open"
        elif any(kw in all_content_lower for kw in ['reject', 'denied', 'cancel']):
            return "Rejected"
    
    # ============================================================
    # STEP 3: LOOK FOR "COLUMN_NAME: VALUE" PATTERNS
    # ============================================================
    
    # Try exact column name match first
    extract_pattern = rf'{re.escape(column_name)}\s*[:\-=]\s*([^\n,;]+)'
    match = re.search(extract_pattern, all_content, re.IGNORECASE)
    if match:
        value = match.group(1).strip()
        if 0 < len(value) < 100:
            return value
    
    # Try variations (singular/plural, with/without spaces)
    column_words = column_name.split()
    if len(column_words) >= 1:
        # Try first word
        first_word = column_words[0]
        if len(first_word) >= 2:
            pattern = rf'{re.escape(first_word)}\s*[:\-=]\s*([^\n,;]+)'
            match = re.search(pattern, all_content, re.IGNORECASE)
            if match:
                value = match.group(1).strip()
                if 0 < len(value) < 100:
                    return value
    
    # ============================================================
    # STEP 4: KEYWORD EXISTENCE CHECK (Yes/No)
    # ============================================================
    
    # Word boundary search for the column name
    search_terms = [column_name]
    # Also try individual words if column has multiple words
    if ' ' in column_name:
        search_terms.extend(column_name.split())
    
    for term in search_terms:
        if len(term) >= 2:
            pattern = r'\b' + re.escape(term.lower()) + r'\b'
            if re.search(pattern, all_content_lower):
                return "Yes"
    
    return ""  # Not found


def extract_prefix_numbers(prefix: str, content: str) -> str:
    """
    Extract all numbers associated with a prefix from content.
    
    Works with ANY prefix:
    - "MI" → finds MI-12345, MI #001, MI: 123, MI123
    - "Building" → finds Building 5, Building-12, Building: 100
    - "Work Order" → finds Work Order 12345, WO-001
    
    Args:
        prefix: The identifier prefix (e.g., "MI", "Building", "Work Order")
        content: The email content to search
    
    Returns:
        Semicolon-separated list of found references
    """
    extracted = []
    prefix_clean = prefix.strip()
    
    # Build search variations
    variations = [prefix_clean]
    
    # Add common abbreviations for multi-word prefixes
    if ' ' in prefix_clean:
        words = prefix_clean.split()
        # Add acronym (first letters)
        acronym = ''.join(w[0].upper() for w in words if w)
        if len(acronym) >= 2:
            variations.append(acronym)
        # Add first word
        variations.append(words[0])
    
    # Search for each variation
    for var in variations:
        # Multiple patterns to catch different formats
        patterns = [
            # PREFIX followed by separator then numbers: MI-12345, MI #001, MI: 123
            rf'(?<![A-Za-z]){re.escape(var)}\s*[#:\-_/\.]*\s*(\d[\d\-]*\d|\d+)',
            # PREFIX immediately followed by numbers: MI12345
            rf'(?<![A-Za-z]){re.escape(var)}(\d+)',
        ]
        
        for pattern in patterns:
            try:
                matches = re.findall(pattern, content, re.IGNORECASE)
                for match in matches:
                    # Clean and format the result
                    num = match.strip('-').strip()
                    if num and len(num) >= 1:
                        # Format as PREFIX-NUMBER
                        formatted = f"{var.upper()}-{num}"
                        formatted = re.sub(r'-+', '-', formatted)  # Clean double dashes
                        if formatted not in extracted:
                            extracted.append(formatted)
            except re.error:
                continue
    
    # Return unique results, limited to 10
    if extracted:
        return "; ".join(extracted[:10])
    
    return ""


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
        schema: Schema dict with {columns: [{name, type, extract_type}, ...]}
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
        col_type = col.get("extract_type", col.get("type", "auto"))  # Get column type
        
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
        
        # Priority 3: Smart search - use column name and TYPE for extraction
        else:
            smart_value = smart_search_column(email_data, col_name, col_type)
            record[col_name] = smart_value
            
            if explain and smart_value:
                explanations[col_name] = {
                    "matched": True,
                    "match_type": "smart_search",
                    "matched_term": col_name,
                    "extract_type": col_type,
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
