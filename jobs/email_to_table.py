"""
Core job: Extract emails into tabular data based on schema and rules.
"""
import re
from typing import Dict, List, Tuple

# Import synonym resolver
try:
    from .synonym_resolver import get_synonyms, columns_match, get_resolver
except ImportError:
    from jobs.synonym_resolver import get_synonyms, columns_match, get_resolver


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
    Extract numbers/IDs - handles various formats:
    - "MI #12345" → MI-12345
    - "4000390042 ZM03" → 4000390042 ZM03 (number + code)
    - "TORFER0048" → TORFER0048 (letters + numbers mixed)
    - "B89" → B89 (short code)
    
    NEVER returns "Yes".
    """
    column_lower = column_name.lower().strip()
    
    # Get the prefix (part before number/id/#/etc.)
    prefix = column_name
    for suffix in ['number', 'num', '#', 'no.', 'no', 'id', 'code', 'ref', 'reference']:
        prefix = re.sub(rf'\s*{re.escape(suffix)}\s*$', '', prefix, flags=re.IGNORECASE)
    prefix = prefix.strip()
    
    # STRATEGY 1: Look for "ColumnName: value" pattern first (most reliable)
    patterns_with_label = [
        rf'{re.escape(column_name)}\s*[:\-=]\s*([^\n,;]+)',  # Order: 4000390042 ZM03
        rf'{re.escape(prefix)}\s*[:\-=]\s*([^\n,;]+)',  # Order: value
    ]
    
    for pattern in patterns_with_label:
        match = re.search(pattern, content, re.IGNORECASE)
        if match:
            value = match.group(1).strip()
            # Clean up - remove trailing punctuation but keep alphanumeric + spaces
            value = re.sub(r'[,;:\s]+$', '', value)
            if value and len(value) < 50:
                return value
    
    # STRATEGY 2: If we have a prefix, search for PREFIX + NUMBER
    if prefix:
        extracted = extract_prefix_numbers(prefix, content)
        if extracted:
            return extracted
    
    # STRATEGY 3: If column name is just a short code (MI, QE, MM), search directly
    column_stripped = re.sub(r'[#\s\.\-_]+', '', column_name)
    if len(column_stripped) <= 6 and column_stripped.isupper():
        extracted = extract_prefix_numbers(column_stripped, content)
        if extracted:
            return extracted
    
    # STRATEGY 4: Look for alphanumeric codes that match the column pattern
    # e.g., column "Equipment" might have values like "TORFER0048"
    if column_lower in ['equipment', 'order', 'work order', 'building', 'code', 'id', 'reference']:
        # Look for alphanumeric patterns: letters+numbers or numbers+letters
        alphanumeric_patterns = [
            r'\b([A-Z]{2,}[0-9]{2,})\b',  # TORFER0048, ZM03
            r'\b([0-9]{6,}\s*[A-Z]{2,}[0-9]*)\b',  # 4000390042 ZM03
            r'\b([A-Z][0-9]{2,})\b',  # B89, M123
        ]
        for pattern in alphanumeric_patterns:
            matches = re.findall(pattern, content)
            if matches:
                # Return first few unique matches
                unique = list(dict.fromkeys(matches))[:5]
                return "; ".join(unique)
    
    return ""


def smart_search_column(email_data: Dict, column_name: str, extract_type: str = "auto", use_synonyms: bool = True) -> str:
    """
    FULLY INTELLIGENT column extraction - works with ANY column name!
    
    Now supports:
    - SYNONYMS: "QC" matches "Quality Criticality" (can be disabled)
    - Explicit TYPE to prevent wrong extractions
    - Row association: values on same line stay together
    
    Types:
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
        extract_type: The expected value type
        use_synonyms: If True, use synonym matching
    
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
    
    # Get synonyms only if enabled
    column_synonyms = get_synonyms(column_name) if use_synonyms else [column_name]
    
    # ============================================================
    # EXPLICIT TYPE HANDLING - User selected a specific type
    # This prevents "Yes" being returned when expecting a number
    # ============================================================
    
    if extract_type == "number":
        # USER WANTS NUMBERS ONLY - never return "Yes"
        # Try column name and all synonyms
        for name in [column_name] + column_synonyms:
            result = _extract_number_only(name, all_content)
            if result:
                return result
        return ""
    
    elif extract_type == "date":
        # USER WANTS DATES ONLY
        email_date = email_data.get("date", "")
        if email_date:
            return str(email_date)
        date_pattern = r'\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4}|\d{4}[/\-]\d{1,2}[/\-]\d{1,2}'
        match = re.search(date_pattern, all_content)
        return match.group() if match else ""
    
    elif extract_type == "time":
        # USER WANTS TIME ONLY
        # Look for time patterns: 10:30, 2:45 PM, 14:30:00, etc.
        time_patterns = [
            r'\b(\d{1,2}:\d{2}(?::\d{2})?\s*(?:AM|PM|am|pm)?)\b',  # 10:30 AM, 2:45pm, 14:30:00
            r'\b(\d{1,2}\s*(?:AM|PM|am|pm))\b',  # 10 AM, 2pm
        ]
        for pattern in time_patterns:
            match = re.search(pattern, all_content)
            if match:
                return match.group(1)
        return ""
    
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
        # USER WANTS TEXT - extract value after "Column: value" or from table
        # Uses SYNONYMS: "QC" matches "Quality Criticality"
        
        # Build patterns for column name AND all its synonyms
        search_names = [column_name] + column_synonyms
        
        for name in search_names:
            patterns = [
                # Standard "Label: Value" format
                rf'{re.escape(name)}\s*[:\-=]\s*([^\n\t]+)',
                # Table format with tab separator
                rf'{re.escape(name)}\t+([^\t\n]+)',
                # Table format with multiple spaces
                rf'{re.escape(name)}\s{{2,}}([^\s].*?)(?:\s{{2,}}|$)',
            ]
            
            for pattern in patterns:
                try:
                    match = re.search(pattern, all_content, re.IGNORECASE)
                    if match:
                        value = match.group(1).strip()
                        # Clean up extra whitespace but keep the value
                        value = re.sub(r'\s+', ' ', value).strip()
                        if 0 < len(value) < 200:
                            return value
                except re.error:
                    continue
        
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
    
    # Time columns (check first, before date)
    if any(kw in column_lower for kw in ['time', 'hour', 'clock', 'start time', 'end time', 'meeting time']):
        time_patterns = [
            r'\b(\d{1,2}:\d{2}(?::\d{2})?\s*(?:AM|PM|am|pm)?)\b',
            r'\b(\d{1,2}\s*(?:AM|PM|am|pm))\b',
        ]
        for pattern in time_patterns:
            match = re.search(pattern, all_content)
            if match:
                return match.group(1)
    
    # Date columns
    if any(kw in column_lower for kw in ['date', 'when', 'received', 'sent', 'due']):
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
    Extract all numbers/codes associated with a prefix from content.
    
    Handles various formats:
    - "MI" → finds MI-12345, MI #001, MI: 123, MI123
    - "Building" → finds Building 5, Building-12, B89
    - "Order" → finds 4000390042 ZM03, Order: 12345
    - "Equipment" → finds TORFER0048, EQ-123
    
    Args:
        prefix: The identifier prefix (e.g., "MI", "Building", "Order")
        content: The email content to search
    
    Returns:
        Semicolon-separated list of found references
    """
    extracted = []
    prefix_clean = prefix.strip()
    prefix_lower = prefix_clean.lower()
    
    # FIRST: Try to find "Prefix: value" or "Prefix\tvalue" patterns (for table data)
    table_patterns = [
        rf'{re.escape(prefix_clean)}\s*[:\t]\s*([^\n\t]+)',  # Order: 4000390042 ZM03
        rf'^{re.escape(prefix_clean)}\s+(.+)$',  # Order    4000390042
    ]
    
    for pattern in table_patterns:
        matches = re.findall(pattern, content, re.IGNORECASE | re.MULTILINE)
        for match in matches:
            value = match.strip()
            # Clean up - take first meaningful part
            value = re.split(r'\s{2,}|\t', value)[0].strip()
            if value and len(value) < 50 and value not in extracted:
                extracted.append(value)
    
    if extracted:
        return "; ".join(extracted[:5])
    
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
            # PREFIX followed by separator then numbers/alphanumeric: MI-12345, MI #001
            rf'(?<![A-Za-z]){re.escape(var)}\s*[#:\-_/\.]*\s*(\d[\d\-A-Za-z]*)',
            # PREFIX immediately followed by numbers: MI12345
            rf'(?<![A-Za-z]){re.escape(var)}(\d+)',
            # Number followed by PREFIX: 4000390042 ZM03
            rf'(\d{{6,}}\s*{re.escape(var)}[0-9]*)',
            # Short letter-number codes: B89, M01
            rf'\b({re.escape(var[0]) if var else ""}[0-9]{{2,}})\b',
        ]
        
        for pattern in patterns:
            try:
                matches = re.findall(pattern, content, re.IGNORECASE)
                for match in matches:
                    # Clean and format the result
                    value = match.strip('-').strip()
                    if value and len(value) >= 1:
                        # Don't double-prefix
                        if not value.upper().startswith(var.upper()):
                            formatted = f"{var.upper()}-{value}"
                        else:
                            formatted = value
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
    explain: bool = False,
    use_synonyms: bool = True
) -> Tuple[Dict[str, str], Dict]:
    """
    Convert email to record dict (NEW canonical format).
    
    Args:
        email_data: Email dict
        schema: Schema dict with {columns: [{name, type, extract_type}, ...]}
        rules: List of keyword/regex rules
        explain: If True, return explanation of which rules matched
        use_synonyms: If True, use synonym matching (QC = Quality Criticality)
    
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
        
        # Priority 1: Direct value from data (Excel/CSV already has values in columns!)
        # Check if email_data already has this column's value directly
        # Optionally uses synonym matching: "QC" matches "Quality Criticality"
        direct_value = None
        matched_key = None
        
        # Get synonyms if enabled
        col_synonyms = get_synonyms(col_name) if use_synonyms else [col_name]
        
        for key in email_data.keys():
            # Direct match (always check)
            if key.lower() == col_name.lower() or key == col_name:
                direct_value = email_data.get(key, "")
                matched_key = key
                break
            
            # Synonym match (only if enabled)
            if use_synonyms:
                if columns_match(key, col_name):
                    direct_value = email_data.get(key, "")
                    matched_key = key
                    break
                
                # Check against all synonyms
                for syn in col_synonyms:
                    if key.lower() == syn.lower() or syn.lower() in key.lower() or key.lower() in syn.lower():
                        direct_value = email_data.get(key, "")
                        matched_key = key
                        break
                
                if direct_value:
                    break
        
        if direct_value and str(direct_value).strip():
            record[col_name] = str(direct_value).strip()
            if explain:
                explanations[col_name] = {
                    "matched": True,
                    "match_type": "direct_column",
                    "matched_term": col_name,
                    "search_in": ["data_row"],
                    "snippet": str(direct_value)[:50],
                }
        
        # Priority 2: Standard email fields (Subject, From, To, Date, Body, etc.)
        elif col_name in standard_fields and standard_fields[col_name]:
            record[col_name] = standard_fields[col_name]
            if explain:
                explanations[col_name] = {
                    "matched": True,
                    "match_type": "standard_field",
                    "matched_term": col_name,
                    "search_in": ["email_headers"],
                    "snippet": str(standard_fields[col_name])[:50],
                }
        
        # Priority 3: Explicit rule results
        elif col_name in rule_results and rule_results[col_name]:
            record[col_name] = rule_results[col_name]
        
        # Priority 4: Smart search - use column name and TYPE for extraction
        else:
            smart_value = smart_search_column(email_data, col_name, col_type, use_synonyms)
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
