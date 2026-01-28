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


# Confidence helpers -------------------------------------------------
def _confidence(source: str, matched: bool = True) -> int:
    """
    Simple confidence scale (0-10):
      direct              -> 10
      standard_field      -> 9
      rule                -> 8
      synonym_direct      -> 7
      smart_heading       -> 6
      smart_fallback      -> 3
      missing / no match  -> 0
    """
    if not matched:
        return 0
    scale = {
        "direct": 10,
        "standard_field": 9,
        "rule": 8,
        "synonym_direct": 7,
        "smart_heading": 6,
        "smart_fallback": 3,
    }
    return scale.get(source, 5)


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


def _ai_assist(email_data: Dict, column_name: str, extract_type: str) -> Tuple[str, Dict[str, int]]:
    """
    Placeholder for optional AI assistance. Currently returns no value but
    preserves the shape for future integration with a real model/API.
    """
    # TODO: integrate with external AI provider if configured.
    return "", {"source": "ai_assist", "confidence": 0}


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
                "confidence": _confidence("rule", matched),
                "source": "rule",
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


def smart_search_column(email_data: Dict, column_name: str, extract_type: str = "auto", use_synonyms: bool = True) -> Tuple[str, Dict[str, int]]:
    """
    FULLY INTELLIGENT column extraction - works with ANY column name!
    
    Now supports:
    - SYNONYMS (optional)
    - Explicit type hints that bias the search
    - Confidence metadata for each match

    Returns:
        (value, {"source": str, "confidence": int})
    """
    def make_meta(source: str, confidence: int) -> Dict[str, int]:
        return {"source": source, "confidence": confidence}

    # Check if we're focusing on a specific order number
    focus_order = email_data.get("_focus_order")
    
    # Collect contextual text
    subject = str(email_data.get("subject", ""))
    body = str(email_data.get("body", ""))
    from_addr = str(email_data.get("from", ""))
    to_addr = str(email_data.get("to", ""))
    attachments = " ".join(email_data.get("attachments", []))
    attachment_text = str(email_data.get("attachment_text", ""))
    
    all_content = f"{subject}\n{body}\n{from_addr}\n{to_addr}\n{attachments}\n{attachment_text}"
    
    # If focusing on specific order, extract context around it
    if focus_order:
        # Find the order number in content and extract surrounding context (400 chars before/after)
        pattern = rf'.{{0,400}}{re.escape(focus_order)}.{{0,400}}'
        match = re.search(pattern, all_content, re.IGNORECASE)
        if match:
            all_content = match.group(0)  # Use only context around this order
    
    all_content_lower = all_content.lower()
    column_lower = column_name.lower().strip()
    column_synonyms = get_synonyms(column_name) if use_synonyms else [column_name]

    # Type-specific extractors
    if extract_type == "number":
        for name in [column_name] + column_synonyms:
            result = _extract_number_only(name, all_content)
            if result:
                return result, make_meta("smart_number", 7)
        return "", make_meta("smart_number", 0)

    if extract_type == "mixed":
        for name in [column_name] + (column_synonyms if use_synonyms else []):
            patterns = [
                rf'{re.escape(name)}\s*[:\-=]\s*([A-Za-z0-9][\w\s\-]*)',
                rf'{re.escape(name)}\t+([^\t\n]+)',
            ]
            for pattern in patterns:
                try:
                    match = re.search(pattern, all_content, re.IGNORECASE)
                    if match:
                        value = re.sub(r'\s+', ' ', match.group(1).strip())
                        if value:
                            return value, make_meta("smart_label_match", 7)
                except re.error:
                    continue
        alphanumeric_patterns = [
            r'\b(\d{6,}\s*[A-Z]{2,}\d*)\b',
            r'\b([A-Z]{2,}\d{3,})\b',
            r'\b([A-Z]\d{2,})\b',
        ]
        for pattern in alphanumeric_patterns:
            matches = re.findall(pattern, all_content)
            if matches:
                unique = list(dict.fromkeys(matches))
                return "; ".join(unique[:5]), make_meta("smart_alphanumeric", 6)
        return "", make_meta("smart_alphanumeric", 0)

    if extract_type == "date":
        email_date = email_data.get("date", "")
        if email_date:
            return str(email_date), make_meta("smart_date", 8)
        pattern = r'\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4}|\d{4}[/\-]\d{1,2}[/\-]\d{1,2}'
        match = re.search(pattern, all_content)
        return (match.group(), make_meta("smart_date", 7)) if match else ("", make_meta("smart_date", 0))

    if extract_type == "time":
        patterns = [
            r'\b(\d{1,2}:\d{2}(?::\d{2})?\s*(?:AM|PM|am|pm)?)\b',
            r'\b(\d{1,2}\s*(?:AM|PM|am|pm))\b',
        ]
        for pattern in patterns:
            match = re.search(pattern, all_content)
            if match:
                return match.group(1), make_meta("smart_time", 6)
        return "", make_meta("smart_time", 0)

    if extract_type == "amount":
        amount_pattern = r'[\$€£]\s*[\d,]+\.?\d*|\d{1,3}(?:,\d{3})*(?:\.\d{2})?'
        match = re.search(amount_pattern, all_content)
        return (match.group(), make_meta("smart_amount", 7)) if match else ("", make_meta("smart_amount", 0))

    if extract_type == "yesno":
        pattern = r'\b' + re.escape(column_lower) + r'\b'
        return ("Yes", make_meta("smart_yesno", 5)) if re.search(pattern, all_content_lower) else ("No", make_meta("smart_yesno", 5))

    if extract_type == "text":
        candidates = [column_name] + (column_synonyms if use_synonyms else [])
        for name in candidates:
            patterns = [
                rf'{re.escape(name)}\s*[:\-=]\s*([^\n\t]+)',
                rf'{re.escape(name)}\t+([^\t\n]+)',
                rf'{re.escape(name)}\s{{2,}}([^\s].*?)(?:\s{{2,}}|$)',
            ]
            for pattern in patterns:
                try:
                    match = re.search(pattern, all_content, re.IGNORECASE)
                    if match:
                        value = re.sub(r'\s+', ' ', match.group(1).strip())
                        if value:
                            return value, make_meta("smart_label_text", 6)
                except re.error:
                    continue
        return "", make_meta("smart_text", 0)

    number_indicators = ['#', 'number', 'no.', 'no', 'id', 'num', 'code', 'ref', 'reference']
    is_number_column = any(ind in column_lower for ind in number_indicators)
    column_stripped = re.sub(r'[#\s\.\-_]+', '', column_name)
    is_abbreviation = len(column_stripped) <= 5 and column_stripped.isupper()
    if is_number_column or is_abbreviation:
        result = _extract_number_only(column_name, all_content)
        if result:
            return result, make_meta("smart_number", 7)
        return "", make_meta("smart_number", 0)

    if any(kw in column_lower for kw in ['time', 'hour', 'clock', 'start time', 'end time', 'meeting time']):
        patterns = [
            r'\b(\d{1,2}:\d{2}(?::\d{2})?\s*(?:AM|PM|am|pm)?)\b',
            r'\b(\d{1,2}\s*(?:AM|PM|am|pm))\b',
        ]
        for pattern in patterns:
            match = re.search(pattern, all_content)
            if match:
                return match.group(1), make_meta("smart_time", 6)

    if any(kw in column_lower for kw in ['date', 'when', 'received', 'sent', 'due']):
        email_date = email_data.get("date", "")
        if email_date:
            return str(email_date), make_meta("smart_date", 8)
        pattern = r'\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4}|\d{4}[/\-]\d{1,2}[/\-]\d{1,2}'
        match = re.search(pattern, all_content)
        if match:
            return match.group(), make_meta("smart_date", 7)

    if any(kw in column_lower for kw in ['amount', 'total', 'price', 'cost', 'qty', 'quantity', 'dollar', 'payment']):
        amount_pattern = r'[\$€£]\s*[\d,]+\.?\d*|\d{1,3}(?:,\d{3})*(?:\.\d{2})?'
        match = re.search(amount_pattern, all_content)
        if match:
            return match.group(), make_meta("smart_amount", 6)

    if any(kw in column_lower for kw in ['priority', 'urgent', 'importance']):
        if any(kw in all_content_lower for kw in ['urgent', 'asap', 'critical', 'high priority', 'important', 'immediately']):
            return "High", make_meta("smart_priority", 6)
        if any(kw in all_content_lower for kw in ['low priority', 'fyi', 'when possible', 'no rush']):
            return "Low", make_meta("smart_priority", 6)
        return "Normal", make_meta("smart_priority", 4)

    if any(kw in column_lower for kw in ['status', 'state']):
        if any(kw in all_content_lower for kw in ['complete', 'done', 'finished', 'resolved', 'closed', 'approved']):
            return "Complete", make_meta("smart_status", 6)
        if any(kw in all_content_lower for kw in ['pending', 'waiting', 'in progress', 'ongoing', 'review']):
            return "Pending", make_meta("smart_status", 5)
        if any(kw in all_content_lower for kw in ['open', 'new', 'unread', 'submitted']):
            return "Open", make_meta("smart_status", 5)
        if any(kw in all_content_lower for kw in ['reject', 'denied', 'cancel']):
            return "Rejected", make_meta("smart_status", 5)

    search_names = [column_name] + (column_synonyms if use_synonyms else [])
    for name in search_names:
        pattern = rf'{re.escape(name)}\s*[:\-=]\s*([^\n,;]+)'
        match = re.search(pattern, all_content, re.IGNORECASE)
        if match:
            value = re.sub(r'\s+', ' ', match.group(1).strip())
            if value:
                return value, make_meta("smart_label_match", 5)

    if ' ' in column_name:
        first_word = column_name.split()[0]
        if len(first_word) >= 2:
            pattern = rf'{re.escape(first_word)}\s*[:\-=]\s*([^\n,;]+)'
            match = re.search(pattern, all_content, re.IGNORECASE)
            if match:
                value = re.sub(r'\s+', ' ', match.group(1).strip())
                if value:
                    return value, make_meta("smart_label_match", 4)

    search_terms = [column_name]
    if ' ' in column_name:
        search_terms.extend(column_name.split())
    for term in search_terms:
        if len(term) >= 2:
            pattern = r'\b' + re.escape(term.lower()) + r'\b'
            if re.search(pattern, all_content_lower):
                return "Yes", make_meta("smart_keyword", 3)

    return "", make_meta("smart_search", 0)


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


def expand_order_numbers_to_rows(
    email_data: Dict,
    schema: Dict,
    rules: List[Dict],
    explain: bool,
    use_synonyms: bool,
    use_ai: bool,
    ai_threshold: int
) -> List[Tuple[Dict[str, str], Dict]]:
    """
    When multiple order numbers are found, create separate records for each.
    This ensures each order number gets its own row with properly filled columns.
    
    Returns:
        List of (record, explanation) tuples, one per order number found
    """
    # First, check if we have table data (excel_rows or body table)
    excel_rows = email_data.get("excel_rows", [])
    if excel_rows:
        # Already expanded - process each row separately
        records = []
        for row_data in excel_rows:
            merged_data = dict(email_data)
            merged_data.update(row_data)
            record, explanation = email_to_record(
                merged_data, schema, rules, explain, use_synonyms, use_ai, ai_threshold
            )
            records.append((record, explanation))
        return records
    
    # Otherwise, try to find order numbers and expand them
    # Find the "order" or "work order" column
    order_column = None
    for col in schema.get("columns", []):
        col_name = col.get("name", "").lower()
        if col_name in ["order", "work order", "wo", "wo number", "order number"]:
            order_column = col.get("name")
            break
    
    if not order_column:
        # No order column - return single record as normal
        record, explanation = email_to_record(
            email_data, schema, rules, explain, use_synonyms, use_ai, ai_threshold
        )
        return [(record, explanation)]
    
    # Extract order numbers - but DON'T concatenate!
    all_content = f"{email_data.get('subject', '')}\n{email_data.get('body', '')}\n{' '.join(email_data.get('attachment_text', []))}"
    
    # Try to find order numbers in table format first
    order_numbers = []
    
    # Strategy 1: Look for table rows with order numbers
    # Pattern: Order | Other columns... or Order: value
    table_patterns = [
        rf'{re.escape(order_column)}\s*[:\t|]\s*([^\n\t|]+)',  # Order: 4000390042 or Order | 4000390042
        rf'{re.escape(order_column)}\s+(\d{{6,}})',  # Order    4000390042
    ]
    
    for pattern in table_patterns:
        matches = re.findall(pattern, all_content, re.IGNORECASE | re.MULTILINE)
        for match in matches:
            order_val = match.strip().split()[0] if isinstance(match, str) else str(match).strip().split()[0]
            # Clean up - take first meaningful part
            order_val = re.split(r'\s{2,}|\t|;', order_val)[0].strip()
            if order_val and len(order_val) >= 4 and order_val not in order_numbers:
                order_numbers.append(order_val)
    
    # Strategy 2: Look for standalone order number patterns (if no table matches)
    if not order_numbers:
        patterns = [
            r'\b(\d{6,})\b',  # 4000390042
            r'\b(\d{6,}\s+[A-Z]{2,}\d*)\b',  # 4000390042 ZM03
        ]
        for pattern in patterns:
            matches = re.findall(pattern, all_content)
            if matches:
                for m in matches:
                    order_val = m.strip() if isinstance(m, str) else str(m).strip()
                    if len(order_val) >= 6 and order_val not in order_numbers:
                        order_numbers.append(order_val)
    
    # Remove duplicates but preserve order
    order_numbers = list(dict.fromkeys(order_numbers))
    
    if len(order_numbers) <= 1:
        # Single or no order number - return normal record
        record, explanation = email_to_record(
            email_data, schema, rules, explain, use_synonyms, use_ai, ai_threshold
        )
        return [(record, explanation)]
    
    # Multiple order numbers found - create separate records
    records = []
    for order_num in order_numbers:
        # Create a modified email_data with this specific order number
        # This helps extraction find related fields for THIS order number
        modified_data = dict(email_data)
        
        # Add order number to body context (helps smart_search find related fields)
        modified_data["_focus_order"] = order_num
        
        # Extract record
        record, explanation = email_to_record(
            modified_data, schema, rules, explain, use_synonyms, use_ai, ai_threshold
        )
        
        # Force the order number into the record
        record[order_column] = order_num
        
        records.append((record, explanation))
    
    return records


def email_to_record(
    email_data: Dict,
    schema: Dict,
    rules: List[Dict],
    explain: bool = False,
    use_synonyms: bool = True,
    use_ai: bool = False,
    ai_threshold: int = 0
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
                exact_match = matched_key and matched_key.lower() == col_name.lower()
                source_tag = "direct" if exact_match else "synonym_direct"
                explanations[col_name] = {
                    "matched": True,
                    "match_type": source_tag,
                    "matched_term": matched_key or col_name,
                    "search_in": ["data_row"],
                    "snippet": str(direct_value)[:50],
                    "confidence": _confidence("direct" if exact_match else "synonym_direct"),
                    "source": source_tag,
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
                    "confidence": 9,
                    "source": "standard_field",
                }
        
        # Priority 3: Explicit rule results
        elif col_name in rule_results and rule_results[col_name]:
            record[col_name] = rule_results[col_name]
            if explain:
                explanations[col_name] = {
                    "matched": True,
                    "match_type": "rule",
                    "matched_term": col_name,
                    "search_in": ["rules"],
                    "snippet": str(rule_results[col_name])[:50],
                    "confidence": _confidence("rule"),
                    "source": "rule",
                }
        
        # Priority 4: Smart search - use column name and TYPE for extraction
        else:
            smart_value, smart_meta = smart_search_column(email_data, col_name, col_type, use_synonyms)
            smart_conf = smart_meta.get("confidence", 0)
            smart_source = smart_meta.get("source", "smart_search")
            smart_matched = bool(smart_value)

            # Optional AI assist if low confidence
            if use_ai and ai_threshold and smart_conf < ai_threshold:
                ai_value, ai_meta = _ai_assist(email_data, col_name, col_type)
                if ai_value:
                    smart_value = ai_value
                    smart_conf = ai_meta.get("confidence", smart_conf)
                    smart_source = ai_meta.get("source", "ai_assist")
                    smart_matched = True

            record[col_name] = smart_value
            
            if explain:
                explanations[col_name] = {
                    "matched": smart_matched,
                    "match_type": smart_source,
                    "matched_term": col_name,
                    "extract_type": col_type,
                    "search_in": ["all"],
                    "snippet": f"Found '{col_name}' in email content" if smart_matched else None,
                    "confidence": smart_conf,
                    "source": smart_source,
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
