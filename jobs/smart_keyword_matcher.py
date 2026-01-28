"""
Smart keyword matching - finds dates, numbers, and semantic matches.
"""
import re
from typing import List, Dict, Any
from datetime import datetime


def extract_dates(text: str) -> List[str]:
    """
    Extract all date patterns from text.
    
    Supports formats:
    - 2024-01-15, 01/15/2024, 15-Jan-2024
    - January 15, 2024
    - 15th January 2024
    """
    dates = []
    
    # Common date patterns
    patterns = [
        r'\d{4}-\d{2}-\d{2}',  # 2024-01-15
        r'\d{2}/\d{2}/\d{4}',  # 01/15/2024
        r'\d{2}-\d{2}-\d{4}',  # 01-15-2024
        r'\d{1,2}[/-]\w{3}[/-]\d{4}',  # 15-Jan-2024
        r'\w+ \d{1,2},? \d{4}',  # January 15, 2024
        r'\d{1,2}(?:st|nd|rd|th)? \w+ \d{4}',  # 15th January 2024
    ]
    
    for pattern in patterns:
        matches = re.findall(pattern, text, re.IGNORECASE)
        dates.extend(matches)
    
    return list(set(dates))  # Remove duplicates


def extract_numbers(text: str, include_currency: bool = True) -> List[str]:
    """
    Extract all numbers from text.
    
    Args:
        text: Text to search
        include_currency: Include currency amounts ($100, €50)
    
    Returns:
        List of found numbers
    """
    numbers = []
    
    # Currency patterns
    if include_currency:
        currency_patterns = [
            r'\$[\d,]+\.?\d*',  # $1,234.56
            r'€[\d,]+\.?\d*',  # €1,234.56
            r'£[\d,]+\.?\d*',  # £1,234.56
            r'[\d,]+\.?\d*\s*(?:USD|EUR|GBP)',  # 1,234.56 USD
        ]
        
        for pattern in currency_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            numbers.extend(matches)
    
    # Regular numbers
    number_patterns = [
        r'\d+\.\d+',  # Decimals: 123.45
        r'\d{1,3}(?:,\d{3})+',  # Thousands: 1,234,567
        r'\b\d+\b',  # Integers: 123
    ]
    
    for pattern in number_patterns:
        matches = re.findall(pattern, text)
        numbers.extend(matches)
    
    return list(set(numbers))


def extract_emails(text: str) -> List[str]:
    """Extract all email addresses from text."""
    pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    return re.findall(pattern, text)


def extract_phones(text: str) -> List[str]:
    """Extract all phone numbers from text."""
    patterns = [
        r'\+?\d{1,3}[-.\s]?\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}',  # +1-555-123-4567
        r'\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}',  # (555) 123-4567
    ]
    
    phones = []
    for pattern in patterns:
        matches = re.findall(pattern, text)
        phones.extend(matches)
    
    return list(set(phones))


def smart_keyword_match(content: str, keyword: str) -> Dict[str, Any]:
    """
    Smart matching for a keyword - finds exact match + related data.
    
    Args:
        content: Text to search
        keyword: Keyword to find
    
    Returns:
        Dict with:
        - matched: bool
        - exact_matches: List of exact keyword matches
        - related_data: List of related extracted data (dates, numbers, etc.)
        - match_type: Type of match found
    """
    keyword_lower = keyword.lower().strip()
    content_lower = content.lower()
    
    result = {
        "matched": False,
        "exact_matches": [],
        "related_data": [],
        "match_type": None,
    }
    
    # Check for exact keyword match
    if keyword_lower in content_lower:
        result["matched"] = True
        result["match_type"] = "exact"
        
        # Find all occurrences
        pattern = r'\b' + re.escape(keyword) + r'\b'
        matches = re.findall(pattern, content, re.IGNORECASE)
        result["exact_matches"] = matches
    
    # Smart extraction based on keyword type
    
    # Date-related keywords
    if keyword_lower in ['date', 'datetime', 'time', 'when', 'schedule', 'deadline']:
        dates = extract_dates(content)
        if dates:
            result["matched"] = True
            result["related_data"].extend(dates)
            result["match_type"] = "date_extraction"
    
    # Number/Amount keywords
    elif keyword_lower in ['amount', 'price', 'cost', 'total', 'number', 'count', 'quantity']:
        numbers = extract_numbers(content, include_currency=True)
        if numbers:
            result["matched"] = True
            result["related_data"].extend(numbers)
            result["match_type"] = "number_extraction"
    
    # Email keywords
    elif keyword_lower in ['email', 'contact', 'address']:
        emails = extract_emails(content)
        if emails:
            result["matched"] = True
            result["related_data"].extend(emails)
            result["match_type"] = "email_extraction"
    
    # Phone keywords
    elif keyword_lower in ['phone', 'telephone', 'mobile', 'cell']:
        phones = extract_phones(content)
        if phones:
            result["matched"] = True
            result["related_data"].extend(phones)
            result["match_type"] = "phone_extraction"
    
    # Invoice/Reference number patterns
    elif keyword_lower in ['invoice', 'reference', 'order', 'ticket', 'id']:
        # Look for common patterns
        patterns = [
            r'\b(?:INV|REF|ORD|TKT|ID)[-\s]?\d{4,}',
            r'\b[A-Z]{2,3}-\d{4,}',
        ]
        
        for pattern in patterns:
            matches = re.findall(pattern, content, re.IGNORECASE)
            if matches:
                result["matched"] = True
                result["related_data"].extend(matches)
                result["match_type"] = "pattern_extraction"
    
    return result


def apply_smart_rules(content: str, keywords: List[str]) -> Dict[str, Any]:
    """
    Apply smart matching for multiple keywords.
    
    Args:
        content: Text to search
        keywords: List of keywords
    
    Returns:
        Dict mapping keyword to match result
    """
    results = {}
    
    for keyword in keywords:
        results[keyword] = smart_keyword_match(content, keyword)
    
    return results
