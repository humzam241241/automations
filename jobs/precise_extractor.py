"""
Precise Extractor - Hard-coded rules for known patterns.
Enforces:
- Order Type: ZM# format (ZM03, ZM05, etc.)
- Work Order: 4000# format (4000390042, etc.)
- Equipment: 3000# format (300020888, etc.)
"""
import re
from typing import Dict, List, Tuple, Optional


# Hard-coded pattern rules (always enforced)
KNOWN_PATTERNS = {
    "order_type": {
        "patterns": [
            r'\b([A-Z]{2,3}\d{1,3})\b',  # ZM03, ZM05, ZMO3
            r'\b([A-Z]{2,}\d{1,3})\b',    # Any 2+ letters + 1-3 digits
        ],
        "keywords": ["order type", "type", "order type code", "type code"],
        "validation": lambda v: len(v) >= 3 and len(v) <= 8 and bool(re.match(r'^[A-Z]{2,}\d{1,3}$', v))
    },
    "work_order": {
        "patterns": [
            r'\b(4\d{9,})\b',  # 4000390042, 4000278194
            r'\b(4\d{8,})\b',  # Allow 8+ digits starting with 4
        ],
        "keywords": ["order", "work order", "wo", "wo number", "order number"],
        "validation": lambda v: v.startswith('4') and len(v) >= 9 and v.isdigit()
    },
    "equipment": {
        "patterns": [
            r'\b(3\d{8,})\b',  # 300020888, 300020881
            r'\b(3\d{7,})\b',  # Allow 7+ digits starting with 3
        ],
        "keywords": ["equipment", "equip", "eq", "equipment id", "equipment number"],
        "validation": lambda v: v.startswith('3') and len(v) >= 8 and v.isdigit()
    },
}


def extract_by_known_pattern(column_name: str, content: str) -> Optional[str]:
    """
    Extract value using hard-coded known patterns.
    Returns the first valid match based on column name.
    """
    column_lower = column_name.lower().strip()
    
    # Determine which pattern type to use based on column name
    pattern_type = None
    for ptype, config in KNOWN_PATTERNS.items():
        if any(kw in column_lower for kw in config["keywords"]):
            pattern_type = ptype
            break
    
    if not pattern_type:
        return None
    
    config = KNOWN_PATTERNS[pattern_type]
    
    # Strategy 1: Look for "ColumnName: value" pattern
    for keyword in config["keywords"]:
        label_pattern = rf'{re.escape(keyword)}\s*[:\-=]\s*([^\n,;]+)'
        match = re.search(label_pattern, content, re.IGNORECASE)
        if match:
            value = match.group(1).strip()
            # Extract the pattern from the value
            for pattern in config["patterns"]:
                pattern_match = re.search(pattern, value)
                if pattern_match:
                    extracted = pattern_match.group(1)
                    if config["validation"](extracted):
                        return extracted
    
    # Strategy 2: Search for pattern in content
    for pattern in config["patterns"]:
        matches = re.findall(pattern, content)
        for match in matches:
            if isinstance(match, tuple):
                match = match[0] if match else ""
            if match and config["validation"](match):
                return match
    
    return None


def extract_order_type(content: str) -> Optional[str]:
    """Extract order type (ZM03, ZM05, etc.) from content."""
    # Look for ZM# pattern
    patterns = [
        r'\b(ZM\d{1,3})\b',      # ZM03, ZM05
        r'\b(ZMO\d{1,2})\b',      # ZMO3, ZMO5
        r'\b([A-Z]{2,3}\d{1,3})\b',  # Any 2-3 letters + digits
    ]
    
    for pattern in patterns:
        matches = re.findall(pattern, content, re.IGNORECASE)
        for match in matches:
            if isinstance(match, tuple):
                match = match[0] if match else ""
            match = match.upper()
            # Validate: 2-3 letters + 1-3 digits, total 3-8 chars
            if re.match(r'^[A-Z]{2,3}\d{1,3}$', match) and 3 <= len(match) <= 8:
                return match
    
    return None


def extract_work_order(content: str) -> Optional[str]:
    """Extract work order number (4000#) from content."""
    # Look for 4 followed by 8+ digits
    patterns = [
        r'\b(4\d{9,})\b',  # 4000390042
        r'\b(4\d{8,})\b',   # 400039004 (8+ digits)
    ]
    
    for pattern in patterns:
        matches = re.findall(pattern, content)
        for match in matches:
            if isinstance(match, tuple):
                match = match[0] if match else ""
            # Validate: starts with 4, 9+ digits
            if match.startswith('4') and len(match) >= 9 and match.isdigit():
                return match
    
    return None


def extract_equipment(content: str) -> Optional[str]:
    """Extract equipment ID (3000#) from content."""
    # Look for 3 followed by 7+ digits
    patterns = [
        r'\b(3\d{8,})\b',  # 300020888
        r'\b(3\d{7,})\b',  # 30002088 (7+ digits)
    ]
    
    for pattern in patterns:
        matches = re.findall(pattern, content)
        for match in matches:
            if isinstance(match, tuple):
                match = match[0] if match else ""
            # Validate: starts with 3, 8+ digits
            if match.startswith('3') and len(match) >= 8 and match.isdigit():
                return match
    
    return None


def extract_precise_value(column_name: str, content: str, context: str = "") -> Optional[str]:
    """
    Extract value using precise, hard-coded rules.
    Returns the most accurate match based on column name and known patterns.
    """
    column_lower = column_name.lower().strip()
    
    # Check for order type
    if any(kw in column_lower for kw in ["order type", "type", "order type code"]):
        # First try label-value pattern
        for keyword in ["order type", "type", "order type code"]:
            pattern = rf'{re.escape(keyword)}\s*[:\-=]\s*([^\n,;]+)'
            match = re.search(pattern, content, re.IGNORECASE)
            if match:
                value = match.group(1).strip()
                order_type = extract_order_type(value)
                if order_type:
                    return order_type
        
        # Then search in content
        order_type = extract_order_type(content)
        if order_type:
            return order_type
    
    # Check for work order
    if any(kw in column_lower for kw in ["order", "work order", "wo", "wo number", "order number"]):
        # First try label-value pattern
        for keyword in ["order", "work order", "wo"]:
            pattern = rf'{re.escape(keyword)}\s*[:\-=]\s*([^\n,;]+)'
            match = re.search(pattern, content, re.IGNORECASE)
            if match:
                value = match.group(1).strip()
                wo = extract_work_order(value)
                if wo:
                    return wo
        
        # Then search in content
        wo = extract_work_order(content)
        if wo:
            return wo
    
    # Check for equipment
    if any(kw in column_lower for kw in ["equipment", "equip", "eq", "equipment id", "equipment number"]):
        # First try label-value pattern
        for keyword in ["equipment", "equip", "eq"]:
            pattern = rf'{re.escape(keyword)}\s*[:\-=]\s*([^\n,;]+)'
            match = re.search(pattern, content, re.IGNORECASE)
            if match:
                value = match.group(1).strip()
                eq = extract_equipment(value)
                if eq:
                    return eq
        
        # Then search in content
        eq = extract_equipment(content)
        if eq:
            return eq
    
    # Fallback to generic pattern matching
    return extract_by_known_pattern(column_name, content)
