"""
Pattern Learner - Automatically learns number patterns and data types from context.
Analyzes extracted data across multiple emails to infer:
- 4000# = Work Orders
- 3000# = Equipment IDs  
- ZMO# = Order Types
etc.
"""
import re
import json
from collections import Counter, defaultdict
from typing import Dict, List, Tuple, Optional
from pathlib import Path


class PatternLearner:
    """
    Learns patterns from extracted data to improve future extraction.
    """
    
    def __init__(self, cache_file: Optional[str] = None):
        """
        Initialize pattern learner.
        
        Args:
            cache_file: Path to save/load learned patterns (optional)
        """
        self.cache_file = cache_file or "config/learned_patterns.json"
        self.patterns = self._load_patterns()
        
        # Pattern types we can learn
        self.pattern_types = {
            "work_order": ["order", "work order", "wo", "wo number"],
            "equipment": ["equipment", "equip", "eq", "equipment id"],
            "order_type": ["order type", "type", "order type code"],
            "building": ["building", "bldg"],
            "room": ["room"],
        }
    
    def _load_patterns(self) -> Dict:
        """Load learned patterns from cache."""
        try:
            path = Path(self.cache_file)
            if path.exists():
                with path.open('r') as f:
                    return json.load(f)
        except Exception as e:
            print(f"Could not load patterns: {e}")
        
        return {
            "number_patterns": {},  # {pattern: {column_type: confidence}}
            "prefix_patterns": {},  # {prefix: column_type}
            "context_patterns": {},  # {context_word: column_type}
        }
    
    def _save_patterns(self):
        """Save learned patterns to cache."""
        try:
            path = Path(self.cache_file)
            path.parent.mkdir(parents=True, exist_ok=True)
            with path.open('w') as f:
                json.dump(self.patterns, f, indent=2)
        except Exception as e:
            print(f"Could not save patterns: {e}")
    
    def analyze_records(self, records: List[Dict], schema_columns: List[Dict]):
        """
        Analyze extracted records to learn patterns.
        
        Args:
            records: List of extracted records
            schema_columns: Column definitions from schema
        """
        # Build column type mapping
        column_types = {}
        for col in schema_columns:
            col_name = col.get("name", "").lower()
            col_type = col.get("extract_type", col.get("type", "auto"))
            
            # Map column to pattern type
            for pattern_type, keywords in self.pattern_types.items():
                if any(kw in col_name for kw in keywords):
                    column_types[col_name] = pattern_type
                    break
        
        # Analyze number patterns
        number_patterns = defaultdict(lambda: defaultdict(int))
        
        for record in records:
            for col_name, value in record.items():
                if not value or not isinstance(value, str):
                    continue
                
                col_name_lower = col_name.lower()
                pattern_type = column_types.get(col_name_lower)
                
                if not pattern_type:
                    continue
                
                # Extract number patterns
                numbers = self._extract_numbers_from_value(value)
                for num in numbers:
                    pattern = self._classify_number_pattern(num)
                    if pattern:
                        number_patterns[pattern][pattern_type] += 1
        
        # Update learned patterns
        for pattern, type_counts in number_patterns.items():
            if pattern not in self.patterns["number_patterns"]:
                self.patterns["number_patterns"][pattern] = {}
            
            for pattern_type, count in type_counts.items():
                current_conf = self.patterns["number_patterns"][pattern].get(pattern_type, 0)
                # Weighted average (favor recent data)
                self.patterns["number_patterns"][pattern][pattern_type] = int(
                    (current_conf * 0.7) + (count * 0.3)
                )
        
        self._save_patterns()
    
    def _extract_numbers_from_value(self, value: str) -> List[str]:
        """Extract number patterns from a value."""
        numbers = []
        
        # Pure numbers (6+ digits)
        matches = re.findall(r'\b(\d{6,})\b', value)
        numbers.extend(matches)
        
        # Numbers with prefixes (4000390042 ZM03)
        matches = re.findall(r'\b(\d{6,}\s*[A-Z]{2,}\d*)\b', value)
        numbers.extend(matches)
        
        # Alphanumeric codes (TORFER0048)
        matches = re.findall(r'\b([A-Z]{2,}\d{3,})\b', value)
        numbers.extend(matches)
        
        return numbers
    
    def _classify_number_pattern(self, number: str) -> Optional[str]:
        """
        Classify a number into a pattern.
        
        Examples:
        - "4000390042" → "4000#" (Work Order pattern)
        - "300020888" → "3000#" (Equipment pattern)
        - "ZM03" → "ZM#" (Order Type pattern)
        """
        # Extract leading digits
        digit_match = re.match(r'^(\d{1,4})', number)
        if digit_match:
            leading_digits = digit_match.group(1)
            # Classify by leading digits
            if leading_digits.startswith('4'):
                return "4000#"
            elif leading_digits.startswith('3'):
                return "3000#"
            elif leading_digits.startswith('2'):
                return "2000#"
            elif leading_digits.startswith('1'):
                return "1000#"
        
        # Extract letter prefix
        letter_match = re.match(r'^([A-Z]{2,})', number)
        if letter_match:
            prefix = letter_match.group(1)
            return f"{prefix}#"
        
        return None
    
    def predict_column_type(self, value: str, context: str = "") -> Optional[str]:
        """
        Predict what column type a value belongs to based on learned patterns.
        
        Args:
            value: The value to classify
            context: Surrounding text context
        
        Returns:
            Predicted column type (e.g., "work_order", "equipment")
        """
        numbers = self._extract_numbers_from_value(value)
        if not numbers:
            return None
        
        scores = defaultdict(float)
        
        for num in numbers:
            pattern = self._classify_number_pattern(num)
            if not pattern:
                continue
            
            # Check learned patterns
            if pattern in self.patterns["number_patterns"]:
                for col_type, confidence in self.patterns["number_patterns"][pattern].items():
                    scores[col_type] += confidence
        
        # Also check context
        context_lower = context.lower()
        for pattern_type, keywords in self.pattern_types.items():
            for keyword in keywords:
                if keyword in context_lower:
                    scores[pattern_type] += 5
        
        if scores:
            # Return highest scoring type
            return max(scores.items(), key=lambda x: x[1])[0]
        
        return None
    
    def get_pattern_hint(self, column_name: str, value: str) -> Optional[str]:
        """
        Get a pattern-based extraction hint for a column.
        
        Returns:
            Regex pattern hint or None
        """
        predicted_type = self.predict_column_type(value, column_name)
        if not predicted_type:
            return None
        
        # Return pattern based on type
        patterns = {
            "work_order": r'\b(4\d{9,})\b',  # 4000# pattern
            "equipment": r'\b(3\d{8,})\b',   # 3000# pattern
            "order_type": r'\b([A-Z]{2,}\d{1,3})\b',  # ZM03 pattern
        }
        
        return patterns.get(predicted_type)
