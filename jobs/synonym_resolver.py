"""
Synonym Resolver - Maps column names to their synonyms/aliases.
Allows "QC" to match "Quality Criticality", etc.
"""
import json
import os
from pathlib import Path
from typing import List, Set, Dict, Optional


class SynonymResolver:
    """Resolves column names to their synonyms for flexible matching."""
    
    def __init__(self, config_path: str = None):
        """
        Initialize the synonym resolver.
        
        Args:
            config_path: Path to synonyms.json (default: config/synonyms.json)
        """
        self.synonyms: Dict[str, List[str]] = {}
        self.abbreviations: Dict[str, str] = {}
        
        # Find config file
        if config_path is None:
            # Try relative paths
            possible_paths = [
                Path("config/synonyms.json"),
                Path(__file__).parent.parent / "config" / "synonyms.json",
            ]
            for p in possible_paths:
                if p.exists():
                    config_path = str(p)
                    break
        
        if config_path and os.path.exists(config_path):
            self._load_config(config_path)
        else:
            # Use built-in defaults
            self._load_defaults()
    
    def _load_config(self, config_path: str):
        """Load synonyms from config file."""
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                self.synonyms = data.get("synonyms", {})
                self.abbreviations = data.get("abbreviations", {})
        except Exception as e:
            print(f"Warning: Could not load synonyms config: {e}")
            self._load_defaults()
    
    def _load_defaults(self):
        """Load default synonym mappings."""
        self.synonyms = {
            "qc": ["quality criticality", "quality crit", "criticality"],
            "quality criticality": ["qc", "quality crit"],
            "wo": ["work order", "work order number", "workorder"],
            "work order": ["wo", "wo #", "work order number"],
            "order": ["order number", "order #", "order no"],
            "equipment": ["equipment id", "equip", "eq"],
            "building": ["bldg", "building number", "bldg #"],
            "date": ["due date", "finish date", "completion date"],
            "status": ["completion status", "state"],
            "description": ["desc", "details"],
            "comments": ["notes", "remarks"],
        }
        self.abbreviations = {
            "qc": "Quality Criticality",
            "wo": "Work Order",
            "eq": "Equipment",
            "bldg": "Building",
        }
    
    def get_synonyms(self, column_name: str) -> List[str]:
        """
        Get all synonyms for a column name.
        
        Args:
            column_name: The column name to find synonyms for
        
        Returns:
            List of synonyms (including the original name)
        """
        name_lower = column_name.lower().strip()
        
        # Start with the original name
        synonyms = [column_name, name_lower]
        
        # Look up direct synonyms
        if name_lower in self.synonyms:
            synonyms.extend(self.synonyms[name_lower])
        
        # Check if any synonym key contains our name
        for key, values in self.synonyms.items():
            if name_lower in values or key in name_lower or name_lower in key:
                synonyms.append(key)
                synonyms.extend(values)
        
        # Remove duplicates while preserving order
        seen = set()
        unique = []
        for s in synonyms:
            s_lower = s.lower().strip()
            if s_lower not in seen:
                seen.add(s_lower)
                unique.append(s)
        
        return unique
    
    def expand_name(self, abbreviation: str) -> str:
        """
        Expand an abbreviation to its full name.
        
        Args:
            abbreviation: Short form (e.g., "QC")
        
        Returns:
            Full name (e.g., "Quality Criticality") or original if not found
        """
        abbr_lower = abbreviation.lower().strip()
        return self.abbreviations.get(abbr_lower, abbreviation)
    
    def matches(self, column_name: str, search_term: str) -> bool:
        """
        Check if a column name matches a search term (including synonyms).
        
        Args:
            column_name: The column name from data
            search_term: The term being searched for
        
        Returns:
            True if they match (directly or via synonym)
        """
        col_lower = column_name.lower().strip()
        search_lower = search_term.lower().strip()
        
        # Direct match
        if col_lower == search_lower:
            return True
        
        # Partial match
        if search_lower in col_lower or col_lower in search_lower:
            return True
        
        # Synonym match
        col_synonyms = set(s.lower() for s in self.get_synonyms(column_name))
        search_synonyms = set(s.lower() for s in self.get_synonyms(search_term))
        
        return bool(col_synonyms & search_synonyms)
    
    def find_matching_column(self, search_term: str, available_columns: List[str]) -> Optional[str]:
        """
        Find a column that matches the search term (using synonyms).
        
        Args:
            search_term: The column name to search for
            available_columns: List of actual column names in the data
        
        Returns:
            The matching column name from available_columns, or None
        """
        search_lower = search_term.lower().strip()
        search_synonyms = self.get_synonyms(search_term)
        
        for col in available_columns:
            col_lower = col.lower().strip()
            
            # Direct match
            if col_lower == search_lower:
                return col
            
            # Check against synonyms
            for syn in search_synonyms:
                if col_lower == syn.lower() or syn.lower() in col_lower or col_lower in syn.lower():
                    return col
        
        return None


# Global instance for convenience
_resolver = None

def get_resolver() -> SynonymResolver:
    """Get or create the global synonym resolver."""
    global _resolver
    if _resolver is None:
        _resolver = SynonymResolver()
    return _resolver


def get_synonyms(column_name: str) -> List[str]:
    """Get synonyms for a column name."""
    return get_resolver().get_synonyms(column_name)


def columns_match(col1: str, col2: str) -> bool:
    """Check if two column names match (including synonyms)."""
    return get_resolver().matches(col1, col2)
