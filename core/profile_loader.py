"""
Load and validate profile configurations.
"""
import json
from pathlib import Path
from typing import Dict, List, Optional


class ProfileLoader:
    """Load profile JSON configurations."""
    
    def __init__(self, profiles_dir: str = "profiles"):
        self.profiles_dir = Path(profiles_dir)
    
    def list_profiles(self) -> List[str]:
        """List available profile names."""
        if not self.profiles_dir.exists():
            return []
        
        profiles = []
        for path in self.profiles_dir.glob("*.json"):
            profiles.append(path.stem)
        return sorted(profiles)
    
    def load_profile(self, profile_name: str) -> Optional[Dict]:
        """
        Load a profile by name.
        
        Args:
            profile_name: Profile name (without .json extension)
        
        Returns:
            Profile dict or None if not found
        """
        path = self.profiles_dir / f"{profile_name}.json"
        if not path.exists():
            return None
        
        with path.open("r", encoding="utf-8") as f:
            return json.load(f)
    
    def save_profile(self, profile: Dict) -> None:
        """
        Save a profile.
        
        Args:
            profile: Profile dict (must have "name" key)
        """
        profile_name = profile.get("name")
        if not profile_name:
            raise ValueError("Profile must have a 'name' field")
        
        self.profiles_dir.mkdir(exist_ok=True)
        path = self.profiles_dir / f"{profile_name}.json"
        
        with path.open("w", encoding="utf-8") as f:
            json.dump(profile, f, indent=2)
    
    def validate_profile(self, profile: Dict) -> List[str]:
        """
        Validate a profile configuration.
        
        Args:
            profile: Profile dict
        
        Returns:
            List of validation error messages (empty if valid)
        """
        errors = []
        
        if not profile.get("name"):
            errors.append("Missing 'name' field")
        
        if not profile.get("input_source"):
            errors.append("Missing 'input_source' field")
        
        input_source = profile.get("input_source")
        if input_source not in ["graph", "local_eml", "local_csv"]:
            errors.append(f"Invalid input_source: {input_source}")
        
        if not profile.get("schema"):
            errors.append("Missing 'schema' field")
        
        if not profile.get("output"):
            errors.append("Missing 'output' field")
        
        return errors
