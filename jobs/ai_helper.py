"""
Simple AI helper that optionally calls an external endpoint to disambiguate low-confidence columns.
Configure via EMAIL_AI_ENDPOINT and EMAIL_AI_API_KEY environment variables.
"""
import json
import logging
import os
from typing import Dict, Optional

try:
    import requests
except ImportError:  # pragma: no cover
    requests = None


class AIHelper:
    def __init__(self, endpoint: Optional[str] = None, api_key: Optional[str] = None):
        self.endpoint = endpoint or os.getenv("EMAIL_AI_ENDPOINT")
        self.api_key = api_key or os.getenv("EMAIL_AI_API_KEY")

    def is_configured(self) -> bool:
        return bool(self.endpoint) and requests is not None

    def suggest(self, column_name: str, email_data: Dict[str, str]) -> Optional[Dict[str, str]]:
        if not self.is_configured():
            return None

        payload = {
            "column": column_name,
            "subject": email_data.get("subject", ""),
            "body": email_data.get("body", ""),
            "attachments": email_data.get("attachments", []),
        }
        headers = {"Content-Type": "application/json"}
        if self.api_key:
            headers["Authorization"] = f"Bearer {self.api_key}"

        try:
            response = requests.post(self.endpoint, headers=headers, json=payload, timeout=20)
            response.raise_for_status()
            data = response.json()
            value = data.get("value")
            if not value:
                return None
            return {
                "value": value,
                "confidence": data.get("confidence", 5),
                "snippet": data.get("snippet") or value[:100],
            }
        except Exception as exc:  # pragma: no cover
            logging.warning("AI helper request failed: %s", exc)
            return None
    
    def classify_value(self, value: str, column_name: str, context: str) -> Optional[str]:
        """
        Use AI to classify a value into the correct column type.
        
        Example: "4000390042" + context "Work Order" → returns "4000390042"
        
        Args:
            value: The value to classify
            column_name: Name of the column
            context: Surrounding text context
        
        Returns:
            Classified value or None
        """
        if not self.is_configured():
            return None
        
        payload = {
            "action": "classify_value",
            "value": value,
            "column": column_name,
            "context": context[:1000],  # Limit context size
        }
        headers = {"Content-Type": "application/json"}
        if self.api_key:
            headers["Authorization"] = f"Bearer {self.api_key}"
        
        try:
            response = requests.post(self.endpoint, headers=headers, json=payload, timeout=20)
            response.raise_for_status()
            data = response.json()
            classified_value = data.get("value")
            if classified_value and classified_value.strip():
                return classified_value.strip()
        except Exception as exc:
            logging.warning("AI value classification failed: %s", exc)
        
        return None