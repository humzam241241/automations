"""
Column suggester: derive probable column names from a sample of emails.
Heuristics:
 - Collect existing keys in email dicts (non-empty, non-standard)
 - Scan body/subject/attachments for Label: value patterns and count label frequency
 - Return suggestions with confidence, example text, and synonyms
"""
import re
from collections import Counter
from typing import Dict, List

from jobs.synonym_resolver import get_synonyms

STANDARD_KEYS = {"subject", "from", "to", "date", "body", "attachments", "attachment_text", "id"}


class ColumnSuggester:
    def __init__(self, emails: List[Dict], max_samples: int = 50):
        self.emails = emails[:max_samples]

    @staticmethod
    def _lines_from_email(email: Dict) -> List[str]:
        parts = []
        if email.get("subject"):
            parts.append(str(email["subject"]))
        if email.get("body"):
            parts.append(str(email["body"]))
        attachments = email.get("attachments", [])
        if attachments:
            parts.extend([str(a) for a in attachments])
        attachment_text = email.get("attachment_text", [])
        if attachment_text:
            parts.extend([str(a) for a in attachment_text])
        combined = "\n".join(parts)
        return [line.strip() for line in combined.splitlines() if line.strip()]

    @staticmethod
    def _label_from_line(line: str) -> str:
        if ":" in line:
            label = line.split(":", 1)[0].strip()
            if 2 <= len(label) <= 30:
                return label
        match = re.match(r"([A-Za-z\s]{3,30})\s+", line)
        if match:
            candidate = match.group(1).strip()
            if candidate:
                return candidate
        return ""

    def _extract_candidates(self) -> (Counter, Dict[str, str]):
        counter = Counter()
        examples: Dict[str, str] = {}

        for email in self.emails:
            # existing non-standard keys
            for k, v in email.items():
                if not v or k.lower() in STANDARD_KEYS:
                    continue
                key = k.strip()
                counter[key.lower()] += 1
                if key.lower() not in examples:
                    examples[key.lower()] = str(v)[:120]

            # parse lines for labels
            lines = self._lines_from_email(email)
            for line in lines:
                label = self._label_from_line(line)
                if not label:
                    continue
                key = label.lower()
                counter[key] += 1
                if key not in examples:
                    examples[key] = line.strip()[:120]

        return counter, examples

    def suggest(self, max_suggestions: int = 10) -> List[Dict]:
        counter, examples = self._extract_candidates()
        suggestions = []

        for label, freq in counter.most_common():
            if len(suggestions) >= max_suggestions:
                break
            suggestion = {
                "name": label.title(),
                "confidence": min(9, 3 + freq),
                "example": examples.get(label, ""),
                "synonyms": get_synonyms(label.title())
            }
            suggestions.append(suggestion)

        return suggestions


def suggest_columns(emails: List[Dict], top_n: int = 10) -> List[Dict]:
    """Backwards-compatible helper to get suggestions from a list of emails."""
    suggester = ColumnSuggester(emails, max_samples=50)
    return suggester.suggest(max_suggestions=top_n)
