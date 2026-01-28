"""
Column suggester: derive probable column names from a sample of emails.
Heuristics:
 - Collect existing keys in email dicts (non-empty, non-standard)
 - Scan body/subject for Label: value patterns and count label frequency
"""
import re
from collections import Counter
from typing import List, Dict

STANDARD_KEYS = {"subject", "from", "to", "date", "body", "attachments", "attachment_text", "id"}


def suggest_columns(emails: List[Dict], top_n: int = 10) -> List[Dict]:
    label_counter = Counter()
    key_counter = Counter()

    for email in emails:
        # Existing keys
        for k, v in email.items():
            if not v or k.lower() in STANDARD_KEYS:
                continue
            key_counter[k.strip()] += 1

        text_parts = [
            str(email.get("subject", "")),
            str(email.get("body", "")),
            " ".join(email.get("attachment_text", [])) if isinstance(email.get("attachment_text", []), list) else str(email.get("attachment_text", "")),
        ]
        text = "\n".join(text_parts)

        # Find Label: value patterns
        for match in re.finditer(r"([A-Za-z][A-Za-z0-9 _]{2,40})\s*[:\-]\s*[A-Za-z0-9]", text):
            label = match.group(1).strip()
            # normalize multiple spaces
            label = re.sub(r"\s+", " ", label)
            if len(label) <= 40:
                label_counter[label] += 1

    combined = Counter()
    combined.update(key_counter)
    combined.update(label_counter)

    suggestions = []
    for name, count in combined.most_common(top_n):
        confidence = min(10, max(3, count))  # simple scaling
        suggestions.append({
            "name": name,
            "count": count,
            "confidence": confidence,
            "synonyms": []
        })

    return suggestions
"\"\"\"Suggest column names based on recurring headings in sample emails.\"\"\"\n+import re\n+from collections import Counter\n+from typing import Dict, List\n+\n+from jobs.synonym_resolver import get_synonyms\n+\n+\n+class ColumnSuggester:\n+    def __init__(self, emails: List[Dict], max_samples: int = 20):\n+        self.emails = emails[:max_samples]\n+\n+    def _extract_candidates(self) -> Counter:\n+        counter = Counter()\n+        examples: Dict[str, str] = {}\n+\n+        for email in self.emails:\n+            lines = self._lines_from_email(email)\n+            for line in lines:\n+                label = self._label_from_line(line)\n+                if not label:\n+                    continue\n+                key = label.lower()\n+                counter[key] += 1\n+                if key not in examples:\n+                    examples[key] = line.strip()\n+\n+        return counter, examples\n+\n+    @staticmethod\n+    def _lines_from_email(email: Dict) -> List[str]:\n+        parts = []\n+        if email.get(\"subject\"):\n+            parts.append(email[\"subject\"])\n+        if email.get(\"body\"):\n+            parts.append(email[\"body\"])\n+        attachments = email.get(\"attachments\", [])\n+        if attachments:\n+            parts.extend(attachments)\n+        attachment_text = email.get(\"attachment_text\", [])\n+        if attachment_text:\n+            parts.extend(attachment_text)\n+        combined = \"\\n\".join(parts)\n+        return [line.strip() for line in combined.splitlines() if line.strip()]\n+\n+    @staticmethod\n+    def _label_from_line(line: str) -> str:\n+        if \":\" in line:\n+            label = line.split(\":\", 1)[0].strip()\n+            if len(label) >= 2 and len(label) <= 30:\n+                return label\n+        match = re.match(r\"([A-Za-z\\s]{3,30})\\s+\", line)\n+        if match:\n+            candidate = match.group(1).strip()\n+            if candidate:\n+                return candidate\n+        return \"\"\n+\n+    def suggest(self, max_suggestions: int = 10) -> List[Dict]:\n+        counter, examples = self._extract_candidates()\n+        suggestions = []\n+\n+        for label, freq in counter.most_common():\n+            if len(suggestions) >= max_suggestions:\n+                break\n+            suggestion = {\n+                \"name\": label.title(),\n+                \"confidence\": min(9, 3 + freq),\n+                \"example\": examples.get(label, \"\"),\n+                \"synonyms\": get_synonyms(label.title())\n+            }\n+            suggestions.append(suggestion)\n+\n+        return suggestions
