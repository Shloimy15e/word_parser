"""
Letter/correspondence document format handler.
"""

import re
from typing import Dict, Any

from word_parser.core.formats._base import (
    DocumentFormat,
    FormatRegistry,
    Document,
    remove_page_markings,
)


@FormatRegistry.register
class LetterFormat(DocumentFormat):
    """
    Letter/correspondence document format.

    Structure:
    - H1: Collection name
    - H2: Category (e.g., "מכתבים")
    - H3: Recipient or subject
    - H4: Date

    Detection: Document starts with greeting pattern or contains
    letter markers like "ב״ה", date, "כבוד", etc.
    """

    @classmethod
    def get_format_name(cls) -> str:
        return "letter"

    @classmethod
    def get_priority(cls) -> int:
        return 40

    @classmethod
    def detect(cls, doc: Document, context: Dict[str, Any]) -> bool:
        """Detect letter format from document content."""
        if context.get("mode") == "letter":
            return True

        if not doc.paragraphs:
            return False

        # Check first few paragraphs for letter patterns
        first_paras = [p.text.strip() for p in doc.paragraphs[:5] if p.text.strip()]

        letter_patterns = [
            r'^ב["\']ה',  # ב"ה at start
            r"^כבוד",  # כבוד (honorific)
            r"^לכבוד",  # לכבוד
            r"^שלום",  # שלום greeting
            r"^ידידי",  # ידידי (my friend)
            r"הנדון:",  # הנדון: (subject)
        ]
        for para in first_paras:
            for pattern in letter_patterns:
                if re.match(pattern, para):
                    return True

        return False

    @classmethod
    def get_optional_context(cls) -> Dict[str, Any]:
        return {
            "book": None,
            "category": "מכתבים",
            "recipient": None,
            "date": None,
        }

    def process(self, doc: Document, context: Dict[str, Any]) -> Document:
        """Process letter document."""
        book = context.get("book", "")
        category = context.get("category", "מכתבים")
        recipient = context.get("recipient")
        date = context.get("date")

        # Remove page markings and merge split paragraphs
        doc = remove_page_markings(doc)

        # Try to extract recipient and date from document if not provided
        if not recipient or not date:
            extracted = self._extract_letter_info(doc)
            recipient = recipient or extracted.get("recipient")
            date = date or extracted.get("date")

        doc.set_headings(h1=book, h2=category, h3=recipient, h4=date)

        return doc

    def _extract_letter_info(self, doc: Document) -> Dict[str, str]:
        """Try to extract recipient and date from letter content."""
        info = {}

        for para in doc.paragraphs[:10]:
            txt = para.text.strip()

            # Look for recipient pattern
            recipient_match = re.match(r"^(?:לכבוד|כבוד)\s+(.+?)(?:\s+שליט״א)?$", txt)
            if recipient_match and "recipient" not in info:
                info["recipient"] = recipient_match.group(1)

            # Look for date patterns
            date_match = re.search(r"(\d{1,2}[./]\d{1,2}[./]\d{2,4})", txt)
            if date_match and "date" not in info:
                info["date"] = date_match.group(1)

            # Hebrew date pattern
            heb_date_match = re.search(r"([א-ת]+'?\s+[א-ת]+\s+תש[א-ת\"']+)", txt)
            if heb_date_match and "date" not in info:
                info["date"] = heb_date_match.group(1)

        return info

