"""
Siman/Halacha document format handler.
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
class SimanFormat(DocumentFormat):
    """
    Siman/Halacha document format.

    Structure:
    - H1: Book (e.g., "שולחן ערוך")
    - H2: Section (e.g., "אורח חיים")
    - H3: Siman number (e.g., "סימן א")
    - H4: Seif number (optional)

    Used for halachic works organized by siman numbers.
    """

    @classmethod
    def get_format_name(cls) -> str:
        return "siman"

    @classmethod
    def get_priority(cls) -> int:
        return 45

    @classmethod
    def detect(cls, doc: Document, context: Dict[str, Any]) -> bool:
        """Detect siman format from context or content."""
        if context.get("mode") == "siman":
            return True

        # Check filename for siman pattern
        filename = context.get("filename", "").lower()
        if re.match(r"^siman\d+", filename) or re.match(r"^סימן", filename):
            return True

        # Check document for siman markers
        for para in doc.paragraphs[:20]:
            txt = para.text.strip()
            if re.match(r"^סימן\s+[א-ת]+", txt):
                return True

        return False

    @classmethod
    def get_optional_context(cls) -> Dict[str, Any]:
        return {
            "book": None,
            "section": None,
            "siman": None,
            "seif": None,
            "filename": None,
        }

    def process(self, doc: Document, context: Dict[str, Any]) -> Document:
        """Process siman-based document."""
        book = context.get("book", "")
        section = context.get("section", "")
        siman = context.get("siman")
        seif = context.get("seif")
        filename = context.get("filename", "")

        # Remove page markings and merge split paragraphs
        doc = remove_page_markings(doc)

        # Try to extract siman from filename
        if not siman and filename:
            siman_match = re.match(r"^siman(\d+)", filename.lower())
            if siman_match:
                from word_parser.core.processing import number_to_hebrew_gematria

                num = int(siman_match.group(1))
                siman = f"סימן {number_to_hebrew_gematria(num)}"

        doc.set_headings(h1=book, h2=section, h3=siman, h4=seif)

        return doc

