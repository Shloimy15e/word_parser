"""
Talmud/Daf-style document format handler.
"""

import re
from typing import Dict, Any

from word_parser.core.formats._base import (
    DocumentFormat,
    FormatRegistry,
    Document,
    is_old_header,
    should_start_content,
    extract_daf_headings,
    remove_page_markings,
)


@FormatRegistry.register
class DafFormat(DocumentFormat):
    """
    Talmud/Daf-style document format.

    Structure:
    - H1: Book (parent folder or --book arg)
    - H2: Tractate/Masechet (folder name)
    - H3: Perek (extracted from filename, e.g., "פרק א")
    - H4: Chelek/Section (optional, from filename suffix)

    Filename patterns:
    - PEREK1 → H3: "פרק א"
    - PEREK1A → H3: "פרק א", H4: "חלק א"
    - MEKOROS → H3: "מקורות"
    - HAKDOMO → H3: "הקדמה"
    """

    @classmethod
    def get_format_name(cls) -> str:
        return "daf"

    @classmethod
    def get_priority(cls) -> int:
        return 50

    @classmethod
    def detect(cls, doc: Document, context: Dict[str, Any]) -> bool:
        """Detect daf format from context or filename patterns."""
        # Check if explicitly requested
        if context.get("mode") == "daf":
            return True

        # Check filename for perek or chelek pattern
        filename = context.get("filename", "").lower()
        if re.match(r"^perek\d+", filename):
            return True
        if re.match(r"^(?:chelek|חלק)\d+", filename):
            return True
        if re.match(r"^me?koros", filename):
            return True
        if re.match(r"^hakdomo", filename):
            return True

        return False

    @classmethod
    def get_optional_context(cls) -> Dict[str, Any]:
        return {
            "book": None,  # Can be derived from parent folder
            "folder": None,  # Tractate name
            "filename": None,
            "filter_headers": True,
        }

    def process(self, doc: Document, context: Dict[str, Any]) -> Document:
        """Process document with daf/perek structure."""
        book = context.get("book", "")
        # Support both "folder" and "daf_folder" for compatibility
        folder = context.get("folder") or context.get("daf_folder", "")
        filename = context.get("filename", "")

        # Remove page markings and merge split paragraphs
        doc = remove_page_markings(doc)

        # Extract H3 and H4 from filename
        heading3, heading4 = extract_daf_headings(filename)

        # Set headings
        doc.set_headings(h1=book, h2=folder, h3=heading3, h4=heading4)

        # Filter old headers if requested
        if context.get("filter_headers", True):
            doc = self._filter_headers(doc)

        return doc

    def _filter_headers(self, doc: Document) -> Document:
        """Remove old header paragraphs from document."""
        in_header_section = True
        filtered_paragraphs = []

        for para in doc.paragraphs:
            txt = para.text.strip()

            if in_header_section:
                if txt and should_start_content(txt):
                    in_header_section = False
                    filtered_paragraphs.append(para)
                elif txt and is_old_header(txt):
                    continue
                elif not txt:
                    continue
            else:
                if txt and is_old_header(txt):
                    continue
                filtered_paragraphs.append(para)

        doc.paragraphs = filtered_paragraphs
        return doc

