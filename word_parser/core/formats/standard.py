"""
Standard Torah document format handler.
"""

from typing import Dict, Any, List

from word_parser.core.formats._base import (
    DocumentFormat,
    FormatRegistry,
    Document,
    is_old_header,
    should_start_content,
    extract_year,
    extract_heading4_info,
    remove_page_markings,
)


@FormatRegistry.register
class StandardFormat(DocumentFormat):
    """
    Standard Torah document format.

    Structure:
    - H1: Book (e.g., "ליקוטי שיחות")
    - H2: Sefer (e.g., "סדר בראשית")
    - H3: Parshah (e.g., "פרשת בראשית")
    - H4: Year or subsection (e.g., "תש״נ")

    Used for most Torah commentaries organized by parshah.
    """

    @classmethod
    def get_format_name(cls) -> str:
        return "standard"

    @classmethod
    def get_priority(cls) -> int:
        return 0  # Default/fallback format

    @classmethod
    def detect(cls, doc: Document, context: Dict[str, Any]) -> bool:
        # Standard format is the default - always matches as fallback
        return True

    @classmethod
    def get_required_context(cls) -> List[str]:
        return ["book"]

    @classmethod
    def get_optional_context(cls) -> Dict[str, Any]:
        return {
            "sefer": None,
            "parshah": None,
            "filename": None,
            "skip_parshah_prefix": False,
            "filter_headers": True,
            "use_filename_for_h4": False,  # If True, use clean filename instead of extracted year
        }

    def process(self, doc: Document, context: Dict[str, Any]) -> Document:
        """Process document with standard parshah structure."""
        from pathlib import Path
        
        book = context.get("book", "")
        sefer = context.get("sefer", "")
        parshah = context.get("parshah", "")
        filename = context.get("filename", "")
        skip_prefix = context.get("skip_parshah_prefix", False)

        # Remove page markings and merge split paragraphs
        doc = remove_page_markings(doc)

        # Determine H4: use filename directly if option is set, otherwise extract year
        use_filename = context.get("use_filename_for_h4", False)
        if use_filename:
            # Use filename from context directly (already cleaned when option is set)
            h4 = filename if filename else None
        elif filename and not context.get("year"):
            # Try to extract year from filename if not provided
            year = extract_year(filename)
            heading4_info = extract_heading4_info(filename)
            h4 = year or heading4_info or filename
        else:
            h4 = context.get("year") or filename

        # Set headings
        h3 = parshah if skip_prefix else f"פרשת {parshah}" if parshah else None
        doc.set_headings(h1=book, h2=sefer, h3=h3, h4=h4)

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

