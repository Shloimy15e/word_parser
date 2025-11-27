"""
Formatted format handler - for already-formatted files.
"""

from typing import Dict, Any

from word_parser.core.formats._base import (
    DocumentFormat,
    FormatRegistry,
    Document,
    HeadingLevel,
)


@FormatRegistry.register
class FormattedFormat(DocumentFormat):
    """
    Format handler for already-formatted files.
    
    This format is used when processing files that have already been formatted
    (e.g., files with "-formatted" in the filename or files with paragraph-level
    headings already set). It extracts headings from existing paragraph styles
    without re-processing the document structure.
    
    Structure:
    - Headings are extracted from paragraph-level heading styles (Heading 1-4)
    - Document-level headings are set from the first occurrence of each heading level
    - Minimal processing - preserves existing formatting
    """

    @classmethod
    def get_format_name(cls) -> str:
        return "formatted"

    @classmethod
    def get_priority(cls) -> int:
        return 20  # High priority - check before other formats

    @classmethod
    def detect(cls, doc: Document, context: Dict[str, Any]) -> bool:
        """Detect if this is an already-formatted file."""
        # Check if filename contains "-formatted"
        input_path = context.get("input_path", "")
        if input_path and "-formatted" in str(input_path):
            return True
        
        # Check if document already has paragraph-level headings set
        # (from Word styles like "Heading 1", "Heading 2", etc.)
        has_headings = any(
            para.heading_level != HeadingLevel.NORMAL 
            for para in doc.paragraphs
        )
        
        # Also check if explicitly requested
        if context.get("format") == "formatted" or context.get("mode") == "formatted":
            return True
            
        return has_headings

    @classmethod
    def get_optional_context(cls) -> Dict[str, Any]:
        return {
            "book": None,
            "sefer": None,
            "parshah": None,
            "filename": None,
            "filter_headers": False,  # Don't filter headers in formatted files
        }

    def process(self, doc: Document, context: Dict[str, Any]) -> Document:
        """Process already-formatted document - extract headings from paragraph styles."""
        # Extract headings from paragraph-level headings
        # Use first occurrence of each heading level
        h1 = None
        h2 = None
        h3 = None
        h4 = None
        
        for para in doc.paragraphs:
            if para.heading_level == HeadingLevel.HEADING_1 and h1 is None:
                h1 = para.text.strip()
            elif para.heading_level == HeadingLevel.HEADING_2 and h2 is None:
                h2 = para.text.strip()
            elif para.heading_level == HeadingLevel.HEADING_3 and h3 is None:
                h3 = para.text.strip()
            elif para.heading_level == HeadingLevel.HEADING_4 and h4 is None:
                h4 = para.text.strip()
        
        # Use context values if provided, otherwise use extracted headings
        book = context.get("book") or h1 or ""
        sefer = context.get("sefer") or h2
        parshah = context.get("parshah") or h3
        filename = context.get("filename") or h4
        
        # Set document-level headings
        doc.set_headings(h1=book, h2=sefer, h3=parshah, h4=filename)
        
        # No additional processing needed - file is already formatted
        # Paragraph-level headings are already set by the reader from Word styles
        
        return doc

