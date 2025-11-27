"""
Minimal format handler - only cleans markers, leaves everything else as-is.
"""

import re
from typing import Dict, Any

from word_parser.core.formats._base import (
    DocumentFormat,
    FormatRegistry,
    Document,
    HeadingLevel,
    remove_page_markings,
)


@FormatRegistry.register
class MinimalFormat(DocumentFormat):
    """
    Minimal format handler - only cleans markers.
    
    Structure:
    - Removes @ markers (like @99, @88, etc.)
    - Removes דף markers (like דף א, דף ב, etc.)
    - Leaves everything else as-is (no heading detection, no merging)
    """

    @classmethod
    def get_format_name(cls) -> str:
        return "minimal"

    @classmethod
    def get_priority(cls) -> int:
        return 15

    @classmethod
    def detect(cls, doc: Document, context: Dict[str, Any]) -> bool:
        return (
            context.get("mode") == "minimal"
            or context.get("format") == "minimal"
        )

    @classmethod
    def get_optional_context(cls) -> Dict[str, Any]:
        return {
            "book": None,
            "sefer": None,
        }

    def process(self, doc: Document, context: Dict[str, Any]) -> Document:
        """Process document - only clean markers and remove old headings, leave everything else as-is."""
        # Remove page markings first
        doc = remove_page_markings(doc)
        
        # Clean @ markers
        self._clean_at_markers(doc)
        
        # Clean דף markers
        self._clean_daf_markers(doc)
        
        # Remove old headings (H1 and H2)
        self._remove_old_headings(doc)
        
        return doc

    def _remove_old_headings(self, doc: Document) -> None:
        """
        Remove old heading levels (H1 and H2) from paragraphs.
        Converts them to normal paragraphs.
        """
        print(f"Minimal format: removing old headings from {len(doc.paragraphs)} paragraphs")
        removed_h1_count = 0
        removed_h2_count = 0
        for para in doc.paragraphs:
            if para.heading_level == HeadingLevel.HEADING_1:
                para.heading_level = HeadingLevel.NORMAL
                removed_h1_count += 1
                print(f"  -> Removed H1 from paragraph: '{para.text[:50] if para.text else ''}'")
            elif para.heading_level == HeadingLevel.HEADING_2:
                para.heading_level = HeadingLevel.NORMAL
                removed_h2_count += 1
                print(f"  -> Removed H2 from paragraph: '{para.text[:50] if para.text else ''}'")
        
        if removed_h1_count > 0 or removed_h2_count > 0:
            print(f"Minimal format: removed H1 from {removed_h1_count} paragraph(s), H2 from {removed_h2_count} paragraph(s)")

    def _clean_at_markers(self, doc: Document) -> None:
        """
        Remove @ markers (like @99, @88, @22, etc.) from all paragraph text.
        """
        print(f"Minimal format: cleaning @ markers from {len(doc.paragraphs)} paragraphs")
        cleaned_count = 0
        for para in doc.paragraphs:
            if para.text:
                original_text = para.text
                # Remove @ followed by one or more digits
                cleaned = re.sub(r"@[0-9]+", "", para.text)
                if cleaned != original_text:
                    para.text = cleaned
                    # Clean up any double spaces that might result
                    para.text = re.sub(r"\s+", " ", para.text)
                    para.text = para.text.strip()
                    cleaned_count += 1
        
        if cleaned_count > 0:
            print(f"Minimal format: cleaned @ markers from {cleaned_count} paragraph(s)")

    def _clean_daf_markers(self, doc: Document) -> None:
        """
        Remove "דף _hebrew_letters_" patterns from paragraphs.
        """
        print(f"Minimal format: cleaning דף markers from {len(doc.paragraphs)} paragraphs")
        
        # Pattern to match "דף" followed by one or more Hebrew letters
        daf_pattern = re.compile(r'דף\s+[א-ת]+')
        
        cleaned_count = 0
        for para in doc.paragraphs:
            if para.text:
                original_text = para.text
                # Remove דף pattern from the text
                cleaned_text = daf_pattern.sub('', para.text).strip()
                
                # Clean up any double spaces that might result
                cleaned_text = re.sub(r'\s+', ' ', cleaned_text)
                
                if cleaned_text != original_text:
                    para.text = cleaned_text
                    cleaned_count += 1
        
        if cleaned_count > 0:
            print(f"Minimal format: cleaned דף markers from {cleaned_count} paragraph(s)")

