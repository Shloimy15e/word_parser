"""
H2-only format handler - documents that only have Heading 2s (one-line sentences).
"""

import re
from typing import Dict, Any
from pathlib import Path

from word_parser.core.formats._base import (
    DocumentFormat,
    FormatRegistry,
    Document,
    HeadingLevel,
    remove_page_markings,
)


@FormatRegistry.register
class H2OnlyFormat(DocumentFormat):
    """
    Format for documents that only have Heading 2s (one-line sentences).
    
    Structure:
    - H1: Book (from context or filename)
    - H2: One-line sentences detected from content
    - No H3 or H4
    
    Detection: Explicit format selection only.
    """
    
    @classmethod
    def get_format_name(cls) -> str:
        return "h2-only"
    
    @classmethod
    def get_priority(cls) -> int:
        return 15
    
    @classmethod
    def detect(cls, doc: Document, context: Dict[str, Any]) -> bool:
        return (
            context.get("mode") == "h2-only"
            or context.get("format") == "h2-only"
        )
    
    @classmethod
    def get_optional_context(cls) -> Dict[str, Any]:
        return {
            "book": None,
            "filename": None,
        }
    
    def process(self, doc: Document, context: Dict[str, Any]) -> Document:
        """Process document with H2-only structure."""
        # Get H1 from context or filename
        book = context.get("book", "")
        if not book:
            filename = context.get("filename", "")
            if filename:
                # Try to extract book name from filename
                book = Path(filename).stem
        
        # Remove page markings first
        doc = remove_page_markings(doc)
        
        # Set base headings - H1 from context/filename, no H2 at document level
        doc.set_headings(h1=book, h2=None, h3=None, h4=None)
        
        # Detect one-line sentences and mark them as H2
        self._detect_h2_sentences(doc)
        
        return doc
    
    def _is_one_line_sentence(self, text: str) -> bool:
        """
        Check if text is a one-line sentence that should be H2.
        Criteria:
        - Single line (no newlines)
        - Contains Hebrew text
        - Reasonable length (not too short, not too long)
        - Looks like a sentence/heading (not just a word or marker)
        """
        if not text:
            return False
        
        text = text.strip()
        
        # Must be single line
        if "\n" in text:
            return False
        
        # Must have Hebrew content
        if not any("\u0590" <= c <= "\u05ff" for c in text):
            return False
        
        # Should be reasonably long (at least 5 chars) but not too long (< 200 chars)
        if len(text) < 5 or len(text) >= 200:
            return False
        
        # Should not be just a marker or single word
        # Check if it has multiple words (spaces) or punctuation that suggests a sentence
        has_spaces = ' ' in text
        has_punctuation = any(c in text for c in '.,;:!?')
        
        # If it's a single word without punctuation, it's probably not a sentence
        if not has_spaces and not has_punctuation:
            return False
        
        # Skip common markers
        markers = ("h", "q", "Y", "*", "***", "* * *")
        if text in markers:
            return False
        
        return True
    
    def _detect_h2_sentences(self, doc: Document) -> None:
        """Detect one-line sentences and mark them as H2."""
        print(f"H2-only format: detecting H2 sentences in {len(doc.paragraphs)} paragraphs")
        
        for para in doc.paragraphs:
            txt = para.text.strip()
            
            # Skip empty paragraphs
            if not txt:
                continue
            
            # Skip if already has a heading level
            if para.heading_level != HeadingLevel.NORMAL:
                continue
            
            # Check if this is a one-line sentence
            if self._is_one_line_sentence(txt):
                para.heading_level = HeadingLevel.HEADING_2
                print(f"  -> Detected H2: '{txt[:50]}'")

