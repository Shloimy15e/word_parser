"""
Folder-filename format handler - documents with folder-based structure.
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
from word_parser.core.processing import is_valid_gematria_number


@FormatRegistry.register
class FolderFilenameFormat(DocumentFormat):
    """
    Format for documents with folder-based structure.
    
    Structure:
    - H1: Folder name (parent directory)
    - H2: Filename (without extension)
    - H3: One-line sentences detected from content
    - H4: One-line sentences that follow an H3 (consecutive headings)
    
    Detection: Explicit format selection only.
    """
    
    @classmethod
    def get_format_name(cls) -> str:
        return "folder-filename"
    
    @classmethod
    def get_priority(cls) -> int:
        return 15
    
    @classmethod
    def detect(cls, doc: Document, context: Dict[str, Any]) -> bool:
        return (
            context.get("mode") == "folder-filename"
            or context.get("format") == "folder-filename"
        )
    
    @classmethod
    def get_optional_context(cls) -> Dict[str, Any]:
        return {
            "book": None,  # Can override folder name
            "filename": None,  # Can override filename
            "input_path": None,  # Used to extract folder/filename
        }
    
    def process(self, doc: Document, context: Dict[str, Any]) -> Document:
        """Process document with folder-filename structure."""
        # Get H1 from folder name (parent directory of input_path)
        book = context.get("book", "")
        if not book:
            input_path = context.get("input_path", "")
            if input_path:
                try:
                    folder_name = Path(input_path).parent.name
                    if folder_name:
                        book = folder_name
                except Exception:
                    pass
        
        # Get H2 from filename (stem of input_path)
        filename_h2 = context.get("filename", "")
        if not filename_h2:
            input_path = context.get("input_path", "")
            if input_path:
                try:
                    filename_h2 = Path(input_path).stem
                except Exception:
                    pass
        
        # Remove page markings first
        doc = remove_page_markings(doc)
        
        # Set base headings - H1 from folder, H2 from filename
        doc.set_headings(h1=book, h2=filename_h2, h3=None, h4=None)
        
        # Detect one-line sentences and mark them as H3
        self._detect_h3_sentences(doc)
        
        # Remove paragraphs that exactly match H1 or H2 (except actual heading paragraphs)
        self._remove_duplicate_headings(doc, book, filename_h2)
        
        return doc
    
    def _is_one_line_sentence(self, text: str) -> bool:
        """
        Check if text is a one-line sentence that should be H3.
        Criteria:
        - Single line (no newlines)
        - Contains Hebrew text
        - Reasonable length (not too short, not too long)
        - Looks like a sentence/heading (not just a word or marker)
        - Does NOT start with a siman marking (e.g., "ריב.")
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
        
        # Skip if starts with siman marking (Hebrew letters followed by period)
        # Pattern: 1-4 Hebrew letters (valid gematria) followed by period and space
        siman_match = re.match(r"^([א-ת]{1,4})\.\s+", text)
        if siman_match:
            siman_text = siman_match.group(1)
            if is_valid_gematria_number(siman_text):
                return False  # This is a siman marking, not a heading
        
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
    
    def _detect_h3_sentences(self, doc: Document) -> None:
        """Detect one-line sentences and mark them as H3 or H4."""
        print(f"Folder-filename format: detecting H3/H4 sentences in {len(doc.paragraphs)} paragraphs")
        
        prev_heading_level = None
        
        for para in doc.paragraphs:
            txt = para.text.strip()
            
            # Skip empty paragraphs
            if not txt:
                prev_heading_level = None
                continue
            
            # Track the previous paragraph's heading level
            current_heading_level = para.heading_level
            
            # If already has a heading level, update prev_heading_level and continue
            if para.heading_level != HeadingLevel.NORMAL:
                prev_heading_level = para.heading_level
                continue
            
            # Check if this is a one-line sentence
            if self._is_one_line_sentence(txt):
                # If previous paragraph was H3, mark this as H4
                if prev_heading_level == HeadingLevel.HEADING_3:
                    para.heading_level = HeadingLevel.HEADING_4
                    print(f"  -> Detected H4 (follows H3): '{txt[:50]}'")
                else:
                    para.heading_level = HeadingLevel.HEADING_3
                    print(f"  -> Detected H3: '{txt[:50]}'")
                prev_heading_level = para.heading_level
            else:
                # Not a heading, reset prev_heading_level
                prev_heading_level = None
    
    def _remove_duplicate_headings(self, doc: Document, h1: str, h2: str) -> None:
        """
        Remove paragraphs that exactly match H1 or H2 text.
        Preserves paragraphs that are already marked as headings.
        """
        if not h1 and not h2:
            return
        
        filtered_paragraphs = []
        removed_count = 0
        
        for para in doc.paragraphs:
            txt = para.text.strip()
            
            # Keep paragraphs that are already marked as headings
            if para.heading_level != HeadingLevel.NORMAL:
                filtered_paragraphs.append(para)
                continue
            
            # Remove content paragraphs that exactly match H1 or H2
            if h1 and txt == h1:
                removed_count += 1
                print(f"  -> Removed duplicate H1: '{txt[:50]}'")
                continue
            
            if h2 and txt == h2:
                removed_count += 1
                print(f"  -> Removed duplicate H2: '{txt[:50]}'")
                continue
            
            # Keep all other paragraphs
            filtered_paragraphs.append(para)
        
        if removed_count > 0:
            print(f"Folder-filename format: removed {removed_count} duplicate heading paragraph(s)")
        
        doc.paragraphs = filtered_paragraphs

