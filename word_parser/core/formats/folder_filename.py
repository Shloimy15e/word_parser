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

        # Detect footnotes section and mark it
        footnote_start_idx = self._detect_footnotes_start(doc)

        # Detect one-line sentences and mark them as H3 (only in main content, not footnotes)
        self._detect_h3_sentences(doc, footnote_start_idx)

        # Remove paragraphs that exactly match H1 or H2 (except actual heading paragraphs and footnotes)
        self._remove_duplicate_headings(doc, book, filename_h2, footnote_start_idx)

        return doc

    def _is_one_line_sentence(self, text: str) -> bool:
        """
        Check if text is a one-line SINGLE SENTENCE that should be H3.
        Simple rule: one-line SINGLE SENTENCE = heading (unless it's a list item).

        Criteria:
        - Single line (no newlines)
        - Contains Hebrew text
        - Reasonable length (not too short, not too long)
        - Must be a SINGLE SENTENCE (ends with sentence punctuation, not multiple sentences)
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
        siman_match = False  # re.match(r"^([א-ת]{1,4})\.\s+", text)
        if siman_match:
            siman_text = siman_match.group(1)
            if is_valid_gematria_number(siman_text):
                return False  # This is a siman marking, not a heading

        # ENFORCE: Must be a SINGLE SENTENCE
        # Count sentence-ending punctuation marks (. ! ?)
        sentence_endings = text.count('.') + text.count('!') + text.count('?')
        
        # If it has multiple sentence endings, it's multiple sentences - NOT a heading
        if sentence_endings > 1:
            return False
        
        # If it has no sentence-ending punctuation, it might not be a complete sentence
        # But allow it if it's short and looks like a heading phrase
        if sentence_endings == 0:
            # Allow short phrases without punctuation (like "סדר ליל שבת קודש")
            if len(text) > 80:  # Too long to be a heading without punctuation
                return False
        
        # Should be reasonably long (at least 3 chars) but not too long
        # For a true single sentence heading, limit to ~150 chars max
        if len(text) < 3 or len(text) > 150:
            return False

        # Skip common markers
        markers = ("h", "q", "Y", "*", "***", "* * *")
        if text in markers:
            return False

        # If it passes all checks above, it's a one-line SINGLE SENTENCE = heading
        return True

    def _detect_footnotes_start(self, doc: Document) -> int:
        """
        Detect where footnotes start in the document.
        Returns the index of the first footnote paragraph, or len(doc.paragraphs) if no footnotes found.

        Footnotes are detected by:
        1. Horizontal separator lines (---, ___, ===, or similar repeating characters)
        2. Paragraphs starting with footnote markers like (א), (ב), א., ב., etc.
        """
        for i, para in enumerate(doc.paragraphs):
            txt = para.text.strip()

            # Check for horizontal separator line
            # Pattern: 3+ repeating dashes, underscores, equals, or similar characters
            if re.match(r"^[-=_~\.]{3,}$", txt):
                # Found separator - footnotes start after this
                if i + 1 < len(doc.paragraphs):
                    print(
                        f"Folder-filename format: detected footnotes separator at paragraph {i+1}"
                    )
                    return i + 1

            # Check for footnote marker patterns
            # Pattern 1: (א), (ב), etc. - Hebrew letter in parentheses
            if re.match(r"^\([א-ת]\)", txt):
                print(
                    f"Folder-filename format: detected footnotes starting at paragraph {i+1}"
                )
                return i

            # Pattern 2: א., ב., etc. - Hebrew letter followed by period
            if re.match(r"^[א-ת]\.\s", txt):
                # Make sure it's not a siman marking (which would be longer)
                # If the paragraph is short or looks like a footnote reference, it's a footnote
                if len(txt) < 100:  # Footnotes are typically shorter
                    print(
                        f"Folder-filename format: detected footnotes starting at paragraph {i+1}"
                    )
                    return i

        # No footnotes detected
        return len(doc.paragraphs)

    def _detect_h3_sentences(
        self, doc: Document, footnote_start_idx: int = None
    ) -> None:
        """
        Detect one-line sentences and mark them as H3 or H4.
        Skips footnotes section if footnote_start_idx is provided.
        """
        if footnote_start_idx is None:
            footnote_start_idx = len(doc.paragraphs)

        print(
            f"Folder-filename format: detecting H3/H4 sentences in {len(doc.paragraphs)} paragraphs (footnotes start at {footnote_start_idx})"
        )

        prev_heading_level = None

        for i, para in enumerate(doc.paragraphs):
            # Skip footnotes section
            if i >= footnote_start_idx:
                continue

            txt = para.text.strip()

            # Skip empty paragraphs
            if not txt:
                prev_heading_level = None
                continue

            # if its בס"ד, skip it
            if txt == 'בס"ד':
                prev_heading_level = None
                continue
            
            # If already has a heading level, update prev_heading_level and continue
            if para.heading_level != HeadingLevel.NORMAL:
                prev_heading_level = para.heading_level
                continue

            # Skip numbered list items - they should remain as list items with their formatting
            # "List Paragraph" style can be headings, but actual numbered list items cannot
            is_numbered = para.is_numbered_list_item()
            if is_numbered:
                print(f"  -> Skipping numbered list item: '{txt[:50]}'")
                prev_heading_level = None
                continue
            elif txt.startswith('יט.'):
                # Debug: check why this wasn't detected
                print(f"  -> DEBUG: Text starts with 'יט.' but is_numbered_list_item() returned False")
                print(f"     Style: {para.style_name}, Metadata: {para.metadata}")
                print(f"     Full text: '{txt[:100]}'")

            # Check if this is a one-line sentence (heading)
            # Simple rule: one-line sentence = heading (unless it's a list item, which we already skipped)
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

    def _remove_duplicate_headings(
        self, doc: Document, h1: str, h2: str, footnote_start_idx: int = None
    ) -> None:
        """
        Remove paragraphs that exactly match H1 or H2 text.
        Preserves paragraphs that are already marked as headings and footnotes.
        """
        if not h1 and not h2:
            return

        if footnote_start_idx is None:
            footnote_start_idx = len(doc.paragraphs)

        filtered_paragraphs = []
        removed_count = 0

        for i, para in enumerate(doc.paragraphs):
            # Always keep footnotes intact
            if i >= footnote_start_idx:
                filtered_paragraphs.append(para)
                continue

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
            print(
                f"Folder-filename format: removed {removed_count} duplicate heading paragraph(s)"
            )

        doc.paragraphs = filtered_paragraphs
