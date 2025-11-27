"""
Perek-H3 format handler (like perek-h2 but H2 from context/folder, H3 without bold requirement).
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
class PerekH3Format(DocumentFormat):
    """
    Perek-H3 format (like perek-h2 but H2 from context/folder, H3 without bold requirement).
    
    Structure:
    - H1: Book (from context)
    - H2: Sefer (from context arg or folder name, NOT from document)
    - H3: Underlined short sentences OR just short sentences (without bold requirement)
    - No H4s
    
    Detection: Explicit format selection only.
    """

    @classmethod
    def get_format_name(cls) -> str:
        return "perek-h3"

    @classmethod
    def get_priority(cls) -> int:
        return 15  # Same priority as perek-h2

    @classmethod
    def detect(cls, doc: Document, context: Dict[str, Any]) -> bool:
        return (
            context.get("mode") == "perek-h3"
            or context.get("format") == "perek-h3"
        )

    @classmethod
    def get_optional_context(cls) -> Dict[str, Any]:
        return {
            "book": None,
            "sefer": None,
            "filename": None,
        }

    def process(self, doc: Document, context: Dict[str, Any]) -> Document:
        book = context.get("book", "")
        # H2 comes from context (sefer arg) or folder name, not from document
        sefer = context.get("sefer")
        if not sefer:
            # Try to get from input_path folder name
            input_path = context.get("input_path", "")
            if input_path:
                folder_name = Path(input_path).parent.name
                if folder_name:
                    sefer = folder_name

        # Remove page markings first
        doc = remove_page_markings(doc)
        
        # Remove תמונה patterns (same as perek-h2)
        doc.paragraphs = [para for para in doc.paragraphs if not re.match(r"^\{?\s*תמונה\s*(?:המשך\s+)?([א-ת]{1,3}|\d{1,3})\s*(?:[א-ת'\s–-]+(?:\{?\s*תמונה\s*[א-ת\d\s–-]+\}?)?)*\s*\}?$", para.text.strip())]

        # Remove "בס"ד" header and book name header (usually first 1-2 paragraphs)
        # Filter out "בס"ד" and book name (if it matches the book from context)
        filtered_paragraphs = []
        skip_count = 0
        for i, para in enumerate(doc.paragraphs):
            text = para.text.strip()
            # Skip "בס"ד" (with or without quotes/variations)
            if re.match(r'^בס["\']?ד["\']?$', text):
                skip_count += 1
                continue
            # Skip book name if it matches context book (usually first or second paragraph after בס"ד)
            if book and i < 3 and text == book:
                skip_count += 1
                continue
            # Skip if it's the book name from context (allow some variations)
            if book and i < 3 and book in text and len(text) <= len(book) + 5:
                skip_count += 1
                continue
            filtered_paragraphs.append(para)
        doc.paragraphs = filtered_paragraphs

        # Set base headings - H2 from context/folder, not from document
        doc.set_headings(h1=book, h2=sefer)

        # Scan for H3 patterns (including bold sentences)
        self._apply_perek_h3_headings(doc)

        return doc

    def _is_short_sentence(self, text: str) -> bool:
        """
        Check if text is a short sentence.
        A short sentence is:
        - Not too long (< 60 characters, similar to should_start_content check)
        - Contains Hebrew text
        - Is a single line (no newlines)
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
        
        # Must be reasonably short (less than 60 chars)
        if len(text) >= 60:
            return False
        
        # Should not be empty or just whitespace
        if not text:
            return False
        
        return True

    def _is_bold(self, para, allow_partial: bool = True) -> bool:
        """Check if paragraph is bold (any run is bold)."""
        if not para.runs:
            return False
        if allow_partial:
            return any(run.style.bold for run in para.runs if run.style.bold is not None)
        else:
            return all(run.style.bold for run in para.runs if run.style.bold is not None)

    def _apply_perek_h3_headings(self, doc: Document) -> None:
        """Apply perek-h3 heading detection logic - H3 without bold requirement."""
        print(f"Perek-H3 headings: processing {len(doc.paragraphs)} paragraphs")

        i = 0
        while i < len(doc.paragraphs):
            para = doc.paragraphs[i]
            text = para.text.strip()

            # Skip empty paragraphs
            if not text:
                i += 1
                continue

            # Skip if already a heading
            if para.heading_level != HeadingLevel.NORMAL:
                i += 1
                continue

            # Check for H3: bold short sentences (first priority - like perek-h2 but without underline requirement)
            if self._is_bold(para) and self._is_short_sentence(text):
                # Check if the next paragraph is also a bold short sentence
                if i + 1 < len(doc.paragraphs):
                    next_para = doc.paragraphs[i + 1]
                    next_text = next_para.text.strip()
                    if (next_para.heading_level == HeadingLevel.NORMAL and 
                        self._is_bold(next_para) and 
                        self._is_short_sentence(next_text)):
                        # Combine the two paragraphs into one H3
                        combined_text = text + " " + next_text
                        para.text = combined_text
                        para.heading_level = HeadingLevel.HEADING_3
                        # Remove the next paragraph from the list
                        doc.paragraphs.pop(i + 1)
                        print(f"Found H3 (bold, combined): {combined_text[:50]}")
                        i += 1
                        continue
                
                # Check if there's following content (like perek-h2 logic)
                has_following_content = False
                if i + 1 < len(doc.paragraphs):
                    next_para = doc.paragraphs[i + 1]
                    next_text = re.sub(r"@\d+", "", next_para.text.strip())
                    if (next_text and 
                        len(next_text) > 3 and
                        not next_text.startswith("פרק") and
                        not next_text in ("*", " *", "* ", " * ") and
                        next_para.heading_level == HeadingLevel.NORMAL):
                        has_following_content = True
                
                if has_following_content:
                    para.heading_level = HeadingLevel.HEADING_3
                    print(f"Found H3 (bold short sentence): {text[:50]}")
                    i += 1
                    continue

            # Check for H3: underlined short sentences (without bold requirement)
            # If followed by another such heading, combine them
            is_underlined = para.runs and any(run.style.underline for run in para.runs if run.style.underline is not None)
            if is_underlined and self._is_short_sentence(text):
                # Check if the next paragraph is also an underlined short sentence
                if i + 1 < len(doc.paragraphs):
                    next_para = doc.paragraphs[i + 1]
                    next_text = next_para.text.strip()
                    next_is_underlined = next_para.runs and any(run.style.underline for run in next_para.runs if run.style.underline is not None)
                    if (next_para.heading_level == HeadingLevel.NORMAL and 
                        next_is_underlined and 
                        self._is_short_sentence(next_text)):
                        # Combine the two paragraphs into one H3
                        combined_text = text + " " + next_text
                        para.text = combined_text
                        para.heading_level = HeadingLevel.HEADING_3
                        # Remove the next paragraph from the list
                        doc.paragraphs.pop(i + 1)
                        print(f"Found H3 (underlined, combined): {combined_text[:50]}")
                        i += 1
                        continue
                para.heading_level = HeadingLevel.HEADING_3
                print(f"Found H3 (underlined): {text[:50]}")
                i += 1
                continue
            
            # Check for H3: just short sentences (without bold or underline requirement)
            if self._is_short_sentence(text):
                # Only set as H3 if:
                # 1. Not already a heading
                # 2. Does not start with asterisk
                # 3. Previous paragraph is NOT H3 (no consecutive H3s) and NOT an asterisk
                # 4. Followed by content paragraph
                # Check if paragraph starts with asterisk
                starts_with_asterisk = text.strip().startswith("*")
                
                # Check previous paragraph is not H3 and not an asterisk
                prev_is_h3 = False
                prev_is_asterisk = False
                if i > 0:
                    prev_para = doc.paragraphs[i - 1]
                    prev_text = prev_para.text.strip()
                    # Only block if previous is H3 (consecutive H3s not allowed)
                    if prev_para.heading_level == HeadingLevel.HEADING_3:
                        prev_is_h3 = True
                    # Check if previous paragraph is just an asterisk (or asterisk with spaces)
                    if prev_text in ("*", " *", "* ", " * "):
                        prev_is_asterisk = True
                
                # Check if there's a next paragraph with NORMAL content
                has_following_content = False
                if i + 1 < len(doc.paragraphs):
                    next_para = doc.paragraphs[i + 1]
                    next_text = re.sub(r"@\d+", "", next_para.text.strip())
                    # Remove noise such as @0075
                    # Next paragraph should be:
                    # - Not empty (must have actual text)
                    # - Not already a heading
                    # - Not bold (if bold and short, it could become a heading)
                    # - Actual content paragraph with substantial text (at least a few characters)
                    if (next_text and 
                        len(next_text) > 3 and  # Must have substantial content, not just a few chars
                        not next_text.startswith("פרק") and
                        not next_text in ("*", " *", "* ", " * ") and
                        next_para.heading_level == HeadingLevel.NORMAL):
                        # Check if next para is bold - if so, it might be a heading candidate
                        next_is_bold = next_para.runs and any(run.style.bold for run in next_para.runs if run.style.bold is not None)
                        if not next_is_bold or len(next_text) > 60:  # Allow if not bold or if it's long enough
                            has_following_content = True
                
                # Only set as H3 if:
                # - Doesn't start with asterisk
                # - Previous is not H3 and not asterisk
                # - There's following content
                if not starts_with_asterisk and not prev_is_h3 and not prev_is_asterisk and has_following_content:
                    para.heading_level = HeadingLevel.HEADING_3
                    print(f"Found H3 (short sentence): {text[:50]}")

            i += 1

