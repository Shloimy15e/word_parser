"""
Perek-H2 format handler.
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
class PerekH2Format(DocumentFormat):
    """
    Perek-H2 format.
    
    Structure:
    - H1: Book (from context)
    - H2: Derived from lines matching "פרק *" followed by short sentences
    - H3: Bold single-line sentences (short sentences that are bold)
    - No H4s
    
    Detection: Explicit format selection only.
    """

    @classmethod
    def get_format_name(cls) -> str:
        return "perek-h2"

    @classmethod
    def get_priority(cls) -> int:
        return 15  # Higher than special-heading (10) to ensure it takes precedence

    @classmethod
    def detect(cls, doc: Document, context: Dict[str, Any]) -> bool:
        return (
            context.get("mode") == "perek-h2"
            or context.get("format") == "perek-h2"
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

        # Remove page markings first
        doc = remove_page_markings(doc)
        
        #Remove תמונה ** (number or letter up to 3 characters can optionally be followed by space or words/letters such as תמונה 144 א - ב)
        # Examples
        # מונה 9 – א' תמונה 9 – ב'
        # {תמונה 58}
        # תמונה 8 – המשך
        # תמונה 149 א – ב – ג
        # תמונה 143 א – ב
        # תמונה 41 – א – ב – ג – ד
        # תמונה 154 א – ב
        # תמונה 42 – א - ב
        # תמונה 236 א – ב
        # תמונה 32 – א' תמונה 32 – ב'
        # תמונה המשך 35
        # תמונה 127 א – ב – ג- ד-
        # תמונה 39 – א – ב – ג- ד- ה
        # תמונה 47 – א – ב – ג –
        # תמונה המשך 37 
        # תמונה 130 א – ב
        # {תמונה 3} {תמונה 3 – ב}
        # תמונה 147 א – ב
        doc.paragraphs = [para for para in doc.paragraphs if not re.match(r"^\{?\s*תמונה\s*(?:המשך\s+)?([א-ת]{1,3}|\d{1,3})\s*(?:[א-ת'\s–-]+(?:\{?\s*תמונה\s*[א-ת\d\s–-]+\}?)?)*\s*\}?$", para.text.strip())]

        # Set base headings - H2 will be derived from document content, not from context
        doc.set_headings(h1=book, h2=None)

        # Scan for perek-h2 patterns
        self._apply_perek_h2_headings(doc)

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

    def _apply_perek_h2_headings(self, doc: Document) -> None:
        """Apply perek-h2 heading detection logic."""
        print(f"Perek-H2 headings: processing {len(doc.paragraphs)} paragraphs")

        # Pattern for "פרק *" - matches "פרק" followed by optional space and any text
        perek_pattern = re.compile(r"^פרק\s+.*$")

        i = 0
        while i < len(doc.paragraphs):
            para = doc.paragraphs[i]
            text = para.text.strip()

            # Skip empty paragraphs
            if not text:
                i += 1
                continue

            # Check for H2: "פרק *" - ONLY the "פרק *" line itself becomes H2
            if perek_pattern.match(text):
                # The "פרק *" paragraph itself becomes H2
                para.heading_level = HeadingLevel.HEADING_2
                # Don't set document-level H2 - the paragraph itself is the heading
                print(f"Found H2 (פרק *): {text[:50]}")
                i += 1
                continue

            # Check for H3: bold and underlined (regardless of length) and short sentence
            # If its followed by another such heading, they should be combined into one H3
            if para.heading_level == HeadingLevel.NORMAL:
                is_underlined = para.runs and any(run.style.underline for run in para.runs if run.style.underline is not None)
                if self._is_bold(para) and is_underlined and self._is_short_sentence(text):
                    # Check if the next paragraph is also a bold and underlined short sentence
                    if i + 1 < len(doc.paragraphs):
                        next_para = doc.paragraphs[i + 1]
                        next_text = next_para.text.strip()
                        next_is_underlined = next_para.runs and any(run.style.underline for run in next_para.runs if run.style.underline is not None)
                        if (next_para.heading_level == HeadingLevel.NORMAL and 
                            self._is_bold(next_para) and 
                            next_is_underlined and 
                            self._is_short_sentence(next_text)):
                            # Combine the two paragraphs into one H3
                            combined_text = text + " " + next_text
                            para.text = combined_text
                            para.heading_level = HeadingLevel.HEADING_3
                            # Remove the next paragraph from the list
                            doc.paragraphs.pop(i + 1)
                            print(f"Found H3 (bold and underlined, combined): {combined_text[:50]}")
                            i += 1
                            continue
                    para.heading_level = HeadingLevel.HEADING_3
                    print(f"Found H3 (bold and underlined): {text[:50]}")
                    i += 1
                    continue
            
            # Check for H3: bold single-line sentence
            if self._is_bold(para) and self._is_short_sentence(text):
                # Only set as H3 if:
                # 1. Not already a heading
                # 2. Does not start with asterisk
                # 3. Previous paragraph is NOT H3 (no consecutive H3s) and NOT an asterisk
                # 4. Followed by content paragraph
                if para.heading_level == HeadingLevel.NORMAL:                    
                    # Check if paragraph starts with asterisk
                    starts_with_asterisk = text.strip().startswith("*")
                    
                    # Check previous paragraph is not H3 and not an asterisk
                    # Note: Previous can be H2 (like "פרק א") - that's fine
                    prev_is_h3 = False
                    prev_is_asterisk = False
                    if i > 0:
                        prev_para = doc.paragraphs[i - 1]
                        prev_text = prev_para.text.strip()
                        # Only block if previous is H3 (consecutive H3s not allowed)
                        # H2 before H3 is fine
                        if prev_para.heading_level == HeadingLevel.HEADING_3:
                            prev_is_h3 = True
                        # Check if previous paragraph is just an asterisk (or asterisk with spaces)
                        if prev_text in ("*", " *", "* ", " * "):
                            prev_is_asterisk = True
                    
                    # Check if there's a next paragraph with NORMAL content (means a paragraph with text)
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
                            next_para.heading_level == HeadingLevel.NORMAL and
                            # עזר בקודש needs this to be false for the first one
                            not self._is_bold(next_para)):
                            has_following_content = True
                    
                    # Only set as H3 if:
                    # - Doesn't start with asterisk
                    # - Previous is not H3 and not asterisk
                    # - There's following content
                    if not starts_with_asterisk and not prev_is_h3 and not prev_is_asterisk and has_following_content:
                        para.heading_level = HeadingLevel.HEADING_3
                        print(f"Found H3 (bold short sentence): {text[:50]}")

            i += 1

