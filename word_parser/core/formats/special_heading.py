"""
Special heading format handler.
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
class SpecialHeadingFormat(DocumentFormat):
    """
    Special heading format.

    Structure:
    - H3: Determined by a preceding line with pattern:
        1. Hebrew word followed by a period (e.g. "מילה.")
        2. OR "–– heb_word ––"
        3. OR "heb_word – [heb_word]"
        4. OR "[heb_word] – heb_word"

    The line AFTER the pattern becomes Heading 3.
    """

    @classmethod
    def get_format_name(cls) -> str:
        return "special-heading"

    @classmethod
    def get_priority(cls) -> int:
        return 10

    @classmethod
    def detect(cls, doc: Document, context: Dict[str, Any]) -> bool:
        return (
            context.get("mode") == "special-heading"
            or context.get("format") == "special-heading"
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
        sefer = context.get("sefer", "")

        # Remove page markings first
        doc = remove_page_markings(doc)

        # Set base headings
        doc.set_headings(h1=book, h2=sefer)

        # Scan for special heading patterns
        self._apply_special_headings(doc)

        return doc

    def _apply_special_headings(self, doc: Document) -> None:
        print(f"Special headings: {len(doc.paragraphs)}")

        # Pattern 1: Hebrew word followed by period
        pattern1 = re.compile(r"^[\u0590-\u05ff]+\.$")

        # Pattern 2: –– heb_word ––
        pattern2 = re.compile(r"^[-–—]+\s*[\u0590-\u05ff]+\s*[-–—]+$")

        # Pattern 3: heb_word – [heb_word_or_letter]
        pattern3 = re.compile(r"^[\u0590-\u05ff]+\s*[-–—]+\s*\[[\u0590-\u05ff]+\]\s*$")

        # Pattern 4: [heb_word_or_letter] – heb_word
        pattern4 = re.compile(r"^\[[\u0590-\u05ff]+\]\s*[-–—]+\s*[\u0590-\u05ff]+$")

        # Pattern 5: Single or double Hebrew letter (e.g., "ב", "א", "ג", "יא", "יב", "יג")
        pattern5 = re.compile(r"^[\u0590-\u05ff]{1,2}$")

        i = 0
        while i < len(doc.paragraphs) - 1:
            para = doc.paragraphs[i]
            text = para.text.strip()

            # Skip empty paragraphs
            if not text:
                i += 1
                continue

            is_match = False
            matched_pattern = None
            if pattern1.match(text):
                is_match = True
                matched_pattern = "Pattern 1 (heb_word.)"
            elif pattern2.match(text):
                is_match = True
                matched_pattern = "Pattern 2 (–– heb_word ––)"
            elif pattern3.match(text):
                is_match = True
                matched_pattern = "Pattern 3 (heb_word – [heb_word])"
            elif pattern4.match(text):
                is_match = True
                matched_pattern = "Pattern 4 ([heb_word] – heb_word)"
            elif pattern5.match(text):
                is_match = True
                matched_pattern = "Pattern 5 (single Hebrew letter)"

            if is_match:
                print(f"Found match at paragraph {i} ({matched_pattern}): '{text}'")
                # The NEXT paragraph is the heading
                if i + 1 < len(doc.paragraphs):
                    next_para = doc.paragraphs[i + 1]
                    next_text = next_para.text.strip()
                    if next_text:  # Ensure next para is not empty
                        next_para.heading_level = HeadingLevel.HEADING_3
                        print(f"  -> Set paragraph {i+1} as H3: '{next_text[:50]}'")
                    else:
                        print(f"  -> Warning: Next paragraph is empty, skipping")
                else:
                    print(f"  -> Warning: No next paragraph available")

            i += 1

