"""
Pound format handler - uses # markers to determine heading structure.
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
class PoundFormat(DocumentFormat):
    """
    Pound format - uses # markers to determine heading structure.
    
    Structure:
    - If # is followed by a sentence (heading text):
        - That sentence becomes Heading 3
        - All paragraphs until next # become Heading 4
    - If # is followed by something else (like a siman/letter):
        - All paragraphs until next # become Heading 3
    """

    @classmethod
    def get_format_name(cls) -> str:
        return "pound"

    @classmethod
    def get_priority(cls) -> int:
        return 15

    @classmethod
    def detect(cls, doc: Document, context: Dict[str, Any]) -> bool:
        return (
            context.get("mode") == "pound"
            or context.get("format") == "pound"
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

        # First, detect all heading 3 patterns (like special-heading patterns)
        self._apply_special_headings(doc)
        
        # Then, apply pound logic to adjust headings
        self._apply_pound_headings(doc)

        return doc
    
    def _apply_special_headings(self, doc: Document) -> None:
        """Apply special-heading patterns to detect H3 candidates."""
        print(f"Pound format - detecting H3 patterns: {len(doc.paragraphs)} paragraphs")

        # Pattern 1: Hebrew word followed by period
        pattern1 = re.compile(r"^[\u0590-\u05ff]+\.$")

        # Pattern 2: –– heb_word ––
        pattern2 = re.compile(r"^[-–—]+\s*[\u0590-\u05ff]+\s*[-–—]+$")

        # Pattern 3: heb_word – [heb_word_or_letter]
        pattern3 = re.compile(r"^[\u0590-\u05ff]+\s*[-–—]+\s*\[[\u0590-\u05ff]+\]\s*$")

        # Pattern 4: [heb_word_or_letter] – heb_word
        pattern4 = re.compile(r"^\[[\u0590-\u05ff]+\]\s*[-–—]+\s*[\u0590-\u05ff]+$")

        # Pattern 5: Valid gematria (Hebrew letters used for numbering: א-ט, י-צ, ק-ת)
        # Valid gematria letters: א, ב, ג, ד, ה, ו, ז, ח, ט (1-9)
        #                          י, כ, ל, מ, נ, ן, ס, ע, פ, צ (10-90)
        #                          ק, ר, ש, ת (100-400)
        gematria_letters = r"[\u05d0-\u05d8\u05d9\u05db\u05dc\u05de\u05e0\u05df\u05e1\u05e2\u05e4\u05e6\u05e7\u05e8\u05e9\u05ea]"
        pattern5 = re.compile(r"^" + gematria_letters + r"{1,3}$")

        i = 0
        while i < len(doc.paragraphs) - 1:
            para = doc.paragraphs[i]
            text = para.text.strip()

            # Skip empty paragraphs
            if not text:
                i += 1
                continue

            is_match = False
            if pattern1.match(text) or pattern2.match(text) or pattern3.match(text) or pattern4.match(text) or pattern5.match(text):
                is_match = True

            if is_match:
                # The NEXT paragraph becomes H3 (if it's not empty)
                if i + 1 < len(doc.paragraphs):
                    next_para = doc.paragraphs[i + 1]
                    next_text = next_para.text.strip()
                    if next_text and next_para.heading_level == HeadingLevel.NORMAL:
                        next_para.heading_level = HeadingLevel.HEADING_3
                        print(f"Detected H3 pattern at paragraph {i} ('{text}') -> paragraph {i+1} set as H3: '{next_text[:50]}'")

            i += 1

    def _is_sentence(self, text: str) -> bool:
        """Check if text looks like a sentence (heading) rather than just a marker."""
        text = text.strip()
        if not text:
            return False
        
        # If it's just a single Hebrew letter or two letters (like א, ב, יא), it's not a sentence
        if re.match(r'^[\u0590-\u05ff]{1,2}$', text):
            return False
        
        # If it's very short (less than 10 chars) and doesn't have spaces, probably not a sentence
        if len(text) < 10 and ' ' not in text:
            return False
        
        # If it has multiple words (spaces) or is reasonably long, it's likely a sentence
        if ' ' in text or len(text) > 15:
            return True
        
        return False

    def _apply_pound_headings(self, doc: Document) -> None:
        print(f"Pound format: processing {len(doc.paragraphs)} paragraphs")

        i = 0
        while i < len(doc.paragraphs):
            para = doc.paragraphs[i]
            text = para.text.strip()

            # Look for paragraphs containing #
            if '#' in text:
                # Check if this paragraph ends with # or has # followed by text
                # Pattern: text ending with #, or # followed by text
                if text.endswith('#'):
                    # This paragraph ends with #, check the next paragraph
                    if i + 1 < len(doc.paragraphs):
                        next_para = doc.paragraphs[i + 1]
                        next_text = next_para.text.strip()
                        
                        if self._is_sentence(next_text):
                            # Next is a sentence -> it's H3, downgrade all existing H3s to H4 until next #
                            next_para.heading_level = HeadingLevel.HEADING_3
                            print(f"Found # at paragraph {i} ('{text}'), next paragraph is sentence -> H3: '{next_text[:50]}'")
                            
                            # Mark the # paragraph for removal
                            para.text = ""  # Mark for removal
                            
                            # Downgrade all existing H3s to H4 until next #
                            i += 2  # Skip the # para and the H3 para
                            while i < len(doc.paragraphs):
                                current_para = doc.paragraphs[i]
                                current_text = current_para.text.strip()
                                
                                # Stop if we hit another #
                                if '#' in current_text:
                                    break
                                
                                # Skip empty paragraphs
                                if not current_text:
                                    i += 1
                                    continue
                                
                                # If this paragraph is already H3, downgrade it to H4
                                if current_para.heading_level == HeadingLevel.HEADING_3:
                                    current_para.heading_level = HeadingLevel.HEADING_4
                                    print(f"  -> Downgraded paragraph {i} from H3 to H4: '{current_text[:50]}'")
                                
                                i += 1
                            continue
                        else:
                            # Next is not a sentence (like a letter/siman) -> keep existing H3s as H3
                            print(f"Found # at paragraph {i} ('{text}'), next paragraph is not sentence -> keeping H3s as H3")
                            # Mark the # paragraph for removal
                            para.text = ""  # Mark for removal
                            i += 1
                            continue
                else:
                    # # is in the middle or beginning of text
                    # Check if there's text after #
                    parts = text.split('#', 1)
                    if len(parts) > 1 and parts[1].strip():
                        text_after_pound = parts[1].strip()
                        
                        if self._is_sentence(text_after_pound):
                            # Text after # is a sentence -> it's H3, downgrade all existing H3s to H4 until next #
                            heading_para = doc.paragraphs[i]
                            heading_para.text = text_after_pound
                            heading_para.heading_level = HeadingLevel.HEADING_3
                            print(f"Found # in paragraph {i} with sentence -> H3: '{text_after_pound[:50]}'")
                            
                            # Downgrade all existing H3s to H4 until next #
                            i += 1
                            while i < len(doc.paragraphs):
                                current_para = doc.paragraphs[i]
                                current_text = current_para.text.strip()
                                
                                # Stop if we hit another #
                                if '#' in current_text:
                                    break
                                
                                # Skip empty paragraphs
                                if not current_text:
                                    i += 1
                                    continue
                                
                                # If this paragraph is already H3, downgrade it to H4
                                if current_para.heading_level == HeadingLevel.HEADING_3:
                                    current_para.heading_level = HeadingLevel.HEADING_4
                                    print(f"  -> Downgraded paragraph {i} from H3 to H4: '{current_text[:50]}'")
                                
                                i += 1
                            continue
                        else:
                            # Text after # is not a sentence -> keep existing H3s as H3
                            heading_para = doc.paragraphs[i]
                            heading_para.text = text_after_pound
                            heading_para.heading_level = HeadingLevel.HEADING_3
                            print(f"Found # in paragraph {i} without sentence -> H3: '{text_after_pound[:50]}'")
                            # Just continue, existing H3s remain H3
                            i += 1
                            continue
                    else:
                        # Paragraph is just # or starts with # but has no text after
                        # Mark for removal
                        para.text = ""

            i += 1
        
        # Remove all # symbols from remaining paragraphs and remove empty paragraphs
        new_paragraphs = []
        for para in doc.paragraphs:
            # Remove # symbols from text
            if para.text:
                para.text = para.text.replace('#', '')
            # Only keep non-empty paragraphs
            if para.text.strip():
                new_paragraphs.append(para)
        doc.paragraphs = new_paragraphs

