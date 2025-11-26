"""
Built-in document format handlers.

This module provides format handlers for common Hebrew document structures:
- StandardFormat: Basic parshah/sefer structure
- DafFormat: Talmud-style daf/perek structure
- MultiParshahFormat: Single document with multiple sections
- LetterFormat: Correspondence format
"""

import re
from typing import Dict, Any, List

from word_parser.core.document import Document
from word_parser.core.formats import DocumentFormat, FormatRegistry
from word_parser.core.processing import (
    is_old_header,
    should_start_content,
    extract_year,
    extract_heading4_info,
    extract_daf_headings,
    detect_parshah_boundary,
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
        }

    def process(self, doc: Document, context: Dict[str, Any]) -> Document:
        """Process document with standard parshah structure."""
        book = context.get("book", "")
        sefer = context.get("sefer", "")
        parshah = context.get("parshah", "")
        filename = context.get("filename", "")
        skip_prefix = context.get("skip_parshah_prefix", False)

        # Remove page markings and merge split paragraphs
        doc = remove_page_markings(doc)

        # Try to extract year from filename if not provided
        if filename and not context.get("year"):
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


@FormatRegistry.register
class DafFormat(DocumentFormat):
    """
    Talmud/Daf-style document format.

    Structure:
    - H1: Book (parent folder or --book arg)
    - H2: Tractate/Masechet (folder name)
    - H3: Perek (extracted from filename, e.g., "פרק א")
    - H4: Chelek/Section (optional, from filename suffix)

    Filename patterns:
    - PEREK1 → H3: "פרק א"
    - PEREK1A → H3: "פרק א", H4: "חלק א"
    - MEKOROS → H3: "מקורות"
    - HAKDOMO → H3: "הקדמה"
    """

    @classmethod
    def get_format_name(cls) -> str:
        return "daf"

    @classmethod
    def get_priority(cls) -> int:
        return 50

    @classmethod
    def detect(cls, doc: Document, context: Dict[str, Any]) -> bool:
        """Detect daf format from context or filename patterns."""
        # Check if explicitly requested
        if context.get("mode") == "daf":
            return True

        # Check filename for perek pattern
        filename = context.get("filename", "").lower()
        if re.match(r"^perek\d+", filename):
            return True
        if re.match(r"^me?koros", filename):
            return True
        if re.match(r"^hakdomo", filename):
            return True

        return False

    @classmethod
    def get_optional_context(cls) -> Dict[str, Any]:
        return {
            "book": None,  # Can be derived from parent folder
            "folder": None,  # Tractate name
            "filename": None,
            "filter_headers": True,
        }

    def process(self, doc: Document, context: Dict[str, Any]) -> Document:
        """Process document with daf/perek structure."""
        book = context.get("book", "")
        folder = context.get("folder", "")
        filename = context.get("filename", "")

        # Remove page markings and merge split paragraphs
        doc = remove_page_markings(doc)

        # Extract H3 and H4 from filename
        heading3, heading4 = extract_daf_headings(filename)

        # Set headings
        doc.set_headings(h1=book, h2=folder, h3=heading3, h4=heading4)

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


@FormatRegistry.register
class MultiParshahFormat(DocumentFormat):
    """
    Multi-parshah document format.

    A single document containing multiple sections, where each section
    is marked by a parshah boundary (e.g., "פרשת פנחס") or list item.
    Each section becomes a new section with its own headings.

    Structure:
    - H1: Book
    - H2: Sefer
    - H3: Section name (from parshah boundary text)

    Detection: Document contains parshah boundary patterns or list-style
    paragraphs that serve as section markers.
    """

    @classmethod
    def get_format_name(cls) -> str:
        return "multi-parshah"

    @classmethod
    def get_priority(cls) -> int:
        return 60

    @classmethod
    def detect(cls, doc: Document, context: Dict[str, Any]) -> bool:
        """Detect multi-parshah format from document structure."""
        # Check if explicitly requested
        if context.get("mode") == "multi-parshah":
            return True

        # Count parshah boundaries
        boundary_count = 0
        for p in doc.paragraphs:
            is_boundary, _, _ = detect_parshah_boundary(p.text)
            if is_boundary:
                boundary_count += 1

        if boundary_count >= 2:
            return True

        # Count list items - if multiple, likely multi-parshah
        list_count = sum(1 for p in doc.paragraphs if p.is_list_item())
        return list_count >= 3

    @classmethod
    def get_required_context(cls) -> List[str]:
        return ["book", "sefer"]

    @classmethod
    def get_optional_context(cls) -> Dict[str, Any]:
        return {
            "skip_parshah_prefix": False,
            "special_heading": False,
            "font_size_heading": False,
        }

    def process(self, doc: Document, context: Dict[str, Any]) -> Document:
        """Process multi-parshah document into sections."""
        # For multi-parshah, we mark section boundaries in the paragraphs
        book = context.get("book", "")
        sefer = context.get("sefer", "")
        special_heading = context.get("special_heading", False)
        font_size_heading = context.get("font_size_heading", False)

        # Remove page markings and merge split paragraphs
        doc = remove_page_markings(doc)

        doc.set_headings(h1=book, h2=sefer)
        doc.metadata.extra["is_multi_parshah"] = True

        # Detect parshah boundaries and mark paragraphs with their section
        self._mark_parshah_sections(doc, special_heading, font_size_heading)

        return doc

    def _mark_parshah_sections(
        self,
        doc: Document,
        special_heading: bool = False,
        font_size_heading: bool = False,
    ) -> None:
        """Mark each paragraph with its parshah section for chunk title generation."""
        current_parshah = None
        prev_text = None
        prev_para = None
        section_index = 0  # Index within current parshah

        # Markers that appear before parshah headings and should be removed
        PARSHAH_MARKERS = {"*", "ה", "***", "* * *", "", "h"}

        # Patterns for special heading mode
        pattern1 = re.compile(r"^[\u0590-\u05ff]+\.$")
        pattern2 = re.compile(r"^[-–—]+\s*[\u0590-\u05ff]+\s*[-–—]+$")
        # Also allow [א] – תשצח . Also allow optional trailing period before or after the heb word such as [א] - תתיג.
        pattern3 = re.compile(
            r"^\s*\.*\s*[\u0590-\u05ff]+\s*[-–—]+\s*\[[\u0590-\u05ff]+\]\s*$"
        )
        pattern4 = re.compile(
            r"^\s*\[[\u0590-\u05ff]+\]\s*[-–—]+\s*[\u0590-\u05ff]+\s*\.*\s*$"
        )

        i = 0
        while i < len(doc.paragraphs):
            para = doc.paragraphs[i]
            text = para.text.strip()
            current_parshah = text

            # Font size heading mode: detect size 14 standalone sentences
            if font_size_heading:
                # Check if this paragraph is size 14 and standalone
                is_size_14_heading = False
                if text and len(para.runs) > 0:
                    # Check if all runs are size 14
                    all_size_14 = all(
                        run.style.font_size == 14.0
                        for run in para.runs
                        if run.style.font_size is not None
                    )
                    # Check if has at least one run with size 14
                    has_size_14 = any(
                        run.style.font_size == 14.0
                        for run in para.runs
                        if run.style.font_size is not None
                    )

                    if all_size_14 and has_size_14:
                        is_size_14_heading = True

                if is_size_14_heading:
                    print(f"Found size 14 heading: {text}")
                    # Look ahead for heading line two
                    if i < len(doc.paragraphs) - 1:
                        next_para = doc.paragraphs[i + 1]
                        next_text = next_para.text.strip()
                        if next_text:
                            # Check if next paragraph is size 14
                            all_size_14 = all(
                                run.style.font_size == 14.0
                                for run in next_para.runs
                                if run.style.font_size is not None
                            )
                            has_size_14 = any(
                                run.style.font_size == 14.0
                                for run in next_para.runs
                                if run.style.font_size is not None
                            )
                            if all_size_14 and has_size_14:
                                # Next paragraph is size 14, so its part of the heading
                                # Combine both lines into the parshah name
                                combined_heading = text + "\n" + next_text
                                current_parshah = combined_heading
                                section_index = 0
                                
                                # Mark the first paragraph as parshah boundary with combined heading
                                para.metadata["is_parshah_boundary"] = True
                                para.metadata["parshah_name"] = combined_heading
                                para.metadata["current_parshah"] = current_parshah
                                para.metadata["section_index"] = section_index
                                
                                # Mark the second paragraph to be skipped in output
                                next_para.metadata["is_parshah_marker"] = True
                                next_para.metadata["current_parshah"] = current_parshah
                                
                                prev_text = text
                                prev_para = para
                                i += 2  # Skip both paragraphs (we've processed them)
                                continue

                    section_index = 0
                    para.metadata["is_parshah_boundary"] = True
                    para.metadata["parshah_name"] = current_parshah
                    para.metadata["current_parshah"] = current_parshah
                    para.metadata["section_index"] = section_index

                    prev_text = text
                    prev_para = para
                    i += 1
                    continue
                else:
                    # Not a heading, increment section index
                    if not para.metadata.get("is_parshah_marker"):
                        section_index += 1
                    para.metadata["current_parshah"] = current_parshah
                    para.metadata["section_index"] = section_index
                    prev_text = text
                    prev_para = para
                    i += 1
                    continue

            if special_heading:
                # Check for marker
                if (
                    pattern1.match(text)
                    or pattern2.match(text)
                    or pattern3.match(text)
                    or pattern4.match(text)
                ):
                    print(f"Found marker")
                    # Found marker. The NEXT line is the heading.
                    para.metadata["is_parshah_marker"] = True
                    para.metadata["current_parshah"] = current_parshah

                    # Look ahead for heading
                    if i + 1 < len(doc.paragraphs):
                        # Heading Line 1
                        heading_para = doc.paragraphs[i + 1]
                        heading_text = heading_para.text.strip()

                        # Look ahead for Heading Line 2 (subtitle)
                        extra_text = None
                        consumed_extra = False

                        if i + 2 < len(doc.paragraphs):
                            next_para = doc.paragraphs[i + 2]
                            next_text = next_para.text.strip()
                            # If next line is short (not content), append it to title
                            if next_text and not should_start_content(next_text):
                                extra_text = next_text
                                next_para.metadata["is_parshah_marker"] = (
                                    True  # Skip in output
                                )
                                next_para.metadata["current_parshah"] = (
                                    current_parshah  # Temp
                                )
                                consumed_extra = True

                        # Set heading info
                        parshah_name = heading_text
                        if extra_text:
                            parshah_name += " " + extra_text

                        current_parshah = parshah_name
                        section_index = 0

                        heading_para.metadata["is_parshah_boundary"] = True
                        heading_para.metadata["parshah_name"] = parshah_name
                        heading_para.metadata["current_parshah"] = current_parshah
                        heading_para.metadata["section_index"] = section_index

                        # Advance index
                        # We processed: i (marker), i+1 (heading)
                        # And optionally i+2 (extra)
                        i += 2
                        if consumed_extra:
                            i += 1

                        prev_text = heading_text
                        prev_para = heading_para
                        continue
                    else:
                        # Marker at end of file
                        i += 1
                        continue

            # Standard logic (or special mode non-marker lines)
            if not special_heading:
                is_boundary, parshah_name, _ = detect_parshah_boundary(text, prev_text)

                if is_boundary:
                    current_parshah = parshah_name
                    section_index = 0
                    para.metadata["is_parshah_boundary"] = True
                    para.metadata["parshah_name"] = parshah_name

                    if prev_para is not None and prev_text in PARSHAH_MARKERS:
                        prev_para.metadata["is_parshah_marker"] = True
                else:
                    if not para.metadata.get("is_parshah_marker"):
                        section_index += 1
            else:
                # Special mode, non-marker line
                if not para.metadata.get("is_parshah_marker"):
                    section_index += 1

            para.metadata["current_parshah"] = current_parshah
            para.metadata["section_index"] = section_index

            prev_text = text
            prev_para = para
            i += 1


@FormatRegistry.register
class LetterFormat(DocumentFormat):
    """
    Letter/correspondence document format.

    Structure:
    - H1: Collection name
    - H2: Category (e.g., "מכתבים")
    - H3: Recipient or subject
    - H4: Date

    Detection: Document starts with greeting pattern or contains
    letter markers like "ב״ה", date, "כבוד", etc.
    """

    @classmethod
    def get_format_name(cls) -> str:
        return "letter"

    @classmethod
    def get_priority(cls) -> int:
        return 40

    @classmethod
    def detect(cls, doc: Document, context: Dict[str, Any]) -> bool:
        """Detect letter format from document content."""
        if context.get("mode") == "letter":
            return True

        if not doc.paragraphs:
            return False

        # Check first few paragraphs for letter patterns
        first_paras = [p.text.strip() for p in doc.paragraphs[:5] if p.text.strip()]

        letter_patterns = [
            r'^ב["\']ה',  # ב"ה at start
            r"^כבוד",  # כבוד (honorific)
            r"^לכבוד",  # לכבוד
            r"^שלום",  # שלום greeting
            r"^ידידי",  # ידידי (my friend)
            r"הנדון:",  # הנדון: (subject)
        ]

        for para in first_paras:
            for pattern in letter_patterns:
                if re.match(pattern, para):
                    return True

        return False

    @classmethod
    def get_optional_context(cls) -> Dict[str, Any]:
        return {
            "book": None,
            "category": "מכתבים",
            "recipient": None,
            "date": None,
        }

    def process(self, doc: Document, context: Dict[str, Any]) -> Document:
        """Process letter document."""
        book = context.get("book", "")
        category = context.get("category", "מכתבים")
        recipient = context.get("recipient")
        date = context.get("date")

        # Remove page markings and merge split paragraphs
        doc = remove_page_markings(doc)

        # Try to extract recipient and date from document if not provided
        if not recipient or not date:
            extracted = self._extract_letter_info(doc)
            recipient = recipient or extracted.get("recipient")
            date = date or extracted.get("date")

        doc.set_headings(h1=book, h2=category, h3=recipient, h4=date)

        return doc

    def _extract_letter_info(self, doc: Document) -> Dict[str, str]:
        """Try to extract recipient and date from letter content."""
        info = {}

        for para in doc.paragraphs[:10]:
            txt = para.text.strip()

            # Look for recipient pattern
            recipient_match = re.match(r"^(?:לכבוד|כבוד)\s+(.+?)(?:\s+שליט״א)?$", txt)
            if recipient_match and "recipient" not in info:
                info["recipient"] = recipient_match.group(1)

            # Look for date patterns
            date_match = re.search(r"(\d{1,2}[./]\d{1,2}[./]\d{2,4})", txt)
            if date_match and "date" not in info:
                info["date"] = date_match.group(1)

            # Hebrew date pattern
            heb_date_match = re.search(r"([א-ת]+'?\s+[א-ת]+\s+תש[א-ת\"']+)", txt)
            if heb_date_match and "date" not in info:
                info["date"] = heb_date_match.group(1)

        return info


@FormatRegistry.register
class SimanFormat(DocumentFormat):
    """
    Siman/Halacha document format.

    Structure:
    - H1: Book (e.g., "שולחן ערוך")
    - H2: Section (e.g., "אורח חיים")
    - H3: Siman number (e.g., "סימן א")
    - H4: Seif number (optional)

    Used for halachic works organized by siman numbers.
    """

    @classmethod
    def get_format_name(cls) -> str:
        return "siman"

    @classmethod
    def get_priority(cls) -> int:
        return 45

    @classmethod
    def detect(cls, doc: Document, context: Dict[str, Any]) -> bool:
        """Detect siman format from context or content."""
        if context.get("mode") == "siman":
            return True

        # Check filename for siman pattern
        filename = context.get("filename", "").lower()
        if re.match(r"^siman\d+", filename) or re.match(r"^סימן", filename):
            return True

        # Check document for siman markers
        for para in doc.paragraphs[:20]:
            txt = para.text.strip()
            if re.match(r"^סימן\s+[א-ת]+", txt):
                return True

        return False

    @classmethod
    def get_optional_context(cls) -> Dict[str, Any]:
        return {
            "book": None,
            "section": None,
            "siman": None,
            "seif": None,
            "filename": None,
        }

    def process(self, doc: Document, context: Dict[str, Any]) -> Document:
        """Process siman-based document."""
        book = context.get("book", "")
        section = context.get("section", "")
        siman = context.get("siman")
        seif = context.get("seif")
        filename = context.get("filename", "")

        # Remove page markings and merge split paragraphs
        doc = remove_page_markings(doc)

        # Try to extract siman from filename
        if not siman and filename:
            siman_match = re.match(r"^siman(\d+)", filename.lower())
            if siman_match:
                from word_parser.core.processing import number_to_hebrew_gematria

                num = int(siman_match.group(1))
                siman = f"סימן {number_to_hebrew_gematria(num)}"

        doc.set_headings(h1=book, h2=section, h3=siman, h4=seif)

        return doc


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
        from word_parser.core.document import HeadingLevel

        print(f"Special headings: {len(doc.paragraphs)}")

        # Pattern 1: Hebrew word followed by period
        # ^[\u0590-\u05ff]+\.$
        pattern1 = re.compile(r"^[\u0590-\u05ff]+\.$")

        # Pattern 2: –– heb_word ––
        # ^[-–—]+\s*[\u0590-\u05ff]+\s*[-–—]+$
        pattern2 = re.compile(r"^[-–—]+\s*[\u0590-\u05ff]+\s*[-–—]+$")

        # Pattern 3: heb_word – [heb_word_or_letter]
        # ^[\u0590-\u05ff]+\s*[-–—]+\s*\[[\u0590-\u05ff]+\]\s*$ e.g. [א] – תשצח
        pattern3 = re.compile(r"^[\u0590-\u05ff]+\s*[-–—]+\s*\[[\u0590-\u05ff]+\]\s*$")

        # Pattern 4: [heb_word_or_letter] – heb_word
        # ^\[[\u0590-\u05ff]+\]\s*[-–—]+\s*[\u0590-\u05ff]+$ e.g. [א] – תשצח
        pattern4 = re.compile(r"^\[[\u0590-\u05ff]+\]\s*[-–—]+\s*[\u0590-\u05ff]+$")

        i = 0
        while i < len(doc.paragraphs) - 1:
            para = doc.paragraphs[i]
            text = para.text.strip()

            is_match = False
            if pattern1.match(text):
                is_match = True
            elif pattern2.match(text):
                is_match = True
            elif pattern3.match(text):
                is_match = True
            elif pattern4.match(text):
                is_match = True

            if is_match:
                print(f"Found match at paragraph {i}")
                # The NEXT paragraph is the heading
                next_para = doc.paragraphs[i + 1]
                if next_para.text.strip():  # Ensure next para is not empty
                    next_para.heading_level = HeadingLevel.HEADING_3

            i += 1
