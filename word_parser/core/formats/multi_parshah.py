"""
Multi-parshah document format handler.
"""

import re
from typing import Dict, Any, List

from word_parser.core.formats._base import (
    DocumentFormat,
    FormatRegistry,
    Document,
    should_start_content,
    detect_parshah_boundary,
    remove_page_markings,
)


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
                    print("Found marker")
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

