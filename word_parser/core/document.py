"""
Unified document model for representing parsed documents.

This module provides a format-agnostic representation of documents that can be
created from any input format and written to any output format.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from enum import Enum, auto
from typing import List, Optional, Dict, Any


class HeadingLevel(Enum):
    """Document heading levels."""
    HEADING_1 = 1
    HEADING_2 = 2
    HEADING_3 = 3
    HEADING_4 = 4
    NORMAL = 0


class Alignment(Enum):
    """Paragraph alignment options."""
    LEFT = auto()
    CENTER = auto()
    RIGHT = auto()
    JUSTIFY = auto()


@dataclass
class RunStyle:
    """Formatting options for a text run within a paragraph."""
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[bool] = None
    font_size: Optional[float] = None  # in points
    font_name: Optional[str] = None
    color_rgb: Optional[tuple] = None  # (R, G, B) tuple
    highlight_color: Optional[str] = None
    all_caps: Optional[bool] = None
    small_caps: Optional[bool] = None
    strike: Optional[bool] = None
    superscript: Optional[bool] = None
    subscript: Optional[bool] = None


@dataclass
class TextRun:
    """A run of text with consistent formatting within a paragraph."""
    text: str
    style: RunStyle = field(default_factory=RunStyle)
    footnote_id: Optional[int] = None  # ID of footnote referenced by this run


@dataclass
class ParagraphFormat:
    """Formatting options for a paragraph."""
    alignment: Alignment = Alignment.RIGHT
    right_to_left: bool = True
    left_indent: Optional[float] = None  # in points
    right_indent: Optional[float] = None
    first_line_indent: Optional[float] = None
    space_before: Optional[float] = None
    space_after: Optional[float] = None
    line_spacing: Optional[float] = None
    line_spacing_rule: Optional[str] = None
    keep_together: Optional[bool] = None
    keep_with_next: Optional[bool] = None
    page_break_before: Optional[bool] = None
    widow_control: Optional[bool] = None


@dataclass
class Footnote:
    """A footnote in a document."""
    id: int  # Unique footnote ID
    paragraphs: List[Paragraph] = field(default_factory=list)  # Footnote content as paragraphs
    original_id: Optional[int] = None  # Original footnote ID from source document
    
    @property
    def text(self) -> str:
        """Get the full text of the footnote."""
        return "\n".join(p.text for p in self.paragraphs if not p.is_empty())


@dataclass
class Paragraph:
    """A paragraph in a document."""
    runs: List[TextRun] = field(default_factory=list)
    format: ParagraphFormat = field(default_factory=ParagraphFormat)
    style_name: Optional[str] = None
    heading_level: HeadingLevel = HeadingLevel.NORMAL
    metadata: Dict[str, Any] = field(default_factory=dict)  # For format-specific data
    
    @property
    def text(self) -> str:
        """Get the full text of the paragraph."""
        return "".join(run.text for run in self.runs)
    
    @text.setter
    def text(self, value: str):
        """Set paragraph text (replaces all runs with a single run)."""
        self.runs = [TextRun(text=value)]
    
    def add_run(self, text: str, style: Optional[RunStyle] = None) -> TextRun:
        """Add a text run to the paragraph."""
        run = TextRun(text=text, style=style or RunStyle())
        self.runs.append(run)
        return run
    
    def is_empty(self) -> bool:
        """Check if paragraph has no text content."""
        return not self.text.strip()
    
    def is_list_item(self) -> bool:
        """Check if paragraph is a list item."""
        return self.style_name and 'list' in self.style_name.lower()
    
    def is_numbered_list_item(self) -> bool:
        """
        Check if paragraph is an actual numbered list item (not just List Paragraph style).
        Distinguishes between:
        - "List Paragraph" style (can be headings) - returns False
        - Actual numbered/bulleted list items - returns True
        """
        # Check if it has actual numbering properties in metadata
        if self.metadata.get('is_numbered_list'):
            return True
        
        # If style is "List Paragraph" (exact match), it's NOT a numbered list item
        if self.style_name and self.style_name.lower() == 'list paragraph':
            return False
        
        # Check text content: if it starts with Hebrew gematria number followed by period,
        # it's a numbered list item (e.g., "יט. אמר בכל פסקא...")
        from word_parser.core.processing import is_valid_gematria_number
        
        text = self.text.strip()
        # Pattern: 1-4 Hebrew letters (gematria) followed by period and optional space
        # This catches patterns like "יט. " or "יט.אמר" (with or without space)
        siman_match = re.match(r"^([א-ת]{1,4})\.\s*", text)
        if siman_match:
            siman_text = siman_match.group(1)
            if is_valid_gematria_number(siman_text):
                # Make sure there's actual content after the number (not just "יט.")
                remaining_text = text[len(siman_match.group(0)):].strip()
                if remaining_text:  # There's content after the number
                    print(f"DEBUG: Detected numbered list item by text pattern: '{text[:50]}' (gematria: {siman_text})")
                    return True  # This is a numbered list item, not a heading
        
        # If style contains "list" but is not "List Paragraph", it's likely a numbered list
        #if self.style_name and 'list' in self.style_name.lower():
           # return True
        
        return False


@dataclass
class DocumentMetadata:
    """Metadata for a document."""
    book: Optional[str] = None  # H1
    sefer: Optional[str] = None  # H2
    parshah: Optional[str] = None  # H3
    subsection: Optional[str] = None  # H4
    year: Optional[str] = None
    date: Optional[str] = None
    source_file: Optional[str] = None
    extra: Dict[str, Any] = field(default_factory=dict)


@dataclass
class Document:
    """
    A format-agnostic document representation.
    
    This class serves as the intermediary between input readers and output writers.
    Input readers convert their specific format to this representation, and output
    writers convert from this representation to their target format.
    """
    paragraphs: List[Paragraph] = field(default_factory=list)
    footnotes: List[Footnote] = field(default_factory=list)  # Document footnotes
    metadata: DocumentMetadata = field(default_factory=DocumentMetadata)
    
    # Heading content (separate from body paragraphs for clarity)
    heading1: Optional[str] = None
    heading2: Optional[str] = None
    heading3: Optional[str] = None
    heading4: Optional[str] = None
    
    def add_paragraph(self, text: str = "", 
                      heading_level: HeadingLevel = HeadingLevel.NORMAL,
                      format: Optional[ParagraphFormat] = None) -> Paragraph:
        """Add a paragraph to the document."""
        para = Paragraph(
            heading_level=heading_level,
            format=format or ParagraphFormat()
        )
        if text:
            para.add_run(text)
        self.paragraphs.append(para)
        return para
    
    def get_body_paragraphs(self) -> List[Paragraph]:
        """Get all non-heading paragraphs."""
        return [p for p in self.paragraphs if p.heading_level == HeadingLevel.NORMAL]
    
    def get_headings(self) -> List[Paragraph]:
        """Get all heading paragraphs."""
        return [p for p in self.paragraphs if p.heading_level != HeadingLevel.NORMAL]
    
    def get_text_content(self) -> str:
        """Get all body text as a single string."""
        return "\n\n".join(p.text for p in self.get_body_paragraphs() if not p.is_empty())
    
    def set_headings(self, h1: str = None, h2: str = None, 
                     h3: str = None, h4: str = None):
        """Set document headings."""
        self.heading1 = h1
        self.heading2 = h2
        self.heading3 = h3
        self.heading4 = h4
        
        # Also update metadata
        if h1:
            self.metadata.book = h1
        if h2:
            self.metadata.sefer = h2
        if h3:
            self.metadata.parshah = h3
        if h4:
            self.metadata.subsection = h4
    
    def add_footnote(self, footnote: Footnote) -> None:
        """Add a footnote to the document."""
        self.footnotes.append(footnote)
    
    def get_footnote_by_id(self, footnote_id: int) -> Optional[Footnote]:
        """Get a footnote by its ID."""
        for fn in self.footnotes:
            if fn.id == footnote_id:
                return fn
        return None
