"""
Core processing utilities for document parsing.
"""

from word_parser.core.document import Document, Paragraph, HeadingLevel, RunStyle
from word_parser.core.processing import (
    is_old_header,
    should_start_content,
    extract_heading4_info,
    extract_daf_headings,
    extract_year,
    extract_year_from_text,
    number_to_hebrew_gematria,
    detect_parshah_boundary,
    is_valid_gematria_number,
    sanitize_xml_text,
)
from word_parser.core.formats import DocumentFormat, FormatRegistry

# Import format handlers to register them
from word_parser.core import format_handlers

__all__ = [
    # Document model
    "Document",
    "Paragraph",
    "HeadingLevel",
    "RunStyle",
    # Processing functions
    "is_old_header",
    "should_start_content",
    "extract_heading4_info",
    "extract_daf_headings",
    "extract_year",
    "extract_year_from_text",
    "number_to_hebrew_gematria",
    "detect_parshah_boundary",
    "is_valid_gematria_number",
    "sanitize_xml_text",
    # Format system
    "DocumentFormat",
    "FormatRegistry",
]
