"""
Built-in document format handlers.

This module provides format handlers for common Hebrew document structures.
All formats are now organized in the word_parser.core.formats package.

For backward compatibility, this module imports all formats from the new location.
"""

# Import all formats from the formats package to ensure they're registered
# This maintains backward compatibility for any code that imports from format_handlers
from word_parser.core.formats import (
    StandardFormat,
    DafFormat,
    MultiParshahFormat,
    LetterFormat,
    SimanFormat,
    SpecialHeadingFormat,
    PoundFormat,
    PerekH2Format,
    PerekH3Format,
    FormattedFormat,
    H2OnlyFormat,
    FolderFilenameFormat,
    FolderTitleFormat,
)

# Re-export for backward compatibility
__all__ = [
    "StandardFormat",
    "DafFormat",
    "MultiParshahFormat",
    "LetterFormat",
    "SimanFormat",
    "SpecialHeadingFormat",
    "PoundFormat",
    "PerekH2Format",
    "PerekH3Format",
    "FormattedFormat",
    "H2OnlyFormat",
    "FolderFilenameFormat",
    "FolderTitleFormat",
]
