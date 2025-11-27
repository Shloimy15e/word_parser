"""
Shared imports and utilities for format handlers.
"""

import re
from typing import Dict, Any, List

from word_parser.core.document import Document, HeadingLevel
# Import from formats.py module directly to avoid circular imports
# We need to import the module file, not the package
# Use the same module name as __init__.py to ensure we get the same FormatRegistry instance
import sys
from pathlib import Path
import importlib.util

# Get the path to formats.py (sibling to formats/ directory)
_formats_file = Path(__file__).parent.parent / "formats.py"
if _formats_file.exists():
    # Use the same module name as __init__.py to ensure singleton behavior
    _mod_name = "word_parser.core._formats_base"
    if _mod_name in sys.modules:
        # Already loaded, reuse it
        _formats_mod = sys.modules[_mod_name]
    else:
        # Load formats.py as a module
        spec = importlib.util.spec_from_file_location(_mod_name, _formats_file)
        _formats_mod = importlib.util.module_from_spec(spec)
        sys.modules[_mod_name] = _formats_mod
        spec.loader.exec_module(_formats_mod)
    DocumentFormat = _formats_mod.DocumentFormat
    FormatRegistry = _formats_mod.FormatRegistry
else:
    raise ImportError(f"Could not find formats.py at {_formats_file}")

from word_parser.core.processing import (
    is_old_header,
    should_start_content,
    extract_year,
    extract_heading4_info,
    extract_daf_headings,
    detect_parshah_boundary,
    remove_page_markings,
)

__all__ = [
    "re",
    "Dict",
    "Any",
    "List",
    "Document",
    "HeadingLevel",
    "DocumentFormat",
    "FormatRegistry",
    "is_old_header",
    "should_start_content",
    "extract_year",
    "extract_heading4_info",
    "extract_daf_headings",
    "detect_parshah_boundary",
    "remove_page_markings",
]

