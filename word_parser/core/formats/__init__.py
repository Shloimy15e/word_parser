"""
Format handlers package.

This package contains individual format handler modules.
All formats are automatically registered when this package is imported.
"""

# First, import DocumentFormat and FormatRegistry from formats.py module (not package) to avoid circular imports
# This must be done BEFORE importing any format classes, as they depend on these base classes
import sys
import importlib.util
from pathlib import Path

# Import from formats.py file directly (sibling to formats/ directory)
# Use a consistent module name so _base.py can reuse the same instance
_mod_name = "word_parser.core._formats_base"
_formats_file = Path(__file__).parent.parent / "formats.py"
if _formats_file.exists():
    if _mod_name in sys.modules:
        # Already loaded, reuse it
        _formats_base = sys.modules[_mod_name]
    else:
        # Load formats.py as a module
        spec = importlib.util.spec_from_file_location(_mod_name, _formats_file)
        _formats_base = importlib.util.module_from_spec(spec)
        sys.modules[_mod_name] = _formats_base
        spec.loader.exec_module(_formats_base)
    DocumentFormat = _formats_base.DocumentFormat
    FormatRegistry = _formats_base.FormatRegistry
else:
    raise ImportError(f"Could not find formats.py at {_formats_file}")

# Now import all formats to ensure they're registered
# These imports will use DocumentFormat and FormatRegistry from _base.py, which imports from formats.py
from word_parser.core.formats.standard import StandardFormat
from word_parser.core.formats.daf import DafFormat
from word_parser.core.formats.multi_parshah import MultiParshahFormat
from word_parser.core.formats.letter import LetterFormat
from word_parser.core.formats.siman import SimanFormat
from word_parser.core.formats.special_heading import SpecialHeadingFormat
from word_parser.core.formats.pound import PoundFormat
from word_parser.core.formats.perek_h2 import PerekH2Format
from word_parser.core.formats.perek_h3 import PerekH3Format
from word_parser.core.formats.formatted import FormattedFormat
from word_parser.core.formats.h2_only import H2OnlyFormat
from word_parser.core.formats.folder_filename import FolderFilenameFormat
from word_parser.core.formats.minimal import MinimalFormat
from word_parser.core.formats.haus_bachur import HausBachurFormat


__all__ = [
    "DocumentFormat",
    "FormatRegistry",
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
    "MinimalFormat",
    "HausBachurFormat",
]
