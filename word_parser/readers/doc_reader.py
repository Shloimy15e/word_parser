"""
Reader for Microsoft Word .doc files (legacy format).

Requires pywin32 on Windows for COM automation.
"""

import tempfile
from pathlib import Path
from typing import List

from word_parser.core.document import Document
from word_parser.readers.base import InputReader, ReaderRegistry
from word_parser.readers.docx_reader import DocxReader

# Try to import pywin32
try:
    import win32com.client
    WORD_AVAILABLE = True
except ImportError:
    WORD_AVAILABLE = False


@ReaderRegistry.register
class DocReader(InputReader):
    """
    Reader for Microsoft Word .doc files (legacy binary format).
    
    This reader converts .doc to .docx using Word COM automation,
    then reads the result using DocxReader.
    
    Requires:
    - Windows OS
    - Microsoft Word installed
    - pywin32 package installed
    """
    
    @classmethod
    def get_extensions(cls) -> List[str]:
        return ['.doc']
    
    @classmethod
    def supports_file(cls, file_path: Path) -> bool:
        if not WORD_AVAILABLE:
            return False
        return file_path.suffix.lower() == '.doc'
    
    @classmethod
    def get_priority(cls) -> int:
        return 90  # High priority for .doc files
    
    @classmethod
    def is_available(cls) -> bool:
        """Check if this reader can be used (pywin32 installed)."""
        return WORD_AVAILABLE
    
    def read(self, file_path: Path) -> Document:
        """Read a .doc file by converting to .docx first."""
        if not WORD_AVAILABLE:
            raise RuntimeError(
                "pywin32 not installed. Cannot convert .doc files. "
                "Install with: pip install pywin32"
            )
        
        # Convert to temporary .docx
        temp_docx = self._convert_to_docx(file_path)
        
        try:
            # Use DocxReader to read the converted file
            docx_reader = DocxReader()
            doc = docx_reader.read(temp_docx)
            doc.metadata.source_file = str(file_path)
            return doc
        finally:
            # Clean up temp file
            if temp_docx.exists():
                temp_docx.unlink()
    
    def _convert_to_docx(self, doc_path: Path) -> Path:
        """
        Convert .doc file to .docx using Word COM automation.
        Returns path to temporary .docx file.
        """
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        
        try:
            # Open the .doc file
            doc = word.Documents.Open(str(doc_path.absolute()))
            
            # Create temp file for .docx
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
            temp_file.close()
            
            # Save as .docx (format 16 = wdFormatXMLDocument)
            doc.SaveAs(temp_file.name, FileFormat=16)
            doc.Close()
            
            return Path(temp_file.name)
        finally:
            word.Quit()
