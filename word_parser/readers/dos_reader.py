"""
Reader for DOS-encoded Hebrew text files (CP862).
"""

from pathlib import Path
from typing import List

from word_parser.core.document import Document, Alignment
from word_parser.core.processing import clean_dos_text, sanitize_xml_text
from word_parser.readers.base import InputReader, ReaderRegistry


@ReaderRegistry.register
class DosReader(InputReader):
    """
    Reader for DOS-encoded Hebrew text files.
    
    These files use CP862 (Hebrew DOS) encoding and typically have
    no file extension. The reader detects them by checking for:
    - No file extension
    - Content that decodes successfully as CP862
    - Presence of Hebrew characters after decoding
    """
    
    @classmethod
    def get_extensions(cls) -> List[str]:
        # DOS files typically have no extension
        return []
    
    @classmethod
    def supports_file(cls, file_path: Path) -> bool:
        """Check if file is a DOS-encoded Hebrew text file."""
        # Must be a file (not directory)
        if not file_path.is_file():
            return False
        
        # DOS files typically have no extension
        if file_path.suffix:
            return False
        
        return cls._is_dos_encoded(file_path)
    
    @classmethod
    def _is_dos_encoded(cls, file_path: Path) -> bool:
        """
        Check if a file is a DOS-encoded Hebrew text file (CP862).
        """
        try:
            with open(file_path, "rb") as f:
                raw_data = f.read(2048)  # Read first 2KB for detection
            
            # File must have some content
            if len(raw_data) == 0:
                return False
            
            # Try to decode as CP862 (Hebrew DOS)
            try:
                text = raw_data.decode('cp862', errors='strict')
                # Check if it contains Hebrew characters
                hebrew_chars = sum(1 for c in text if '\u0590' <= c <= '\u05FF')
                total_chars = len([c for c in text if c.isprintable() and not c.isspace()])
                
                # If more than 5% Hebrew characters, likely a DOS Hebrew file
                if total_chars > 0 and hebrew_chars > total_chars * 0.05:
                    return True
            except (UnicodeDecodeError, UnicodeError):
                # If strict decoding fails, try with errors='ignore'
                try:
                    text = raw_data.decode('cp862', errors='ignore')
                    hebrew_chars = sum(1 for c in text if '\u0590' <= c <= '\u05FF')
                    # More lenient check with ignore errors
                    if hebrew_chars > 10:  # At least 10 Hebrew characters
                        return True
                except:
                    pass
            
            return False
        except Exception:
            return False
    
    @classmethod
    def get_priority(cls) -> int:
        # Lower priority since detection is based on content, not extension
        return 50
    
    def read(self, file_path: Path) -> Document:
        """Read a DOS-encoded Hebrew text file."""
        # Read the raw file
        with open(file_path, "rb") as f:
            raw_data = f.read()
        
        # Decode from CP862 (Hebrew DOS encoding)
        text = raw_data.decode('cp862', errors='ignore')
        
        # Clean DOS formatting codes and garbage
        text = clean_dos_text(text)
        
        # Sanitize text to remove invalid XML characters
        text = sanitize_xml_text(text)
        
        # Create document
        doc = Document()
        doc.metadata.source_file = str(file_path)
        
        # Split into paragraphs and add to document
        paragraphs = text.split('\n')
        for para_text in paragraphs:
            para_text = para_text.strip()
            if para_text:
                para = doc.add_paragraph(para_text)
                para.format.alignment = Alignment.RIGHT
                para.format.right_to_left = True
        
        return doc
