"""
Input readers package for document parsing.

This package provides a plugin-based architecture for reading different document formats.
To add a new input format:

1. Create a new reader class inheriting from InputReader
2. Implement the required methods (read, supports_file, get_extensions)
3. Register it with @ReaderRegistry.register decorator or ReaderRegistry.register_reader()

Example:
    from word_parser.readers.base import InputReader, ReaderRegistry
    
    @ReaderRegistry.register
    class MyFormatReader(InputReader):
        @classmethod
        def get_extensions(cls) -> list:
            return ['.myformat']
        
        @classmethod
        def supports_file(cls, file_path: Path) -> bool:
            return file_path.suffix.lower() == '.myformat'
        
        def read(self, file_path: Path) -> Document:
            # Parse and return Document
            ...
"""

from word_parser.readers.base import InputReader, ReaderRegistry
from word_parser.readers.docx_reader import DocxReader
from word_parser.readers.doc_reader import DocReader
from word_parser.readers.idml_reader import IdmlReader
from word_parser.readers.dos_reader import DosReader
from word_parser.readers.rtf_reader import RtfReader

__all__ = [
    "InputReader",
    "ReaderRegistry",
    "DocxReader",
    "DocReader",
    "IdmlReader",
    "DosReader",
    "RtfReader",
]
