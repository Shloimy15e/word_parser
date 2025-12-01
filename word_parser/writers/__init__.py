"""
Output writers package for document generation.

This package provides a plugin-based architecture for writing documents to different formats.
To add a new output format:

1. Create a new writer class inheriting from OutputWriter
2. Implement the required methods (write, get_format_name, get_extension)
3. Register it with @WriterRegistry.register decorator or WriterRegistry.register_writer()

Example:
    from word_parser.writers.base import OutputWriter, WriterRegistry
    
    @WriterRegistry.register
    class MyFormatWriter(OutputWriter):
        @classmethod
        def get_format_name(cls) -> str:
            return 'myformat'
        
        @classmethod
        def get_extension(cls) -> str:
            return '.myformat'
        
        def write(self, doc: Document, output_path: Path, **options) -> None:
            # Write document to file
            ...
"""

from word_parser.writers.base import OutputWriter, WriterRegistry
from word_parser.writers.docx_writer import DocxWriter
from word_parser.writers.json_writer import JsonWriter
from word_parser.writers.seif_footnotes_writer import SeifFootnotesWriter

__all__ = [
    "OutputWriter",
    "WriterRegistry",
    "DocxWriter",
    "JsonWriter",
    "SeifFootnotesWriter",
]
