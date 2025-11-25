"""
Word Parser - A modular document processing library for Hebrew text documents.

This package provides a plugin-based architecture for reading various document formats,
processing them according to different document schemas, and writing to different output formats.

## Adding new capabilities:

1. To add a new input format (file type):
   - Create a new reader class inheriting from InputReader
   - Implement the required methods
   - Register it with the reader registry

2. To add a new output format:
   - Create a new writer class inheriting from OutputWriter
   - Implement the required methods
   - Register it with the writer registry

3. To add a new document format (schema/structure):
   - Create a new format class inheriting from DocumentFormat
   - Implement detect() and process() methods
   - Register it with the format registry

Example:
    from word_parser import ReaderRegistry, WriterRegistry, FormatRegistry
    
    # Get all available formats
    print(ReaderRegistry.get_supported_extensions())
    print(WriterRegistry.get_supported_formats())
    print(FormatRegistry.list_formats())
    
    # Process a document
    reader = ReaderRegistry.get_reader_for_file(input_path)
    doc = reader.read(input_path)
    
    # Auto-detect and apply document format
    format_handler = FormatRegistry.detect_format(doc, context)
    doc = format_handler.process(doc, context)
    
    # Write output
    writer = WriterRegistry.get_writer('json')
    writer.write(doc, output_path)
"""

from word_parser.core.document import Document, Paragraph, HeadingLevel
from word_parser.core.formats import DocumentFormat, FormatRegistry
from word_parser.readers import ReaderRegistry
from word_parser.writers import WriterRegistry

__version__ = "2.1.0"
__all__ = [
    "Document",
    "Paragraph", 
    "HeadingLevel",
    "DocumentFormat",
    "FormatRegistry",
    "ReaderRegistry",
    "WriterRegistry",
]
