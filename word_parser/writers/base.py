"""
Base classes for output writers.

This module provides the abstract base class for all output writers and the
registry system for managing them.
"""

from abc import ABC, abstractmethod
from pathlib import Path
from typing import Dict, List, Optional, Type, Any

from word_parser.core.document import Document


class OutputWriter(ABC):
    """
    Abstract base class for document output writers.
    
    Subclasses must implement:
    - get_format_name(): Return the format name (e.g., 'json', 'docx')
    - get_extension(): Return the file extension for this format
    - write(doc, output_path, **options): Write the document to a file
    
    Optionally override:
    - get_default_options(): Return default options for this writer
    """
    
    @classmethod
    @abstractmethod
    def get_format_name(cls) -> str:
        """
        Return the format name for this writer.
        
        This is used to identify the writer (e.g., 'json', 'docx', 'html').
        """
        pass
    
    @classmethod
    @abstractmethod
    def get_extension(cls) -> str:
        """
        Return the file extension for this format.
        
        Should include the dot (e.g., '.json', '.docx').
        """
        pass
    
    @abstractmethod
    def write(self, doc: Document, output_path: Path, **options) -> None:
        """
        Write the document to the specified path.
        
        Args:
            doc: Document object to write
            output_path: Path for the output file
            **options: Format-specific options
            
        Raises:
            ValueError: If the document cannot be written
            IOError: If file writing fails
        """
        pass
    
    @classmethod
    def get_default_options(cls) -> Dict[str, Any]:
        """
        Return default options for this writer.
        
        Override to provide format-specific defaults.
        """
        return {}
    
    @classmethod
    def get_name(cls) -> str:
        """Return human-readable name for this writer."""
        return cls.__name__
    
    @classmethod
    def get_description(cls) -> str:
        """Return description of what this writer produces."""
        return cls.__doc__ or f"Writer for {cls.get_format_name()} format"


class WriterRegistry:
    """
    Registry for managing output writers.
    
    Provides methods for:
    - Registering writers
    - Finding the appropriate writer for a format
    - Listing all supported formats
    """
    
    _writers: Dict[str, Type[OutputWriter]] = {}
    
    @classmethod
    def register(cls, writer_class: Type[OutputWriter]) -> Type[OutputWriter]:
        """
        Register a writer class with the registry.
        
        Can be used as a decorator:
            @WriterRegistry.register
            class MyWriter(OutputWriter):
                ...
        
        Or called directly:
            WriterRegistry.register(MyWriter)
        """
        format_name = writer_class.get_format_name()
        cls._writers[format_name] = writer_class
        return writer_class
    
    @classmethod
    def register_writer(cls, writer_class: Type[OutputWriter]) -> None:
        """Register a writer class (alternative to decorator syntax)."""
        cls.register(writer_class)
    
    @classmethod
    def unregister(cls, format_name: str) -> None:
        """Unregister a writer by format name."""
        if format_name in cls._writers:
            del cls._writers[format_name]
    
    @classmethod
    def get_writer(cls, format_name: str) -> Optional[OutputWriter]:
        """
        Get a writer instance for the specified format.
        
        Args:
            format_name: Format name (e.g., 'json', 'docx')
            
        Returns:
            Writer instance or None if format not supported
        """
        writer_cls = cls._writers.get(format_name.lower())
        if writer_cls:
            return writer_cls()
        return None
    
    @classmethod
    def get_writer_for_extension(cls, extension: str) -> Optional[OutputWriter]:
        """
        Get a writer instance for the specified file extension.
        
        Args:
            extension: File extension with or without dot (e.g., '.json' or 'json')
        """
        if not extension.startswith('.'):
            extension = '.' + extension
        extension = extension.lower()
        
        for writer_cls in cls._writers.values():
            if writer_cls.get_extension().lower() == extension:
                return writer_cls()
        
        return None
    
    @classmethod
    def get_supported_formats(cls) -> List[str]:
        """Get list of all supported format names."""
        return sorted(cls._writers.keys())
    
    @classmethod
    def get_supported_extensions(cls) -> List[str]:
        """Get list of all supported file extensions."""
        return sorted(w.get_extension() for w in cls._writers.values())
    
    @classmethod
    def get_all_writers(cls) -> Dict[str, Type[OutputWriter]]:
        """Get dictionary of all registered writers."""
        return cls._writers.copy()
    
    @classmethod
    def list_writers(cls) -> List[Dict]:
        """Get list of writer info for display."""
        return [
            {
                "name": writer_cls.get_name(),
                "format": writer_cls.get_format_name(),
                "extension": writer_cls.get_extension(),
                "description": writer_cls.get_description()
            }
            for writer_cls in cls._writers.values()
        ]
