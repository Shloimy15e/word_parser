"""
Base classes for input readers.

This module provides the abstract base class for all input readers and the
registry system for managing them.
"""

from abc import ABC, abstractmethod
from pathlib import Path
from typing import Dict, List, Optional, Type

from word_parser.core.document import Document


class InputReader(ABC):
    """
    Abstract base class for document input readers.
    
    Subclasses must implement:
    - get_extensions(): Return list of file extensions this reader handles
    - supports_file(file_path): Check if this reader can handle a specific file
    - read(file_path): Parse the file and return a Document
    
    Optionally override:
    - get_priority(): Return priority for reader selection (higher = preferred)
    """
    
    @classmethod
    @abstractmethod
    def get_extensions(cls) -> List[str]:
        """
        Return list of file extensions this reader supports.
        
        Extensions should include the dot (e.g., ['.docx', '.doc'])
        Return empty list for readers that detect files by content rather than extension.
        """
        pass
    
    @classmethod
    @abstractmethod
    def supports_file(cls, file_path: Path) -> bool:
        """
        Check if this reader can handle the given file.
        
        This method may check file extension, magic bytes, or other properties.
        """
        pass
    
    @abstractmethod
    def read(self, file_path: Path) -> Document:
        """
        Read and parse the file, returning a Document object.
        
        Args:
            file_path: Path to the input file
            
        Returns:
            Document object containing the parsed content
            
        Raises:
            ValueError: If the file cannot be parsed
            FileNotFoundError: If the file doesn't exist
        """
        pass
    
    @classmethod
    def get_priority(cls) -> int:
        """
        Return priority for reader selection.
        
        Higher values = higher priority. When multiple readers support a file,
        the one with highest priority is used.
        
        Default is 0. Override to change priority.
        """
        return 0
    
    @classmethod
    def get_name(cls) -> str:
        """Return human-readable name for this reader."""
        return cls.__name__
    
    @classmethod
    def get_description(cls) -> str:
        """Return description of what this reader handles."""
        return cls.__doc__ or f"Reader for {cls.get_extensions()}"


class ReaderRegistry:
    """
    Registry for managing input readers.
    
    Provides methods for:
    - Registering readers
    - Finding the appropriate reader for a file
    - Listing all supported formats
    """
    
    _readers: Dict[str, Type[InputReader]] = {}
    
    @classmethod
    def register(cls, reader_class: Type[InputReader]) -> Type[InputReader]:
        """
        Register a reader class with the registry.
        
        Can be used as a decorator:
            @ReaderRegistry.register
            class MyReader(InputReader):
                ...
        
        Or called directly:
            ReaderRegistry.register(MyReader)
        """
        name = reader_class.get_name()
        cls._readers[name] = reader_class
        return reader_class
    
    @classmethod
    def register_reader(cls, reader_class: Type[InputReader]) -> None:
        """Register a reader class (alternative to decorator syntax)."""
        cls.register(reader_class)
    
    @classmethod
    def unregister(cls, name: str) -> None:
        """Unregister a reader by name."""
        if name in cls._readers:
            del cls._readers[name]
    
    @classmethod
    def get_reader_for_file(cls, file_path: Path) -> Optional[InputReader]:
        """
        Get an appropriate reader instance for the given file.
        
        Returns the reader with highest priority that supports the file,
        or None if no reader supports it.
        """
        file_path = Path(file_path)
        
        # Find all readers that support this file
        supporting_readers = [
            reader_cls for reader_cls in cls._readers.values()
            if reader_cls.supports_file(file_path)
        ]
        
        if not supporting_readers:
            return None
        
        # Sort by priority (highest first) and return an instance
        supporting_readers.sort(key=lambda r: r.get_priority(), reverse=True)
        return supporting_readers[0]()
    
    @classmethod
    def get_reader_by_extension(cls, extension: str) -> Optional[Type[InputReader]]:
        """
        Get reader class for a specific extension.
        
        Args:
            extension: File extension with or without dot (e.g., '.docx' or 'docx')
        """
        if not extension.startswith('.'):
            extension = '.' + extension
        extension = extension.lower()
        
        for reader_cls in cls._readers.values():
            if extension in [ext.lower() for ext in reader_cls.get_extensions()]:
                return reader_cls
        
        return None
    
    @classmethod
    def get_supported_extensions(cls) -> List[str]:
        """Get list of all supported file extensions."""
        extensions = set()
        for reader_cls in cls._readers.values():
            extensions.update(reader_cls.get_extensions())
        return sorted(extensions)
    
    @classmethod
    def get_all_readers(cls) -> Dict[str, Type[InputReader]]:
        """Get dictionary of all registered readers."""
        return cls._readers.copy()
    
    @classmethod
    def list_readers(cls) -> List[Dict]:
        """Get list of reader info for display."""
        return [
            {
                "name": reader_cls.get_name(),
                "extensions": reader_cls.get_extensions(),
                "priority": reader_cls.get_priority(),
                "description": reader_cls.get_description()
            }
            for reader_cls in cls._readers.values()
        ]
