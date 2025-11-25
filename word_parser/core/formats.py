"""
Document format/schema detection and processing.

This module provides a plugin-based architecture for handling different document
structures within the same file type. For example:
- Torah commentary format (parshah-based with years)
- Talmud format (daf-based with perek structure)
- Letter format (date, recipient, body)
- Multi-parshah format (single document with multiple sections)

To add a new document format:

1. Create a new format class inheriting from DocumentFormat
2. Implement the required methods (detect, process)
3. Register it with @FormatRegistry.register decorator
"""

from abc import ABC, abstractmethod
from typing import Dict, List, Optional, Type, Any

from word_parser.core.document import Document


class DocumentFormat(ABC):
    """
    Abstract base class for document format handlers.
    
    A format handler is responsible for:
    1. Detecting if a document matches its expected structure
    2. Processing the document (extracting headings, chunking content, etc.)
    
    Subclasses must implement:
    - get_format_name(): Return unique format identifier
    - detect(doc, context): Check if document matches this format
    - process(doc, context): Process document according to format rules
    """
    
    @classmethod
    @abstractmethod
    def get_format_name(cls) -> str:
        """
        Return unique identifier for this format.
        
        Examples: 'parshah', 'daf', 'letter', 'multi-parshah'
        """
        pass
    
    @classmethod
    def get_description(cls) -> str:
        """Return human-readable description of this format."""
        return cls.__doc__ or f"Format: {cls.get_format_name()}"
    
    @classmethod
    def get_priority(cls) -> int:
        """
        Return detection priority (higher = checked first).
        
        When auto-detecting format, formats are checked in priority order.
        First format where detect() returns True is used.
        """
        return 0
    
    @classmethod
    @abstractmethod
    def detect(cls, doc: Document, context: Dict[str, Any]) -> bool:
        """
        Check if the document matches this format.
        
        Args:
            doc: The document to check
            context: Additional context (filename, folder structure, CLI args, etc.)
            
        Returns:
            True if document appears to match this format
        """
        pass
    
    @abstractmethod
    def process(self, doc: Document, context: Dict[str, Any]) -> Document:
        """
        Process the document according to this format's rules.
        
        This typically involves:
        - Setting appropriate headings (H1-H4)
        - Filtering/transforming content
        - Setting metadata
        
        Args:
            doc: The document to process
            context: Additional context (book name, sefer, etc.)
            
        Returns:
            The processed document (may be modified in place)
        """
        pass
    
    @classmethod
    def get_required_context(cls) -> List[str]:
        """
        Return list of required context keys for this format.
        
        Used for validation before processing.
        """
        return []
    
    @classmethod
    def get_optional_context(cls) -> Dict[str, Any]:
        """
        Return dict of optional context keys with default values.
        """
        return {}


class FormatRegistry:
    """
    Registry for managing document format handlers.
    
    Provides methods for:
    - Registering formats
    - Auto-detecting document format
    - Getting format by name
    """
    
    _formats: Dict[str, Type[DocumentFormat]] = {}
    
    @classmethod
    def register(cls, format_class: Type[DocumentFormat]) -> Type[DocumentFormat]:
        """
        Register a format class with the registry.
        
        Can be used as a decorator:
            @FormatRegistry.register
            class MyFormat(DocumentFormat):
                ...
        """
        name = format_class.get_format_name()
        cls._formats[name] = format_class
        return format_class
    
    @classmethod
    def unregister(cls, name: str) -> None:
        """Unregister a format by name."""
        if name in cls._formats:
            del cls._formats[name]
    
    @classmethod
    def get_format(cls, name: str) -> Optional[DocumentFormat]:
        """
        Get a format instance by name.
        
        Returns None if format not found.
        """
        format_cls = cls._formats.get(name)
        if format_cls:
            return format_cls()
        return None
    
    @classmethod
    def detect_format(cls, doc: Document, context: Dict[str, Any]) -> Optional[DocumentFormat]:
        """
        Auto-detect the appropriate format for a document.
        
        Checks formats in priority order, returns first match.
        Returns None if no format matches.
        """
        # Sort formats by priority (highest first)
        sorted_formats = sorted(
            cls._formats.values(),
            key=lambda f: f.get_priority(),
            reverse=True
        )
        
        for format_cls in sorted_formats:
            if format_cls.detect(doc, context):
                return format_cls()
        
        return None
    
    @classmethod
    def get_all_formats(cls) -> Dict[str, Type[DocumentFormat]]:
        """Get dictionary of all registered formats."""
        return cls._formats.copy()
    
    @classmethod
    def list_formats(cls) -> List[Dict]:
        """Get list of format info for display."""
        return [
            {
                "name": fmt.get_format_name(),
                "priority": fmt.get_priority(),
                "description": fmt.get_description(),
                "required_context": fmt.get_required_context(),
            }
            for fmt in cls._formats.values()
        ]
