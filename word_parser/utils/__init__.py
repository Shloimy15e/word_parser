"""
Utility functions for file handling.
"""

from pathlib import Path
from typing import List, Optional

from word_parser.readers.base import ReaderRegistry


def get_processable_files(directory: Path, all_types: bool = False) -> List[Path]:
    """
    Get files to process from a directory.
    Priority order: .docx > .doc > .rtf > .idml > DOS-encoded (no extension)
    
    Args:
        directory: Directory to search for files
        all_types: If True, returns all supported file types. If False (default),
                   returns only files of ONE type (the highest priority type found).
    
    Returns:
        List of file paths to process
    """
    from word_parser.readers.dos_reader import DosReader
    
    files_by_type = {
        'docx': list(directory.glob("*.docx")),
        'doc': list(directory.glob("*.doc")),
        'rtf': list(directory.glob("*.rtf")),
        'idml': list(directory.glob("*.idml")),
        'dos': []
    }
    
    # Find DOS-encoded files (no extension)
    for file in directory.iterdir():
        if file.is_file() and not file.suffix and DosReader.supports_file(file):
            files_by_type['dos'].append(file)
    
    # If all_types is True, return all files from all types
    if all_types:
        all_files = []
        for file_type in ['docx', 'doc', 'rtf', 'idml', 'dos']:
            all_files.extend(files_by_type[file_type])
        return all_files
    
    # Return files in priority order (only one type)
    for file_type in ['docx', 'doc', 'rtf', 'idml', 'dos']:
        if files_by_type[file_type]:
            return files_by_type[file_type]
    
    return []


def get_reader_for_file(file_path: Path):
    """
    Get the appropriate reader for a file.
    Returns a reader instance or None if no reader supports the file.
    """
    return ReaderRegistry.get_reader_for_file(file_path)


def can_process_file(file_path: Path) -> bool:
    """Check if we have a reader that can process this file."""
    reader = ReaderRegistry.get_reader_for_file(file_path)
    return reader is not None


def get_file_stem(file_path: Path) -> str:
    """
    Get the stem of a filename.
    For files with extension: returns stem
    For DOS files (no extension): returns full name
    """
    if file_path.suffix:
        return file_path.stem
    return file_path.name
