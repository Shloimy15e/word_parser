"""
Reader for Adobe InDesign IDML files.
"""

import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import List

from word_parser.core.document import Document, Paragraph, Alignment
from word_parser.readers.base import InputReader, ReaderRegistry


@ReaderRegistry.register
class IdmlReader(InputReader):
    """
    Reader for Adobe InDesign IDML files.
    
    IDML (InDesign Markup Language) files are ZIP archives containing
    XML files with text content in the Stories folder.
    """
    
    @classmethod
    def get_extensions(cls) -> List[str]:
        return ['.idml']
    
    @classmethod
    def supports_file(cls, file_path: Path) -> bool:
        if file_path.suffix.lower() != '.idml':
            return False
        # Verify it's a valid ZIP/IDML file
        try:
            with zipfile.ZipFile(file_path, 'r') as zf:
                # Check for typical IDML structure
                return any(name.startswith('Stories/') for name in zf.namelist())
        except (zipfile.BadZipFile, IOError):
            return False
    
    @classmethod
    def get_priority(cls) -> int:
        return 80
    
    def read(self, file_path: Path) -> Document:
        """Read an IDML file and return a Document object."""
        texts = self._extract_text_from_idml(file_path)
        
        if not texts:
            raise ValueError(f"No text content found in IDML file: {file_path}")
        
        doc = Document()
        doc.metadata.source_file = str(file_path)
        
        # Add extracted text as paragraphs
        for text in texts:
            if text.strip():
                para = doc.add_paragraph(text)
                para.format.alignment = Alignment.RIGHT
                para.format.right_to_left = True
        
        return doc
    
    def _extract_text_from_idml(self, idml_path: Path) -> List[str]:
        """
        Extract text content from an IDML file.
        IDML is a ZIP archive containing XML files.
        Returns a list of text content strings.
        """
        texts = []
        
        try:
            with zipfile.ZipFile(idml_path, 'r') as zip_file:
                # IDML files contain Stories folder with XML files containing text
                story_files = [
                    name for name in zip_file.namelist() 
                    if name.startswith('Stories/') and name.endswith('.xml')
                ]
                
                for story_file in story_files:
                    with zip_file.open(story_file) as f:
                        tree = ET.parse(f)
                        root = tree.getroot()
                        
                        # Extract all text content from the XML
                        for elem in root.iter():
                            if elem.text and elem.text.strip():
                                text = elem.text.strip()
                                # Filter out standalone "0" or other noise characters
                                if text != "0" and len(text) > 0:
                                    texts.append(text)
                            if elem.tail and elem.tail.strip():
                                tail = elem.tail.strip()
                                if tail != "0" and len(tail) > 0:
                                    texts.append(tail)
        except Exception as e:
            print(f"Warning: Error extracting text from IDML: {e}")
        
        return texts
