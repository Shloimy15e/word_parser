"""
Writer for JSON output format.
"""

import json
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, List

from word_parser.core.document import Document
from word_parser.core.processing import is_old_header, should_start_content
from word_parser.writers.base import OutputWriter, WriterRegistry


@WriterRegistry.register
class JsonWriter(OutputWriter):
    """
    Writer for JSON output format.
    
    Produces a structured JSON file with book metadata and chunks of text.
    """
    
    @classmethod
    def get_format_name(cls) -> str:
        return 'json'
    
    @classmethod
    def get_extension(cls) -> str:
        return '.json'
    
    @classmethod
    def get_default_options(cls) -> Dict[str, Any]:
        return {
            'filter_headers': True,
            'indent': 2,
            'ensure_ascii': False,
        }
    
    def write(self, doc: Document, output_path: Path, **options) -> None:
        """
        Write document to a JSON file.
        
        Options:
            filter_headers: Skip old header paragraphs
            indent: JSON indentation level
            ensure_ascii: Whether to escape non-ASCII characters
        """
        opts = {**self.get_default_options(), **options}
        
        # Build JSON structure
        json_data = self._build_json_structure(doc, opts)
        
        # Write file
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(
                json_data, 
                f, 
                ensure_ascii=opts.get('ensure_ascii', False), 
                indent=opts.get('indent', 2)
            )
    
    def _build_json_structure(self, doc: Document, opts: Dict[str, Any]) -> Dict:
        """Build the JSON data structure from a Document."""
        current_date = datetime.now().strftime("%Y-%m-%d")
        
        # Determine book name
        book_parts = []
        if doc.heading1:
            book_parts.append(doc.heading1)
        if doc.heading2:
            book_parts.append(doc.heading2)
        book_name = " - ".join(book_parts) if book_parts else ""
        
        # Build base chunk title from H3 (and optionally H4)
        base_chunk_title = doc.heading3 or ""
        if doc.heading4 and doc.heading4 != doc.heading3:
            base_chunk_title = f"{base_chunk_title} - {doc.heading4}"
        
        # Build metadata
        metadata = {
            "date": current_date,
        }
        if doc.heading1:
            metadata["book"] = doc.heading1
        if doc.heading3:
            metadata["section"] = doc.heading3
        if doc.heading4 and doc.heading4 != doc.heading3:
            metadata["subsection"] = doc.heading4
        
        # Add any extra metadata
        if doc.metadata.extra:
            metadata.update(doc.metadata.extra)
        
        json_data = {
            "book_name_he": book_name,
            "book_name_en": "",
            "book_metadata": metadata,
            "chunks": self._build_chunks(doc, base_chunk_title, opts),
        }
        
        return json_data
    
    def _build_chunks(self, doc: Document, base_chunk_title: str, 
                      opts: Dict[str, Any]) -> List[Dict]:
        """Build the chunks array from document paragraphs."""
        chunks = []
        filter_headers = opts.get('filter_headers', True)
        is_multi_parshah = doc.metadata.extra.get('is_multi_parshah', False)
        
        in_header_section = filter_headers
        chunk_id = 1
        
        for para in doc.paragraphs:
            txt = para.text.strip()
            
            # Skip parshah boundary lines and marker lines in multi-parshah mode
            if is_multi_parshah:
                if para.metadata.get('is_parshah_boundary'):
                    continue
                if para.metadata.get('is_parshah_marker'):
                    continue
            
            # Header filtering logic
            if filter_headers and in_header_section:
                if txt and should_start_content(txt):
                    in_header_section = False
                elif txt and is_old_header(txt):
                    continue
                elif not txt:
                    continue
                else:
                    continue
            
            if filter_headers and txt and is_old_header(txt):
                continue
            
            # Skip paragraphs that are only markers (ה, *, ***, * * *)
            if txt in ('h', '*', '***', '* * *'):
                continue
            
            # Only add non-empty paragraphs as chunks
            if txt:
                # Build chunk title based on multi-parshah or regular mode
                if is_multi_parshah:
                    current_parshah = para.metadata.get('current_parshah', '')
                    section_index = para.metadata.get('section_index', chunk_id)
                    if current_parshah:
                        chunk_title = f"{current_parshah} {section_index}"
                    else:
                        chunk_title = str(chunk_id)
                else:
                    chunk_title = f"{base_chunk_title} {chunk_id}" if base_chunk_title else str(chunk_id)
                
                chunk = {
                    "chunk_id": chunk_id,
                    "chunk_metadata": {"chunk_title": chunk_title},
                    "text": txt
                }
                chunks.append(chunk)
                chunk_id += 1
        
        return chunks
