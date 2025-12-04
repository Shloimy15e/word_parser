"""
Writer for JSON output format.
"""

import json
import re
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, List

from word_parser.core.document import Document, HeadingLevel
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
        return "json"

    @classmethod
    def get_extension(cls) -> str:
        return ".json"

    @classmethod
    def get_default_options(cls) -> Dict[str, Any]:
        return {
            "filter_headers": True,
            "indent": 2,
            "ensure_ascii": False,
            "chunking_strategy": "paragraph",
        }

    def write(self, doc: Document, output_path: Path, **options) -> None:
        """
        Write document to a JSON file.

        Options:
            filter_headers: Skip old header paragraphs
            indent: JSON indentation level
            ensure_ascii: Whether to escape non-ASCII characters
            chunking_strategy: 'paragraph', 'h4', 'h3', or 'chunk' (chunks within each H3 by asterisk markers)
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
                ensure_ascii=opts.get("ensure_ascii", False),
                indent=opts.get("indent", 2),
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
            "chunks": self._build_chunks(doc, opts),
        }

        return json_data

    def _build_chunks(
        self, doc: Document, opts: Dict[str, Any]
    ) -> List[Dict]:
        """Build the chunks array from document paragraphs."""
        strategy = opts.get("chunking_strategy", "paragraph")

        if strategy == "h3":
            chunks = self._build_chunks_h3(doc, opts)
        elif strategy == "h4":
            chunks = self._build_chunks_h4(doc, opts)
        elif strategy == "chunk":
            chunks = self._build_chunks_asterisk(doc, opts)
        else:
            chunks = self._build_chunks_paragraph(doc, opts)
        
        # Add footnotes as chunks
        if doc.footnotes:
            footnote_chunks = self._build_footnote_chunks(doc, chunks)
            chunks.extend(footnote_chunks)
        
        return chunks
    
    def _build_footnote_chunks(self, doc: Document, existing_chunks: List[Dict]) -> List[Dict]:
        """Build chunks from footnotes."""
        footnote_chunks = []
        chunk_id = len(existing_chunks) + 1
        
        for footnote in doc.footnotes:
            footnote_text = footnote.text.strip()
            if footnote_text and not self._is_single_word_or_letter(footnote_text):
                chunk_title = f"הערה {footnote.id}"
                chunk = {
                    "chunk_id": chunk_id,
                    "chunk_metadata": {
                        "chunk_title": chunk_title,
                        "footnote_id": footnote.id,
                        "is_footnote": True,
                    },
                    "text": footnote_text,
                }
                footnote_chunks.append(chunk)
                chunk_id += 1
        
        return footnote_chunks

    def _is_single_word_or_letter(self, text: str) -> bool:
        """Check if text is a single word or single letter (should be skipped)."""
        if not text:
            return True
        
        # Remove whitespace and check length
        cleaned = text.strip()
        if not cleaned:
            return True
        
        # Check if it's a single letter (Hebrew or English)
        if len(cleaned) == 1:
            return True
        
        # Check if it's a single word (no spaces, no punctuation that creates multiple tokens)
        # Split by whitespace and common punctuation
        words = re.split(r'[\s\.,;:!?\-–—]+', cleaned)
        # Filter out empty strings
        words = [w for w in words if w]
        
        # If only one word remains after splitting, it's a single word
        if len(words) == 1:
            return True
        
        return False

    def _build_chunks_paragraph(
        self, doc: Document, opts: Dict[str, Any]
    ) -> List[Dict]:
        """Original paragraph-based chunking."""
        chunks = []
        filter_headers = opts.get("filter_headers", True)
        is_multi_parshah = doc.metadata.extra.get("is_multi_parshah", False)

        in_header_section = filter_headers
        chunk_id = 1
        
        # Track current H3 and H4 from paragraph-level headings
        current_h3 = doc.heading3 or ""  # Start with document-level H3 if available
        current_h4 = doc.heading4 or ""   # Start with document-level H4 if available
        
        # Track index for each unique heading3+heading4 combination (chunk title)
        chunk_title_index_map = {}  # Maps (h3, h4) tuple to current index

        for para in doc.paragraphs:
            txt = para.text.strip()

            # Handle H3 paragraphs - update current H3, reset H4
            if para.heading_level == HeadingLevel.HEADING_3:
                current_h3 = txt
                current_h4 = ""
                continue

            # Handle H4 paragraphs - update current H4
            if para.heading_level == HeadingLevel.HEADING_4:
                current_h4 = txt
                continue

            # Skip other heading paragraphs (H1, H2)
            if para.heading_level != HeadingLevel.NORMAL:
                continue

            # Skip parshah boundary lines and marker lines in multi-parshah mode
            if is_multi_parshah:
                if para.metadata.get("is_parshah_boundary"):
                    continue
                if para.metadata.get("is_parshah_marker"):
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
            if txt in ("h", "*", "***", "* * *", "q", "Y"):
                continue

            # Only add non-empty paragraphs as chunks, skip single word/letter chunks
            if txt and not self._is_single_word_or_letter(txt):
                # Build chunk title based on multi-parshah or regular mode
                if is_multi_parshah:
                    current_parshah = para.metadata.get("current_parshah", "")
                    section_index = para.metadata.get("section_index", chunk_id)
                    if current_parshah:
                        chunk_title = f"{current_parshah} {section_index}"
                    else:
                        chunk_title = str(chunk_id)
                else:
                    # Build chunk title with current H3 and H4 (from paragraph-level headings)
                    # Track index for the entire chunk title (heading3+heading4 combination)
                    heading_key = (current_h3, current_h4)
                    if heading_key not in chunk_title_index_map:
                        chunk_title_index_map[heading_key] = 0
                    chunk_title_index_map[heading_key] += 1
                    title_index = chunk_title_index_map[heading_key]
                    
                    title_parts = []
                    if current_h3:
                        title_parts.append(current_h3)
                    if current_h4:
                        title_parts.append(current_h4)
                    
                    if title_parts:
                        chunk_title = " - ".join(title_parts) + f" {title_index}"
                    else:
                        chunk_title = str(chunk_id)

                chunk = {
                    "chunk_id": chunk_id,
                    "chunk_metadata": {"chunk_title": chunk_title},
                    "text": txt,
                }
                chunks.append(chunk)
                chunk_id += 1

        return chunks

    def _build_chunks_h4(self, doc: Document, opts: Dict[str, Any]) -> List[Dict]:
        """Chunk by Heading 4 (or H3 if H4 is missing)."""
        chunks = []
        filter_headers = opts.get("filter_headers", True)
        is_multi_parshah = doc.metadata.extra.get("is_multi_parshah", False)

        current_chunk_text = []
        current_h3 = doc.heading3 or ""
        current_h4 = None
        h4_index = 1  # Index of H4 within current H3

        chunk_id = 1

        def flush_current_chunk():
            """Flush the current chunk if it has content."""
            nonlocal chunk_id, current_chunk_text, current_h3, current_h4, h4_index
            if current_chunk_text:
                chunk_text = "\n".join(current_chunk_text)
                # Skip single word/letter chunks
                if not self._is_single_word_or_letter(chunk_text):
                    # Build chunk title: "H3 - H4" or just "H3" or "H4" (no index needed)
                    if current_h3:
                        if current_h4:
                            chunk_title = f"{current_h3} - {current_h4}"
                        else:
                            chunk_title = current_h3
                    elif current_h4:
                        chunk_title = current_h4
                    else:
                        chunk_title = str(chunk_id)
                    
                    chunks.append(
                        {
                            "chunk_id": chunk_id,
                            "chunk_metadata": {"chunk_title": chunk_title},
                            "text": chunk_text,
                        }
                    )
                    chunk_id += 1
                current_chunk_text = []

        for para in doc.paragraphs:
            txt = para.text.strip()

            # Handle H3 paragraphs - flush chunk, update H3, reset H4
            if para.heading_level == HeadingLevel.HEADING_3:
                flush_current_chunk()
                current_h3 = txt
                current_h4 = None
                h4_index = 1
                continue

            # Handle H4 paragraphs - flush chunk, update H4, increment index for next H4
            if para.heading_level == HeadingLevel.HEADING_4:
                flush_current_chunk()
                # If we already had an H4, increment index; otherwise this is the first H4 (index 1)
                if current_h4 is not None:
                    h4_index += 1
                current_h4 = txt
                continue

            # Skip other heading paragraphs (H1, H2)
            if para.heading_level != HeadingLevel.NORMAL:
                continue

            # Check for H4 change (or H3 change which implies H4 reset)
            # In multi-parshah, H3 changes at boundaries
            if is_multi_parshah and para.metadata.get("is_parshah_boundary"):
                flush_current_chunk()
                current_h3 = para.metadata.get("parshah_name", "")
                current_h4 = None
                h4_index = 1
                continue

            # For now, just accumulate text, filtering markers
            if is_multi_parshah and para.metadata.get("is_parshah_marker"):
                continue

            if filter_headers and is_old_header(txt):
                continue

            if txt and txt not in ("h", "*", "***", "* * *", "q", "Y"):
                current_chunk_text.append(txt)

        # Flush final chunk
        flush_current_chunk()

        return chunks

    def _build_chunks_h3(self, doc: Document, opts: Dict[str, Any]) -> List[Dict]:
        """Chunk by Heading 3."""
        chunks = []
        filter_headers = opts.get("filter_headers", True)
        is_multi_parshah = doc.metadata.extra.get("is_multi_parshah", False)

        current_chunk_text = []
        current_title = doc.heading3 or ""
        current_h2 = doc.heading2 or ""  # Track H2 (perek) for chunk title

        chunk_id = 1

        def flush_current_chunk():
            """Flush the current chunk if it has content."""
            nonlocal chunk_id, current_chunk_text, current_title, current_h2
            if current_chunk_text:
                chunk_text = "\n".join(current_chunk_text)
                # Skip single word/letter chunks
                if not self._is_single_word_or_letter(chunk_text):
                    # Build chunk title: H2 (perek) + H3 if both exist
                    if current_h2 and current_title:
                        chunk_title = f"{current_h2} - {current_title}"
                    elif current_title:
                        chunk_title = current_title
                    elif current_h2:
                        chunk_title = current_h2
                    else:
                        chunk_title = ""
                    
                    chunks.append(
                        {
                            "chunk_id": chunk_id,
                            "chunk_metadata": {"chunk_title": chunk_title},
                            "text": chunk_text,
                        }
                    )
                    chunk_id += 1
                current_chunk_text = []

        for para in doc.paragraphs:
            txt = para.text.strip()

            # Handle H2 paragraphs (perek) - update current_h2
            if para.heading_level == HeadingLevel.HEADING_2:
                current_h2 = txt
                continue

            # Handle H3 paragraphs - flush previous chunk and start new one
            if para.heading_level == HeadingLevel.HEADING_3:
                flush_current_chunk()
                current_title = txt
                continue

            # Skip other heading paragraphs (for combined documents)
            if para.heading_level != HeadingLevel.NORMAL:
                continue

            # In multi-parshah, H3 changes at boundaries
            if is_multi_parshah and para.metadata.get("is_parshah_boundary"):
                flush_current_chunk()
                current_title = para.metadata.get("parshah_name", "")
                continue

            if is_multi_parshah and para.metadata.get("is_parshah_marker"):
                continue

            if filter_headers and is_old_header(txt):
                continue

            if txt and txt not in ("h", "*", "***", "* * *", "q", "Y"):
                current_chunk_text.append(txt)

        # Flush final chunk
        flush_current_chunk()

        return chunks

    def _build_chunks_asterisk(self, doc: Document, opts: Dict[str, Any]) -> List[Dict]:
        """
        Chunk by asterisk markers within each H3 section.
        
        Within each H3 section, paragraphs are grouped into chunks.
        Asterisk markers (*, ***, * * *, etc.) indicate chunk boundaries.
        The asterisk markers themselves are skipped (not included in chunks).
        """
        chunks = []
        filter_headers = opts.get("filter_headers", True)
        is_multi_parshah = doc.metadata.extra.get("is_multi_parshah", False)

        current_chunk_text = []
        current_h3_title = doc.heading3 or ""
        chunk_id = 1
        chunk_index_within_h3 = 1  # Index of chunk within current H3

        in_header_section = filter_headers

        def is_asterisk_marker(text: str) -> bool:
            """Check if text is an asterisk marker or single letter/word."""
            txt = text.strip()
            # Check for common asterisk patterns
            if txt in ("*", "***", "* * *", "**", "* *", "* * * *", "q", "h", "Y"):
                return True
            if txt.startswith("*") and len(txt) <= 10:
                return True
            # Check for single letter/word (Hebrew or short word)
            # Single Hebrew letter or very short word (1-3 characters)
            if len(txt) <= 3 and txt:
                return True
            return False

        def flush_current_chunk():
            """Flush the current chunk if it has content."""
            nonlocal chunk_id, chunk_index_within_h3
            if current_chunk_text:
                chunk_text = "\n".join(current_chunk_text)
                # Skip single word/letter chunks
                if not self._is_single_word_or_letter(chunk_text):
                    # Build chunk title: just H3 name (no index needed)
                    if current_h3_title:
                        chunk_title = current_h3_title
                    else:
                        chunk_title = str(chunk_id)
                    
                    chunks.append(
                        {
                            "chunk_id": chunk_id,
                            "chunk_metadata": {"chunk_title": chunk_title},
                            "text": chunk_text,
                        }
                    )
                    chunk_id += 1
                    chunk_index_within_h3 += 1
                current_chunk_text.clear()

        for para in doc.paragraphs:
            txt = para.text.strip()

            # Skip heading paragraphs (for combined documents)
            if para.heading_level != HeadingLevel.NORMAL:
                continue

            # Handle multi-parshah mode: H3 changes at boundaries
            if is_multi_parshah:
                if para.metadata.get("is_parshah_boundary"):
                    # Flush current chunk before changing H3
                    flush_current_chunk()
                    # Update H3 title and reset chunk index
                    current_h3_title = para.metadata.get("parshah_name", "")
                    chunk_index_within_h3 = 1
                    continue
                
                if para.metadata.get("is_parshah_marker"):
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

            # Check if this is an asterisk marker (chunk boundary)
            if is_asterisk_marker(txt):
                # Flush current chunk and start a new one
                flush_current_chunk()
                continue

            # Skip other marker patterns
            if txt in ("h", "ה", "q"):
                continue

            # Add paragraph to current chunk
            if txt:
                current_chunk_text.append(txt)

        # Flush final chunk
        flush_current_chunk()

        return chunks
