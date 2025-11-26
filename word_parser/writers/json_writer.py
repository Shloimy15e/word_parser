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

    def _build_chunks(
        self, doc: Document, base_chunk_title: str, opts: Dict[str, Any]
    ) -> List[Dict]:
        """Build the chunks array from document paragraphs."""
        strategy = opts.get("chunking_strategy", "paragraph")

        if strategy == "h3":
            return self._build_chunks_h3(doc, opts)
        elif strategy == "h4":
            return self._build_chunks_h4(doc, opts)
        elif strategy == "chunk":
            return self._build_chunks_asterisk(doc, opts)
        else:
            return self._build_chunks_paragraph(doc, base_chunk_title, opts)

    def _build_chunks_paragraph(
        self, doc: Document, base_chunk_title: str, opts: Dict[str, Any]
    ) -> List[Dict]:
        """Original paragraph-based chunking."""
        chunks = []
        filter_headers = opts.get("filter_headers", True)
        is_multi_parshah = doc.metadata.extra.get("is_multi_parshah", False)

        in_header_section = filter_headers
        chunk_id = 1

        for para in doc.paragraphs:
            txt = para.text.strip()

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
            if txt in ("h", "*", "***", "* * *", "q"):
                continue

            # Only add non-empty paragraphs as chunks
            if txt:
                # Build chunk title based on multi-parshah or regular mode
                if is_multi_parshah:
                    current_parshah = para.metadata.get("current_parshah", "")
                    section_index = para.metadata.get("section_index", chunk_id)
                    if current_parshah:
                        chunk_title = f"{current_parshah} {section_index}"
                    else:
                        chunk_title = str(chunk_id)
                else:
                    chunk_title = (
                        f"{base_chunk_title} {chunk_id}"
                        if base_chunk_title
                        else str(chunk_id)
                    )

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
        current_title = doc.heading3 or ""
        if doc.heading4 and doc.heading4 != doc.heading3:
            current_title = doc.heading4

        chunk_id = 1

        for para in doc.paragraphs:
            txt = para.text.strip()

            # Check for H4 change (or H3 change which implies H4 reset)
            # In multi-parshah, H3 changes at boundaries
            if is_multi_parshah and para.metadata.get("is_parshah_boundary"):
                # Flush current chunk
                if current_chunk_text:
                    chunks.append(
                        {
                            "chunk_id": chunk_id,
                            "chunk_metadata": {"chunk_title": current_title},
                            "text": "\n".join(current_chunk_text),
                        }
                    )
                    chunk_id += 1
                    current_chunk_text = []

                # Update title
                current_title = para.metadata.get("parshah_name", "")
                continue

            # TODO: Detect H4 changes in standard docs if we had H4 detection logic in paragraphs
            # Currently H4 is mostly from filename or specific formats.
            # If the format handler marked paragraphs with heading levels, we could use that.

            # For now, just accumulate text, filtering markers
            if is_multi_parshah and para.metadata.get("is_parshah_marker"):
                continue

            if filter_headers and is_old_header(txt):
                continue

            if txt and txt not in ("h", "*", "***", "* * *", "q"):
                current_chunk_text.append(txt)

        # Flush final chunk
        if current_chunk_text:
            chunks.append(
                {
                    "chunk_id": chunk_id,
                    "chunk_metadata": {"chunk_title": current_title},
                    "text": "\n".join(current_chunk_text),
                }
            )

        return chunks

    def _build_chunks_h3(self, doc: Document, opts: Dict[str, Any]) -> List[Dict]:
        """Chunk by Heading 3."""
        chunks = []
        filter_headers = opts.get("filter_headers", True)
        is_multi_parshah = doc.metadata.extra.get("is_multi_parshah", False)

        current_chunk_text = []
        current_title = doc.heading3 or ""

        chunk_id = 1

        for para in doc.paragraphs:
            txt = para.text.strip()

            # In multi-parshah, H3 changes at boundaries
            if is_multi_parshah and para.metadata.get("is_parshah_boundary"):
                # Flush current chunk
                if current_chunk_text:
                    chunks.append(
                        {
                            "chunk_id": chunk_id,
                            "chunk_metadata": {"chunk_title": current_title},
                            "text": "\n".join(current_chunk_text),
                        }
                    )
                    chunk_id += 1
                    current_chunk_text = []

                # Update title
                current_title = para.metadata.get("parshah_name", "")
                continue

            if is_multi_parshah and para.metadata.get("is_parshah_marker"):
                continue

            if filter_headers and is_old_header(txt):
                continue

            if txt and txt not in ("h", "*", "***", "* * *", "q"):
                current_chunk_text.append(txt)

        # Flush final chunk
        if current_chunk_text:
            chunks.append(
                {
                    "chunk_id": chunk_id,
                    "chunk_metadata": {"chunk_title": current_title},
                    "text": "\n".join(current_chunk_text),
                }
            )

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
            if txt in ("*", "***", "* * *", "**", "* *", "* * * *", "q", "h"):
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
                # Build chunk title: H3 name + chunk number within H3
                if current_h3_title:
                    chunk_title = f"{current_h3_title} {chunk_index_within_h3}"
                else:
                    chunk_title = str(chunk_id)
                
                chunks.append(
                    {
                        "chunk_id": chunk_id,
                        "chunk_metadata": {"chunk_title": chunk_title},
                        "text": "\n".join(current_chunk_text),
                    }
                )
                chunk_id += 1
                chunk_index_within_h3 += 1
                current_chunk_text.clear()

        for para in doc.paragraphs:
            txt = para.text.strip()

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
