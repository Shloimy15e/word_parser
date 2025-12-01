"""
Writer for merging content and footnotes files by matching seif markers.

This writer takes two files:
1. Content file - contains the main text with footnote references like #(א)
2. Footnotes file - contains footnotes with seif markers (Hebrew letters like א, ב, ג)

The writer matches footnotes to content by seif markers and merges them.

Usage example:
    from word_parser.writers import WriterRegistry
    from pathlib import Path
    
    writer = WriterRegistry.get_writer("seif-footnotes")
    writer.write(
        doc=None,  # Ignored, we read from files instead
        output_path=Path("output.docx"),
        content_file="content.docx",
        footnotes_file="footnotes.docx",
        output_format="docx"
    )

Content file format:
    Paragraphs with footnote references like:
    "א. מיום חמשה עשר באב ואילך היתה ניכרת מאוד... #(א)"

Footnotes file format:
    Paragraphs starting with seif markers like:
    "א. Footnote text here..."
    "ב. Another footnote..."
"""

import re
from pathlib import Path
from typing import Dict, Any, List

from word_parser.core.document import Document, Paragraph, TextRun, Footnote
from word_parser.readers import ReaderRegistry
from word_parser.writers.base import OutputWriter, WriterRegistry


@WriterRegistry.register
class SeifFootnotesWriter(OutputWriter):
    """
    Writer that merges content and footnotes files by matching seif markers.
    
    The content file should have footnote references like #(א), #(ב), etc.
    The footnotes file should have footnotes starting with seif markers like "א. ", "ב. ", etc.
    """

    @classmethod
    def get_format_name(cls) -> str:
        return "seif-footnotes"

    @classmethod
    def get_extension(cls) -> str:
        return ".docx"  # Output format is docx (or could be configurable)

    @classmethod
    def get_default_options(cls) -> Dict[str, Any]:
        return {
            "content_file": None,  # Path to content file
            "footnotes_file": None,  # Path to footnotes file
            "output_format": "docx",  # Output format (docx, json, etc.)
        }

    def write(self, doc: Document, output_path: Path, **options) -> None:
        """
        Write merged document to output file.
        
        This method expects options to include:
        - content_file: Path to content file (required)
        - footnotes_file: Path to footnotes file (required)
        - output_format: Format for output (default: docx)
        
        Note: The doc parameter is ignored - we read from content_file and footnotes_file instead.
        """
        opts = {**self.get_default_options(), **options}
        
        content_file = opts.get("content_file")
        footnotes_file = opts.get("footnotes_file")
        output_format = opts.get("output_format", "docx")
        
        if not content_file:
            raise ValueError("content_file option is required")
        if not footnotes_file:
            raise ValueError("footnotes_file option is required")
        
        content_path = Path(content_file)
        footnotes_path = Path(footnotes_file)
        
        if not content_path.exists():
            raise FileNotFoundError(f"Content file not found: {content_path}")
        if not footnotes_path.exists():
            raise FileNotFoundError(f"Footnotes file not found: {footnotes_path}")
        
        # Read both files
        content_doc = self._read_file(content_path)
        footnotes_doc = self._read_file(footnotes_path)
        
        # Parse seif markers and footnotes
        footnotes_by_seif = self._parse_footnotes_by_seif(footnotes_doc)
        
        # Merge footnotes into content
        merged_doc = self._merge_footnotes(content_doc, footnotes_by_seif)
        
        # Write merged document using the appropriate writer
        writer = WriterRegistry.get_writer(output_format)
        if not writer:
            raise ValueError(f"Unknown output format: {output_format}")
        
        # Only pass safe options that don't modify formatting or headings
        # We preserve everything exactly as-is, so we only pass chunking_strategy for JSON output
        safe_options = {}
        if output_format == "json" and "chunking_strategy" in opts:
            safe_options["chunking_strategy"] = opts["chunking_strategy"]
        
        writer.write(merged_doc, output_path, **safe_options)

    def _read_file(self, file_path: Path) -> Document:
        """Read a file using the appropriate reader."""
        reader = ReaderRegistry.get_reader_for_file(file_path)
        if not reader:
            raise ValueError(f"No reader found for file: {file_path}")
        return reader.read(file_path)

    def _parse_footnotes_by_seif(self, footnotes_doc: Document) -> Dict[str, List[Paragraph]]:
        """
        Parse footnotes from footnotes document, organized by seif marker.
        
        Returns a dictionary mapping seif markers (e.g., "א", "ב") to lists of paragraphs.
        """
        footnotes_by_seif: Dict[str, List[Paragraph]] = {}
        
        # Pattern to match seif marker at start of paragraph: Hebrew letter(s) followed by period
        seif_pattern = re.compile(r"^([א-ת]{1,4})\.\s*")
        
        current_seif = None
        current_footnote_paras: List[Paragraph] = []
        
        for para in footnotes_doc.paragraphs:
            text = para.text.strip()
            
            # Check if this paragraph starts with a seif marker
            match = seif_pattern.match(text)
            if match:
                # Save previous footnote if exists
                if current_seif and current_footnote_paras:
                    if current_seif not in footnotes_by_seif:
                        footnotes_by_seif[current_seif] = []
                    footnotes_by_seif[current_seif].extend(current_footnote_paras)
                
                # Start new footnote
                current_seif = match.group(1)
                
                # Create new paragraph without seif marker, preserving all runs and formatting
                new_para = Paragraph()
                new_para.format = para.format
                new_para.style_name = para.style_name
                new_para.heading_level = para.heading_level
                new_para.metadata = para.metadata.copy()
                
                # Remove seif marker from first run if it exists
                seif_end_pos = match.end()
                if para.runs:
                    first_run = para.runs[0]
                    first_run_text = first_run.text
                    
                    # Check if seif marker is in first run
                    if len(first_run_text) >= seif_end_pos:
                        # Seif marker is in first run, remove it
                        remaining_text = first_run_text[seif_end_pos:]
                        if remaining_text.strip():
                            new_run = TextRun(
                                text=remaining_text,
                                style=first_run.style,
                                footnote_id=first_run.footnote_id
                            )
                            new_para.runs.append(new_run)
                        
                        # Add remaining runs
                        for run in para.runs[1:]:
                            new_para.runs.append(run)
                    else:
                        # Seif marker spans multiple runs, handle more carefully
                        # For now, just remove from first run and add rest
                        if len(first_run_text) < seif_end_pos:
                            # Skip first run entirely if seif marker consumes it
                            for run in para.runs[1:]:
                                new_para.runs.append(run)
                        else:
                            remaining_text = first_run_text[seif_end_pos:]
                            if remaining_text.strip():
                                new_run = TextRun(
                                    text=remaining_text,
                                    style=first_run.style,
                                    footnote_id=first_run.footnote_id
                                )
                                new_para.runs.append(new_run)
                            for run in para.runs[1:]:
                                new_para.runs.append(run)
                else:
                    # No runs, create from text
                    remaining_text = text[seif_end_pos:].strip()
                    if remaining_text:
                        new_para.add_run(remaining_text)
                
                if new_para.runs:  # Only add if paragraph has content
                    current_footnote_paras = [new_para]
                else:
                    current_footnote_paras = []
            elif current_seif:
                # Continue current footnote
                if text:  # Only add non-empty paragraphs
                    current_footnote_paras.append(para)
            else:
                # No seif marker yet, skip until we find one
                continue
        
        # Save last footnote
        if current_seif and current_footnote_paras:
            if current_seif not in footnotes_by_seif:
                footnotes_by_seif[current_seif] = []
            footnotes_by_seif[current_seif].extend(current_footnote_paras)
        
        return footnotes_by_seif

    def _merge_footnotes(self, content_doc: Document, footnotes_by_seif: Dict[str, List[Paragraph]]) -> Document:
        """
        Merge footnotes into content document by replacing footnote references with footnote text.
        
        Footnote references in content are like #(א), #(ב), etc.
        
        This method preserves all headings, metadata, and formatting exactly as-is from the content document.
        Only footnote references are replaced with actual footnote objects.
        """
        merged_doc = Document()
        
        # Copy metadata and headings EXACTLY as-is (no modifications)
        merged_doc.metadata = content_doc.metadata
        merged_doc.heading1 = content_doc.heading1
        merged_doc.heading2 = content_doc.heading2
        merged_doc.heading3 = content_doc.heading3
        merged_doc.heading4 = content_doc.heading4
        
        # Pattern to match footnote references: #(א), #(ב), etc.
        footnote_ref_pattern = re.compile(r"#\(([א-ת]{1,4})\)")
        
        # Track footnote IDs for proper footnote references
        footnote_id_counter = 1
        
        for para in content_doc.paragraphs:
            # Create new paragraph
            new_para = Paragraph()
            new_para.format = para.format
            new_para.style_name = para.style_name
            new_para.heading_level = para.heading_level
            new_para.metadata = para.metadata.copy()
            
            # Process runs in the paragraph
            for run in para.runs:
                text = run.text
                
                # Find all footnote references in this run
                matches = list(footnote_ref_pattern.finditer(text))
                
                if not matches:
                    # No footnote references, just copy the run
                    new_run = TextRun(text=text, style=run.style, footnote_id=run.footnote_id)
                    new_para.runs.append(new_run)
                else:
                    # Split text by footnote references and insert footnotes
                    last_end = 0
                    for match in matches:
                        # Add text before the reference
                        if match.start() > last_end:
                            before_text = text[last_end:match.start()]
                            if before_text:
                                new_run = TextRun(text=before_text, style=run.style, footnote_id=run.footnote_id)
                                new_para.runs.append(new_run)
                        
                        # Get seif marker
                        seif_marker = match.group(1)
                        
                        # Find corresponding footnote
                        if seif_marker in footnotes_by_seif:
                            footnote_paras = footnotes_by_seif[seif_marker]
                            
                            # Create footnote object
                            footnote = Footnote(
                                id=footnote_id_counter,
                                paragraphs=footnote_paras.copy(),
                                original_id=footnote_id_counter
                            )
                            merged_doc.add_footnote(footnote)
                            
                            # Add footnote reference run
                            # For now, we'll add the footnote reference inline
                            # In docx output, this will be converted to proper footnote reference
                            new_run = TextRun(
                                text="",  # Empty text, footnote reference will be added by docx writer
                                style=run.style,
                                footnote_id=footnote_id_counter
                            )
                            new_para.runs.append(new_run)
                            
                            footnote_id_counter += 1
                        else:
                            # Footnote not found, keep the reference as text (or remove it)
                            # Option: keep as text, or remove silently
                            # For now, we'll remove it silently
                            pass
                        
                        last_end = match.end()
                    
                    # Add remaining text after last reference
                    if last_end < len(text):
                        remaining_text = text[last_end:]
                        if remaining_text:
                            new_run = TextRun(text=remaining_text, style=run.style, footnote_id=run.footnote_id)
                            new_para.runs.append(new_run)
            
            merged_doc.paragraphs.append(new_para)
        
        return merged_doc

