"""
Reader for Microsoft Word .docx files.
"""

from pathlib import Path
from typing import List

from docx import Document as DocxDocument
from docx.enum.text import WD_ALIGN_PARAGRAPH

from word_parser.core.document import (
    Document, Paragraph, TextRun, RunStyle, ParagraphFormat, 
    HeadingLevel, Alignment
)
from word_parser.readers.base import InputReader, ReaderRegistry


@ReaderRegistry.register
class DocxReader(InputReader):
    """Reader for Microsoft Word .docx files (Open XML format)."""
    
    @classmethod
    def get_extensions(cls) -> List[str]:
        return ['.docx']
    
    @classmethod
    def supports_file(cls, file_path: Path) -> bool:
        return file_path.suffix.lower() == '.docx'
    
    @classmethod
    def get_priority(cls) -> int:
        return 100  # Highest priority for .docx files
    
    def read(self, file_path: Path) -> Document:
        """Read a .docx file and return a Document object."""
        source = DocxDocument(str(file_path))
        doc = Document()
        doc.metadata.source_file = str(file_path)
        
        for src_para in source.paragraphs:
            para = self._convert_paragraph(src_para)
            doc.paragraphs.append(para)
        
        return doc
    
    def _convert_paragraph(self, src_para) -> Paragraph:
        """Convert a python-docx paragraph to our Paragraph model."""
        para = Paragraph()
        
        # Convert alignment
        if src_para.alignment == WD_ALIGN_PARAGRAPH.CENTER:
            para.format.alignment = Alignment.CENTER
        elif src_para.alignment == WD_ALIGN_PARAGRAPH.LEFT:
            para.format.alignment = Alignment.LEFT
        elif src_para.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
            para.format.alignment = Alignment.JUSTIFY
        else:
            para.format.alignment = Alignment.RIGHT
        
        # Copy paragraph format
        pf = src_para.paragraph_format
        if pf.left_indent is not None:
            para.format.left_indent = pf.left_indent.pt if hasattr(pf.left_indent, 'pt') else pf.left_indent
        if pf.right_indent is not None:
            para.format.right_indent = pf.right_indent.pt if hasattr(pf.right_indent, 'pt') else pf.right_indent
        if pf.first_line_indent is not None:
            para.format.first_line_indent = pf.first_line_indent.pt if hasattr(pf.first_line_indent, 'pt') else pf.first_line_indent
        if pf.space_before is not None:
            para.format.space_before = pf.space_before.pt if hasattr(pf.space_before, 'pt') else pf.space_before
        if pf.space_after is not None:
            para.format.space_after = pf.space_after.pt if hasattr(pf.space_after, 'pt') else pf.space_after
        para.format.line_spacing = pf.line_spacing
        para.format.keep_together = pf.keep_together
        para.format.keep_with_next = pf.keep_with_next
        para.format.page_break_before = pf.page_break_before
        para.format.widow_control = pf.widow_control
        
        # Store style name
        if src_para.style:
            para.style_name = src_para.style.name
            # Detect heading level from style
            if para.style_name == "Heading 1":
                para.heading_level = HeadingLevel.HEADING_1
            elif para.style_name == "Heading 2":
                para.heading_level = HeadingLevel.HEADING_2
            elif para.style_name == "Heading 3":
                para.heading_level = HeadingLevel.HEADING_3
            elif para.style_name == "Heading 4":
                para.heading_level = HeadingLevel.HEADING_4
        
        # Convert runs
        for src_run in src_para.runs:
            run = self._convert_run(src_run)
            para.runs.append(run)
        
        return para
    
    def _convert_run(self, src_run) -> TextRun:
        """Convert a python-docx run to our TextRun model."""
        style = RunStyle(
            bold=src_run.font.bold,
            italic=src_run.font.italic,
            underline=src_run.font.underline,
            font_size=src_run.font.size.pt if src_run.font.size else None,
            font_name=src_run.font.name,
            color_rgb=(src_run.font.color.rgb.red, src_run.font.color.rgb.green, src_run.font.color.rgb.blue) if src_run.font.color.rgb else None,
            all_caps=src_run.font.all_caps,
            small_caps=src_run.font.small_caps,
            strike=src_run.font.strike,
            superscript=src_run.font.superscript,
            subscript=src_run.font.subscript
        )
        
        return TextRun(text=src_run.text, style=style)
