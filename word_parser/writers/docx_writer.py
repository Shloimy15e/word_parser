"""
Writer for Microsoft Word .docx files.
"""

from pathlib import Path
from typing import Dict, Any

from docx import Document as DocxDocument
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

from word_parser.core.document import Document, Paragraph, HeadingLevel, Alignment
from word_parser.core.processing import is_old_header, should_start_content
from word_parser.writers.base import OutputWriter, WriterRegistry


@WriterRegistry.register
class DocxWriter(OutputWriter):
    """Writer for Microsoft Word .docx files (Open XML format)."""
    
    @classmethod
    def get_format_name(cls) -> str:
        return 'docx'
    
    @classmethod
    def get_extension(cls) -> str:
        return '.docx'
    
    @classmethod
    def get_default_options(cls) -> Dict[str, Any]:
        return {
            'skip_parshah_prefix': False,
            'filter_headers': True,
            'add_blank_lines': True,
        }
    
    def write(self, doc: Document, output_path: Path, **options) -> None:
        """
        Write document to a .docx file.
        
        Options:
            skip_parshah_prefix: Don't add 'פרשת' prefix to H3
            filter_headers: Skip old header paragraphs
            add_blank_lines: Add blank line after each paragraph
        """
        opts = {**self.get_default_options(), **options}
        
        # Create new document
        new_doc = DocxDocument()
        self._configure_styles(new_doc)
        
        # Add headings
        self._add_headings(new_doc, doc, opts.get('skip_parshah_prefix', False))
        
        # Process body paragraphs
        self._add_body_paragraphs(new_doc, doc, opts)
        
        # Save
        output_path.parent.mkdir(parents=True, exist_ok=True)
        new_doc.save(str(output_path))
    
    def _configure_styles(self, docx_doc: DocxDocument) -> None:
        """Configure document styles for Hebrew text."""
        styles = docx_doc.styles
        
        def style_config(style_name, size, rgb, bold=True, space_after=6):
            try:
                s = styles[style_name]
                s.font.size = Pt(size)
                s.font.color.rgb = RGBColor(*rgb)
                s.font.bold = bold
                s.paragraph_format.space_after = Pt(space_after)
            except KeyError:
                pass
        
        style_config("Heading 1", 16, (0x2F, 0x54, 0x96), space_after=6)
        style_config("Heading 2", 13, (0x44, 0x72, 0xC4), space_after=4)
        style_config("Heading 3", 12, (0x1F, 0x37, 0x63), space_after=4)
        style_config("Heading 4", 11, (0x2F, 0x54, 0x96), space_after=4)
        
        try:
            normal = styles["Normal"]
            normal.font.size = Pt(12)
            normal.paragraph_format.space_after = Pt(0)
            normal.paragraph_format.line_spacing = 1.15
        except KeyError:
            pass
    
    def _add_headings(self, docx_doc: DocxDocument, doc: Document, 
                      skip_parshah_prefix: bool) -> None:
        """Add document headings."""
        headings = []
        
        if doc.heading1:
            headings.append(("Heading 1", doc.heading1))
        if doc.heading2:
            headings.append(("Heading 2", doc.heading2))
        if doc.heading3:
            h3_text = doc.heading3 if skip_parshah_prefix else f"פרשת {doc.heading3}"
            headings.append(("Heading 3", h3_text))
        if doc.heading4:
            headings.append(("Heading 4", doc.heading4))
        
        for level, text in headings:
            if text:
                p = docx_doc.add_paragraph(text, style=level)
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p.paragraph_format.right_to_left = True
    
    def _add_body_paragraphs(self, docx_doc: DocxDocument, doc: Document, 
                             opts: Dict[str, Any]) -> None:
        """Add body paragraphs with formatting."""
        filter_headers = opts.get('filter_headers', True)
        add_blank_lines = opts.get('add_blank_lines', True)
        is_multi_parshah = doc.metadata.extra.get('is_multi_parshah', False)
        
        in_header_section = filter_headers
        current_parshah = None  # Track current parshah for multi-parshah mode
        
        for para in doc.paragraphs:
            txt = para.text.strip()
            
            # Handle multi-parshah mode
            if is_multi_parshah:
                # Skip parshah marker lines (*, ה, etc.)
                if para.metadata.get('is_parshah_marker'):
                    continue
                
                # Check if this is a parshah boundary line (skip it, we'll add our own heading)
                if para.metadata.get('is_parshah_boundary'):
                    parshah_name = para.metadata.get('parshah_name', '')
                    if parshah_name and parshah_name != current_parshah:
                        current_parshah = parshah_name
                        # Add the parshah as a Heading 3
                        h3 = docx_doc.add_paragraph(f"פרשת {parshah_name}", style="Heading 3")
                        h3.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                        h3.paragraph_format.right_to_left = True
                    continue  # Skip the original boundary paragraph
                
                # Check if parshah changed (for paragraphs that aren't boundaries)
                para_parshah = para.metadata.get('current_parshah')
                if para_parshah and para_parshah != current_parshah:
                    current_parshah = para_parshah
                    # This shouldn't normally happen, but handle it just in case
            
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
            if txt in ('h'):
                continue
            
            # Create new paragraph
            new_p = docx_doc.add_paragraph()
            
            # Set alignment
            if para.format.alignment == Alignment.CENTER:
                new_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif para.format.alignment == Alignment.LEFT:
                new_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            elif para.format.alignment == Alignment.JUSTIFY:
                new_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            else:
                new_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            
            # Set RTL
            new_p.paragraph_format.right_to_left = True
            
            # Copy paragraph formatting
            pf = new_p.paragraph_format
            if para.format.left_indent is not None:
                pf.left_indent = Pt(para.format.left_indent)
            if para.format.right_indent is not None:
                pf.right_indent = Pt(para.format.right_indent)
            if para.format.first_line_indent is not None:
                pf.first_line_indent = Pt(para.format.first_line_indent)
            if para.format.space_before is not None:
                pf.space_before = Pt(para.format.space_before)
            if para.format.space_after is not None:
                pf.space_after = Pt(para.format.space_after)
            if para.format.line_spacing is not None:
                pf.line_spacing = para.format.line_spacing
            if para.format.keep_together is not None:
                pf.keep_together = para.format.keep_together
            if para.format.keep_with_next is not None:
                pf.keep_with_next = para.format.keep_with_next
            if para.format.page_break_before is not None:
                pf.page_break_before = para.format.page_break_before
            if para.format.widow_control is not None:
                pf.widow_control = para.format.widow_control
            
            # Copy runs
            for run in para.runs:
                new_r = new_p.add_run(run.text)
                
                if run.style.bold is not None:
                    new_r.font.bold = run.style.bold
                if run.style.italic is not None:
                    new_r.font.italic = run.style.italic
                if run.style.underline is not None:
                    new_r.font.underline = run.style.underline
                if run.style.font_size is not None:
                    new_r.font.size = Pt(run.style.font_size)
                if run.style.font_name is not None:
                    new_r.font.name = run.style.font_name
                if run.style.color_rgb is not None:
                    new_r.font.color.rgb = RGBColor(*run.style.color_rgb)
                if run.style.all_caps is not None:
                    new_r.font.all_caps = run.style.all_caps
                if run.style.small_caps is not None:
                    new_r.font.small_caps = run.style.small_caps
                if run.style.strike is not None:
                    new_r.font.strike = run.style.strike
                if run.style.superscript is not None:
                    new_r.font.superscript = run.style.superscript
                if run.style.subscript is not None:
                    new_r.font.subscript = run.style.subscript
            
            # Add blank line after non-empty paragraphs
            if add_blank_lines and txt:
                docx_doc.add_paragraph()
