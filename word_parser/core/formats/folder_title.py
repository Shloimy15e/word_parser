"""
Folder-title format handler - uses folder structure and font-size detection for headings.
"""

from typing import Dict, Any, List
from pathlib import Path

from word_parser.core.formats._base import (
    DocumentFormat,
    FormatRegistry,
    Document,
    HeadingLevel,
    remove_page_markings,
)


@FormatRegistry.register
class FolderTitleFormat(DocumentFormat):
    """
    Format for documents where headings are determined by folder structure and font styling.

    Structure:
    - H1: Parent folder (e.g., "דולה ומשקה") - from path
    - H2: Subfolder name (e.g., "חלק א") - from path
    - H3: Document title - detected by bold + font size 21
    - H4: Subtitles - detected by bold + font size 17

    This format is useful for documents organized in nested folders where:
    - The folder hierarchy provides the book/section context
    - The actual headings within the document are styled with specific font sizes
    """

    @classmethod
    def get_format_name(cls) -> str:
        return "folder-title"

    @classmethod
    def get_priority(cls) -> int:
        return 12

    @classmethod
    def detect(cls, doc: Document, context: Dict[str, Any]) -> bool:
        return (
            context.get("mode") == "folder-title"
            or context.get("format") == "folder-title"
        )

    @classmethod
    def get_optional_context(cls) -> Dict[str, Any]:
        return {
            "book": None,  # Can override H1 (parent folder)
            "sefer": None,  # Can override H2 (subfolder)
            "input_path": None,  # Used to extract folder structure
            "h3_font_size": 21.0,  # Font size for H3 (title)
            "h4_font_size": 17.0,  # Font size for H4 (subtitle)
            "require_bold": True,  # Whether headings must be bold
        }

    def process(self, doc: Document, context: Dict[str, Any]) -> Document:
        """Process document with folder-based structure and font-size heading detection."""
        input_path = context.get("input_path", "")
        
        # Extract H1 and H2 from folder structure
        # For folder-title format, ALWAYS prefer path-based folder names over context values
        # since the folder structure is the primary source of heading hierarchy
        h1 = context.get("book", "")
        h2 = ""  # Always derive from path, not from context
        
        if input_path:
            try:
                path = Path(input_path)
                # Get parent folders
                # Structure: .../H1_folder/H2_subfolder/file.docx
                # Example: docs/סדר בראשית/פרשת נח/מאמר.rtf
                #   - file: מאמר.rtf
                #   - parent (H2): פרשת נח (immediate parent = subfolder)
                #   - grandparent (H1): סדר בראשית (parent of parent = main folder)
                parent = path.parent  # Immediate parent folder (subfolder for H2)
                grandparent = parent.parent  # Parent of parent folder (for H1)
                
                # Always use folder names from path for H2 (subfolder)
                if parent.name:
                    h2 = parent.name
                    print(f"FolderTitleFormat: H2 from subfolder: '{h2}'")
                # Use grandparent for H1 only if not explicitly provided via book
                if not h1 and grandparent.name:
                    h1 = grandparent.name
                    print(f"FolderTitleFormat: H1 from parent folder: '{h1}'")
                    
                print(f"FolderTitleFormat: Path structure - file: {path.name}, parent: {parent.name}, grandparent: {grandparent.name}")
            except Exception as e:
                print(f"FolderTitleFormat: Error extracting folder names: {e}")

        # Get font size thresholds
        h3_size = context.get("h3_font_size", 21.0)
        h4_size = context.get("h4_font_size", 17.0)
        require_bold = context.get("require_bold", True)

        # IMPORTANT: Detect headings BEFORE removing page markings
        # This ensures headings are identified before any content is removed
        self._detect_headings_by_font_size(doc, h3_size, h4_size, require_bold)
        
        # Now remove page markings (this won't affect detected headings)
        doc = remove_page_markings(doc)

        # Extract first H3 text for document-level heading
        h3_text = None
        for para in doc.paragraphs:
            if para.heading_level == HeadingLevel.HEADING_3:
                h3_text = self._clean_heading_text(para.text.strip())
                # Also clean the paragraph text itself
                para.text = h3_text
                break

        # Set document-level headings
        doc.set_headings(h1=h1, h2=h2, h3=h3_text, h4=None)

        return doc

    def _clean_heading_text(self, text: str) -> str:
        """
        Clean heading text by removing individual garbage characters.
        
        Removes garbage characters like dots, parentheses, brackets that appear
        as artifacts from shape/textbox extraction, but preserves sentence structure.
        Only removes these characters when they appear:
        - At the start or end of the text
        - In sequences (multiple consecutive garbage chars)
        """
        import re
        if not text:
            return text
        
        # Remove leading garbage (dots, parentheses, spaces)
        text = re.sub(r'^[\s\.\(\)\[\]\{\}]+', '', text)
        # Remove trailing garbage
        text = re.sub(r'[\s\.\(\)\[\]\{\}]+$', '', text)
        
        # Remove sequences of garbage chars in the middle (3+ consecutive)
        # This catches things like "...()()" but preserves legitimate punctuation
        text = re.sub(r'[\.\(\)\[\]\{\}]{3,}', '', text)
        
        # Clean up any double/triple spaces left behind
        text = re.sub(r'\s{2,}', ' ', text)
        
        return text.strip()

    def _detect_headings_by_font_size(
        self,
        doc: Document,
        h3_size: float,
        h4_size: float,
        require_bold: bool,
    ) -> None:
        """
        Detect headings based on font size and bold formatting.
        Combines consecutive paragraphs of the same heading level into one.
        
        Args:
            doc: Document to process
            h3_size: Font size for H3 headings (e.g., 21)
            h4_size: Font size for H4 headings (e.g., 17)
            require_bold: Whether headings must be bold
        """
        # First pass: detect heading levels for each paragraph
        for para in doc.paragraphs:
            if not para.runs:
                continue
            
            text = para.text.strip()
            if not text:
                continue
            
            # Check if paragraph has consistent font size across all runs
            font_sizes = []
            is_all_bold = True
            
            for run in para.runs:
                if run.text.strip():  # Only consider runs with actual text
                    if run.style.font_size is not None:
                        font_sizes.append(run.style.font_size)
                    if run.style.bold is not True:
                        is_all_bold = False
            
            if not font_sizes:
                continue
            
            # Check if all runs have the same font size
            if len(set(font_sizes)) != 1:
                continue  # Mixed font sizes - not a heading
            
            font_size = font_sizes[0]
            
            # Check bold requirement
            if require_bold and not is_all_bold:
                continue
            
            # Determine heading level based on font size
            # Use a tolerance of 0.5 for font size matching
            if abs(font_size - h3_size) <= 0.5:
                para.heading_level = HeadingLevel.HEADING_3
                print(f"FolderTitleFormat: Detected H3 (size {font_size}): '{text[:50]}'")
            elif abs(font_size - h4_size) <= 0.5:
                para.heading_level = HeadingLevel.HEADING_4
                print(f"FolderTitleFormat: Detected H4 (size {font_size}): '{text[:50]}'")
        
        # Second pass: combine consecutive headings of the same level
        self._combine_consecutive_headings(doc)

    def _combine_consecutive_headings(self, doc: Document) -> None:
        """
        Combine consecutive paragraphs with the same heading level (H3 or H4) into one.
        The text is joined with a space.
        """
        if not doc.paragraphs:
            return
        
        paragraphs_to_remove = []
        i = 0
        
        while i < len(doc.paragraphs):
            para = doc.paragraphs[i]
            
            # Only process H3 and H4 headings
            if para.heading_level not in (HeadingLevel.HEADING_3, HeadingLevel.HEADING_4):
                i += 1
                continue
            
            # Look ahead for consecutive paragraphs with the same heading level
            j = i + 1
            combined_text_parts = [para.text.strip()]
            
            while j < len(doc.paragraphs):
                next_para = doc.paragraphs[j]
                
                # Check if next paragraph has the same heading level
                if next_para.heading_level == para.heading_level:
                    next_text = next_para.text.strip()
                    if next_text:
                        combined_text_parts.append(next_text)
                    paragraphs_to_remove.append(j)
                    j += 1
                else:
                    break
            
            # If we found consecutive headings, combine them
            if len(combined_text_parts) > 1:
                combined_text = " ".join(combined_text_parts)
                para.text = combined_text
                # Also update runs to reflect combined text
                if para.runs:
                    para.runs[0].text = combined_text
                    para.runs = para.runs[:1]  # Keep only first run
                print(f"FolderTitleFormat: Combined {len(combined_text_parts)} consecutive H{3 if para.heading_level == HeadingLevel.HEADING_3 else 4} headings: '{combined_text[:50]}'")
            
            i = j
        
        # Remove the merged paragraphs (in reverse order to maintain indices)
        for idx in reversed(paragraphs_to_remove):
            del doc.paragraphs[idx]