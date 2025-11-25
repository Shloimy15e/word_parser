"""
Reader for Rich Text Format (.rtf) files.
"""

import re
from pathlib import Path
from typing import List, Optional, Dict

from word_parser.core.document import Document, Paragraph, Alignment, TextRun, RunStyle
from word_parser.readers.base import InputReader, ReaderRegistry


# Mapping of RTF charset IDs to Python codec names
RTF_CHARSET_MAP = {
    0: 'cp1252',      # ANSI
    1: 'cp1252',      # DEFAULT (assume Western)
    2: 'ascii',       # SYMBOL
    77: 'mac-roman',  # MAC
    128: 'cp932',     # SHIFTJIS (Japanese)
    129: 'cp949',     # HANGUL (Korean)
    130: 'johab',     # JOHAB (Korean)
    134: 'gb2312',    # GB2312 (Chinese Simplified)
    136: 'big5',      # BIG5 (Chinese Traditional)
    161: 'cp1253',    # GREEK
    162: 'cp1254',    # TURKISH
    163: 'cp1258',    # VIETNAMESE
    177: 'cp1255',    # HEBREW
    178: 'cp1256',    # ARABIC
    186: 'cp1257',    # BALTIC
    204: 'cp1251',    # RUSSIAN
    222: 'cp874',     # THAI
    238: 'cp1250',    # EASTEUROPE
    254: 'cp437',     # PC437
    255: 'cp850',     # OEM
}


@ReaderRegistry.register
class RtfReader(InputReader):
    """
    Reader for Rich Text Format (.rtf) files.
    
    Parses RTF files and extracts text content with basic formatting.
    Handles Hebrew text and RTL paragraphs correctly.
    """
    
    @classmethod
    def get_extensions(cls) -> List[str]:
        return ['.rtf']
    
    @classmethod
    def supports_file(cls, file_path: Path) -> bool:
        if file_path.suffix.lower() != '.rtf':
            return False
        # Verify it's a valid RTF file by checking the header
        try:
            with open(file_path, 'rb') as f:
                header = f.read(5)
                return header == b'{\\rtf'
        except (IOError, OSError):
            return False
    
    @classmethod
    def get_priority(cls) -> int:
        return 85  # Between .doc (90) and .idml (80)
    
    def read(self, file_path: Path) -> Document:
        """Read an RTF file and return a Document object."""
        with open(file_path, 'rb') as f:
            rtf_content = f.read()
        
        # Decode as latin-1 to preserve byte values (RTF is ASCII with \' escapes)
        rtf_text = rtf_content.decode('latin-1')
        
        doc = Document()
        doc.metadata.source_file = str(file_path)
        
        # Detect the character encoding from the font table
        self._default_charset = self._detect_charset(rtf_text)
        
        # Parse RTF and extract text
        paragraphs = self._parse_rtf(rtf_text)
        
        for para_text, is_bold, is_italic, style_name in paragraphs:
            if para_text.strip():
                para = doc.add_paragraph()
                para.format.alignment = Alignment.RIGHT
                para.format.right_to_left = True
                para.format.style_name = style_name
                
                # Add text with formatting
                style = RunStyle(bold=is_bold, italic=is_italic)
                para.runs.append(TextRun(text=para_text, style=style))
        
        return doc
    
    def _detect_charset(self, rtf_text: str) -> str:
        """Detect the character encoding from the RTF font table."""
        # Look for \fcharsetN in the fonttbl
        charset_match = re.search(r'\\fcharset(\d+)', rtf_text)
        if charset_match:
            charset_id = int(charset_match.group(1))
            return RTF_CHARSET_MAP.get(charset_id, 'cp1252')
        return 'cp1252'  # Default to Western
    
    def _parse_rtf(self, rtf_text: str) -> List[tuple]:
        """
        Parse RTF content and extract paragraphs with basic formatting.
        
        Returns list of tuples: (text, is_bold, is_italic, style_name)
        """
        paragraphs = []
        
        # Remove RTF header and preamble (fonttbl, colortbl, stylesheet, info, etc.)
        # These are groups at the start of the document that we want to skip
        rtf_text = re.sub(r'^\{\\rtf\d?', '', rtf_text)
        
        # Skip document metadata groups (fonttbl, colortbl, stylesheet, info)
        rtf_text = re.sub(r'\{\\fonttbl[^{}]*(?:\{[^{}]*\}[^{}]*)*\}', '', rtf_text)
        rtf_text = re.sub(r'\{\\colortbl[^{}]*\}', '', rtf_text)
        rtf_text = re.sub(r'\{\\stylesheet[^{}]*(?:\{[^{}]*\}[^{}]*)*\}', '', rtf_text)
        rtf_text = re.sub(r'\{\\info[^{}]*(?:\{[^{}]*\}[^{}]*)*\}', '', rtf_text)
        
        # Track formatting state
        current_text = []
        is_bold = False
        is_italic = False
        current_style = None
        pending_hex_bytes = []  # Buffer for multi-byte hex sequences
        group_depth = 0  # Track nested groups
        skip_group = False  # Whether to skip content in current group
        
        # Process the RTF content
        i = 0
        while i < len(rtf_text):
            char = rtf_text[i]
            
            if char == '\\':
                # First, flush any pending hex bytes
                if pending_hex_bytes:
                    decoded = self._decode_hex_bytes(pending_hex_bytes)
                    current_text.append(decoded)
                    pending_hex_bytes = []
                
                # Control word or symbol
                control, end_pos = self._parse_control_word(rtf_text, i)
                
                if control == 'par' or control == 'line':
                    # Paragraph break
                    text = ''.join(current_text).strip()
                    if text:
                        paragraphs.append((text, is_bold, is_italic, current_style))
                    current_text = []
                elif control == 'b':
                    is_bold = True
                elif control == 'b0':
                    is_bold = False
                elif control == 'i':
                    is_italic = True
                elif control == 'i0':
                    is_italic = False
                elif control.startswith('s') and control[1:].isdigit():
                    # Style reference like \s0, \s1
                    pass  # We'll get the style from stylesheet
                elif control == "'":
                    # Hex character - collect bytes
                    if end_pos + 2 <= len(rtf_text):
                        hex_val = rtf_text[end_pos:end_pos + 2]
                        try:
                            byte_val = int(hex_val, 16)
                            pending_hex_bytes.append(byte_val)
                        except ValueError:
                            pass
                        end_pos += 2
                elif control.startswith('u') and len(control) > 1 and control[1:].lstrip('-').isdigit():
                    # Unicode character \uN
                    try:
                        unicode_val = int(control[1:])
                        if unicode_val < 0:
                            unicode_val += 65536
                        current_text.append(chr(unicode_val))
                    except ValueError:
                        pass
                elif control == '\\':
                    current_text.append('\\')
                elif control == '{':
                    current_text.append('{')
                elif control == '}':
                    current_text.append('}')
                elif control == 'tab':
                    current_text.append('\t')
                elif control == '~':
                    current_text.append('\u00A0')  # Non-breaking space
                
                i = end_pos
            elif char == '{':
                # Start of group - flush hex bytes
                if pending_hex_bytes:
                    decoded = self._decode_hex_bytes(pending_hex_bytes)
                    current_text.append(decoded)
                    pending_hex_bytes = []
                i += 1
            elif char == '}':
                # End of group - flush hex bytes
                if pending_hex_bytes:
                    decoded = self._decode_hex_bytes(pending_hex_bytes)
                    current_text.append(decoded)
                    pending_hex_bytes = []
                i += 1
            elif char == '\r' or char == '\n':
                # Ignore line breaks in RTF (they're not significant)
                i += 1
            else:
                # Regular character - flush hex bytes first
                if pending_hex_bytes:
                    decoded = self._decode_hex_bytes(pending_hex_bytes)
                    current_text.append(decoded)
                    pending_hex_bytes = []
                current_text.append(char)
                i += 1
        
        # Flush any remaining hex bytes
        if pending_hex_bytes:
            decoded = self._decode_hex_bytes(pending_hex_bytes)
            current_text.append(decoded)
        
        # Don't forget the last paragraph
        text = ''.join(current_text).strip()
        if text:
            paragraphs.append((text, is_bold, is_italic, current_style))
        
        return paragraphs
    
    def _decode_hex_bytes(self, byte_list: List[int]) -> str:
        """Decode a sequence of hex bytes using the detected charset."""
        try:
            byte_array = bytes(byte_list)
            return byte_array.decode(self._default_charset)
        except (UnicodeDecodeError, AttributeError):
            # Fallback: try cp1255 (Hebrew) then latin-1
            try:
                return bytes(byte_list).decode('cp1255')
            except UnicodeDecodeError:
                return bytes(byte_list).decode('latin-1', errors='replace')
    
    def _parse_control_word(self, rtf_text: str, start: int) -> tuple:
        """
        Parse an RTF control word starting at position start.
        Returns (control_word, end_position).
        """
        i = start + 1  # Skip the backslash
        
        if i >= len(rtf_text):
            return ('', i)
        
        # Check for control symbol (single character)
        char = rtf_text[i]
        if not char.isalpha():
            # Control symbol like \\ \{ \} \' etc.
            return (char, i + 1)
        
        # Control word (letters followed by optional number)
        word_start = i
        while i < len(rtf_text) and rtf_text[i].isalpha():
            i += 1
        
        # Optional numeric parameter (can be negative)
        if i < len(rtf_text) and (rtf_text[i].isdigit() or rtf_text[i] == '-'):
            while i < len(rtf_text) and (rtf_text[i].isdigit() or rtf_text[i] == '-'):
                i += 1
        
        control_word = rtf_text[word_start:i]
        
        # Skip optional space delimiter
        if i < len(rtf_text) and rtf_text[i] == ' ':
            i += 1
        
        return (control_word, i)
