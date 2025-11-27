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
        """Detect the character encoding from the RTF header and font table."""
        # First check for \ansicpgN in the RTF header (most reliable)
        ansi_match = re.search(r'\\ansicpg(\d+)', rtf_text)
        if ansi_match:
            ansi_cp = int(ansi_match.group(1))
            # Map ANSI code page to charset
            if ansi_cp == 1255:
                return 'cp1255'  # Hebrew
            elif ansi_cp == 1252:
                return 'cp1252'  # Western
            elif ansi_cp == 1251:
                return 'cp1251'  # Russian
            elif ansi_cp == 1250:
                return 'cp1250'  # Eastern Europe
            elif ansi_cp == 1256:
                return 'cp1256'  # Arabic
            elif ansi_cp == 1253:
                return 'cp1253'  # Greek
            elif ansi_cp == 1254:
                return 'cp1254'  # Turkish
            elif ansi_cp == 1257:
                return 'cp1257'  # Baltic
            elif ansi_cp == 1258:
                return 'cp1258'  # Vietnamese
            # For other code pages, try to map
            return f'cp{ansi_cp}'
        
        # Fallback: Look for \fcharsetN in the fonttbl (prefer Hebrew if found)
        charset_matches = re.findall(r'\\fcharset(\d+)', rtf_text)
        if charset_matches:
            # Check if any font uses Hebrew charset (177)
            for charset_id_str in charset_matches:
                charset_id = int(charset_id_str)
                if charset_id == 177:  # Hebrew
                    return 'cp1255'
            # Use the first charset found
            charset_id = int(charset_matches[0])
            return RTF_CHARSET_MAP.get(charset_id, 'cp1252')
        
        return 'cp1252'  # Default to Western
    
    def _parse_rtf(self, rtf_text: str) -> List[tuple]:
        """
        Parse RTF content and extract paragraphs with basic formatting.
        
        Returns list of tuples: (text, is_bold, is_italic, style_name)
        """
        paragraphs = []
        
        # Get charset for decoding hex sequences
        charset = getattr(self, '_default_charset', 'cp1255')  # Default to Hebrew if not set
        
        # Skip RTF header and all metadata until we find actual content
        # Look for \pard which typically marks the start of document content
        content_start = rtf_text.find('\\pard')
        if content_start == -1:
            # Fallback: look for \par or \line which indicate paragraphs
            content_start = rtf_text.find('\\par')
            if content_start == -1:
                content_start = rtf_text.find('\\line')
        
        if content_start > 0:
            # Skip everything before content starts
            rtf_text = rtf_text[content_start:]
        
        # Also trim from the end - look for closing braces that might indicate end of content
        # Find the last meaningful paragraph marker before trailing metadata
        last_par_pos = rtf_text.rfind('\\par')
        if last_par_pos > len(rtf_text) * 0.8:  # If it's in the last 20%, might be trailing metadata
            # Check if there's Hebrew content after this point
            remaining = rtf_text[last_par_pos:]
            if not any('\u0590' <= c <= '\u05ff' for c in remaining[:500]):  # Check first 500 chars
                # No Hebrew content after last par, trim it
                rtf_text = rtf_text[:last_par_pos + 4]  # Keep the \par itself
        
        # Metadata keywords that indicate groups we should skip entirely
        metadata_keywords = [
            'fonttbl', 'colortbl', 'stylesheet', 'info', 'revtbl', 
            'rsidtbl', 'mmathPr', 'xmlnstbl', 'wgrffmtfilter',
            'pnseclvl', 'defchp', 'defpap', 'paperw', 'margl', 'margr',
            'margt', 'margb', 'gutter', 'ltrsect', 'widowctrl', 'ftnbj',
            'aenddoc', 'trackmoves', 'trackformatting', 'donotembedsysfont',
            'relyonvml', 'donotembedlingdata', 'grfdocevents', 'validatexml',
            'showplaceholdtext', 'ignoremixedcontent', 'saveinvalidxml',
            'showxmlerrors', 'horzdoc', 'dghspace', 'dgvspace', 'dghorigin',
            'dgvorigin', 'dghshow', 'dgvshow', 'jcompress', 'viewkind',
            'viewscale', 'rsidroot', 'fet', 'ilfomacatclnup', 'sectd',
            'pgnrestart', 'linex', 'endnhere', 'sectdefaultcl', 'sftnbj'
        ]
        
        # Track formatting state
        current_text = []
        is_bold = False
        is_italic = False
        current_style = None
        pending_hex_bytes = []  # Buffer for multi-byte hex sequences
        
        # Stack to track formatting state for nested groups
        # Each entry is (is_bold, is_italic, current_style)
        format_stack = []
        
        # Track if we're inside a metadata group (skip all content)
        skip_group_depth = 0
        in_metadata_group = False
        
        # Process the RTF content
        i = 0
        while i < len(rtf_text):
            char = rtf_text[i]
            
            if char == '\\':
                # First, flush any pending hex bytes
                if pending_hex_bytes:
                    decoded = self._decode_hex_bytes(pending_hex_bytes)
                    if not in_metadata_group:
                        current_text.append(decoded)
                    pending_hex_bytes = []
                
                # Control word or symbol
                control, end_pos = self._parse_control_word(rtf_text, i)
                
                # Check if this starts a metadata group
                if any(control.startswith(keyword) for keyword in metadata_keywords):
                    # Look backwards for opening brace
                    brace_pos = i - 1
                    while brace_pos >= 0 and rtf_text[brace_pos] in ' \r\n\t':
                        brace_pos -= 1
                    if brace_pos >= 0 and rtf_text[brace_pos] == '{':
                        # We're entering a metadata group - skip it
                        in_metadata_group = True
                        skip_group_depth = 1
                        # Skip to the matching closing brace
                        j = brace_pos + 1
                        while j < len(rtf_text) and skip_group_depth > 0:
                            if rtf_text[j] == '{':
                                skip_group_depth += 1
                            elif rtf_text[j] == '}':
                                skip_group_depth -= 1
                            j += 1
                        i = j
                        in_metadata_group = False
                        continue
                
                # Only process control words if not in metadata group
                if not in_metadata_group:
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
                        # Hex character - decode immediately (RTF uses \'XX for non-ASCII chars)
                        if end_pos + 2 <= len(rtf_text):
                            hex_val = rtf_text[end_pos:end_pos + 2]
                            try:
                                byte_val = int(hex_val, 16)
                                # Decode immediately using detected charset
                                decoded_char = bytes([byte_val]).decode(charset, errors='replace')
                                current_text.append(decoded_char)
                            except (ValueError, UnicodeDecodeError):
                                # Fallback: try cp1255 (Hebrew) then latin-1
                                try:
                                    decoded_char = bytes([byte_val]).decode('cp1255', errors='replace')
                                    current_text.append(decoded_char)
                                except (UnicodeDecodeError, ValueError):
                                    # Last resort: try latin-1
                                    try:
                                        decoded_char = bytes([byte_val]).decode('latin-1', errors='replace')
                                        current_text.append(decoded_char)
                                    except (UnicodeDecodeError, ValueError):
                                        # Final fallback: replacement character
                                        current_text.append('\ufffd')
                            end_pos += 2
                            i = end_pos
                            continue
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
                # Start of group - save current formatting state
                if not in_metadata_group:
                    format_stack.append((is_bold, is_italic, current_style))
                # Flush hex bytes
                if pending_hex_bytes:
                    decoded = self._decode_hex_bytes(pending_hex_bytes)
                    if not in_metadata_group:
                        current_text.append(decoded)
                    pending_hex_bytes = []
                i += 1
            elif char == '}':
                # End of group - restore formatting state
                if not in_metadata_group and format_stack:
                    is_bold, is_italic, current_style = format_stack.pop()
                # Flush hex bytes
                if pending_hex_bytes:
                    decoded = self._decode_hex_bytes(pending_hex_bytes)
                    if not in_metadata_group:
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
                    if not in_metadata_group:
                        current_text.append(decoded)
                    pending_hex_bytes = []
                if not in_metadata_group:
                    current_text.append(char)
                i += 1
        
        # Flush any remaining hex bytes
        if pending_hex_bytes:
            decoded = self._decode_hex_bytes(pending_hex_bytes)
            if not in_metadata_group:
                current_text.append(decoded)
        
        # Don't forget the last paragraph
        text = ''.join(current_text).strip()
        if text:
            paragraphs.append((text, is_bold, is_italic, current_style))
        
        # Filter out paragraphs that are clearly metadata (font names, etc.)
        filtered_paragraphs = []
        for para_text, is_bold, is_italic, style_name in paragraphs:
            # Skip paragraphs that are just font names or metadata
            if self._is_metadata_text(para_text):
                continue
            filtered_paragraphs.append((para_text, is_bold, is_italic, style_name))
        
        return filtered_paragraphs
    
    def _is_metadata_text(self, text: str) -> bool:
        """Check if text looks like metadata (font names, URLs, etc.)"""
        text = text.strip()
        if not text:
            return True
        
        # Skip if it's a URL
        if text.startswith('http://') or text.startswith('https://'):
            return True
        
        # Skip if it contains mostly font names or RTF control words
        # Common font names that might leak through
        font_indicators = ['Times New Roman', 'Arial', 'Cambria', 'Aptos', 'David', 
                          'CE', 'Cyr', 'Greek', 'Tur', 'Hebrew', 'Arabic', 'Baltic', 
                          'Vietnamese', 'Display', 'Math']
        font_count = sum(1 for indicator in font_indicators if indicator in text)
        if font_count >= 2:  # Multiple font indicators = likely metadata
            return True
        
        # Skip if it's mostly non-Hebrew and looks like metadata
        # (e.g., "Unknown;", "2450", etc.)
        if len(text) < 50 and not any('\u0590' <= c <= '\u05ff' for c in text):
            # Check if it contains mostly numbers, punctuation, or English words
            if sum(1 for c in text if c.isalnum() or c in ';:()[]{}') / max(len(text), 1) > 0.8:
                return True
        
        return False
    
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
