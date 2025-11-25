"""
Built-in document format handlers.

This module provides format handlers for common Hebrew document structures:
- StandardFormat: Basic parshah/sefer structure
- DafFormat: Talmud-style daf/perek structure  
- MultiParshahFormat: Single document with multiple sections
- LetterFormat: Correspondence format
"""

import re
from typing import Dict, Any, List

from word_parser.core.document import Document, HeadingLevel
from word_parser.core.formats import DocumentFormat, FormatRegistry
from word_parser.core.processing import (
    is_old_header,
    should_start_content,
    extract_year,
    extract_heading4_info,
    extract_daf_headings,
    detect_parshah_boundary,
)


@FormatRegistry.register
class StandardFormat(DocumentFormat):
    """
    Standard Torah document format.
    
    Structure:
    - H1: Book (e.g., "ליקוטי שיחות")
    - H2: Sefer (e.g., "סדר בראשית")
    - H3: Parshah (e.g., "פרשת בראשית")
    - H4: Year or subsection (e.g., "תש״נ")
    
    Used for most Torah commentaries organized by parshah.
    """
    
    @classmethod
    def get_format_name(cls) -> str:
        return 'standard'
    
    @classmethod
    def get_priority(cls) -> int:
        return 0  # Default/fallback format
    
    @classmethod
    def detect(cls, doc: Document, context: Dict[str, Any]) -> bool:
        # Standard format is the default - always matches as fallback
        return True
    
    @classmethod
    def get_required_context(cls) -> List[str]:
        return ['book']
    
    @classmethod
    def get_optional_context(cls) -> Dict[str, Any]:
        return {
            'sefer': None,
            'parshah': None,
            'filename': None,
            'skip_parshah_prefix': False,
            'filter_headers': True,
        }
    
    def process(self, doc: Document, context: Dict[str, Any]) -> Document:
        """Process document with standard parshah structure."""
        book = context.get('book', '')
        sefer = context.get('sefer', '')
        parshah = context.get('parshah', '')
        filename = context.get('filename', '')
        skip_prefix = context.get('skip_parshah_prefix', False)
        
        # Try to extract year from filename if not provided
        if filename and not context.get('year'):
            year = extract_year(filename)
            heading4_info = extract_heading4_info(filename)
            h4 = year or heading4_info or filename
        else:
            h4 = context.get('year') or filename
        
        # Set headings
        h3 = parshah if skip_prefix else f"פרשת {parshah}" if parshah else None
        doc.set_headings(h1=book, h2=sefer, h3=h3, h4=h4)
        
        # Filter old headers if requested
        if context.get('filter_headers', True):
            doc = self._filter_headers(doc)
        
        return doc
    
    def _filter_headers(self, doc: Document) -> Document:
        """Remove old header paragraphs from document."""
        in_header_section = True
        filtered_paragraphs = []
        
        for para in doc.paragraphs:
            txt = para.text.strip()
            
            if in_header_section:
                if txt and should_start_content(txt):
                    in_header_section = False
                    filtered_paragraphs.append(para)
                elif txt and is_old_header(txt):
                    continue
                elif not txt:
                    continue
            else:
                if txt and is_old_header(txt):
                    continue
                filtered_paragraphs.append(para)
        
        doc.paragraphs = filtered_paragraphs
        return doc


@FormatRegistry.register
class DafFormat(DocumentFormat):
    """
    Talmud/Daf-style document format.
    
    Structure:
    - H1: Book (parent folder or --book arg)
    - H2: Tractate/Masechet (folder name)
    - H3: Perek (extracted from filename, e.g., "פרק א")
    - H4: Chelek/Section (optional, from filename suffix)
    
    Filename patterns:
    - PEREK1 → H3: "פרק א"
    - PEREK1A → H3: "פרק א", H4: "חלק א"
    - MEKOROS → H3: "מקורות"
    - HAKDOMO → H3: "הקדמה"
    """
    
    @classmethod
    def get_format_name(cls) -> str:
        return 'daf'
    
    @classmethod
    def get_priority(cls) -> int:
        return 50
    
    @classmethod
    def detect(cls, doc: Document, context: Dict[str, Any]) -> bool:
        """Detect daf format from context or filename patterns."""
        # Check if explicitly requested
        if context.get('mode') == 'daf':
            return True
        
        # Check filename for perek pattern
        filename = context.get('filename', '').lower()
        if re.match(r'^perek\d+', filename):
            return True
        if re.match(r'^me?koros', filename):
            return True
        if re.match(r'^hakdomo', filename):
            return True
        
        return False
    
    @classmethod
    def get_optional_context(cls) -> Dict[str, Any]:
        return {
            'book': None,  # Can be derived from parent folder
            'folder': None,  # Tractate name
            'filename': None,
            'filter_headers': True,
        }
    
    def process(self, doc: Document, context: Dict[str, Any]) -> Document:
        """Process document with daf/perek structure."""
        book = context.get('book', '')
        folder = context.get('folder', '')
        filename = context.get('filename', '')
        
        # Extract H3 and H4 from filename
        heading3, heading4 = extract_daf_headings(filename)
        
        # Set headings
        doc.set_headings(h1=book, h2=folder, h3=heading3, h4=heading4)
        
        # Filter old headers if requested
        if context.get('filter_headers', True):
            doc = self._filter_headers(doc)
        
        return doc
    
    def _filter_headers(self, doc: Document) -> Document:
        """Remove old header paragraphs from document."""
        in_header_section = True
        filtered_paragraphs = []
        
        for para in doc.paragraphs:
            txt = para.text.strip()
            
            if in_header_section:
                if txt and should_start_content(txt):
                    in_header_section = False
                    filtered_paragraphs.append(para)
                elif txt and is_old_header(txt):
                    continue
                elif not txt:
                    continue
            else:
                if txt and is_old_header(txt):
                    continue
                filtered_paragraphs.append(para)
        
        doc.paragraphs = filtered_paragraphs
        return doc


@FormatRegistry.register  
class MultiParshahFormat(DocumentFormat):
    """
    Multi-parshah document format.
    
    A single document containing multiple sections, where each section
    is marked by a parshah boundary (e.g., "פרשת פנחס") or list item.
    Each section becomes a new section with its own headings.
    
    Structure:
    - H1: Book
    - H2: Sefer
    - H3: Section name (from parshah boundary text)
    
    Detection: Document contains parshah boundary patterns or list-style
    paragraphs that serve as section markers.
    """
    
    @classmethod
    def get_format_name(cls) -> str:
        return 'multi-parshah'
    
    @classmethod
    def get_priority(cls) -> int:
        return 60
    
    @classmethod
    def detect(cls, doc: Document, context: Dict[str, Any]) -> bool:
        """Detect multi-parshah format from document structure."""
        # Check if explicitly requested
        if context.get('mode') == 'multi-parshah':
            return True
        
        # Count parshah boundaries
        boundary_count = 0
        for p in doc.paragraphs:
            is_boundary, _, _ = detect_parshah_boundary(p.text)
            if is_boundary:
                boundary_count += 1
        
        if boundary_count >= 2:
            return True
        
        # Count list items - if multiple, likely multi-parshah
        list_count = sum(1 for p in doc.paragraphs if p.is_list_item())
        return list_count >= 3
    
    @classmethod
    def get_required_context(cls) -> List[str]:
        return ['book', 'sefer']
    
    @classmethod
    def get_optional_context(cls) -> Dict[str, Any]:
        return {
            'skip_parshah_prefix': False,
        }
    
    def process(self, doc: Document, context: Dict[str, Any]) -> Document:
        """Process multi-parshah document into sections."""
        # For multi-parshah, we mark section boundaries in the paragraphs
        book = context.get('book', '')
        sefer = context.get('sefer', '')
        
        doc.set_headings(h1=book, h2=sefer)
        doc.metadata.extra['is_multi_parshah'] = True
        
        # Detect parshah boundaries and mark paragraphs with their section
        self._mark_parshah_sections(doc)
        
        return doc
    
    def _mark_parshah_sections(self, doc: Document) -> None:
        """Mark each paragraph with its parshah section for chunk title generation."""
        current_parshah = None
        prev_text = None
        prev_para = None
        section_index = 0  # Index within current parshah
        
        # Markers that appear before parshah headings and should be removed
        PARSHAH_MARKERS = {'*', 'ה', '***', '* * *', '', 'h'}
        
        for para in doc.paragraphs:
            text = para.text.strip()
            
            # Check for parshah boundary
            is_boundary, parshah_name, year = detect_parshah_boundary(text, prev_text)
            
            if is_boundary:
                current_parshah = parshah_name
                section_index = 0  # Reset index for new parshah
                # Mark this paragraph as a boundary (will be skipped in chunks)
                para.metadata['is_parshah_boundary'] = True
                para.metadata['parshah_name'] = parshah_name
                
                # Also mark the previous paragraph if it was a marker (*, ה, etc.)
                if prev_para is not None and prev_text in PARSHAH_MARKERS:
                    prev_para.metadata['is_parshah_marker'] = True
            else:
                section_index += 1
            
            # Store the current parshah context in paragraph metadata
            para.metadata['current_parshah'] = current_parshah
            para.metadata['section_index'] = section_index
            
            prev_text = text
            prev_para = para
    
    def _extract_sections(self, doc: Document) -> List[Dict]:
        """Extract section information from document using parshah boundary detection."""
        sections = []
        current_section = None
        prev_text = None
        
        for i, para in enumerate(doc.paragraphs):
            text = para.text.strip()
            
            # Check for parshah boundary (e.g., "פרשת פנחס" or standalone "פנחס")
            # Pass previous paragraph text for marker detection (* or ה)
            is_boundary, parshah_name, year = detect_parshah_boundary(text, prev_text)
            
            if is_boundary:
                if current_section:
                    sections.append(current_section)
                
                # Create section title
                title = parshah_name or text
                if year:
                    title = f"{title} ({year})"
                
                current_section = {
                    'title': title,
                    'original_text': text,
                    'start_index': i,
                    'year': year,
                    'paragraphs': []
                }
            elif para.is_list_item():
                # Fallback to list item detection
                if current_section:
                    sections.append(current_section)
                current_section = {
                    'title': text,
                    'original_text': text,
                    'start_index': i,
                    'year': None,
                    'paragraphs': []
                }
            elif current_section is not None:
                current_section['paragraphs'].append(i)
            
            # Remember this text for next iteration
            prev_text = text
        
        if current_section:
            sections.append(current_section)
        
        return sections


@FormatRegistry.register
class LetterFormat(DocumentFormat):
    """
    Letter/correspondence document format.
    
    Structure:
    - H1: Collection name
    - H2: Category (e.g., "מכתבים")
    - H3: Recipient or subject
    - H4: Date
    
    Detection: Document starts with greeting pattern or contains
    letter markers like "ב״ה", date, "כבוד", etc.
    """
    
    @classmethod
    def get_format_name(cls) -> str:
        return 'letter'
    
    @classmethod
    def get_priority(cls) -> int:
        return 40
    
    @classmethod
    def detect(cls, doc: Document, context: Dict[str, Any]) -> bool:
        """Detect letter format from document content."""
        if context.get('mode') == 'letter':
            return True
        
        if not doc.paragraphs:
            return False
        
        # Check first few paragraphs for letter patterns
        first_paras = [p.text.strip() for p in doc.paragraphs[:5] if p.text.strip()]
        
        letter_patterns = [
            r'^ב["\']ה',  # ב"ה at start
            r'^כבוד',  # כבוד (honorific)
            r'^לכבוד',  # לכבוד
            r'^שלום',  # שלום greeting
            r'^ידידי',  # ידידי (my friend)
            r'הנדון:',  # הנדון: (subject)
        ]
        
        for para in first_paras:
            for pattern in letter_patterns:
                if re.match(pattern, para):
                    return True
        
        return False
    
    @classmethod
    def get_optional_context(cls) -> Dict[str, Any]:
        return {
            'book': None,
            'category': 'מכתבים',
            'recipient': None,
            'date': None,
        }
    
    def process(self, doc: Document, context: Dict[str, Any]) -> Document:
        """Process letter document."""
        book = context.get('book', '')
        category = context.get('category', 'מכתבים')
        recipient = context.get('recipient')
        date = context.get('date')
        
        # Try to extract recipient and date from document if not provided
        if not recipient or not date:
            extracted = self._extract_letter_info(doc)
            recipient = recipient or extracted.get('recipient')
            date = date or extracted.get('date')
        
        doc.set_headings(h1=book, h2=category, h3=recipient, h4=date)
        
        return doc
    
    def _extract_letter_info(self, doc: Document) -> Dict[str, str]:
        """Try to extract recipient and date from letter content."""
        info = {}
        
        for para in doc.paragraphs[:10]:
            txt = para.text.strip()
            
            # Look for recipient pattern
            recipient_match = re.match(r'^(?:לכבוד|כבוד)\s+(.+?)(?:\s+שליט״א)?$', txt)
            if recipient_match and 'recipient' not in info:
                info['recipient'] = recipient_match.group(1)
            
            # Look for date patterns
            date_match = re.search(r'(\d{1,2}[./]\d{1,2}[./]\d{2,4})', txt)
            if date_match and 'date' not in info:
                info['date'] = date_match.group(1)
            
            # Hebrew date pattern
            heb_date_match = re.search(r"([א-ת]+'?\s+[א-ת]+\s+תש[א-ת\"']+)", txt)
            if heb_date_match and 'date' not in info:
                info['date'] = heb_date_match.group(1)
        
        return info


@FormatRegistry.register
class SimanFormat(DocumentFormat):
    """
    Siman/Halacha document format.
    
    Structure:
    - H1: Book (e.g., "שולחן ערוך")
    - H2: Section (e.g., "אורח חיים")
    - H3: Siman number (e.g., "סימן א")
    - H4: Seif number (optional)
    
    Used for halachic works organized by siman numbers.
    """
    
    @classmethod
    def get_format_name(cls) -> str:
        return 'siman'
    
    @classmethod
    def get_priority(cls) -> int:
        return 45
    
    @classmethod
    def detect(cls, doc: Document, context: Dict[str, Any]) -> bool:
        """Detect siman format from context or content."""
        if context.get('mode') == 'siman':
            return True
        
        # Check filename for siman pattern
        filename = context.get('filename', '').lower()
        if re.match(r'^siman\d+', filename) or re.match(r'^סימן', filename):
            return True
        
        # Check document for siman markers
        for para in doc.paragraphs[:20]:
            txt = para.text.strip()
            if re.match(r'^סימן\s+[א-ת]+', txt):
                return True
        
        return False
    
    @classmethod
    def get_optional_context(cls) -> Dict[str, Any]:
        return {
            'book': None,
            'section': None,
            'siman': None,
            'seif': None,
            'filename': None,
        }
    
    def process(self, doc: Document, context: Dict[str, Any]) -> Document:
        """Process siman-based document."""
        book = context.get('book', '')
        section = context.get('section', '')
        siman = context.get('siman')
        seif = context.get('seif')
        filename = context.get('filename', '')
        
        # Try to extract siman from filename
        if not siman and filename:
            siman_match = re.match(r'^siman(\d+)', filename.lower())
            if siman_match:
                from word_parser.core.processing import number_to_hebrew_gematria
                num = int(siman_match.group(1))
                siman = f"סימן {number_to_hebrew_gematria(num)}"
        
        doc.set_headings(h1=book, h2=section, h3=siman, h4=seif)
        
        return doc
