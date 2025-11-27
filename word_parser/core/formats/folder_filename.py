"""
Folder-filename format handler - documents with folder-based structure.
"""

import re
from typing import Dict, Any
from pathlib import Path

from word_parser.core.formats._base import (
    DocumentFormat,
    FormatRegistry,
    Document,
    HeadingLevel,
    remove_page_markings,
)
from word_parser.core.document import Alignment, Paragraph
from word_parser.core.processing import is_valid_gematria_number


@FormatRegistry.register
class FolderFilenameFormat(DocumentFormat):
    """
    Format for documents with folder-based structure.

    Structure:
    - H1: Folder name (parent directory)
    - H2: Filename (without extension)
    - H3: One-line sentences detected from content
    - Consecutive H3 candidates are merged into a single H3 line

    Detection: Explicit format selection only.
    """

    @classmethod
    def get_format_name(cls) -> str:
        return "folder-filename"

    @classmethod
    def get_priority(cls) -> int:
        return 15

    @classmethod
    def detect(cls, doc: Document, context: Dict[str, Any]) -> bool:
        return (
            context.get("mode") == "folder-filename"
            or context.get("format") == "folder-filename"
        )

    @classmethod
    def get_optional_context(cls) -> Dict[str, Any]:
        return {
            "book": None,  # Can override folder name
            "filename": None,  # Can override filename
            "input_path": None,  # Used to extract folder/filename
        }

    def process(self, doc: Document, context: Dict[str, Any]) -> Document:
        """Process document with folder-filename structure."""
        # Get H1 from folder name (parent directory of input_path)
        book = context.get("book", "")
        if not book:
            input_path = context.get("input_path", "")
            if input_path:
                try:
                    folder_name = Path(input_path).parent.name
                    if folder_name:
                        book = folder_name
                except Exception:
                    pass

        # Get H2 from filename (stem of input_path)
        filename_h2 = context.get("filename", "")
        if not filename_h2:
            input_path = context.get("input_path", "")
            if input_path:
                try:
                    filename_h2 = Path(input_path).stem
                except Exception:
                    pass

        # Track paragraphs with @ markers before any processing
        para_count_before = len(doc.paragraphs)
        paras_with_at_markers = []
        for i, para in enumerate(doc.paragraphs):
            if para.text and '@' in para.text and re.search(r'@\d+', para.text):
                paras_with_at_markers.append((i, para.text[:60]))
        if paras_with_at_markers:
            print(f"Folder-filename format: Found {len(paras_with_at_markers)} paragraph(s) with @ markers before processing")
            for idx, text in paras_with_at_markers[:5]:  # Show first 5
                print(f"  Para {idx}: '{text}'")
        
        # Remove page markings first
        doc = remove_page_markings(doc)
        print(f"Folder-filename format: After remove_page_markings: {len(doc.paragraphs)} paragraphs")
        
        # Clean @ markers from all paragraphs
        self._clean_at_markers(doc)
        print(f"Folder-filename format: After _clean_at_markers: {len(doc.paragraphs)} paragraphs")
        
        # Clean דף markers and merge paragraphs if needed
        self._clean_daf_markers(doc)
        print(f"Folder-filename format: After _clean_daf_markers: {len(doc.paragraphs)} paragraphs")
        
        # Split paragraphs that contain heading-like text embedded within them
        #self._split_embedded_headings(doc)
        print(f"Folder-filename format: After _split_embedded_headings: {len(doc.paragraphs)} paragraphs")

        # Set base headings - H1 from folder, H2 from filename
        doc.set_headings(h1=book, h2=filename_h2, h3=None, h4=None)

        # Convert old headers (H1/H2 from Word styles) to H3
        self._convert_old_headers_to_h3(doc)

        # Detect footnotes section and mark it
        footnote_start_idx = self._detect_footnotes_start(doc)

        # Detect one-line sentences and mark them as H3 (only in main content, not footnotes)
        # This also merges consecutive H3 candidates into one line
        self._detect_h3_sentences(doc, footnote_start_idx)

        # Remove paragraphs that exactly match H1 or H2 (except actual heading paragraphs and footnotes)
        self._remove_duplicate_headings(doc, book, filename_h2, footnote_start_idx)

        return doc

    def _clean_at_markers(self, doc: Document) -> None:
        """
        Remove @ markers (like @99, @88, @22, etc.) from all paragraph text.
        These are formatting markers that should be cleaned out.
        Only removes @ followed by digits, preserving the rest of the text.
        """
        print(f"Folder-filename format: cleaning @ markers from {len(doc.paragraphs)} paragraphs")
        for para_idx, para in enumerate(doc.paragraphs):
            if para.text:
                original_text = para.text
                # Check if this paragraph contains @ markers
                if '@' in original_text and re.search(r'@\d+', original_text):
                    print(f"  -> Found @ marker in paragraph {para_idx}: '{original_text[:60]}'")
                
                # Remove @ followed by one or more digits (e.g., @01, @12, @99, @0075)
                # This pattern only matches @digits, nothing else
                # Be very explicit: @ followed by one or more ASCII digits (0-9) only
                before_clean = para.text
                para.text = re.sub(r"@[0-9]+", "", para.text)
                
                # Debug: show what was removed
                if before_clean != para.text:
                    removed = before_clean.replace(para.text, "")
                    print(f"     Removed: '{removed}' from text")
                # Clean up any double spaces that might result
                para.text = re.sub(r"\s+", " ", para.text)
                para.text = para.text.strip()
                
                # Debug: log what happened
                if '@' in original_text and re.search(r'@\d+', original_text):
                    print(f"     After cleaning: '{para.text[:60]}'")
                    if original_text and not para.text:
                        print(f"     ERROR: Paragraph became empty after removing @ markers!")
                        print(f"     Original: '{original_text}'")
                        # Restore the text if it became empty - this shouldn't happen
                        para.text = original_text.replace('@01', '').replace('@12', '').replace('@13', '').strip()
                        if not para.text:
                            para.text = original_text  # Keep original if still empty after manual cleanup
                            print(f"     Restored original text to prevent loss")
                    elif len(para.text) < len(original_text) - 5:  # Significant reduction
                        print(f"     WARNING: Text significantly reduced (was {len(original_text)}, now {len(para.text)})")

    def _clean_daf_markers(self, doc: Document) -> None:
        """
        Remove "דף _hebrew_letters_" patterns from paragraphs.
        If a paragraph containing such a pattern is between two paragraphs that are NOT
        separately numbered list items, merge those two paragraphs together.
        """
        print(f"Folder-filename format: cleaning דף markers from {len(doc.paragraphs)} paragraphs")
        
        # Pattern to match "דף" followed by one or more Hebrew letters
        daf_pattern = re.compile(r'דף\s+[א-ת]+')
        
        new_paragraphs = []
        i = 0
        merged_count = 0
        cleaned_count = 0
        
        while i < len(doc.paragraphs):
            para = doc.paragraphs[i]
            text = para.text.strip() if para.text else ""
            
            # Check if this paragraph contains a דף marker
            if daf_pattern.search(text):
                print(f"  -> Found דף marker in paragraph {i}: '{text[:60]}'")
                
                # Check if we can merge with previous and next paragraphs
                has_prev = i > 0
                has_next = i + 1 < len(doc.paragraphs)
                
                if has_prev and has_next:
                    # Previous paragraph should already be in new_paragraphs (we process sequentially)
                    prev_para = new_paragraphs[-1] if new_paragraphs else None
                    next_para = doc.paragraphs[i + 1]
                    
                    if prev_para:
                        # Check if both surrounding paragraphs are NOT numbered list items
                        prev_is_numbered = prev_para.is_numbered_list_item()
                        next_is_numbered = next_para.is_numbered_list_item()
                        
                        # Also check if they're not headings
                        prev_is_heading = prev_para.heading_level != HeadingLevel.NORMAL
                        next_is_heading = next_para.heading_level != HeadingLevel.NORMAL
                        
                        # Also check if they're not empty
                        prev_text = prev_para.text.strip() if prev_para.text else ""
                        next_text = next_para.text.strip() if next_para.text else ""
                        
                        # Merge if both are non-numbered, non-heading, non-empty content paragraphs
                        if (not prev_is_numbered and not next_is_numbered and 
                            not prev_is_heading and not next_is_heading and
                            prev_text and next_text):
                            
                            # Merge: append next paragraph text to previous paragraph
                            merged_text = prev_text.rstrip() + " " + next_text.lstrip()
                            
                            # Update the previous paragraph in new_paragraphs
                            new_paragraphs[-1].text = merged_text
                            
                            print(f"     Merged paragraphs {i-1} and {i+1} (removed דף marker paragraph {i})")
                            merged_count += 1
                            cleaned_count += 1
                            
                            # Skip the דף marker paragraph and the next paragraph (already merged)
                            i += 2
                            continue
                
                # If we can't merge, just remove the דף marker from the text
                # Remove the דף pattern from the text
                cleaned_text = daf_pattern.sub('', text).strip()
                
                # Clean up any double spaces that might result
                cleaned_text = re.sub(r'\s+', ' ', cleaned_text)
                
                if cleaned_text:
                    # Keep the paragraph but with cleaned text
                    para.text = cleaned_text
                    new_paragraphs.append(para)
                    print(f"     Removed דף marker, kept paragraph: '{cleaned_text[:50]}'")
                    cleaned_count += 1
                else:
                    # Paragraph became empty after removing דף marker, skip it
                    print(f"     Removed דף marker paragraph (became empty)")
                    cleaned_count += 1
                
                i += 1
            else:
                # No דף marker, keep paragraph as-is
                new_paragraphs.append(para)
                i += 1
        
        doc.paragraphs = new_paragraphs
        
        if cleaned_count > 0:
            print(f"Folder-filename format: cleaned {cleaned_count} דף marker(s), merged {merged_count} paragraph pair(s)")

    def _convert_old_headers_to_h3(self, doc: Document) -> None:
        """
        Convert old headers (H1 and H2 from Word styles) to H3.
        Old headers should become H3 so they can be merged with other H3 candidates.
        """
        converted_h1_count = 0
        converted_h2_count = 0
        for para in doc.paragraphs:
            if para.heading_level == HeadingLevel.HEADING_1:
                para.heading_level = HeadingLevel.HEADING_3
                converted_h1_count += 1
                print(f"  -> Converted H1 to H3: '{para.text[:50] if para.text else ''}'")
            elif para.heading_level == HeadingLevel.HEADING_2:
                para.heading_level = HeadingLevel.HEADING_3
                converted_h2_count += 1
                print(f"  -> Converted H2 to H3: '{para.text[:50] if para.text else ''}'")
        
        if converted_h1_count > 0 or converted_h2_count > 0:
            print(f"Folder-filename format: converted {converted_h1_count} H1 paragraph(s) and {converted_h2_count} H2 paragraph(s) to H3")

    def _split_embedded_headings(self, doc: Document) -> None:
        """
        Split paragraphs that contain heading-like text embedded within them.
        This handles cases where a heading sentence is merged with surrounding text.
        """
        print(f"Folder-filename format: splitting embedded headings in {len(doc.paragraphs)} paragraphs")
        new_paragraphs = []
        split_count = 0
        
        for para_idx, para in enumerate(doc.paragraphs):
            text = para.text.strip()
            if not text:
                new_paragraphs.append(para)
                continue
            
            # DEBUG: Check if this paragraph contains the target heading text
            if "מדת הצדקה" in text or "מדת" in text and "צדקה" in text:
                print(f"  -> Found potential heading text in paragraph {para_idx}: '{text[:80]}'")
                print(f"     Full length: {len(text)}, Has newlines: {'\\n' in text}")
            
            # First, split on explicit line breaks (newlines)
            lines = text.split('\n')
            if len(lines) > 1:
                print(f"  -> Splitting paragraph {para_idx} on {len(lines)} lines")
                # Already has line breaks - might be multiple paragraphs merged
                for line in lines:
                    line = line.strip()
                    if not line:
                        continue
                    # Create a new paragraph for each non-empty line
                    new_para = Paragraph()
                    new_para.text = line
                    new_para.format = para.format
                    new_para.style_name = para.style_name
                    new_para.heading_level = para.heading_level
                    new_para.metadata = para.metadata.copy()
                    new_paragraphs.append(new_para)
            else:
                # Single line - check if it contains multiple sentences where one might be a heading
                # Look for patterns where a heading-like phrase appears
                # Split on sentence boundaries (. ! ?) but only if one part looks like a heading
                sentences = re.split(r'([.!?])\s+', text)
                if len(sentences) > 3:  # More than one sentence (sentences come with punctuation as separate items)
                    # Reconstruct sentences with their punctuation
                    reconstructed = []
                    for i in range(0, len(sentences), 2):
                        if i + 1 < len(sentences):
                            sentence = sentences[i] + sentences[i + 1]
                        else:
                            sentence = sentences[i]
                        if sentence.strip():
                            reconstructed.append(sentence.strip())
                    
                    # Check if any sentence looks like a heading
                    for i, sentence in enumerate(reconstructed):
                        if self._is_one_line_sentence(sentence):
                            print(f"  -> Found heading-like sentence in paragraph {para_idx}, splitting: '{sentence[:50]}'")
                            split_count += 1
                            # This sentence looks like a heading - split it out
                            # Add everything before as a paragraph
                            if i > 0:
                                before_text = ' '.join(reconstructed[:i])
                                if before_text.strip():
                                    before_para = Paragraph()
                                    before_para.text = before_text.strip()
                                    before_para.format = para.format
                                    before_para.style_name = para.style_name
                                    before_para.heading_level = para.heading_level
                                    before_para.metadata = para.metadata.copy()
                                    new_paragraphs.append(before_para)
                            
                            # Add the heading sentence as its own paragraph
                            heading_para = Paragraph()
                            heading_para.text = sentence
                            heading_para.format = para.format
                            heading_para.style_name = para.style_name
                            heading_para.heading_level = para.heading_level
                            heading_para.metadata = para.metadata.copy()
                            new_paragraphs.append(heading_para)
                            
                            # Add everything after as a paragraph
                            if i + 1 < len(reconstructed):
                                after_text = ' '.join(reconstructed[i+1:])
                                if after_text.strip():
                                    after_para = Paragraph()
                                    after_para.text = after_text.strip()
                                    after_para.format = para.format
                                    after_para.style_name = para.style_name
                                    after_para.heading_level = para.heading_level
                                    after_para.metadata = para.metadata.copy()
                                    new_paragraphs.append(after_para)
                            break
                    else:
                        # No heading-like sentence found, keep as-is
                        new_paragraphs.append(para)
                else:
                    # Single sentence or no clear sentence boundaries
                    # Check if the entire paragraph looks like a heading (might be a heading without punctuation)
                    if self._is_one_line_sentence(text):
                        # The whole paragraph is a heading - keep it as-is
                        new_paragraphs.append(para)
                    else:
                        # Check if there's a heading-like phrase embedded (without punctuation)
                        # Look for phrases that match heading patterns
                        words = text.split()
                        # Check if there's a phrase of 3-8 words that looks like a heading
                        for start in range(len(words)):
                            for end in range(start + 3, min(start + 9, len(words) + 1)):
                                phrase = ' '.join(words[start:end])
                                if self._is_one_line_sentence(phrase):
                                    # Found a heading phrase - split it out
                                    print(f"  -> Found heading phrase without punctuation in paragraph {para_idx}: '{phrase[:50]}'")
                                    split_count += 1
                                    # Add text before phrase
                                    if start > 0:
                                        before_text = ' '.join(words[:start])
                                        if before_text.strip():
                                            before_para = Paragraph()
                                            before_para.text = before_text.strip()
                                            before_para.format = para.format
                                            before_para.style_name = para.style_name
                                            before_para.heading_level = para.heading_level
                                            before_para.metadata = para.metadata.copy()
                                            new_paragraphs.append(before_para)
                                    # Add heading phrase
                                    heading_para = Paragraph()
                                    heading_para.text = phrase
                                    heading_para.format = para.format
                                    heading_para.style_name = para.style_name
                                    heading_para.heading_level = para.heading_level
                                    heading_para.metadata = para.metadata.copy()
                                    new_paragraphs.append(heading_para)
                                    # Add text after phrase
                                    if end < len(words):
                                        after_text = ' '.join(words[end:])
                                        if after_text.strip():
                                            after_para = Paragraph()
                                            after_para.text = after_text.strip()
                                            after_para.format = para.format
                                            after_para.style_name = para.style_name
                                            after_para.heading_level = para.heading_level
                                            after_para.metadata = para.metadata.copy()
                                            new_paragraphs.append(after_para)
                                    break
                            else:
                                continue
                            break
                        else:
                            # No heading phrase found, keep as-is
                            new_paragraphs.append(para)
        
        doc.paragraphs = new_paragraphs
        print(f"Folder-filename format: split {split_count} embedded heading(s), total paragraphs now: {len(new_paragraphs)}")

    def _is_one_line_sentence(self, text: str) -> bool:
        """
        Check if text is a one-line SINGLE SENTENCE that should be H3.
        Simple rule: one-line SINGLE SENTENCE = heading (unless it's a list item).

        Criteria:
        - Single line (no newlines)
        - Contains Hebrew text
        - Reasonable length (not too short, not too long)
        - Must be a SINGLE SENTENCE (ends with sentence punctuation, not multiple sentences)
        - Does NOT start with a siman marking (e.g., "ריב.")
        """
        if not text:
            return False

        text = text.strip()

        # Must be single line
        if "\n" in text:
            return False

        # Must have Hebrew content
        if not any("\u0590" <= c <= "\u05ff" for c in text):
            return False

        # Skip if starts with siman marking (Hebrew letters followed by period)
        # Pattern: 1-4 Hebrew letters (valid gematria) followed by period and space
        siman_match = False  # re.match(r"^([א-ת]{1,4})\.\s+", text)
        if siman_match:
            siman_text = siman_match.group(1)
            if is_valid_gematria_number(siman_text):
                return False  # This is a siman marking, not a heading

        # ENFORCE: Must be a SINGLE SENTENCE
        # Count sentence-ending punctuation marks (. ! ?)
        sentence_endings = text.count('.') + text.count('!') + text.count('?')
        
        # If it has multiple sentence endings, it's multiple sentences - NOT a heading
        if sentence_endings > 1:
            return False
        
        # If it has no sentence-ending punctuation, it might be a heading phrase
        # Headings often don't have punctuation (especially centered headings)
        if sentence_endings == 0:
            # Allow phrases without punctuation if they:
            # 1. Are reasonable length (not too long)
            # 2. Have multiple words (spaces) - indicates it's a phrase, not just a word
            # 3. Or are short enough to be a heading
            has_multiple_words = ' ' in text
            # More lenient: allow up to 120 chars for multi-word phrases without punctuation
            if len(text) > 120:  # Too long to be a heading without punctuation
                return False
            # If it's a single word without punctuation and longer than 25 chars, probably not a heading
            if not has_multiple_words and len(text) > 25:
                return False
            # If it has multiple words and reasonable length, it's likely a heading
            if has_multiple_words and len(text) >= 5 and len(text) <= 120:
                return True  # This is likely a heading phrase
        
        # Should be reasonably long (at least 3 chars) but not too long
        # For a true single sentence heading, limit to ~150 chars max
        if len(text) < 3 or len(text) > 150:
            return False

        # Skip common markers
        markers = ("h", "q", "Y", "*", "***", "* * *")
        if text in markers:
            return False

        # If it passes all checks above, it's a one-line SINGLE SENTENCE = heading
        return True

    def _detect_footnotes_start(self, doc: Document) -> int:
        """
        Detect where footnotes start in the document.
        Returns the index of the first footnote paragraph, or len(doc.paragraphs) if no footnotes found.

        Footnotes are detected by:
        1. Horizontal separator lines (---, ___, ===, or similar repeating characters)
        2. Paragraphs starting with footnote markers like (א), (ב), א., ב., etc.
        """
        for i, para in enumerate(doc.paragraphs):
            txt = para.text.strip()

            # Check for horizontal separator line
            # Pattern: 3+ repeating dashes, underscores, equals, or similar characters
            if re.match(r"^[-=_~\.]{3,}$", txt):
                # Found separator - footnotes start after this
                if i + 1 < len(doc.paragraphs):
                    print(
                        f"Folder-filename format: detected footnotes separator at paragraph {i+1}"
                    )
                    return i + 1

            # Check for footnote marker patterns
            # Pattern 1: (א), (ב), etc. - Hebrew letter in parentheses
            if re.match(r"^\([א-ת]\)", txt):
                print(
                    f"Folder-filename format: detected footnotes starting at paragraph {i+1}"
                )
                return i

            # Pattern 2: א., ב., etc. - Hebrew letter followed by period
            if re.match(r"^[א-ת]\.\s", txt):
                # Make sure it's not a siman marking (which would be longer)
                # If the paragraph is short or looks like a footnote reference, it's a footnote
                if len(txt) < 100:  # Footnotes are typically shorter
                    print(
                        f"Folder-filename format: detected footnotes starting at paragraph {i+1}"
                    )
                    return i

        # No footnotes detected
        return len(doc.paragraphs)

    def _detect_h3_sentences(
        self, doc: Document, footnote_start_idx: int = None
    ) -> None:
        """
        Detect one-line sentences and mark them as H3.
        If an H3 is followed by another H3 candidate, merge them into one line.
        Skips footnotes section if footnote_start_idx is provided.
        """
        if footnote_start_idx is None:
            footnote_start_idx = len(doc.paragraphs)

        print(
            f"Folder-filename format: detecting H3 sentences in {len(doc.paragraphs)} paragraphs (footnotes start at {footnote_start_idx})"
        )

        new_paragraphs = []
        i = 0
        merged_count = 0

        while i < len(doc.paragraphs):
            para = doc.paragraphs[i]
            
            # Skip footnotes section - add them as-is
            if i >= footnote_start_idx:
                new_paragraphs.append(para)
                i += 1
                continue

            txt = para.text.strip()

            # Skip empty paragraphs
            if not txt:
                new_paragraphs.append(para)
                i += 1
                continue

            # if its בס"ד, skip it
            if txt == 'בס"ד':
                new_paragraphs.append(para)
                i += 1
                continue
            
            # If already has a heading level (from previous processing, e.g., converted from H1/H2)
            if para.heading_level != HeadingLevel.NORMAL:
                # If it's H3, check if we should merge with previous or next H3
                if para.heading_level == HeadingLevel.HEADING_3:
                    # Check if previous paragraph is H3 - merge with it
                    if new_paragraphs and new_paragraphs[-1].heading_level == HeadingLevel.HEADING_3:
                        # Merge with previous H3
                        prev_para = new_paragraphs[-1]
                        merged_text = prev_para.text.rstrip() + " " + txt.lstrip()
                        prev_para.text = merged_text
                        print(f"  -> ✓ Merged consecutive H3 (from old header): '{txt[:50]}'")
                        merged_count += 1
                        i += 1
                        # Keep merging with next paragraphs if they're also H3 or heading candidates
                        while i < len(doc.paragraphs) and i < footnote_start_idx:
                            next_para = doc.paragraphs[i]
                            next_txt = next_para.text.strip() if next_para.text else ""
                            if not next_txt:
                                break
                            if next_para.is_numbered_list_item():
                                break
                            # Check if next is H3 or heading candidate
                            next_is_centered = next_para.format.alignment == Alignment.CENTER
                            next_is_one_line = self._is_one_line_sentence(next_txt)
                            next_is_centered_short = next_is_centered and len(next_txt.strip()) >= 3 and len(next_txt.strip()) < 200
                            next_is_short_hebrew_phrase = (
                                any("\u0590" <= c <= "\u05ff" for c in next_txt) and
                                len(next_txt.strip()) >= 3 and 
                                len(next_txt.strip()) < 150 and
                                "\n" not in next_txt
                            )
                            next_is_heading = (next_para.heading_level == HeadingLevel.HEADING_3 or 
                                             next_is_one_line or next_is_centered_short or next_is_short_hebrew_phrase)
                            if next_is_heading:
                                # Merge this one too!
                                merged_text = prev_para.text.rstrip() + " " + next_txt.lstrip()
                                prev_para.text = merged_text
                                print(f"  -> ✓ Continued merging H3 (from old header): '{next_txt[:50]}'")
                                merged_count += 1
                                i += 1
                            else:
                                break
                        continue
                    # Check if next paragraph is also H3 - merge them now
                    elif i + 1 < len(doc.paragraphs) and i + 1 < footnote_start_idx:
                        next_para = doc.paragraphs[i + 1]
                        next_txt = next_para.text.strip() if next_para.text else ""
                        if next_txt and next_para.heading_level == HeadingLevel.HEADING_3:
                            # Merge both H3s into one
                            merged_text = txt.rstrip() + " " + next_txt.lstrip()
                            para.text = merged_text
                            para.heading_level = HeadingLevel.HEADING_3
                            print(f"  -> ✓ Merged consecutive H3s (from old headers): '{txt[:30]}... + {next_txt[:30]}'")
                            new_paragraphs.append(para)
                            merged_count += 1
                            i += 2  # Skip both paragraphs
                            continue
                    # No consecutive H3 found, add as-is
                    new_paragraphs.append(para)
                    i += 1
                    continue
                # For other heading levels, keep as-is
                new_paragraphs.append(para)
                i += 1
                continue

            # Skip numbered list items - they should remain as list items with their formatting
            is_numbered = para.is_numbered_list_item()
            if is_numbered:
                print(f"  -> Skipping numbered list item: '{txt[:50]}'")
                new_paragraphs.append(para)
                i += 1
                continue

            # Check if this is a heading candidate
            # Be very lenient: any single-line Hebrew text that's short is likely a heading
            is_centered = para.format.alignment == Alignment.CENTER
            is_one_line_sentence = self._is_one_line_sentence(txt)
            
            # More lenient detection for centered paragraphs - if centered and short, it's likely a heading
            # This catches cases like "(תפלת מוסף ר"ה ויו"כ)" which might not pass _is_one_line_sentence
            is_centered_short = is_centered and txt and len(txt.strip()) >= 3 and len(txt.strip()) < 200
            
            # Also check if it's a short Hebrew phrase (even if not centered) - might be a heading
            # Be more lenient with length - up to 150 chars for Hebrew phrases
            is_short_hebrew_phrase = (
                txt and 
                any("\u0590" <= c <= "\u05ff" for c in txt) and
                len(txt.strip()) >= 3 and 
                len(txt.strip()) < 150 and
                "\n" not in txt
            )
            
            # ANY single-line Hebrew text that's reasonably short should be considered a heading
            is_heading_candidate = is_one_line_sentence or is_centered_short or is_short_hebrew_phrase
            
            if is_heading_candidate:
                # ALWAYS check if the last paragraph in new_paragraphs is an H3 - merge if so
                if new_paragraphs and new_paragraphs[-1].heading_level == HeadingLevel.HEADING_3:
                    # Merge with previous H3
                    prev_para = new_paragraphs[-1]
                    merged_text = prev_para.text.rstrip() + " " + txt.lstrip()
                    prev_para.text = merged_text
                    reason = "centered" if is_centered else ("short Hebrew phrase" if is_short_hebrew_phrase else "one-line sentence")
                    print(f"  -> ✓ Merged H3 candidate into previous H3 ({reason}): '{txt[:50]}'")
                    merged_count += 1
                    i += 1
                    # After merging, check if the NEXT paragraph is also a heading - keep merging!
                    while i < len(doc.paragraphs) and i < footnote_start_idx:
                        next_para = doc.paragraphs[i]
                        next_txt = next_para.text.strip() if next_para.text else ""
                        if not next_txt:
                            break
                        # Skip numbered list items
                        if next_para.is_numbered_list_item():
                            break
                        # Check if next is also a heading candidate or already H3
                        next_is_centered = next_para.format.alignment == Alignment.CENTER
                        next_is_one_line = self._is_one_line_sentence(next_txt)
                        next_is_centered_short = next_is_centered and len(next_txt.strip()) >= 3 and len(next_txt.strip()) < 200
                        next_is_short_hebrew_phrase = (
                            any("\u0590" <= c <= "\u05ff" for c in next_txt) and
                            len(next_txt.strip()) >= 3 and 
                            len(next_txt.strip()) < 150 and
                            "\n" not in next_txt
                        )
                        next_is_heading = (next_para.heading_level == HeadingLevel.HEADING_3 or 
                                         next_is_one_line or next_is_centered_short or next_is_short_hebrew_phrase)
                        if next_is_heading:
                            # Merge this one too!
                            merged_text = prev_para.text.rstrip() + " " + next_txt.lstrip()
                            prev_para.text = merged_text
                            print(f"  -> ✓ Continued merging H3: '{next_txt[:50]}'")
                            merged_count += 1
                            i += 1
                        else:
                            break
                    continue
                else:
                    # Check if the next paragraph is also a heading candidate - if so, merge them now
                    if i + 1 < len(doc.paragraphs) and i + 1 < footnote_start_idx:
                        next_para = doc.paragraphs[i + 1]
                        next_txt = next_para.text.strip() if next_para.text else ""
                        
                        # Check if next paragraph is also a heading candidate or already H3
                        if next_txt:
                            # If next paragraph is already H3, merge immediately
                            if next_para.heading_level == HeadingLevel.HEADING_3:
                                merged_text = txt.rstrip() + " " + next_txt.lstrip()
                                para.text = merged_text
                                para.heading_level = HeadingLevel.HEADING_3
                                reason = "centered" if is_centered else "one-line sentence"
                                print(f"  -> ✓ Merged H3 candidate with existing H3 ({reason}): '{txt[:30]}... + {next_txt[:30]}'")
                                new_paragraphs.append(para)
                                merged_count += 1
                                i += 2  # Skip both paragraphs
                                continue
                            # If next paragraph is NORMAL, check if it's a heading candidate
                            elif next_para.heading_level == HeadingLevel.NORMAL:
                                next_is_centered = next_para.format.alignment == Alignment.CENTER
                                next_is_one_line = self._is_one_line_sentence(next_txt)
                                next_is_centered_short = next_is_centered and len(next_txt.strip()) >= 3 and len(next_txt.strip()) < 150
                                next_is_short_hebrew_phrase = (
                                    any("\u0590" <= c <= "\u05ff" for c in next_txt) and
                                    len(next_txt.strip()) >= 3 and 
                                    len(next_txt.strip()) < 100 and
                                    "\n" not in next_txt
                                )
                                next_is_heading = next_is_one_line or next_is_centered_short or next_is_short_hebrew_phrase
                                
                                if next_is_heading:
                                    # Merge both into one H3
                                    merged_text = txt.rstrip() + " " + next_txt.lstrip()
                                    para.text = merged_text
                                    para.heading_level = HeadingLevel.HEADING_3
                                    reason = "centered" if is_centered else "one-line sentence"
                                    next_reason = "centered" if next_is_centered else "one-line sentence"
                                    print(f"  -> ✓ Detected consecutive H3 candidates, merged into one H3 ({reason} + {next_reason}): '{txt[:30]}... + {next_txt[:30]}'")
                                    new_paragraphs.append(para)
                                    merged_count += 1
                                    i += 2  # Skip both paragraphs
                                    continue
                    
                    # Set as H3 (no consecutive heading found)
                    para.heading_level = HeadingLevel.HEADING_3
                    reason = "centered" if is_centered else "one-line sentence"
                    print(f"  -> ✓ Detected H3 ({reason}): '{txt[:50]}'")
                    new_paragraphs.append(para)
                    i += 1
            else:
                # Not a heading, add as normal paragraph
                new_paragraphs.append(para)
                i += 1

        doc.paragraphs = new_paragraphs
        
        if merged_count > 0:
            print(f"Folder-filename format: merged {merged_count} consecutive H3 paragraph(s)")

    def _remove_duplicate_headings(
        self, doc: Document, h1: str, h2: str, footnote_start_idx: int = None
    ) -> None:
        """
        Remove paragraphs that exactly match H1 or H2 text.
        Preserves paragraphs that are already marked as headings and footnotes.
        """
        if not h1 and not h2:
            return

        if footnote_start_idx is None:
            footnote_start_idx = len(doc.paragraphs)

        filtered_paragraphs = []
        removed_count = 0

        for i, para in enumerate(doc.paragraphs):
            # Always keep footnotes intact
            if i >= footnote_start_idx:
                filtered_paragraphs.append(para)
                continue

            txt = para.text.strip()

            # Keep paragraphs that are already marked as headings
            if para.heading_level != HeadingLevel.NORMAL:
                filtered_paragraphs.append(para)
                continue

            # Remove content paragraphs that exactly match H1 or H2
            if h1 and txt == h1:
                removed_count += 1
                print(f"  -> Removed duplicate H1: '{txt[:50]}'")
                continue

            if h2 and txt == h2:
                removed_count += 1
                print(f"  -> Removed duplicate H2: '{txt[:50]}'")
                continue

            # Keep all other paragraphs
            filtered_paragraphs.append(para)

        if removed_count > 0:
            print(
                f"Folder-filename format: removed {removed_count} duplicate heading paragraph(s)"
            )

        doc.paragraphs = filtered_paragraphs
