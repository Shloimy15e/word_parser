"""
Core processing utilities for Hebrew document parsing.

This module contains functions for:
- Header detection and filtering
- Hebrew gematria conversion
- Year extraction from filenames
- Parshah boundary detection
"""

import re
from typing import Optional, Tuple


# -------------------------------
# Header detection patterns
# -------------------------------
HEADER_HINTS = [
    r"^דברות",
    r"^סדר",
    r"^פרשת",
    r"^שנת",
    r"^תש[\"׳]",
    r"^ס\"ג",
    r"^בעיר",
    r"^ב\"ה",
    r"^ליקוטי",
    r"^במסיבת",
    r"^מוצ\"ש",
    r"^מוצאי",
    r"^מוצש\"ק",
    r"^בבית.*התורה",
    r"^שבת",
    r"^פרשת.*שנת",
    r"^כ\"ק",
    r"לפ\"ק$",
    r"^יום.*פרשת.*שנת",
    r"^יום\s+[א-ת]['\"]",
]


def is_old_header(text: str) -> bool:
    """
    Returns True if the paragraph looks like an old title/header line
    that should be skipped.
    """
    t = text.strip()
    if not t:
        return False  # Empty paragraphs should be preserved, not filtered

    # Single character paragraphs (like *) should be preserved
    if len(t) == 1:
        return False

    # Check against known header patterns
    if any(re.match(p, t) for p in HEADER_HINTS):
        return True

    # Skip short lines without punctuation (likely titles)
    # But NOT if it contains brackets [ ] which might be Torah text or single symbols
    if len(t) < 25 and not re.search(r"[.!?,\[\]\*]", t):
        return True

    return False


def should_start_content(text: str) -> bool:
    """
    Returns True if this paragraph looks like substantive Torah content
    (long paragraph ≥60 chars OR contains Torah markers like brackets),
    signaling we're past the header section.
    """
    t = text.strip()
    # Torah content often has brackets for biblical quotes
    if "[" in t or "]" in t:
        return True
    # Or is a long paragraph
    return len(t) >= 60


# -------------------------------
# Gematria conversion
# -------------------------------
def number_to_hebrew_gematria(num: int) -> str:
    """
    Convert a number to Hebrew gematria notation.
    Examples: 1 → א, 2 → ב, 10 → י, 11 → יא, 20 → כ, 21 → כא, etc.
    """
    if num <= 0:
        return str(num)

    # Hebrew letters and their numeric values
    ones = ["", "א", "ב", "ג", "ד", "ה", "ו", "ז", "ח", "ט"]  # 0-9
    tens = ["", "י", "כ", "ל", "מ", "ן", "ס", "ע", "פ", "צ"]  # 0, 10-90
    hundreds = ["", "ק", "ר", "ש", "ת"]  # 0, 100-400

    result = ""

    # Handle hundreds
    if num >= 100:
        hundreds_digit = min(num // 100, 4)
        result += hundreds[hundreds_digit]
        num %= 100

    # Special cases for 15 and 16 (avoid using God's name)
    if num == 15:
        return result + "טו"
    elif num == 16:
        return result + "טז"

    # Handle tens
    if num >= 10:
        tens_digit = num // 10
        result += tens[tens_digit]
        num %= 10

    # Handle ones
    if num > 0:
        result += ones[num]

    return result if result else str(num)


def is_valid_gematria_number(text: str) -> bool:
    """
    Check if a Hebrew text is a valid gematria number (not a regular word).
    Hebrew alphabet numbering: א=1, ב=2, ג=3... or gematria combinations.
    """
    # ALL single Hebrew letters are valid numbers (Hebrew alphabet numbering)
    if len(text) == 1:
        return True

    # Exclude common Hebrew WORDS that aren't numbers (multi-letter only)
    non_numbers = {
        "מבוא",
        "פרק",
        "חלק",
        "סימן",
        "דרוש",
        "מאמר",
        "שיחה",
        "הקדמה",
        "תוכן",
        "ענין",
        "דבר",
        "מכתב",
        "נושא",
        "הערות",
        "הגהות",
        "ביאור",
        "פסוק",
        "דין",
        "הלכה",
        "מצוה",
        "הערה",
    }
    if text in non_numbers:
        return False

    # For multi-letter: if it's 2-4 letters and not in blacklist, likely a gematria number
    return len(text) <= 4


# -------------------------------
# Heading extraction
# -------------------------------
def extract_heading4_info(filename_stem: str) -> Optional[str]:
    """
    Extract heading 4 information from filename.
    Handles special patterns:
      - "PEREK1" or "perek1" → "פרק א"
      - "PEREK2" → "פרק ב"
      - "PEREK11" → "פרק יא"
      - "PEREK01A" or "perek1a" → "פרק א 1" (letter becomes number, number becomes letter)
      - "MEKOROS" or "MKOROS" → "מקורות"
      - "MEKOROS1" → "מקורות א"
      - "HAKDOMO" or "HAKDOMO1" → "הקדמה" or "הקדמה א"
    Returns the Hebrew string or None if no pattern matched.
    """
    stem = filename_stem.strip().lower()

    # Check for MEKOROS/MKOROS with optional number
    mekoros_match = re.match(r"^me?koros0*(\d*)$", stem, re.IGNORECASE)
    if mekoros_match:
        num_str = mekoros_match.group(1)
        if num_str:
            number = int(num_str)
            hebrew_gematria = number_to_hebrew_gematria(number)
            return f"מקורות {hebrew_gematria}"
        else:
            return "מקורות"

    # Check for HAKDOMO with optional number
    hakdomo_match = re.match(r"^hakdomo0*(\d*)$", stem, re.IGNORECASE)
    if hakdomo_match:
        num_str = hakdomo_match.group(1)
        if num_str:
            number = int(num_str)
            hebrew_gematria = number_to_hebrew_gematria(number)
            return f"הקדמה {hebrew_gematria}"
        else:
            return "הקדמה"

    # Pattern: chelek/חלק followed by number (with optional leading zeros) and optional letter
    chelek_match = re.match(r"^(?:chelek|חלק)0*(\d+)([a-z])?$", stem, re.IGNORECASE)
    if chelek_match:
        number = int(chelek_match.group(1))
        letter = chelek_match.group(2)

        # Convert number to Hebrew gematria
        hebrew_gematria = number_to_hebrew_gematria(number)

        if letter:
            # Convert letter to number (a=1, b=2, etc.)
            letter_num = ord(letter.lower()) - ord("a") + 1
            return f"חלק {hebrew_gematria} {letter_num}"
        else:
            return f"חלק {hebrew_gematria}"

    # Pattern: perek followed by number (with optional leading zeros) and optional letter
    perek_match = re.match(r"^perek0*(\d+)([a-z])?$", stem, re.IGNORECASE)
    if perek_match:
        number = int(perek_match.group(1))
        letter = perek_match.group(2)

        # Convert number to Hebrew gematria
        hebrew_gematria = number_to_hebrew_gematria(number)

        if letter:
            # Convert letter to number (a=1, b=2, etc.)
            letter_num = ord(letter.lower()) - ord("a") + 1
            return f"פרק {hebrew_gematria} {letter_num}"
        else:
            return f"פרק {hebrew_gematria}"

    return None


def extract_daf_headings(filename_stem: str) -> Tuple[Optional[str], Optional[str]]:
    """
    Extract both Heading 3 and Heading 4 from filename for daf mode.
    Examples:
      - "PEREK1A" → (H3: "פרק א", H4: "חלק א")
      - "PEREK1" → (H3: "פרק א", H4: None)
      - "MEKOROS2" → (H3: "מקורות ב", H4: None)
      - "HAKDOMO1" → (H3: "הקדמה א", H4: None)
    Returns tuple: (heading3, heading4)
    """
    stem = filename_stem.strip().lower()

    # Check for MEKOROS/MKOROS with optional number
    mekoros_match = re.match(r"^me?koros0*(\d*)$", stem, re.IGNORECASE)
    if mekoros_match:
        num_str = mekoros_match.group(1)
        if num_str:
            number = int(num_str)
            hebrew_gematria = number_to_hebrew_gematria(number)
            return (f"מקורות {hebrew_gematria}", None)
        else:
            return ("מקורות", None)

    # Check for HAKDOMO with optional number
    hakdomo_match = re.match(r"^hakdomo0*(\d*)$", stem, re.IGNORECASE)
    if hakdomo_match:
        num_str = hakdomo_match.group(1)
        if num_str:
            number = int(num_str)
            hebrew_gematria = number_to_hebrew_gematria(number)
            return (f"הקדמה {hebrew_gematria}", None)
        else:
            return ("הקדמה", None)

    # Pattern: chelek/חלק followed by number (with optional leading zeros) and optional letter
    chelek_match = re.match(r"^(?:chelek|חלק)0*(\d+)([a-z])?$", stem, re.IGNORECASE)
    if chelek_match:
        number = int(chelek_match.group(1))
        letter = chelek_match.group(2)

        # Convert number to Hebrew gematria
        hebrew_gematria = number_to_hebrew_gematria(number)
        heading3 = f"חלק {hebrew_gematria}"

        if letter:
            # Convert letter to Hebrew "חלק" (section)
            letter_gematria = number_to_hebrew_gematria(
                ord(letter.lower()) - ord("a") + 1
            )
            heading4 = f"חלק {letter_gematria}"
            return (heading3, heading4)
        else:
            return (heading3, None)

    # Pattern: perek followed by number (with optional leading zeros) and optional letter
    perek_match = re.match(r"^perek0*(\d+)([a-z])?$", stem, re.IGNORECASE)
    if perek_match:
        number = int(perek_match.group(1))
        letter = perek_match.group(2)

        # Convert number to Hebrew gematria
        hebrew_gematria = number_to_hebrew_gematria(number)
        heading3 = f"פרק {hebrew_gematria}"

        if letter:
            # Convert letter to Hebrew "חלק" (section)
            letter_gematria = number_to_hebrew_gematria(
                ord(letter.lower()) - ord("a") + 1
            )
            heading4 = f"חלק {letter_gematria}"
            return (heading3, heading4)
        else:
            return (heading3, None)

    # Fallback: match everything
    name_match = re.match(r"^(.*)$", stem, re.IGNORECASE)
    if name_match:
        base_name = name_match.group(1)
        # Try to extract trailing number for heading4 if present
        number_match = re.match(r"^(.*?)(\d+)$", stem)
        if number_match:
            number = number_match.group(2)
            number_gematria = number_to_hebrew_gematria(int(number))
            heading4 = f"חלק {number_gematria}"
            return (number_match.group(1), heading4)
        return (base_name, None)

    return (None, None)


# -------------------------------
# Year extraction
# -------------------------------
def extract_year(filename_stem: str) -> Optional[str]:
    """
    Extract year from filename.
    Looks for Hebrew year pattern like תש״כ, תשכ_ז תשכח, תשנ״ט, etc.
    Years always start with תש (taf-shin) and are 3-4 characters long.
    Returns None if no year found (year is optional).
    """
    stem = filename_stem.strip()

    # Split by common separators (including underscore)
    parts = re.split(r"[\s\-–—_]+", stem)
    parts = [p.strip() for p in parts if p.strip()]

    # Look for year pattern: must start with תש and be 3-4 chars total
    year_pattern = r"^תש[\u0590-\u05FF״׳\"]$|^תש[\u0590-\u05FF״׳\"][\u0590-\u05FF״׳\"]$"

    for part in parts:
        if re.match(year_pattern, part) and 3 <= len(part) <= 4:
            return part

    # Fallback: look for תש pattern with correct length
    for part in parts:
        if len(part) >= 3 and len(part) <= 4 and part[0:2] == "תש":
            return part

    return None


def extract_year_from_text(text: str) -> Optional[str]:
    """
    Extract year from text (similar to extract_year but works on paragraph text).
    Looks for Hebrew year pattern like תש״כ, תשכ_ז תשכח, תשנ״ט, etc.
    """
    if not text:
        return None

    # Look for year pattern in the text
    year_pattern = r"תש[\u0590-\u05FF״׳\"][\u0590-\u05FF״׳\"]?"
    matches = re.findall(year_pattern, text)

    for match in matches:
        if 3 <= len(match) <= 4:
            return match

    return None


# -------------------------------
# Parshah boundary detection
# -------------------------------

# List of known parshah names for standalone detection
PARSHAH_NAMES = {
    # בראשית
    "בראשית",
    "נח",
    "לך",
    "לך לך",
    "וירא",
    "חיי שרה",
    "חי שרה",
    "חיי",
    "תולדות",
    "ויצא",
    "וישלח",
    "וישב",
    "מקץ",
    "ויגש",
    "ויחי",
    # שמות
    "שמות",
    "וארא",
    "בא",
    "בשלח",
    "יתרו",
    "משפטים",
    "תרומה",
    "תצוה",
    "כי תשא",
    "תשא",
    "ויקהל",
    "פקודי",
    "ויקהל פקודי",
    # ויקרא
    "ויקרא",
    "צו",
    "שמיני",
    "תזריע",
    "מצורע",
    "תזריע מצורע",
    "אחרי מות",
    "אחרי",
    "קדושים",
    "אחרי קדושים",
    "אחרי מות קדושים",
    "אמור",
    "בהר",
    "בחקתי",
    "בחקותי",
    "בהר בחקתי",
    "בהר בחקותי",
    # במדבר
    "במדבר",
    "נשא",
    "בהעלתך",
    "בהעלותך",
    "שלח",
    "שלח לך",
    "קרח",
    "חקת",
    "בלק",
    "פנחס",
    "פינחס",
    "מטות",
    "מסעי",
    "מטות מסעי",
    # דברים
    "דברים",
    "ואתחנן",
    "עקב",
    "ראה",
    "שפטים",
    "שופטים",
    "כי תצא",
    "תצא",
    "כי תבא",
    "תבוא",
    "כי תבוא",
    "נצבים",
    "וילך",
    "נצבים וילך",
    "האזינו",
    "וזאת הברכה",
    "ברכה",
}


def detect_parshah_boundary(
    text: str, prev_text: str = None, enable_siman_detection: bool = False
) -> Tuple[bool, Optional[str], Optional[str]]:
    """
    Detect if a paragraph indicates the start of a new parshah or section.
    ...
    """
    if not text:
        return (False, None, None)

    # Clean text: strip whitespace AND invisible formatting chars (LRM, RLM, BOM)
    # \u200e = LRM, \u200f = RLM, \ufeff = BOM
    txt = text.strip().strip("\u200e\u200f\ufeff")

    # Skip if too long to be a heading (more than ~50 chars)
    if len(txt) > 50:
        return (False, None, None)

    # Pattern 0: Hebrew letter-number (siman)
    if enable_siman_detection:
        siman_match = re.match(r"^([א-ת]{1,4})[\.\s\t]*$", txt)
        if siman_match and len(txt) <= 10:
            siman = siman_match.group(1)
            if is_valid_gematria_number(siman):
                return (True, f"{siman}.", None)

    # Pattern 1: "פרשת [name]"
    parshah_match = re.match(r"^פרשת\s+([א-ת\s]+?)$", txt)
    if parshah_match:
        parshah_name = parshah_match.group(1).strip()
        parshah_normalized = (
            parshah_name.replace('"', "")
            .replace("'", "")
            .replace("״", "")
            .replace("׳", "")
        )
        if parshah_normalized in PARSHAH_NAMES:
            year = extract_year_from_text(txt)
            return (True, parshah_name, year)

    # Pattern 2: "פרשת [name] - [year]"
    parshah_with_year = re.match(
        r"^פרשת\s+([א-ת\s]+?)(?:\s+שנת|\s*[-–—])\s*(.+?)$", txt
    )
    if parshah_with_year:
        parshah_name = parshah_with_year.group(1).strip()
        parshah_normalized = (
            parshah_name.replace('"', "")
            .replace("'", "")
            .replace("״", "")
            .replace("׳", "")
        )
        if parshah_normalized in PARSHAH_NAMES:
            year_text = parshah_with_year.group(2).strip()
            year = extract_year_from_text(year_text) or extract_year_from_text(txt)
            return (True, parshah_name, year)

    # Pattern 2.5: Standalone parshah name with trailing word
    standalone_with_trailing = re.match(r"^([א-ת\s]+?)\s*[-–—]\s*(.+?)$", txt)
    if standalone_with_trailing:
        parshah_name = standalone_with_trailing.group(1).strip()
        parshah_normalized = (
            parshah_name.replace('"', "")
            .replace("'", "")
            .replace("״", "")
            .replace("׳", "")
        )
        if parshah_normalized in PARSHAH_NAMES:
            return (True, parshah_name, None)

    # Pattern 3: Standalone parshah name
    txt_normalized = (
        txt.replace('"', "").replace("'", "").replace("״", "").replace("׳", "")
    )
    if txt_normalized in PARSHAH_NAMES:
        if prev_text is not None:
            prev_stripped = prev_text.strip()
            # Allow empty, markers, OR parenthesized text (page numbers)
            is_parenthesized = re.match(r"^[\(\)]+.+[\(\)]+$", prev_stripped)
            if prev_stripped in ("", "*", "ה", "***", "* * *") or is_parenthesized:
                return (True, txt_normalized, None)
        else:
            return (True, txt_normalized, None)

    return (False, None, None)


# -------------------------------
# Text sanitization
# -------------------------------
def sanitize_xml_text(text: str) -> str:
    """
    Remove characters that are not valid in XML.
    XML 1.0 valid characters: #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD]
    """

    def is_valid_xml_char(c):
        codepoint = ord(c)
        return (
            codepoint == 0x09  # Tab
            or codepoint == 0x0A  # Line feed
            or codepoint == 0x0D  # Carriage return
            or (0x20 <= codepoint <= 0xD7FF)
            or (0xE000 <= codepoint <= 0xFFFD)
        )

    return "".join(c for c in text if is_valid_xml_char(c))


# -------------------------------
# Page marking detection and removal
# -------------------------------
def is_page_marking(text: str) -> bool:
    """
    Check if a paragraph is a page marking that should be removed.

    Page markings are:
    1. ((xxx)) - Text in double parentheses (e.g., ((תקכב)), ((כב)))
    2. #א or א# - Hebrew letter with # symbol

    Returns True if the text is a page marking.
    """
    txt = text.strip()
    if not txt:
        return False

    # Pattern 1: ((xxx)) - double parentheses
    if re.match(r"^\(\(.+\)\)$", txt):
        return True

    # Pattern 2: #א or א# - Hebrew letter with # (one or more letters)
    if re.match(r"^#[א-ת]+$", txt) or re.match(r"^[א-ת]+#$", txt):
        return True

    return False


def remove_page_markings(doc):
    """
    Remove page markings from document and merge split paragraphs.

    Page markings like ((תקכב)) or #א that appear between content paragraphs
    are removed, and the paragraphs above and below are merged into one.

    Args:
        doc: The document to process (Document instance)

    Returns:
        The document with page markings removed and paragraphs merged
    """
    if not doc.paragraphs:
        return doc

    from word_parser.core.document import HeadingLevel

    new_paragraphs = []
    i = 0

    while i < len(doc.paragraphs):
        current_para = doc.paragraphs[i]
        current_text = current_para.text.strip()

        # Check if current paragraph is a page marking
        if is_page_marking(current_text):
            # This is a page marking - check if we should merge paragraphs
            has_prev = len(new_paragraphs) > 0
            has_next = i + 1 < len(doc.paragraphs)

            if has_prev and has_next:
                prev_para = new_paragraphs[-1]
                next_para = doc.paragraphs[i + 1]

                # Only merge if both surrounding paragraphs are non-empty content (not headings)
                prev_text = prev_para.text.strip()
                next_text = next_para.text.strip()

                is_prev_content = (
                    prev_text
                    and prev_para.heading_level == HeadingLevel.NORMAL
                    and not is_page_marking(prev_text)
                )
                is_next_content = (
                    next_text
                    and next_para.heading_level == HeadingLevel.NORMAL
                    and not is_page_marking(next_text)
                )

                if is_prev_content and is_next_content:
                    # Merge: append next paragraph text to previous paragraph
                    # Add a space between them to ensure proper word separation
                    merged_text = (
                        prev_para.text.rstrip() + " " + next_para.text.lstrip()
                    )
                    prev_para.text = merged_text

                    # Skip the page marking and the next paragraph (already merged)
                    i += 2
                    continue

            # If we didn't merge, just skip the page marking
            i += 1
            continue

        # Normal paragraph - add to results
        new_paragraphs.append(current_para)
        i += 1

    doc.paragraphs = new_paragraphs
    return doc


def clean_dos_text(text: str) -> str:
    """
    Clean DOS text - remove ALL numbers, brackets, and formatting codes.
    Keep ONLY Hebrew text and basic punctuation.
    """
    lines = text.split("\n")
    cleaned_lines = []

    for line in lines:
        line = line.strip()

        # Preserve empty lines
        if not line:
            cleaned_lines.append("")
            continue

        # Skip formatting lines starting with period
        if line.startswith("."):
            continue

        # Must have Hebrew content
        if not any("\u0590" <= c <= "\u05ff" for c in line):
            continue

        temp = line

        # Remove >number< footnote markers (including those with Hebrew letters like >3ט<)
        temp = re.sub(r">[\d\u0590-\u05ff]+<", "", temp)

        # Remove BNARF/OISAR/BSNF markers
        temp = re.sub(r"(BNARF|OISAR|BSNF)\s+[A-Z]\s+\d+[\*]?", "", temp)

        # Remove ALL brackets
        temp = re.sub(r"[<>]", "", temp)

        # Remove ALL numbers (integers and decimals)
        temp = re.sub(r"\d+\.?\d*", "", temp)

        # Remove asterisks
        temp = re.sub(r"\*", "", temp)

        # Remove multiple dashes
        temp = re.sub(r"[-–—]{2,}", "", temp)

        # Remove English letters (codes)
        temp = re.sub(r"[A-Za-z]+", "", temp)

        # Clean up spaces
        temp = re.sub(r"\s+", " ", temp)
        temp = temp.strip()

        # Only keep if has Hebrew
        if temp and any("\u0590" <= c <= "\u05ff" for c in temp):
            cleaned_lines.append(temp)

    return "\n".join(cleaned_lines)
