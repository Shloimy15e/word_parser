import re
import json
import argparse
import tempfile
from pathlib import Path
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

try:
    import win32com.client
    WORD_AVAILABLE = True
except ImportError:
    WORD_AVAILABLE = False


# -------------------------------
# Helper: set up style formatting
# -------------------------------
def configure_styles(doc):
    styles = doc.styles

    def style_config(style_name, size, rgb, bold=True, space_after=6):
        s = styles[style_name]
        s.font.size = Pt(size)
        s.font.color.rgb = RGBColor(*rgb)
        s.font.bold = bold
        s.paragraph_format.space_after = Pt(space_after)

    style_config("Heading 1", 16, (0x2F, 0x54, 0x96), space_after=6)
    style_config("Heading 2", 13, (0x44, 0x72, 0xC4), space_after=4)
    style_config("Heading 3", 12, (0x1F, 0x37, 0x63), space_after=4)
    style_config("Heading 4", 11, (0x2F, 0x54, 0x96), space_after=4)
    
    # Configure Normal style
    normal = styles["Normal"]
    normal.font.size = Pt(12)
    normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.line_spacing = 1.15


# -------------------------------------------------
# Helper: decide which paragraphs are "old headers"
# -------------------------------------------------
HEADER_HINTS = [
    r"^דברות", r"^סדר", r"^פרשת", r"^שנת", r"^תש[\"׳]", 
    r"^ס\"ג", r"^בעיר", r"^ב\"ה", r"^ליקוטי",
    r"^במסיבת", r"^מוצ\"ש", r"^מוצאי", r"^מוצש\"ק", r"^בבית.*התורה",
    r"^שבת", r"^פרשת.*שנת", r"^כ\"ק", r"לפ\"ק$",
    r"^יום.*פרשת.*שנת", r"^יום\s+[א-ת]['\"]"
]

def is_old_header(text):
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


def should_start_content(text):
    """
    Returns True if this paragraph looks like substantive Torah content
    (long paragraph ≥60 chars OR contains Torah markers like brackets),
    signaling we're past the header section.
    """
    t = text.strip()
    # Torah content often has brackets for biblical quotes
    if '[' in t or ']' in t:
        return True
    # Or is a long paragraph
    return len(t) >= 60


# ----------------------------------------
# Helper: convert .doc to .docx using Word COM
# ----------------------------------------
def convert_doc_to_docx(doc_path):
    """
    Convert .doc file to .docx using Word COM automation.
    Returns path to temporary .docx file.
    """
    if not WORD_AVAILABLE:
        raise RuntimeError("pywin32 not installed. Cannot convert .doc files.")
    
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    
    try:
        # Open the .doc file
        doc = word.Documents.Open(str(doc_path.absolute()))
        
        # Create temp file for .docx
        temp_docx = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
        temp_docx.close()
        
        # Save as .docx (format 16 = wdFormatXMLDocument)
        doc.SaveAs(temp_docx.name, FileFormat=16)
        doc.Close()
        
        return Path(temp_docx.name)
    finally:
        word.Quit()


# ----------------------------------------
# Core conversion: one docx → formatted docx
# ----------------------------------------
def extract_year(filename_stem):
    """
    Extract year from filename.
    Looks for Hebrew year pattern like תש״כ, תשכח, etc.
    """
    stem = filename_stem.strip()
    
    # Split by common separators
    parts = re.split(r'[\s\-–—]+', stem)
    parts = [p.strip() for p in parts if p.strip()]
    
    # Look for year pattern (starts with ת and contains Hebrew letters)
    year_pattern = r'^ת[\u0590-\u05FF״׳\"]+'
    
    for part in parts:
        # Check if it's Hebrew and starts with ת
        if re.match(year_pattern, part):
            return part
    
    # Fallback: look for any part with Hebrew characters starting with ת
    for part in parts:
        if part and part[0] == 'ת' and any('\u0590' <= c <= '\u05FF' for c in part):
            return part
    
    # Last resort: return last part if it contains Hebrew
    for part in reversed(parts):
        if any('\u0590' <= c <= '\u05FF' for c in part):
            return part
    
    return parts[-1] if parts else ""


def reformat_docx(input_path, output_path, book, sefer, parshah, year, skip_parshah_prefix=False):
    source = Document(input_path)
    new_doc = Document()
    configure_styles(new_doc)

    # Add document headings
    parshah_heading = parshah if skip_parshah_prefix else f"פרשת {parshah}"
    
    for level, text in [
        ("Heading 1", book),
        ("Heading 2", sefer),
        ("Heading 3", parshah_heading),
        ("Heading 4", year),
    ]:
        if not text:
            continue
        p = new_doc.add_paragraph(text, style=level)
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        # Ensure RTL for Hebrew text
        p.paragraph_format.right_to_left = True

    # Process body text with smart header skipping
    in_header_section = True
    
    for para in source.paragraphs:
        # Get the full paragraph text including ALL characters
        full_text = para.text
        txt = full_text.strip()
        
        # If we're still in the header section
        if in_header_section:
            # Check if this looks like substantial content
            if txt and should_start_content(txt):
                in_header_section = False
                # Fall through to copy this paragraph
            # Skip if it's an old header
            elif txt and is_old_header(txt):
                continue
            # Skip empty paragraphs in header section
            elif not txt:
                continue
            else:
                # Non-header, non-empty text in header section - shouldn't happen but be safe
                continue
        
        # After header section started
        # Skip only matching old headers, preserve EVERYTHING else including empty paragraphs
        if txt and is_old_header(txt):
            continue
            
        # Copy the entire paragraph element to preserve ALL formatting
        # This includes empty paragraphs which create spacing
        new_p = new_doc.add_paragraph()
        
        # Copy paragraph-level formatting
        pf_source = para.paragraph_format
        pf_new = new_p.paragraph_format
        
        # Preserve centered alignment for asterisks, force RTL right alignment for everything else
        if para.alignment == WD_ALIGN_PARAGRAPH.CENTER:
            new_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            new_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        
        # Set RTL direction for Hebrew
        pf_new.right_to_left = True
        
        # Copy all paragraph format attributes
        if pf_source.left_indent is not None:
            pf_new.left_indent = pf_source.left_indent
        if pf_source.right_indent is not None:
            pf_new.right_indent = pf_source.right_indent
        if pf_source.first_line_indent is not None:
            pf_new.first_line_indent = pf_source.first_line_indent
        
        # Always copy spacing - these are critical for layout
        pf_new.space_before = pf_source.space_before
        pf_new.space_after = pf_source.space_after
        pf_new.line_spacing = pf_source.line_spacing
        if pf_source.line_spacing_rule is not None:
            pf_new.line_spacing_rule = pf_source.line_spacing_rule
        
        # Copy keep together settings
        if pf_source.keep_together is not None:
            pf_new.keep_together = pf_source.keep_together
        if pf_source.keep_with_next is not None:
            pf_new.keep_with_next = pf_source.keep_with_next
        if pf_source.page_break_before is not None:
            pf_new.page_break_before = pf_source.page_break_before
        if pf_source.widow_control is not None:
            pf_new.widow_control = pf_source.widow_control
        
        # Copy ALL runs including ones with just symbols/whitespace
        for run in para.runs:
            # Copy the run even if it's just whitespace or symbols
            new_r = new_p.add_run(run.text)
            
            # Copy all font properties
            if run.font.bold is not None:
                new_r.font.bold = run.font.bold
            if run.font.italic is not None:
                new_r.font.italic = run.font.italic
            if run.font.underline is not None:
                new_r.font.underline = run.font.underline
            if run.font.size is not None:
                new_r.font.size = run.font.size
            if run.font.name is not None:
                new_r.font.name = run.font.name
            if run.font.color.rgb is not None:
                new_r.font.color.rgb = run.font.color.rgb
            if run.font.highlight_color is not None:
                new_r.font.highlight_color = run.font.highlight_color
            if run.font.all_caps is not None:
                new_r.font.all_caps = run.font.all_caps
            if run.font.small_caps is not None:
                new_r.font.small_caps = run.font.small_caps
            if run.font.strike is not None:
                new_r.font.strike = run.font.strike
            if run.font.superscript is not None:
                new_r.font.superscript = run.font.superscript
            if run.font.subscript is not None:
                new_r.font.subscript = run.font.subscript
        
        # Add a blank line after each paragraph with content
        if txt:
            new_doc.add_paragraph()

    new_doc.save(output_path)


def convert_to_json(input_path, output_path, book, sefer, parshah, year, skip_parshah_prefix=False):
    """
    Convert docx to JSON structure with chunks.
    Each paragraph becomes a chunk.
    """
    source = Document(input_path)
    
    # Build book name (parshah is the "book" in the JSON structure)
    parshah_name = parshah if skip_parshah_prefix else f"פרשת {parshah}"
    
    # Create JSON structure
    json_data = {
        "book_name_he": parshah_name,
        "book_name_en": parshah,  # Keep Hebrew as fallback for English
        "book_metadata": {
            "sefer_he": sefer,
            "sefer_en": sefer,
            "collection_he": book,
            "collection_en": book,
            "year_he": year,
            "year_en": year,
            "source": "Word Document Conversion"
        },
        "chunks": []
    }
    
    # Process body text with smart header skipping
    in_header_section = True
    chunk_id = 1
    
    for para in source.paragraphs:
        full_text = para.text
        txt = full_text.strip()
        
        # If we're still in the header section
        if in_header_section:
            # Check if this looks like substantial content
            if txt and should_start_content(txt):
                in_header_section = False
                # Fall through to include this paragraph
            # Skip if it's an old header
            elif txt and is_old_header(txt):
                continue
            # Skip empty paragraphs in header section
            elif not txt:
                continue
            else:
                continue
        
        # After header section, skip old headers but keep everything else
        if txt and is_old_header(txt):
            continue
        
        # Only add non-empty paragraphs as chunks
        if txt:
            chunk = {
                "chunk_id": chunk_id,
                "chunk_metadata": {
                    "chunk_title": f"{parshah_name} - קטע {chunk_id}",
                    "sefer": sefer,
                    "parshah": parshah_name,
                    "year": year,
                    "collection": book
                },
                "text": txt
            }
            json_data["chunks"].append(chunk)
            chunk_id += 1
    
    # Write JSON file
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(json_data, f, ensure_ascii=False, indent=2)


# -------------------------------
# Main CLI entry point
# -------------------------------
def main():
    parser = argparse.ArgumentParser(
        description="Reformat Hebrew DOCX files to standardized schema."
    )
    parser.add_argument("--book", required=True, help="Book title (Heading 1)")
    parser.add_argument(
        "--sefer", help="Sefer/tractate title (Heading 2). If not provided, uses folder name."
    )
    parser.add_argument(
        "--parshah", help="Parshah name (Heading 3). If not provided, uses subfolder names."
    )
    parser.add_argument(
        "--skip-parshah-prefix", action="store_true", 
        help="Skip adding 'פרשת' prefix to parshah name in Heading 3"
    )
    parser.add_argument(
        "--json", action="store_true",
        help="Output as JSON structure instead of formatted Word documents"
    )
    parser.add_argument(
        "--docs", default="docs", help="Input folder containing .docx files or subfolders"
    )
    parser.add_argument("--out", default="output", help="Output folder")
    args = parser.parse_args()

    docs_dir = Path(args.docs)
    out_dir = Path(args.out)
    
    # Check if docs_dir exists
    if not docs_dir.exists():
        print(f"Error: Input directory '{docs_dir}' does not exist")
        return
    
    if not docs_dir.is_dir():
        print(f"Error: '{docs_dir}' is not a directory")
        return
    
    out_dir.mkdir(exist_ok=True)
    
    # Create json subdirectory if needed
    if args.json:
        (out_dir / "json").mkdir(exist_ok=True)

    # Check if using folder structure mode (no sefer/parshah specified)
    if not args.sefer and not args.parshah:
        # Use folder name as sefer, subfolders as parshah
        sefer = docs_dir.name
        
        # Get all subdirectories
        subdirs = [d for d in docs_dir.iterdir() if d.is_dir()]
        
        if not subdirs:
            print(f"No subdirectories found in {docs_dir}")
            return
        
        print(f"📚 Processing folder structure: {sefer}\n")
        total_success = 0
        total_files = 0
        
        for subdir in subdirs:
            parshah = subdir.name
            files = list(subdir.glob("*.docx")) + list(subdir.glob("*.doc"))
            
            if not files:
                continue
            
            # Create output subdirectory
            if args.json:
                out_subdir = out_dir / "json" / sefer / parshah
            else:
                out_subdir = out_dir / sefer / parshah
            out_subdir.mkdir(parents=True, exist_ok=True)
            
            print(f"📂 {parshah} ({len(files)} file(s))")
            
            for i, path in enumerate(files, 1):
                temp_docx = None
                try:
                    year = extract_year(Path(path).stem)
                    if not year:
                        print(f"  [{i}/{len(files)}] ⚠️ Skipping {path.name}: cannot extract year")
                        continue
                    
                    if args.json:
                        out_name = f"{Path(path).stem}.json"
                        out_path = out_subdir / out_name
                    else:
                        out_name = f"{Path(path).stem}-formatted.docx"
                        out_path = out_subdir / out_name
                    
                    print(f"  [{i}/{len(files)}] {path.stem} → {out_path.name} ...", end=" ")
                    
                    input_path = path
                    if path.suffix.lower() == '.doc':
                        print("(converting .doc...) ", end="")
                        temp_docx = convert_doc_to_docx(path)
                        input_path = temp_docx
                    
                    if args.json:
                        convert_to_json(input_path, out_path, args.book, sefer, parshah, year, args.skip_parshah_prefix)
                    else:
                        reformat_docx(input_path, out_path, args.book, sefer, parshah, year, args.skip_parshah_prefix)
                    print("✓ done")
                    total_success += 1
                    total_files += 1
                    
                except Exception as e:
                    print(f"⚠️ error: {e}")
                    total_files += 1
                finally:
                    if temp_docx and temp_docx.exists():
                        temp_docx.unlink()
            
            print()
        
        print(f"✅ All done. Successfully processed {total_success}/{total_files} file(s).")
        return
    
    # Original single folder mode
    if not args.sefer or not args.parshah:
        print("Error: Both --sefer and --parshah are required when not using folder structure mode")
        return

    # Collect both .doc and .docx files
    files = list(docs_dir.glob("*.docx")) + list(docs_dir.glob("*.doc"))
    if not files:
        print(f"No .doc or .docx files found in {docs_dir}")
        return

    print(f"📚 Processing {len(files)} file(s)...\n")

    success_count = 0
    for i, path in enumerate(files, 1):
        temp_docx = None
        try:
            # Extract year from ORIGINAL filename (before any conversion)
            year = extract_year(Path(path).stem)
            if not year:
                print(f"[{i}/{len(files)}] ⚠️ Skipping {path.name}: cannot extract year")
                continue
            
            # Output filename based on format
            if args.json:
                out_name = f"{Path(path).stem}.json"
                out_path = out_dir / "json" / out_name
            else:
                out_name = f"{Path(path).stem}-formatted.docx"
                out_path = out_dir / out_name
            
            print(f"[{i}/{len(files)}] Processing {path.stem} → {out_path.name} ...", end=" ")
            
            # Convert .doc to .docx if needed
            input_path = path
            if path.suffix.lower() == '.doc':
                print("(converting .doc...) ", end="")
                temp_docx = convert_doc_to_docx(path)
                input_path = temp_docx
            
            # Pass the year extracted from original filename
            if args.json:
                convert_to_json(input_path, out_path, args.book, args.sefer, args.parshah, year, args.skip_parshah_prefix)
            else:
                reformat_docx(input_path, out_path, args.book, args.sefer, args.parshah, year, args.skip_parshah_prefix)
            print("✓ done")
            success_count += 1
            
        except Exception as e:
            print(f"⚠️ error: {e}")
        finally:
            # Clean up temp file
            if temp_docx and temp_docx.exists():
                temp_docx.unlink()

    print(f"\n✅ All done. Successfully processed {success_count}/{len(files)} file(s).")


if __name__ == "__main__":
    main()
