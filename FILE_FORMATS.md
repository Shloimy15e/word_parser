# File Format Support Guide

## Supported Input Formats

The Word Parser now supports four different file formats for Hebrew documents:

### 1. Microsoft Word - Modern (.docx)

**Extension:** `.docx`

**Description:** The modern Word document format (Office 2007+)

**Processing:** Direct processing, no conversion required

**Requirements:** None

**Example:**
```
פרשת בראשית תשנט.docx
```

### 2. Microsoft Word - Legacy (.doc)

**Extension:** `.doc`

**Description:** Legacy Word document format (Office 97-2003)

**Processing:** Converted to .docx using Word COM automation

**Requirements:** 
- Windows OS
- Microsoft Word installed

**Example:**
```
פרשת בראשית תשנט.doc
```

### 3. Adobe InDesign Markup Language (.idml)

**Extension:** `.idml`

**Description:** XML-based format exported from Adobe InDesign

**Processing:** 
1. Extracts ZIP archive
2. Parses Stories/*.xml files
3. Extracts text content
4. Creates temporary .docx

**Requirements:** None (works on all platforms)

**Limitations:**
- Only text content is extracted
- Complex layouts and formatting may not be preserved
- Images and graphics are not included

**Example:**
```
פרשת בראשית תשנט.idml
```

**Note:** To export IDML from InDesign:
1. Open document in InDesign
2. File → Export
3. Choose "InDesign Markup (IDML)" format

### 4. DOS-Encoded Hebrew Text Files

**Extension:** None (no extension)

**Description:** Plain text files with Hebrew DOS (CP862) encoding

**Processing:**
1. Detects files without extensions
2. Attempts CP862 decoding
3. Validates Hebrew content (>10% Hebrew characters)
4. Creates temporary .docx with RTL formatting

**Requirements:** None (works on all platforms)

**Detection Criteria:**
- File has no extension
- Content can be decoded as CP862
- At least 10% of characters are Hebrew (U+0590 to U+05FF)

**Example:**
```
פרשת_בראשית_תשנט
```

**Common Sources:**
- Old DOS-based word processors
- Legacy backup systems
- Scanned OCR text from DOS systems
- Hebrew DOS text editors

## File Selection Priority

When processing directories, the parser automatically selects **only one file type** per directory. If multiple formats exist, priority is:

1. **`.docx`** (highest priority) - Most reliable, no conversion needed
2. **`.doc`** - Requires conversion but well-supported
3. **`.idml`** - Good for InDesign exports
4. **DOS files** (lowest priority) - Fallback for legacy content

### Example Scenario

If a directory contains:
```
docs/פרשת בראשית/
  ├── תשנט.docx          ← THIS ONE WILL BE PROCESSED
  ├── תשנט.doc           ← Ignored (lower priority)
  ├── תשנט.idml          ← Ignored (lower priority)
  └── תשנט               ← Ignored (lowest priority)
```

Only `תשנט.docx` will be processed because it has the highest priority.

### Rationale

This prevents:
- Duplicate output from the same content
- Processing conflicts
- Inconsistent formatting

If you want to process a different format:
1. Move or rename the higher-priority files
2. Or process them separately

## Usage Examples

### Process Mixed Format Directory

```bash
python main.py --book "ליקוטי שיחות" --docs "docs/סדר בראשית" --out "output"
```

The parser will automatically:
- Scan each parshah subdirectory
- Identify available file formats
- Select highest priority format
- Convert to .docx if needed
- Process through standard pipeline

### Process IDML Files Only

To process only IDML files, ensure no .docx or .doc files exist in the directories:

```bash
# Move other formats temporarily
python main.py --book "ליקוטי שיחות" --docs "docs_idml" --out "output"
```

### Process DOS-Encoded Files

DOS files are processed automatically when no other formats exist:

```bash
python main.py --book "ליקוטי שיחות" --docs "docs_dos" --out "output"
```

### Multi-Parshah Mode with IDML

```bash
python main.py --book "ליקוטי שיחות" --sefer "סדר בראשית" --docs "combined.idml" --out "output" --multi-parshah
```

## Troubleshooting

### IDML Files Not Processing

**Problem:** IDML file is ignored

**Solutions:**
1. Check if .docx or .doc files exist in the same directory
2. Verify the .idml file is valid (can be opened as a ZIP archive)
3. Ensure the IDML contains Stories/*.xml files with text content

### DOS Files Not Detected

**Problem:** DOS-encoded file is skipped

**Solutions:**
1. Verify the file has no extension
2. Check if the file contains Hebrew characters (>10% of content)
3. Try manually decoding: `iconv -f cp862 -t utf8 yourfile`
4. Ensure no other file formats exist in the directory

### .doc Conversion Fails

**Problem:** "pywin32 not installed" or "Cannot convert .doc files"

**Solutions:**
1. Install pywin32: `pip install pywin32`
2. Ensure Microsoft Word is installed (Windows only)
3. Try converting .doc to .docx manually in Word first
4. Check if Word automation is enabled in Windows

### IDML Text Extraction Empty

**Problem:** "No text content found in IDML file"

**Solutions:**
1. Open IDML in InDesign and re-export
2. Check if text is in locked layers
3. Verify text is not threaded outside the document
4. Try exporting as .docx from InDesign instead

## File Format Recommendations

### Best Format for Each Use Case

| Use Case | Recommended Format | Reason |
|----------|-------------------|--------|
| New documents | `.docx` | Native format, no conversion |
| Legacy archives | `.doc` → `.docx` | Convert once, process many times |
| InDesign layouts | `.idml` | Preserves text, easier than PDF |
| DOS archives | DOS → `.docx` | Convert once for future use |
| Long-term storage | `.docx` | Best supported, most portable |

### Converting for Optimal Performance

For best results, pre-convert all files to .docx:

1. **From .doc:**
   - Open in Word
   - Save As → .docx format
   
2. **From .idml:**
   - Use this parser once
   - Keep generated .docx for future processing
   
3. **From DOS:**
   - Use this parser once
   - Keep generated .docx for future processing

## Technical Details

### Encoding Table

| Format | Source Encoding | Target Encoding | Method |
|--------|----------------|-----------------|--------|
| .docx | UTF-8 | UTF-8 | None (direct) |
| .doc | Various | UTF-8 | Word COM |
| .idml | UTF-8 (XML) | UTF-8 | ZIP + XML parse |
| DOS | CP862 | UTF-8 | Python decode |

### CP862 Character Set

CP862 (Hebrew DOS) includes:
- Hebrew letters (א-ת)
- Nikud (vowel points)
- Hebrew punctuation
- Latin letters
- DOS box-drawing characters

### IDML Structure

IDML files are ZIP archives containing:
```
document.idml/
  ├── Stories/
  │   ├── Story_u123.xml    ← Text content here
  │   └── Story_u124.xml
  ├── Spreads/
  └── META-INF/
```

The parser extracts text from all Story_*.xml files.

## Platform Compatibility

| Format | Windows | macOS | Linux |
|--------|---------|-------|-------|
| .docx | ✅ | ✅ | ✅ |
| .doc | ✅ (with Word) | ❌ | ❌ |
| .idml | ✅ | ✅ | ✅ |
| DOS | ✅ | ✅ | ✅ |

**Note:** Only .doc conversion requires Windows + Microsoft Word. All other formats work cross-platform.

