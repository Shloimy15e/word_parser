# Word Parser - Hebrew Document Formatter

A Python tool for reformatting Hebrew Torah documents (Word files) to a standardized schema with consistent styling and structure.

## Features

- **Multiple File Format Support**: Converts Hebrew documents from various formats:
  - Microsoft Word (.docx)
  - Legacy Word (.doc) - using Word COM automation
  - Adobe InDesign Markup Language (.idml)
  - DOS-encoded Hebrew text files (CP862 encoding, no extension)
- **Smart File Selection**: Automatically selects the appropriate file type from each directory (priority: .docx > .doc > .idml > DOS)
- **Automatic Formatting**: Converts documents to a standardized format with consistent headings and styles
- **Smart Header Detection**: Intelligently identifies and removes old headers/metadata while preserving Torah content
- **Folder Structure Processing**: Batch process entire directory structures organized by Sefer/Parshah
- **Year Extraction**: Automatically extracts Hebrew year from filenames (e.g., תש״כ, תשע״ט)
- **RTL Text Handling**: Proper right-to-left formatting for Hebrew text
- **Formatting Preservation**: Maintains character formatting, spacing, and special elements like centered asterisks
- **JSON Export**: Output structured JSON files with each paragraph as a chunk for API integration

## Installation

### Requirements

- Python 3.6+
- Windows OS (required only for .doc file conversion via COM automation)
- Microsoft Word (required only for .doc file conversion)

**Note**: `.idml` and DOS-encoded file support works on all platforms without additional dependencies.

### Install Dependencies

```bash
pip install python-docx pywin32
```

**Dependencies:**
- `python-docx` - For reading and writing .docx files
- `pywin32` - For .doc file conversion (Windows only)

## Usage

### Folder Structure Mode (Recommended)

Process an entire directory structure where folder names represent Sefer and subfolder names represent Parshah:

```bash
python main.py --book "ליקוטי שיחות" --docs "docs/סדר דברים" --out "output"
```

**Directory structure example:**
```
docs/
  סדר דברים/
    שבת שובה/
      תשנ״ט.docx
      תש״ס.docx
    האזינו/
      תשס״א.docx
```

Output will mirror the structure:
```
output/
  סדר דברים/
    שבת שובה/
      תשנ״ט-formatted.docx
      תש״ס-formatted.docx
```

### Single Folder Mode

Process all files in one folder with explicit Sefer and Parshah:

```bash
python main.py --book "ליקוטי שיחות" --sefer "סדר בראשית" --parshah "בראשית" --docs "docs" --out "output"
```

### Skip Parshah Prefix

By default, Heading 3 shows "פרשת [parshah name]". To use just the parshah name:

```bash
python main.py --book "ליקוטי שיחות" --docs "docs/מועדים" --out "output" --skip-parshah-prefix
```

### JSON Output Mode

Output structured JSON files instead of Word documents:

```bash
python main.py --book "ליקוטי שיחות" --docs "docs/סדר בראשית" --out "output" --json
```

**JSON Structure:**
```json
{
  "book_name_he": "פרשת בראשית",
  "book_name_en": "בראשית",
  "book_metadata": {
    "sefer_he": "סדר בראשית",
    "sefer_en": "סדר בראשית",
    "collection_he": "ליקוטי שיחות",
    "collection_en": "ליקוטי שיחות",
    "year_he": "תש״נ",
    "year_en": "תש״נ",
    "source": "Word Document Conversion"
  },
  "chunks": [
    {
      "chunk_id": 1,
      "chunk_metadata": {
        "chunk_title": "פרשת בראשית - קטע 1",
        "sefer": "סדר בראשית",
        "parshah": "פרשת בראשית",
        "year": "תש״נ",
        "collection": "ליקוטי שיחות"
      },
      "text": "paragraph text..."
    }
  ]
}
```

Each paragraph becomes a chunk with metadata for easy API integration.

## Command Line Arguments

| Argument | Required | Description |
|----------|----------|-------------|
| `--book` | Yes | Book title (Heading 1), e.g., "ליקוטי שיחות" |
| `--sefer` | Conditional | Sefer/section name (Heading 2). Required in single folder mode, auto-detected in folder structure mode |
| `--parshah` | Conditional | Parshah name (Heading 3). Required in single folder mode, auto-detected in folder structure mode |
| `--skip-parshah-prefix` | No | Skip adding "פרשת" prefix to parshah name |
| `--json` | No | Output as JSON structure instead of formatted Word documents |
| `--docs` | No | Input directory (default: "docs") |
| `--out` | No | Output directory (default: "output") |

## Document Structure

### Output Document Format

Each processed document contains:

1. **Heading 1** (Book): Large, dark blue (16pt)
2. **Heading 2** (Sefer): Medium blue (13pt)  
3. **Heading 3** (Parshah): Dark navy (12pt)
4. **Heading 4** (Year): Medium blue (11pt)
5. **Body Text**: Formatted Torah content with proper spacing

### Header Detection

The tool automatically removes old metadata/headers by detecting:
- Date patterns (e.g., "מוצ״ש", "מוצאי שבת")
- Format patterns (e.g., "ב״ה", "דברות", "סדר")
- Year patterns (e.g., "תש״", "שנת")
- Location patterns (e.g., "בעיר", "בבית")
- Short title-like lines (< 25 chars without punctuation)

Content detection starts when encountering:
- Paragraphs ≥60 characters
- Text with Torah brackets `[ ]`

## Year Extraction

The tool extracts Hebrew years from filenames using these patterns:
- Full years: תש״כ, תשע״ט, תשנ״ט
- Abbreviated: תשכח, תשעט
- With separators: פרשת בראשית - תש״ס

## File Support

### Input Formats

The parser supports the following file formats:

1. **Microsoft Word Documents**
   - `.docx` - Modern Word format (direct support)
   - `.doc` - Legacy Word format (requires Microsoft Word installed for COM automation)

2. **Adobe InDesign Markup Language**
   - `.idml` - InDesign XML-based format
   - Extracts text content from Stories folder within the IDML archive

3. **DOS-Encoded Hebrew Text Files**
   - No file extension
   - CP862 (Hebrew DOS) encoding
   - Automatically detected by content analysis

### File Selection Priority

When processing directories, the parser automatically selects **only one file type** per directory, in this priority order:
1. `.docx` files (highest priority)
2. `.doc` files
3. `.idml` files
4. DOS-encoded files with no extension (lowest priority)

This ensures consistent processing and avoids duplicate output from the same content in different formats.

### Output Format

- **Output format**: `.docx` (standardized Word document)
- **Note**: All input formats are converted to .docx for processing

## Formatting Preservation

The tool preserves:
- ✓ Bold, italic, underline
- ✓ Font size and color
- ✓ Centered elements (like *)
- ✓ Paragraph spacing
- ✓ Indentation
- ✓ Line spacing
- ✓ Keep together/keep with next settings

## Examples

### Process all Sefarim in folder structure:
```bash
python main.py --book "ליקוטי שיחות" --docs "docs" --out "output"
```

### Process specific Sefer:
```bash
python main.py --book "ליקוטי שיחות" --docs "docs/סדר בראשית" --out "output"
```

### Process Moadim without "פרשת" prefix:
```bash
python main.py --book "ליקוטי שיחות" --docs "docs/מועדים" --out "output" --skip-parshah-prefix
```

### Export to JSON format:
```bash
python main.py --book "ליקוטי שיחות" --docs "docs/סדר בראשית" --out "output" --json
```

## File Format Processing

### IDML Files (Adobe InDesign)

IDML (InDesign Markup Language) files are ZIP archives containing XML files. The parser:
1. Extracts the IDML archive
2. Locates Story files (Stories/*.xml) containing text content
3. Parses XML to extract all text elements
4. Creates a temporary .docx file with the extracted text
5. Processes the .docx file through the standard pipeline

**Example:**
```bash
python main.py --book "ליקוטי שיחות" --docs "docs/סדר בראשית" --out "output"
# Will automatically process any .idml files found
```

### DOS-Encoded Hebrew Files

DOS-encoded files use the CP862 (Hebrew DOS) character encoding and typically have no file extension. The parser:
1. Detects files without extensions
2. Attempts to decode using CP862 encoding
3. Validates Hebrew character content (must be >10% Hebrew characters)
4. Converts to UTF-8 and creates a temporary .docx file
5. Processes through the standard pipeline

**Note**: DOS files are detected automatically - no special flags required.

## Output Formats

### Word Document (.docx)

Default output format with standardized styling:
- Hierarchical headings (Book → Sefer → Parshah → Year)
- Consistent colors and fonts
- Proper RTL text alignment
- Preserved formatting (bold, italic, spacing)

**File naming:** `[original-filename]-formatted.docx`

### JSON Structure

When using `--json` flag, outputs structured JSON files:

**File naming:** `[original-filename].json` (one JSON per input Word file)

**Structure:**
- `book_name_he/en`: Parshah name (the "book" in the chunk structure)
- `book_metadata`: Contains sefer, collection (book), year, and source info
- `chunks`: Array of paragraph chunks with:
  - `chunk_id`: Sequential number
  - `chunk_metadata`: Title and categorization info
  - `text`: The actual paragraph text

Ideal for APIs, search engines, and database imports.

## Error Handling

- Files without extractable years are skipped with warnings
- .doc files that fail conversion are skipped
- Individual file errors don't stop batch processing
- Detailed progress output for all operations

## Technical Details

### Style Configuration

- **Heading 1**: 16pt, RGB(47, 84, 150)
- **Heading 2**: 13pt, RGB(68, 114, 196)
- **Heading 3**: 12pt, RGB(31, 55, 99)
- **Heading 4**: 11pt, RGB(47, 84, 150)
- **Normal**: 12pt, 1.15 line spacing

All headings are right-aligned with RTL direction enabled for proper Hebrew display.

## License

MIT

## Contributing

Contributions welcome! Please open an issue or submit a pull request.
