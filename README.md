# Word Parser - Hebrew Document Formatter

A modular Python tool for reformatting Hebrew Torah documents to a standardized schema with consistent styling and structure. Features a plugin-based architecture for easy extension with new input formats, output formats, and document schemas.

## Features

- **Modular Plugin Architecture**: Easily add new input readers, output writers, and document format handlers
- **Multiple File Format Support**: Converts Hebrew documents from various formats:
  - Microsoft Word (.docx)
  - Legacy Word (.doc) - using Word COM automation
  - Adobe InDesign Markup Language (.idml)
  - DOS-encoded Hebrew text files (CP862 encoding, no extension)
  - Rich Text Format (.rtf)
- **Multiple Output Formats**:
  - Formatted Word documents (.docx)
  - Structured JSON for API integration
- **Document Format/Schema Detection**: Auto-detects or manually specify document structure:
  - Standard Torah format (parshah-based)
  - Daf/Talmud format (perek-based)
  - Siman format (halacha numbering)
  - Multi-parshah format (multiple sections in one file)
  - Letter format (correspondence structure)
- **Smart File Selection**: Automatically selects the appropriate file type from each directory (priority: .docx > .doc > .rtf > .idml > DOS)
- **Automatic Formatting**: Converts documents to a standardized format with consistent headings and styles
- **Smart Header Detection**: Intelligently identifies and removes old headers/metadata while preserving Torah content
- **Folder Structure Processing**: Batch process entire directory structures organized by Sefer/Parshah
- **Year Extraction**: Automatically extracts Hebrew year from filenames (e.g., תש״כ, תשע״ט)
- **RTL Text Handling**: Proper right-to-left formatting for Hebrew text
- **Formatting Preservation**: Maintains character formatting, spacing, and special elements like centered asterisks

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

## Architecture

The project uses a modular plugin-based architecture:

```
word_parser/
├── __init__.py              # Package exports
├── cli.py                   # Command-line interface
├── core/
│   ├── document.py          # Unified Document model
│   ├── processing.py        # Header detection, gematria, etc.
│   ├── formats.py           # DocumentFormat ABC + FormatRegistry
│   └── format_handlers.py   # Built-in format handlers
├── readers/
│   ├── base.py              # InputReader ABC + ReaderRegistry
│   ├── docx_reader.py       # .docx files
│   ├── doc_reader.py        # .doc files (via Word COM)
│   ├── idml_reader.py       # .idml files
│   ├── rtf_reader.py        # .rtf files
│   └── dos_reader.py        # DOS-encoded Hebrew files
├── writers/
│   ├── base.py              # OutputWriter ABC + WriterRegistry
│   ├── docx_writer.py       # Output to .docx
│   └── json_writer.py       # Output to .json
└── utils/
    └── __init__.py          # File handling utilities
```

### List Available Formats

```bash
python -m word_parser.cli --list-formats
```

Output:

```
Supported input formats (file types):
  DocxReader: .docx
  DocReader: .doc
  IdmlReader: .idml
  DosReader: (content-detected)
  RtfReader: .rtf

Supported output formats:
  docx: .docx
  json: .json

Supported document formats (schemas):
  standard: Standard Torah document format.
  daf: Talmud/Daf-style document format.
  multi-parshah: Multi-parshah document format.
  letter: Letter/correspondence document format.
  siman: Siman/Halacha document format.
```

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

### Daf Mode

For Talmud-style documents where filenames like "PEREK1A" map to Hebrew headings:

```bash
python main.py --daf --docs "docs/שס" --out "output"
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

### Specify Document Format

Force a specific document format/schema instead of auto-detection:

```bash
python main.py --book "שולחן ערוך" --docs "docs/אורח חיים" --out "output" --format siman
```

Available document formats:

- `standard` - Torah parshah format (default for most documents)
- `daf` - Talmud/daf format (default in --daf mode)
- `siman` - Halacha siman numbering format
- `multi-parshah` - Multiple sections in one document
- `letter` - Correspondence/letter format

**JSON Structure:**

```json
{
  "book_name_he": "ליקוטי שיחות - סדר בראשית",
  "book_name_en": "",
  "book_metadata": {
    "date": "2025-11-25",
    "book": "ליקוטי שיחות",
    "section": "בראשית"
  },
  "chunks": [
    {
      "chunk_id": 1,
      "chunk_metadata": {
        "chunk_title": "בראשית 1"
      },
      "text": "paragraph text..."
    }
  ]
}
```

## Command Line Arguments


| Argument                | Required    | Description                                                                         |
| ----------------------- | ----------- | ----------------------------------------------------------------------------------- |
| `--book`                | Yes*        | Book title (Heading 1), e.g., "ליקוטי שיחות"                             |
| `--sefer`               | Conditional | Sefer/section name (Heading 2). Auto-detected in folder structure mode              |
| `--parshah`             | Conditional | Parshah name (Heading 3). Auto-detected in folder structure mode                    |
| `--skip-parshah-prefix` | No          | Skip adding "פרשת" prefix to parshah name                                       |
| `--json`                | No          | Output as JSON structure instead of formatted Word documents                        |
| `--format`              | No          | Document format/schema (standard, daf, siman, etc.). Auto-detected if not specified |
| `--daf`                 | No          | Daf mode: Parent folder → H1, Folder → H2, Filename → H3/H4                      |
| `--docs`                | No          | Input directory (default: "docs")                                                   |
| `--out`                 | No          | Output directory (default: "output")                                                |
| `--list-formats`        | No          | List all supported input and output formats                                         |

*Not required in `--daf` mode (uses parent folder name)

## Extending the Parser

### Adding a New Input Format

Create a new reader class in `word_parser/readers/`:

```python
from pathlib import Path
from typing import List
from word_parser.core.document import Document
from word_parser.readers.base import InputReader, ReaderRegistry

@ReaderRegistry.register
class MyFormatReader(InputReader):
    """Reader for .myformat files."""
  
    @classmethod
    def get_extensions(cls) -> List[str]:
        return ['.myformat']
  
    @classmethod
    def supports_file(cls, file_path: Path) -> bool:
        return file_path.suffix.lower() == '.myformat'
  
    @classmethod
    def get_priority(cls) -> int:
        return 50  # Higher = preferred when multiple readers match
  
    def read(self, file_path: Path) -> Document:
        # Parse your format and return a Document object
        doc = Document()
        # ... parse file and add paragraphs ...
        return doc
```

Then import it in `word_parser/readers/__init__.py`:

```python
from word_parser.readers.myformat_reader import MyFormatReader
```

### Adding a New Output Format

Create a new writer class in `word_parser/writers/`:

```python
from pathlib import Path
from typing import Dict, Any
from word_parser.core.document import Document
from word_parser.writers.base import OutputWriter, WriterRegistry

@WriterRegistry.register
class MyFormatWriter(OutputWriter):
    """Writer for .myformat output."""
  
    @classmethod
    def get_format_name(cls) -> str:
        return 'myformat'  # Used with CLI selection
  
    @classmethod
    def get_extension(cls) -> str:
        return '.myformat'
  
    @classmethod
    def get_default_options(cls) -> Dict[str, Any]:
        return {'option1': 'default_value'}
  
    def write(self, doc: Document, output_path: Path, **options) -> None:
        # Write the document to output_path
        opts = {**self.get_default_options(), **options}
        # ... write file ...
```

Then import it in `word_parser/writers/__init__.py`:

```python
from word_parser.writers.myformat_writer import MyFormatWriter
```

### Adding a New Document Format (Schema)

Create a new format handler in `word_parser/core/format_handlers.py` or a new file:

```python
from typing import Dict, Any, List
from word_parser.core.document import Document
from word_parser.core.formats import DocumentFormat, FormatRegistry

@FormatRegistry.register
class MyDocumentFormat(DocumentFormat):
    """
    My custom document format.
  
    Structure:
    - H1: ...
    - H2: ...
    """
  
    @classmethod
    def get_format_name(cls) -> str:
        return 'my-format'
  
    @classmethod
    def get_priority(cls) -> int:
        return 55  # Higher = checked earlier in auto-detection
  
    @classmethod
    def get_required_context(cls) -> List[str]:
        return ['book']  # Required context keys
  
    @classmethod
    def detect(cls, doc: Document, context: Dict[str, Any]) -> bool:
        """Return True if document matches this format."""
        # Check for format-specific patterns
        for para in doc.paragraphs[:5]:
            if 'specific_pattern' in para.text:
                return True
        return False
  
    def process(self, doc: Document, context: Dict[str, Any]) -> Document:
        """Process document according to this format's rules."""
        # Extract headings, filter content, etc.
        # Example: Set H3 based on detected section markers
        for para in doc.paragraphs:
            if self._is_section_header(para):
                doc.heading3 = para.text
                break
        return doc
  
    def _is_section_header(self, para) -> bool:
        # Custom detection logic
        return para.text.startswith('סימן')
```

Then import it in `word_parser/core/__init__.py`:

```python
from word_parser.core.format_handlers import MyDocumentFormat
```

### Using the API Programmatically

```python
from pathlib import Path
from word_parser import ReaderRegistry, WriterRegistry, FormatRegistry

# Read any supported format
input_path = Path("document.docx")
reader = ReaderRegistry.get_reader_for_file(input_path)
doc = reader.read(input_path)

# Set document headings
doc.set_headings(
    h1="ליקוטי שיחות",
    h2="סדר בראשית", 
    h3="בראשית",
    h4="תש״נ"
)

# Optional: Auto-detect or get specific document format
context = {'book': 'ליקוטי שיחות', 'sefer': 'סדר בראשית'}
format_handler = FormatRegistry.detect_format(doc, context)
# Or: format_handler = FormatRegistry.get_format('daf')

if format_handler:
    doc = format_handler.process(doc, context)

# Write to any supported format
writer = WriterRegistry.get_writer('json')
writer.write(doc, Path("output.json"))
```

## Document Model

The `Document` class provides a format-agnostic representation:

```python
from word_parser.core.document import Document, Paragraph, HeadingLevel, Alignment, RunStyle

doc = Document()

# Set headings
doc.heading1 = "Book Title"
doc.heading2 = "Section"
doc.heading3 = "Chapter"
doc.heading4 = "Subsection"

# Add paragraphs
para = doc.add_paragraph("Text content")
para.format.alignment = Alignment.RIGHT
para.format.right_to_left = True

# Add formatted text runs
para.add_run("Bold text", RunStyle(bold=True))
para.add_run("Normal text")
```

## Document Formats (Schemas)

Document formats handle different structural patterns within the same file type. For example, a .docx file might contain a Torah commentary (parshah-based), Talmud notes (daf-based), or halachic content (siman-based).

### Built-in Formats


| Format          | Description          | Heading Structure                                     |
| --------------- | -------------------- | ----------------------------------------------------- |
| `standard`      | Torah parshah format | H1: Book, H2: Sefer, H3: Parshah, H4: Year            |
| `daf`           | Talmud/daf format    | H1: Book, H2: Masechet, H3: Perek, H4: Chelek         |
| `siman`         | Halacha siman format | H1: Book, H2: Section, H3: Siman, H4: Seif            |
| `multi-parshah` | Multiple sections    | H1: Book, H2: Sefer, H3: (from list items)            |
| `letter`        | Correspondence       | H1: Collection, H2: Category, H3: Recipient, H4: Date |

### Auto-Detection

When no `--format` is specified, the system attempts to auto-detect based on:

- Content patterns (e.g., "סימן א", "פרק א", list markers)
- Context (folder structure, filename patterns)
- CLI mode (--daf defaults to 'daf' format)

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
2. **Rich Text Format**

   - `.rtf` - RTF files with Unicode/Hebrew support
3. **Adobe InDesign Markup Language**

   - `.idml` - InDesign XML-based format
   - Extracts text content from Stories folder within the IDML archive
4. **DOS-Encoded Hebrew Text Files**

   - No file extension
   - CP862 (Hebrew DOS) encoding
   - Automatically detected by content analysis

### File Selection Priority

When processing directories, the parser automatically selects **only one file type** per directory, in this priority order:

1. `.docx` files (highest priority)
2. `.doc` files
3. `.rtf` files
4. `.idml` files
5. DOS-encoded files with no extension (lowest priority)

This ensures consistent processing and avoids duplicate output from the same content in different formats.

### Output Format

- **Output format**: `.docx` (standardized Word document) or `.json`
- **Note**: All input formats are converted through a unified Document model

## Formatting Preservation

The tool preserves:

- ✓ Bold, italic, underline
- ✓ Font size and color
- ✓ Centered elements (like *)
- ✓ Paragraph spacing
- ✓ Indentation
- ✓ Line spacing
- ✓ Keep together/keep with next settings

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

### Development Setup

```bash
# Clone the repository
git clone https://github.com/Shloimy15e/word_parser.git
cd word_parser

# Create virtual environment
python -m venv venv
venv\Scripts\activate  # Windows
# or: source venv/bin/activate  # Linux/Mac

# Install dependencies
pip install python-docx pywin32

# Run tests
python -m pytest
```
