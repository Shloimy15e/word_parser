# Daf Mode Usage

The `--daf` mode provides a shifted heading hierarchy for processing daf-based Torah texts.

## Heading Structure

- **Heading 1**: Parent folder name (or `--book` argument if provided)
- **Heading 2**: Folder name
- **Heading 3**: Extracted from filename (e.g., `PEREK1A` → "פרק א")
- **Heading 4**: Extracted from filename letter suffix (e.g., `PEREK1A` → "חלק א") - optional

## Directory Structure Expected

```
docs/אגדות מהרי''ט/      (Parent folder → Heading 1, or use --book)
├── בבא בתרא/            (Folder → Heading 2)
│   ├── PEREK1.docx      (File → H3: "פרק א")
│   ├── PEREK1A.docx     (File → H3: "פרק א", H4: "חלק א")
│   └── PEREK2.docx      (File → H3: "פרק ב")
└── בבא מציעא/           (Folder → Heading 2)
    └── PEREK1.docx      (File → H3: "פרק א")
```

## Filename Parsing

The daf mode intelligently parses filenames to extract Hebrew headings:

| Filename | Heading 3 | Heading 4 |
|----------|-----------|-----------|
| `PEREK1.docx` | פרק א | (none) |
| `PEREK1A.docx` | פרק א | חלק א |
| `PEREK1B.docx` | פרק א | חלק ב |
| `PEREK11.docx` | פרק יא | (none) |
| `MEKOROS.docx` | מקורות | (none) |
| `MEKOROS2.docx` | מקורות ב | (none) |
| `HAKDOMO1.docx` | הקדמה א | (none) |

## Usage Examples

### Basic Usage (Using parent folder name for Heading 1)

```bash
python main.py --daf --docs "docs/אגדות מהרי''ט"
```

This will use the parent folder name (e.g., "אגדות מהרי''ט") as Heading 1.

### With Book Argument (Overrides parent folder name)

```bash
python main.py --daf --book "אגדות מהרי''ט" --docs "docs/אגדות מהרי''ט"
```

This will use "אגדות מהרי''ט" as Heading 1 for all documents, overriding the parent folder name.

### JSON Output

```bash
python main.py --daf --book "אגדות מהרי''ט" --docs "docs/אגדות מהרי''ט" --json
```

This creates JSON files with the metadata:
- `book_name_he`: Heading 2 (folder name, e.g., "בבא בתרא")
- `book_metadata.book`: Heading 1 value
- `book_metadata.section`: Heading 3 value (e.g., "פרק א")
- `book_metadata.subsection`: Heading 4 value (if present, e.g., "חלק א")
- `chunk_metadata.chunk_title`: Combined H3 + H4 (e.g., "פרק א - חלק א" or just "פרק א")

### Custom Output Directory

```bash
python main.py --daf --docs "docs/אגדות מהרי''ט" --out output_daf/
```

### Combine Mode (All files per folder into one document)

```bash
python main.py --daf --book "אגדות מהרי''ט" --docs "docs/אגדות מהרי''ט" --combine-parshah
```

This will combine all files within each folder into a single document. Each file gets its own set of headings (H1-H4) within the combined document.

**Example**: If you have:
```
בבא בתרא/
├── PEREK1.docx
├── PEREK1A.docx
└── PEREK2.docx
```

The output will be a single file: `בבא בתרא-combined.docx` containing:
- H1: אגדות מהרי''ט, H2: בבא בתרא, H3: פרק א
  - [content from PEREK1.docx]
- H1: אגדות מהרי''ט, H2: בבא בתרא, H3: פרק א, H4: חלק א
  - [content from PEREK1A.docx]
- H1: אגדות מהרי''ט, H2: בבא בתרא, H3: פרק ב
  - [content from PEREK2.docx]

## Output Structure

The output maintains a simplified directory structure (no H1 folder since it's shared):

```
output/
├── בבא בתרא/
│   ├── PEREK1-formatted.docx      (H3: פרק א)
│   ├── PEREK1A-formatted.docx     (H3: פרק א, H4: חלק א)
│   └── PEREK2-formatted.docx      (H3: פרק ב)
└── בבא מציעא/
    └── PEREK1-formatted.docx      (H3: פרק א)
```

## Comparison with Regular Mode

| Mode | H1 | H2 | H3 | H4 |
|------|----|----|----|----|
| **Regular** | --book arg | Folder name | Subfolder name | Filename |
| **Daf** | Parent folder (or --book) | Folder name | Parsed from filename | Parsed from filename (optional) |

The daf mode is essentially "one level moved" - it uses only 2 folder levels and parses the filename to extract H3 and optionally H4.

