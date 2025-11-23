# DAF Mode Summary

## Heading Structure

### In DAF Mode (--daf-mode flag):
- **Heading 1**: Book/Collection (from parent folder name when using folder structure)
  - Example: "אגדות מהריט", "תלמוד בבלי"
  - Added once at the beginning, or when it changes
  - **Optional when using folder structure** - automatically derived from parent folder
  
- **Heading 2**: Masechet/Tractate (from folder name)
  - Example: "ביצה", "שבת", etc.
  - Added once at the beginning, or when it changes
  
- **Heading 3**: דף [page number] (extracted from IDML pages)
  - Example: "דף טו", "דף יא", "דף יג"
  - Added automatically from IDML page markers
  
- **Heading 4**: עמוד [column] (extracted from IDML pages)
  - Example: "עמוד א", "עמוד ב"
  - Currently: "עמוד א" is added by default when דף changes

### In Normal Mode (without --daf-mode):
- **Heading 1**: Book (from --book, required)
- **Heading 2**: Sefer (from folder/--sefer)
- **Heading 3**: Parshah (from subfolder/--parshah)
- **Heading 4**: Year or file identifier

## Key Features

1. **Headings only added when they change**: 
   - In combined mode, Heading 1, 2, 3, and 4 are only inserted when their values change
   - This prevents repetition across multiple files

2. **"פרק" paragraphs filtered out in DAF mode**:
   - Chapter designations like "פרק ראשון" are removed since דף/עמוד provide the structure

3. **Page-aware IDML extraction**:
   - Parses IDML spread files to get page names
   - Maps text frames to pages using threading information
   - Inserts page markers at correct boundaries

4. **Automatic folder structure detection**:
   - In DAF mode with folder structure, parent folder → Heading 1, subfolder → Heading 2
   - No need to specify --book when using folder structure in DAF mode

## Folder Structure for DAF Mode

```
אגדות מהריט/          ← Heading 1 (auto-detected)
  ├── ביצה/            ← Heading 2 (masechet)
  │   ├── PEREK1.idml  ← Contains דף/עמוד (Heading 3/4)
  │   ├── PEREK2.idml
  │   └── ...
  ├── שבת/
  └── ...
```

## Usage Examples

### Folder structure with DAF mode (--book NOT needed):
```bash
python main.py --docs "אגדות מהריט" --out "output" --daf-mode
```
This will use:
- "אגדות מהריט" (parent folder) as Heading 1
- Subfolder names (e.g., "ביצה") as Heading 2
- IDML page markers as Heading 3 (דף) and Heading 4 (עמוד)

### Combined mode with DAF (--book NOT needed):
```bash
python main.py --docs "אגדות מהריט" --out "output" --combine-parshah --daf-mode
```
This combines all files per masechet into one document with proper heading hierarchy.

### Single file/folder with explicit parameters (--book required):
```bash
python main.py --book "תלמוד בבלי" --sefer "ביצה" --parshah "PEREK1" --docs "input_folder" --out "output" --daf-mode
```

