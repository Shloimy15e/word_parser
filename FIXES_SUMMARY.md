# Fixes Summary

## Issues Fixed

### 1. "0" Characters Between Paragraphs (IDML Files)

**Problem:** Output documents had "0" characters appearing between paragraphs.

**Root Cause:** IDML files from InDesign contain XML elements with "0" as text content, which were being extracted and included as paragraphs.

**Fix:** Updated `extract_text_from_idml()` function to filter out "0" characters:
```python
# Filter out standalone "0" or other noise characters
if text != "0" and len(text) > 0:
    texts.append(text)
```

**Location:** `main.py` lines 217-245

---

### 2. Hebrew DOS Files Not Being Detected

**Problem:** DOS-encoded Hebrew files (without extensions) were not being found/processed.

**Root Cause:** Detection was too strict - required >10% Hebrew characters and used strict decoding.

**Fixes Applied:**

1. **Lowered threshold:** Changed from 10% to 5% Hebrew characters
2. **Increased sample size:** Read 2KB instead of 1KB for better detection
3. **Added fallback detection:** If strict CP862 decode fails, try with `errors='ignore'` and check for at least 10 Hebrew characters
4. **Better file filtering:** Added explicit check to skip directories
5. **Improved character counting:** Only count printable non-space characters

**Location:** `main.py` lines 151-196

---

### 3. Perek and Special Filename Support

**Added Features:**

1. **Proper Hebrew Gematria:** Numbers now convert correctly (e.g., 11 → יא not כ)
   - Handles ones, tens, hundreds
   - Special cases for 15 (טו) and 16 (טז) to avoid God's name

2. **MEKOROS Support:** Files named "MEKOROS" → "מקורות"

3. **Enhanced Perek Patterns:**
   - PEREK1 → פרק א
   - PEREK11 → פרק יא
   - PEREK1A → פרק א 1

**Functions:**
- `number_to_hebrew_gematria()` - Proper gematria conversion
- `extract_heading4_info()` - Unified extraction for all special patterns

---

## Testing Tools

### 1. Debug DOS Detection

**Script:** `debug_dos_detection.py`

**Usage:**
```bash
python debug_dos_detection.py <directory_path>
```

**What it does:**
- Lists all files in directory
- Shows files by type (.docx, .doc, .idml, no extension)
- Tests DOS detection on each file without extension
- Shows which files would be selected by `get_processable_files()`

**Example:**
```bash
python debug_dos_detection.py "docs/some_folder"
```

### 2. Test Perek Extraction

**Script:** `test_perek.py`

**Usage:**
```bash
python test_perek.py
```

**Tests:**
- Number to Hebrew gematria conversion (1→א, 11→יא, etc.)
- Perek pattern extraction (PEREK1, PEREK1A, etc.)
- MEKOROS detection
- Year extraction (optional)

---

## Technical Details

### DOS File Detection Algorithm

```
1. Check: File has NO extension? (if yes, continue; if no, reject)
2. Check: Is a file, not directory? (if yes, continue; if no, reject)
3. Read first 2KB of file
4. Try to decode as CP862 (strict):
   a. Count Hebrew characters (U+0590 to U+05FF)
   b. Count total printable non-space characters
   c. If Hebrew > 5% of total: ACCEPT as DOS
5. If strict decode fails, try CP862 with errors='ignore':
   a. Count Hebrew characters
   b. If ≥ 10 Hebrew characters: ACCEPT as DOS
6. Otherwise: REJECT
```

### IDML Text Extraction Algorithm

```
1. Open IDML as ZIP archive
2. Find all Stories/*.xml files
3. Parse each XML file
4. For each element:
   a. Extract element.text if present
   b. Extract element.tail if present
   c. Filter out "0" and empty strings
   d. Strip whitespace
   e. Add to texts list
5. Return all extracted texts
```

### Hebrew Gematria Conversion

```
Ones:    א ב ג ד ה ו ז ח ט (1-9)
Tens:    י כ ל מ נ ס ע פ צ (10-90, step 10)
Hundreds: ק ר ש ת (100-400, step 100)

Special cases:
- 15 = טו (not יה)
- 16 = טז (not יו)

Examples:
- 1 = א
- 11 = יא (10+1)
- 21 = כא (20+1)
- 111 = קיא (100+10+1)
```

---

## Usage Examples

### Process IDML Files

```bash
python main.py --book "אגדות מהריט" --docs "docs/idml_files" --out "output"
```

Files like `PEREK1.idml` will be processed with:
- Heading 4: "פרק א"
- No "0" characters in output

### Process DOS Files

```bash
python main.py --book "ספר" --docs "docs/dos_files" --out "output"
```

DOS files (no extension) will be auto-detected if they contain Hebrew text (CP862).

### Debug DOS Detection

```bash
# Check what files are detected in a directory
python debug_dos_detection.py "docs/problematic_folder"

# Output shows:
# - All files by type
# - DOS detection results for each file
# - Which file type would be selected
```

---

## Troubleshooting

### DOS Files Still Not Detected

1. **Check file has no extension:**
   ```bash
   ls -la  # On Windows: dir
   # Look for files without dots
   ```

2. **Run debug script:**
   ```bash
   python debug_dos_detection.py "your_folder"
   ```

3. **Check encoding manually:**
   ```bash
   python
   >>> with open('yourfile', 'rb') as f:
   ...     data = f.read(100)
   >>> data.decode('cp862')  # Should show Hebrew text
   ```

4. **Lower threshold if needed:** Edit `main.py` line 178:
   ```python
   # Change from 0.05 (5%) to 0.01 (1%)
   if total_chars > 0 and hebrew_chars > total_chars * 0.01:
   ```

### "0" Still Appearing

1. **Check source file:** Open original IDML in InDesign
2. **Look for zero characters:** Search for "0" in document
3. **Additional filtering:** Edit `main.py` to filter more patterns:
   ```python
   # Add more filters
   if text not in ["0", "00", "000"] and len(text) > 0:
       texts.append(text)
   ```

### MEKOROS Not Recognized

- Ensure filename is exactly "MEKOROS" (case-insensitive OK)
- Check: `extract_heading4_info("MEKOROS")` returns `"מקורות"`

---

## Summary of Changes

| File | Lines Changed | Description |
|------|---------------|-------------|
| `main.py` | 151-196 | Enhanced DOS detection |
| `main.py` | 217-245 | Filter "0" from IDML |
| `main.py` | 319-350 | Hebrew gematria function |
| `main.py` | 352-375 | Heading 4 extraction |
| `debug_dos_detection.py` | NEW | DOS detection debugger |
| `test_perek.py` | Updated | Test gematria & patterns |

---

## Version

**Date:** November 20, 2024  
**Changes:** Bug fixes for IDML "0" issue and DOS detection  
**Status:** ✅ Ready for testing

