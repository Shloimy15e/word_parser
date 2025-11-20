# Quick Start Guide - Multi-Format Support

## What's New?

The Word Parser now accepts **4 file formats**:

| Format | Extension | Notes |
|--------|-----------|-------|
| Word Modern | `.docx` | Direct processing (fastest) |
| Word Legacy | `.doc` | Requires Windows + MS Word |
| InDesign | `.idml` | Cross-platform, text only |
| DOS Hebrew | (none) | CP862 encoding, auto-detected |

## Quick Usage

### Basic Usage (unchanged)
```bash
python main.py --book "ליקוטי שיחות" --docs "docs/סדר בראשית" --out "output"
```

**New behavior:**
- ✅ Automatically processes all supported formats
- ✅ Selects best format per directory (priority: .docx > .doc > .idml > DOS)
- ✅ Works with existing commands (no changes needed!)

## File Priority

When multiple formats exist in one directory:

```
Priority 1: .docx files    ← Processed first
Priority 2: .doc files     ← Only if no .docx
Priority 3: .idml files    ← Only if no .docx or .doc
Priority 4: DOS files      ← Only if no other formats
```

**Only ONE format per directory is processed** to avoid duplicates.

## Format-Specific Notes

### .docx Files
- ✅ Works on all platforms
- ✅ Fastest processing
- ✅ No conversion needed
- **Recommended format**

### .doc Files  
- ⚠️ Windows only
- ⚠️ Requires Microsoft Word installed
- Converted to .docx automatically
- Slower than .docx

### .idml Files (NEW!)
- ✅ Works on all platforms
- ✅ No extra software needed
- ⚠️ Text only (no images/complex layouts)
- Extract from Adobe InDesign: File → Export → IDML

### DOS Files (NEW!)
- ✅ Works on all platforms
- ✅ No file extension required
- Must be CP862 encoded
- Must contain Hebrew text (>10%)

## Common Scenarios

### Scenario 1: All My Files Are .docx
```bash
# Works exactly as before
python main.py --book "ליקוטי שיחות" --docs "docs" --out "output"
```

### Scenario 2: Mix of .docx and .idml
```bash
# Processes .docx files (higher priority)
# Ignores .idml files automatically
python main.py --book "ליקוטי שיחות" --docs "docs" --out "output"
```

To process .idml instead, remove or move .docx files.

### Scenario 3: Only .idml Files
```bash
# Works automatically - no special flags needed
python main.py --book "ליקוטי שיחות" --docs "docs_idml" --out "output"
```

### Scenario 4: Legacy DOS Files
```bash
# Works if files have no extension and are CP862 encoded
python main.py --book "ליקוטי שיחות" --docs "docs_dos" --out "output"
```

### Scenario 5: Multi-Parshah with IDML
```bash
python main.py --book "ליקוטי שיחות" --sefer "סדר בראשית" \
  --docs "combined.idml" --out "output" --multi-parshah
```

## Testing Your Setup

Run the test suite:
```bash
cd C:\Users\shloi\Desktop\word_parser
.\venv\Scripts\Activate.ps1
python test_new_formats.py
```

Expected output:
```
============================================================
Testing New File Format Support
============================================================

Testing DOS file detection...
  [OK] DOS file detection tests passed

Testing DOS to DOCX conversion...
  [OK] DOS to DOCX conversion tests passed

Testing file selection priority...
  [OK] File priority selection tests passed

============================================================
All tests passed!
============================================================
```

## Troubleshooting

### Problem: IDML files are ignored

**Solution:**
- Check if .docx or .doc files exist in the same directory
- Remove higher-priority formats or move to separate directory

### Problem: DOS files not detected

**Solutions:**
1. Verify file has **no extension**
2. Check encoding: `file yourfile` (should show CP862 or similar)
3. Ensure file contains Hebrew characters
4. Remove other file formats from directory

### Problem: .doc conversion fails

**Solutions:**
1. Install pywin32: `pip install pywin32`
2. Ensure Microsoft Word is installed
3. Try on Windows machine
4. Or convert .doc → .docx manually in Word first

### Problem: "No supported files found"

**Solutions:**
1. Check file extensions (must be .docx, .doc, .idml, or no extension for DOS)
2. Verify directory path is correct
3. Check that files aren't in subdirectories (unless using folder structure mode)

## File Organization Tips

### Best Practice: Separate by Format
```
project/
  ├── docs_docx/          ← Modern Word files
  ├── docs_idml/          ← InDesign exports
  └── docs_dos/           ← Legacy DOS files
```

Process each separately:
```bash
python main.py --book "ליקוטי שיחות" --docs "docs_docx" --out "output"
python main.py --book "ליקוטי שיחות" --docs "docs_idml" --out "output_idml"
```

### Alternative: Convert Everything to .docx First
```bash
# Process once to convert all formats
python main.py --book "ליקוטי שיחות" --docs "mixed_formats" --out "converted"

# Then process the .docx files
python main.py --book "ליקוטי שיחות" --docs "converted" --out "final"
```

## Performance Tips

### Fastest to Slowest
1. `.docx` - Instant (no conversion)
2. `.idml` - Fast (ZIP extraction only)
3. `DOS` - Fast (text conversion only)
4. `.doc` - Slow (requires Word COM automation)

### Optimization
- Pre-convert all files to .docx for maximum speed
- Keep .docx versions after first processing
- Archive original formats separately

## Getting Help

1. **File format issues**: Read `FILE_FORMATS.md`
2. **Technical details**: Read `CHANGELOG.md`
3. **Implementation info**: Read `IMPLEMENTATION_SUMMARY.md`
4. **General usage**: Read `README.md`

## Quick Reference Card

```
┌─────────────────────────────────────────────────────────┐
│ WORD PARSER - Multi-Format Support                      │
├─────────────────────────────────────────────────────────┤
│                                                          │
│ SUPPORTED FORMATS:                                       │
│   ✓ .docx  (all platforms)                              │
│   ✓ .doc   (Windows + Word)                             │
│   ✓ .idml  (all platforms)                              │
│   ✓ DOS    (CP862, no extension, all platforms)         │
│                                                          │
│ FILE PRIORITY:                                           │
│   1. .docx  (highest)                                    │
│   2. .doc                                                │
│   3. .idml                                               │
│   4. DOS    (lowest)                                     │
│                                                          │
│ BASIC USAGE:                                             │
│   python main.py --book "BOOK" \                         │
│                  --docs "INPUT" \                        │
│                  --out "OUTPUT"                          │
│                                                          │
│ TEST INSTALLATION:                                       │
│   python test_new_formats.py                            │
│                                                          │
│ NO CHANGES TO EXISTING COMMANDS REQUIRED!                │
│ All formats work automatically.                          │
│                                                          │
└─────────────────────────────────────────────────────────┘
```

---

**Version:** 2.0.0  
**Date:** November 20, 2024  
**Status:** Ready for use

