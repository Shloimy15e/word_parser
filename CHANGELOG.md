# Changelog

## [Unreleased] - 2024-11-20

### Added

#### 1. **IDML File Support**
- Added support for Adobe InDesign Markup Language (.idml) files
- Implemented `extract_text_from_idml()` function to extract text from IDML ZIP archives
- Implemented `convert_idml_to_docx()` function to convert extracted text to .docx format
- IDML files are automatically detected and processed through the pipeline

#### 2. **DOS-Encoded Hebrew Text File Support**
- Added support for Hebrew DOS-encoded text files (CP862 encoding)
- Implemented `is_dos_encoded_file()` function to detect DOS-encoded files without extensions
- Implemented `convert_dos_to_docx()` function to convert CP862-encoded text to .docx
- Automatic detection based on:
  - No file extension
  - CP862 decodable content
  - >10% Hebrew character content

#### 3. **Smart File Selection**
- Implemented `get_processable_files()` function for intelligent file selection
- Only processes ONE file type per directory based on priority:
  1. `.docx` files (highest priority)
  2. `.doc` files
  3. `.idml` files
  4. DOS-encoded files with no extension (lowest priority)
- Prevents duplicate processing of the same content in different formats

#### 4. **Unified Conversion Function**
- Implemented `convert_to_docx()` function that handles all file formats
- Returns tuple: (path_to_docx, needs_cleanup)
- Automatically determines conversion method based on file extension
- Manages temporary file lifecycle

### Changed

#### File Processing Pipeline
- Updated `combine_parshah_docs()` to use new file selection and conversion functions
- Updated folder structure processing to use `get_processable_files()`
- Updated single folder mode to use `get_processable_files()`
- Updated multi-parshah mode to use `convert_to_docx()`
- All conversion paths now use unified `convert_to_docx()` function

#### Display Names
- Improved file display names to handle files without extensions
- Files without extensions now display their full name instead of stem

### Technical Details

#### New Dependencies
- `zipfile` - For IDML archive extraction (Python standard library)
- `xml.etree.ElementTree` - For XML parsing (Python standard library)

#### File Format Support Matrix

| Format | Extension | Encoding | Detection Method | Conversion Method |
|--------|-----------|----------|------------------|-------------------|
| Modern Word | .docx | UTF-8 | Extension | Direct (no conversion) |
| Legacy Word | .doc | Various | Extension | Word COM automation |
| InDesign | .idml | UTF-8 | Extension | ZIP + XML extraction |
| DOS Text | None | CP862 | Content analysis | Encoding conversion |

#### Encoding Details
- **DOS files**: CP862 (Hebrew DOS codepage) → UTF-8
- **IDML files**: XML UTF-8 → UTF-8 (no encoding conversion needed)
- All output maintained as UTF-8 in .docx format

### Documentation

#### Updated README.md
- Added comprehensive file format support section
- Documented file selection priority algorithm
- Added IDML processing details
- Added DOS-encoded file processing details
- Updated installation requirements
- Clarified platform-specific requirements (Windows for .doc only)

### Notes

#### Backward Compatibility
- ✅ All existing functionality preserved
- ✅ Existing command-line arguments unchanged
- ✅ Output format unchanged
- ✅ No breaking changes

#### Testing Recommendations
1. Test with sample .idml files containing Hebrew text
2. Test with DOS-encoded files (CP862)
3. Test directory with mixed file types to verify priority selection
4. Test multi-parshah mode with different file formats
5. Verify temporary file cleanup after processing

#### Known Limitations
- IDML text extraction may not preserve complex formatting from InDesign
- DOS file detection requires >10% Hebrew characters (may miss files with mostly English)
- .doc conversion still requires Windows + Microsoft Word
- IDML processing extracts only text content, not images or complex layouts

### Future Enhancements
- Add support for RTF files
- Add support for ODT (OpenDocument Text) files
- Improve IDML formatting preservation
- Add configurable DOS encoding detection threshold
- Add support for other Hebrew encodings (ISO-8859-8, Windows-1255)

