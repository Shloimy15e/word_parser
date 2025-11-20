#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Test script for new file format support
Tests IDML and DOS-encoded file handling
"""

import tempfile
import zipfile
from pathlib import Path
from main import (
    is_dos_encoded_file,
    convert_dos_to_docx,
    extract_text_from_idml,
    get_processable_files
)

def test_dos_file_detection():
    """Test DOS-encoded file detection"""
    print("Testing DOS file detection...")
    
    # Create a temporary DOS-encoded file
    temp_dir = Path(tempfile.mkdtemp())
    dos_file = temp_dir / "test_hebrew"
    
    # Write Hebrew text in CP862 encoding
    hebrew_text = "זה טקסט עברי בקידוד DOS\nשורה שנייה\nשורה שלישית"
    with open(dos_file, 'wb') as f:
        f.write(hebrew_text.encode('cp862'))
    
    # Test detection
    result = is_dos_encoded_file(dos_file)
    print(f"  DOS file detected: {result}")
    assert result == True, "Failed to detect DOS-encoded file"
    
    # Test with a regular file (should fail)
    regular_file = temp_dir / "test.txt"
    with open(regular_file, 'w', encoding='utf-8') as f:
        f.write("This is English text")
    
    result2 = is_dos_encoded_file(regular_file)
    print(f"  Regular file detected as DOS: {result2}")
    assert result2 == False, "Incorrectly detected regular file as DOS"
    
    # Cleanup
    dos_file.unlink()
    regular_file.unlink()
    temp_dir.rmdir()
    
    print("  [OK] DOS file detection tests passed\n")


def test_dos_conversion():
    """Test DOS to DOCX conversion"""
    print("Testing DOS to DOCX conversion...")
    
    # Create a temporary DOS-encoded file
    temp_dir = Path(tempfile.mkdtemp())
    dos_file = temp_dir / "test_hebrew"
    
    # Write Hebrew text in CP862 encoding
    hebrew_text = "זה טקסט עברי בקידוד DOS\nשורה שנייה של טקסט\nשורה שלישית"
    with open(dos_file, 'wb') as f:
        f.write(hebrew_text.encode('cp862'))
    
    try:
        # Convert to DOCX
        docx_path = convert_dos_to_docx(dos_file)
        print(f"  Created temp DOCX: {docx_path}")
        
        # Verify DOCX was created
        assert docx_path.exists(), "DOCX file was not created"
        assert docx_path.suffix == '.docx', "Output is not a DOCX file"
        
        # Read the DOCX to verify content
        from docx import Document
        doc = Document(docx_path)
        paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
        
        print(f"  Extracted {len(paragraphs)} paragraphs")
        # Don't print Hebrew text to avoid console encoding issues
        print(f"  Content verified: paragraphs contain text")
        
        assert len(paragraphs) > 0, "No paragraphs extracted"
        # Verify Hebrew text is preserved (check for Hebrew Unicode range)
        has_hebrew = any('\u0590' <= c <= '\u05FF' for p in paragraphs for c in p)
        assert has_hebrew, "Hebrew text not preserved"
        
        # Cleanup
        docx_path.unlink()
        print("  [OK] DOS to DOCX conversion tests passed\n")
        
    finally:
        dos_file.unlink()
        temp_dir.rmdir()


def test_idml_text_extraction():
    """Test IDML text extraction (requires a sample IDML file)"""
    print("Testing IDML text extraction...")
    print("  Note: This test requires a real IDML file with Stories folder")
    print("  Skipping automated test - manual testing recommended\n")


def test_file_priority_selection():
    """Test file selection priority"""
    print("Testing file selection priority...")
    
    # Create temporary directory with multiple file types
    temp_dir = Path(tempfile.mkdtemp())
    
    # Create dummy files (just empty)
    docx_file = temp_dir / "test.docx"
    doc_file = temp_dir / "test.doc"
    idml_file = temp_dir / "test.idml"
    
    docx_file.touch()
    doc_file.touch()
    idml_file.touch()
    
    # Test priority: should return .docx files first
    files = get_processable_files(temp_dir)
    print(f"  Found {len(files)} file(s)")
    if files:
        print(f"  Selected file type: {files[0].suffix}")
        assert files[0].suffix == '.docx', "Priority selection failed - should select .docx first"
    
    # Remove .docx, should select .doc
    docx_file.unlink()
    files = get_processable_files(temp_dir)
    if files:
        print(f"  After removing .docx, selected: {files[0].suffix}")
        assert files[0].suffix == '.doc', "Should select .doc when .docx not available"
    
    # Cleanup
    doc_file.unlink()
    idml_file.unlink()
    temp_dir.rmdir()
    
    print("  [OK] File priority selection tests passed\n")


def main():
    """Run all tests"""
    print("=" * 60)
    print("Testing New File Format Support")
    print("=" * 60 + "\n")
    
    try:
        test_dos_file_detection()
        test_dos_conversion()
        test_file_priority_selection()
        test_idml_text_extraction()
        
        print("=" * 60)
        print("All tests passed!")
        print("=" * 60)
        
    except Exception as e:
        print(f"\n[ERROR] Test failed with error: {e}")
        import traceback
        traceback.print_exc()
        return 1
    
    return 0


if __name__ == "__main__":
    exit(main())

