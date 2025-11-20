#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Debug script to test DOS file detection in a directory
"""

import sys
from pathlib import Path
from main import is_dos_encoded_file, get_processable_files

def debug_directory(directory_path):
    """Check all files in a directory for DOS encoding"""
    directory = Path(directory_path)
    
    if not directory.exists():
        print(f"Error: Directory '{directory}' does not exist")
        return
    
    print(f"Checking directory: {directory}\n")
    print("=" * 70)
    
    # List all files
    all_files = [f for f in directory.iterdir() if f.is_file()]
    print(f"\nTotal files in directory: {len(all_files)}")
    
    # Check files by extension
    docx_files = list(directory.glob("*.docx"))
    doc_files = list(directory.glob("*.doc"))
    idml_files = list(directory.glob("*.idml"))
    no_ext_files = [f for f in all_files if not f.suffix]
    
    print(f"  .docx files: {len(docx_files)}")
    print(f"  .doc files: {len(doc_files)}")
    print(f"  .idml files: {len(idml_files)}")
    print(f"  No extension files: {len(no_ext_files)}")
    
    # Check DOS detection for files without extension
    if no_ext_files:
        print(f"\n" + "=" * 70)
        print(f"Checking {len(no_ext_files)} file(s) without extension for DOS encoding:")
        print("=" * 70)
        
        dos_detected = []
        for file in no_ext_files:
            is_dos = is_dos_encoded_file(file)
            status = "[DOS DETECTED]" if is_dos else "[NOT DOS]"
            print(f"  {status:20} {file.name}")
            if is_dos:
                dos_detected.append(file)
        
        print(f"\nDOS files detected: {len(dos_detected)}")
    
    # Show what get_processable_files returns
    print(f"\n" + "=" * 70)
    print("Files selected by get_processable_files():")
    print("=" * 70)
    
    selected_files = get_processable_files(directory)
    if selected_files:
        file_type = "Unknown"
        if selected_files[0].suffix == '.docx':
            file_type = ".docx"
        elif selected_files[0].suffix == '.doc':
            file_type = ".doc"
        elif selected_files[0].suffix == '.idml':
            file_type = ".idml"
        elif not selected_files[0].suffix:
            file_type = "DOS (no extension)"
        
        print(f"File type selected: {file_type}")
        print(f"Number of files: {len(selected_files)}")
        print("\nFiles:")
        for i, file in enumerate(selected_files, 1):
            print(f"  {i}. {file.name}")
    else:
        print("No processable files found!")
    
    print("\n" + "=" * 70)


def main():
    if len(sys.argv) < 2:
        print("Usage: python debug_dos_detection.py <directory_path>")
        print("\nExample:")
        print("  python debug_dos_detection.py docs/some_folder")
        return 1
    
    directory_path = sys.argv[1]
    debug_directory(directory_path)
    return 0


if __name__ == "__main__":
    exit(main())

