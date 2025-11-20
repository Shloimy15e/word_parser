#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Analyze DOS formatting codes to understand what they mean
"""

import sys
import re
from pathlib import Path

def analyze_dos_file(file_path):
    """Analyze DOS file to understand number code patterns"""
    with open(file_path, 'rb') as f:
        raw = f.read()
    
    text = raw.decode('cp862', errors='ignore')
    lines = text.split('\n')
    
    print(f"Analyzing DOS file\n")
    print("=" * 80)
    print("Lines with >number< codes and Hebrew text:\n")
    
    pattern_count = {}
    
    for i, line in enumerate(lines, start=1):
        # Skip formatting lines
        if line.strip().startswith('.'):
            continue
            
        # Find lines with >number< patterns
        codes = re.findall(r'>(\d+)<', line)
        hebrew = sum(1 for c in line if '\u0590' <= c <= '\u05FF')
        
        if codes and hebrew > 5:
            # Count pattern frequencies
            for code in codes:
                pattern_count[code] = pattern_count.get(code, 0) + 1
            
            # Print first 30 examples
            if len([v for v in pattern_count.values() if v == 1]) < 30:
                # Count Hebrew characters for display
                heb_count = sum(1 for c in line if '\u0590' <= c <= '\u05FF')
                print(f"Line {i:4d}: Codes {codes} (Hebrew chars: {heb_count})")
    
    print("\n" + "=" * 80)
    print("CODE FREQUENCY ANALYSIS:")
    print("=" * 80)
    print(f"{'Code':>6} | {'Count':>6} | Notes")
    print("-" * 80)
    
    # Sort by code number
    for code in sorted(pattern_count.keys(), key=int):
        count = pattern_count[code]
        print(f"  >{code:>3}< | {count:>6} |")
    
    print(f"\nTotal unique codes: {len(pattern_count)}")
    print(f"Total code instances: {sum(pattern_count.values())}")


if __name__ == "__main__":
    file_path = Path(r"docs\אגדות מהריט טאג ספיר\הוריות\PEREK3")
    if file_path.exists():
        analyze_dos_file(file_path)
    else:
        print(f"File not found: {file_path}")

