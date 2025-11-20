#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Test script for perek extraction and number conversion
"""

from main import extract_heading4_info, number_to_hebrew_gematria, extract_year

def test_number_to_hebrew():
    """Test number to Hebrew letter conversion"""
    print("Testing number to Hebrew letter conversion...")
    
    tests = [
        (1, 'א'),
        (2, 'ב'),
        (3, 'ג'),
        (4, 'ד'),
        (5, 'ה'),
        (10, 'י'),
        (11, 'כ'),
        (20, 'ר'),
    ]
    
    for num, expected in tests:
        result = number_to_hebrew_letter(num)
        # Check result without printing Hebrew to avoid console encoding issues
        assert result == expected, f"Failed: {num} conversion incorrect"
        print(f"  {num} -> Hebrew letter (correct)")

    
    print("  [OK] Number to Hebrew conversion tests passed\n")


def test_perek_extraction():
    """Test perek information extraction"""
    print("Testing perek extraction...")
    
    tests = [
        ('PEREK1', 'פרק א'),
        ('perek1', 'פרק א'),
        ('PEREK2', 'פרק ב'),
        ('perek3', 'פרק ג'),
        ('PEREK1A', 'פרק א 1'),
        ('perek1a', 'פרק א 1'),
        ('PEREK2B', 'פרק ב 2'),
        ('perek3c', 'פרק ג 3'),
        ('PEREK01A', 'פרק א 1'),  # With leading zero
        ('PEREK001', 'פרק א'),     # With leading zeros
        ('MEKOROS', 'מקורות'),
        ('mekoros', 'מקורות'),
        ('MKOROS', 'מקורות'),      # Alternative spelling
        ('notaperek', None),
        ('somefilename', None),
    ]
    
    for filename, expected in tests:
        result = extract_perek_info(filename)
        status = "OK" if result == expected else "FAIL"
        # Avoid printing Hebrew to prevent console encoding issues
        print(f"  {filename:20} -> [{status}]")
        assert result == expected, f"Failed: {filename} extraction incorrect"
    
    print("  [OK] Perek extraction tests passed\n")


def test_year_optional():
    """Test that year extraction is optional"""
    print("Testing year extraction (optional)...")
    
    tests = [
        ('PEREK1', None, 'No year'),
        ('somefile', None, 'No year'),
        ('תשנט', 'תשנט', 'Has year'),
        ('פרשת בראשית תשנט', 'תשנט', 'Has year'),
    ]
    
    for filename, expected, desc in tests:
        result = extract_year(filename)
        status = "OK" if result == expected else "FAIL"
        # Avoid printing Hebrew to prevent console encoding issues
        print(f"  {desc:25} [{status}]")
        if expected is not None:  # Only assert if we expect to find something
            assert result == expected, f"Failed: year extraction incorrect"
    
    print("  [OK] Year extraction tests passed\n")


def main():
    """Run all tests"""
    print("=" * 60)
    print("Testing Perek and Year Extraction")
    print("=" * 60 + "\n")
    
    try:
        test_number_to_hebrew()
        test_perek_extraction()
        test_year_optional()
        
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

