#!/usr/bin/env python3
"""
Word Parser - Entry point for the modular document processing CLI.

This file is a thin wrapper that delegates to the modular word_parser package.
All functionality is implemented in the word_parser/ directory.

Usage:
    python main.py --book "ליקוטי שיחות" --docs "docs/סדר בראשית" --out "output"
    python main.py --list-formats
    python main.py --help

For the full modular CLI, you can also run:
    python -m word_parser.cli [args]
"""

from word_parser.cli import main

if __name__ == "__main__":
    main()
