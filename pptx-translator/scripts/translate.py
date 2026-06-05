#!/usr/bin/env python3
"""
Unified Office Document Translator — auto-detects file format and dispatches to
the correct translator.

Supported formats:
  - .pptx  → translate_pptx.py (PowerPoint)
  - .docx  → translate_docx.py (Word)

Usage:
    python translate.py input.pptx --source zh --target en
    python translate.py input.docx --source zh --target en --engine cerebras
"""

import sys
from pathlib import Path


def detect_format(file_path: str) -> str:
    """Detect file format from extension. Returns 'pptx' or 'docx'."""
    ext = Path(file_path).suffix.lower()
    if ext == '.pptx':
        return 'pptx'
    elif ext == '.docx':
        return 'docx'
    else:
        print(f"Error: Unsupported file format '{ext}'. Supported: .pptx, .docx")
        sys.exit(1)


def main():
    if len(sys.argv) < 2:
        print("Usage: python translate.py <input.pptx|input.docx> [options]")
        print("")
        print("Examples:")
        print("  python translate.py input.pptx -s zh -t en")
        print("  python translate.py input.docx -s zh -t en -e cerebras")
        print("")
        print("For full options, run: python translate_pptx.py --help")
        print("                      python translate_docx.py --help")
        sys.exit(1)

    input_file = sys.argv[1]
    fmt = detect_format(input_file)

    # Dispatch to the correct translator
    if fmt == 'pptx':
        from translate_pptx import main as pptx_main
        pptx_main()
    elif fmt == 'docx':
        from translate_docx import main as docx_main
        docx_main()


if __name__ == '__main__':
    main()
