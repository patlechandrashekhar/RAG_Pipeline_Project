#!/usr/bin/env python
"""Test script to verify PDF folder detection functionality."""

import sys
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent / "src"))

from page_indexing_rag.ingestion import get_pdf_folders
from page_indexing_rag.config import PDF_DATA_DIR

def main():
    print(f"Scanning PDF directory: {PDF_DATA_DIR}")
    print("-" * 60)

    folders = get_pdf_folders()

    if not folders:
        print("No PDF files found in any folders.")
    else:
        print(f"Found {len(folders)} folder(s) containing PDFs:")
        print()

        total_pdfs = 0
        for folder_name, folder_info in sorted(folders.items()):
            count = folder_info["count"]
            path = folder_info["path"]
            total_pdfs += count

            if folder_name == "(root)":
                print(f"  [Root Directory]: {count} PDF(s)")
            else:
                print(f"  {folder_name}: {count} PDF(s)")
            print(f"    Path: {path}")
            print()

        print("-" * 60)
        print(f"Total: {total_pdfs} PDF(s) across {len(folders)} folder(s)")

if __name__ == "__main__":
    main()