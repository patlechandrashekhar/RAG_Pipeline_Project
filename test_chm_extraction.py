#!/usr/bin/env python3
"""
Test script to verify CHM extraction is working before running full batch ingestion.
"""

import os
import sys
from pathlib import Path

# Add the src directory to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from batch_ingest_igxl_chm import extract_chm_enhanced

def test_single_chm():
    """Test extraction on a single CHM file."""

    # Path to IGXL CHM files
    igxl_path = Path(r"C:\Users\cpatle2\Desktop\DOC from IFX\UltraFlex Learning\IGXL_Help\IGXL_Help\IGXL_Help")

    # Get first CHM file for testing
    chm_files = list(igxl_path.glob("*.chm"))

    if not chm_files:
        print("❌ No CHM files found in the specified directory!")
        print(f"   Checked: {igxl_path}")
        return

    test_file = chm_files[0]
    print(f"Testing CHM extraction on: {test_file.name}")
    print("-" * 50)

    # Extract documents
    documents = extract_chm_enhanced(test_file)

    if documents:
        print(f"✓ Successfully extracted {len(documents)} documents")
        print(f"\nFirst document preview:")
        print(f"  Source: {documents[0]['source']}")
        print(f"  Topic: {documents[0].get('topic', 'N/A')}")
        print(f"  Extraction method: {documents[0].get('extraction_method', 'N/A')}")
        print(f"  Text length: {len(documents[0]['text'])} characters")
        print(f"  Text preview: {documents[0]['text'][:200]}...")

        # Show statistics
        total_chars = sum(len(doc['text']) for doc in documents)
        print(f"\nStatistics:")
        print(f"  Total documents: {len(documents)}")
        print(f"  Total characters: {total_chars:,}")
        print(f"  Average doc size: {total_chars // len(documents):,} chars")

        # List extraction methods used
        methods = set(doc.get('extraction_method', 'unknown') for doc in documents)
        print(f"  Extraction methods: {', '.join(methods)}")

    else:
        print("❌ No documents extracted from CHM file")
        print("\nTroubleshooting:")
        print("1. Ensure hh.exe is available (Windows Help Compiler)")
        print("2. Check if the CHM file is valid and not corrupted")
        print("3. Try running as administrator if permission issues")

    print("\n" + "=" * 50)
    print("Test complete!")

    if documents:
        print("\n✅ CHM extraction is working! You can proceed with full batch ingestion.")
        print("   Run: python batch_ingest_igxl_chm.py")
    else:
        print("\n⚠️ CHM extraction needs troubleshooting before running full batch.")

if __name__ == "__main__":
    test_single_chm()