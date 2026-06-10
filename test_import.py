#!/usr/bin/env python
"""Test import of new functions."""

import sys
from pathlib import Path

# Add src to path the same way streamlit_app.py does
PROJECT_ROOT = Path(__file__).resolve().parents[0]
SRC_ROOT = PROJECT_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

print(f"Python path: {sys.path[:2]}")
print(f"Importing from: {SRC_ROOT}")

try:
    from page_indexing_rag.ingestion import (
        classify_file,
        collection,
        extract_pdf_pages,
        get_pdf_folders,
        ingest_all_pdfs,
        ingest_pdf,
        ingest_pdfs_from_folder,
        ingest_uploaded_text_file,
        is_pdf_ingested,
    )
    print("SUCCESS: All imports worked!")
    print(f"get_pdf_folders is: {type(get_pdf_folders)}")
    print(f"ingest_pdfs_from_folder is: {type(ingest_pdfs_from_folder)}")
except ImportError as e:
    print(f"ERROR: {e}")
    print("\nTrying alternative import...")
    import page_indexing_rag.ingestion as ing
    print(f"Module attributes: {[x for x in dir(ing) if not x.startswith('_')]}")