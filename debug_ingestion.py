#!/usr/bin/env python3
"""
Debug script for PDF ingestion pipeline issues.
"""

import asyncio
import sys
import os
from pathlib import Path

# Add source to path
PROJECT_ROOT = Path(__file__).resolve().parent
SRC_ROOT = PROJECT_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from page_indexing_rag.config_agentsdk import PDF_DATA_DIR, CHROMA_PATH

print("=" * 70)
print("PDF INGESTION DEBUG")
print("=" * 70)

print(f"\n1. Checking PDF_DATA_DIR configuration:")
print(f"   PDF_DATA_DIR = {PDF_DATA_DIR}")
print(f"   Exists? {PDF_DATA_DIR.exists()}")

if PDF_DATA_DIR.exists():
    pdf_files = list(PDF_DATA_DIR.glob("*.pdf"))
    print(f"   PDF files found: {len(pdf_files)}")
    for pdf in pdf_files[:5]:
        print(f"      - {pdf.name} ({pdf.stat().st_size / 1024:.1f} KB)")

# Check both possible locations
local_pdf_dir = PROJECT_ROOT / "data" / "pdf_data"
legacy_pdf_dir = PROJECT_ROOT.parent / "pdf_data"

print(f"\n2. Checking both possible PDF locations:")
print(f"   Local: {local_pdf_dir}")
print(f"      Exists? {local_pdf_dir.exists()}")
if local_pdf_dir.exists():
    local_pdfs = list(local_pdf_dir.glob("*.pdf"))
    print(f"      PDFs: {len(local_pdfs)}")
    for pdf in local_pdfs[:3]:
        print(f"         - {pdf.name}")

print(f"\n   Legacy: {legacy_pdf_dir}")
print(f"      Exists? {legacy_pdf_dir.exists()}")
if legacy_pdf_dir.exists():
    legacy_pdfs = list(legacy_pdf_dir.glob("*.pdf"))
    print(f"      PDFs: {len(legacy_pdfs)}")
    for pdf in legacy_pdfs[:3]:
        print(f"         - {pdf.name}")

print(f"\n3. Checking ChromaDB path:")
print(f"   CHROMA_PATH = {CHROMA_PATH}")
print(f"   Exists? {CHROMA_PATH.exists()}")

print(f"\n4. Checking environment variables:")
print(f"   AWS_ACCESS_KEY_ID = {'Set' if os.getenv('AWS_ACCESS_KEY_ID') else 'Not set'}")
print(f"   AWS_SECRET_ACCESS_KEY = {'Set' if os.getenv('AWS_SECRET_ACCESS_KEY') else 'Not set'}")
print(f"   ANTHROPIC_API_KEY = {'Set' if os.getenv('ANTHROPIC_API_KEY') else 'Not set'}")
print(f"   OPENAI_API_KEY = {'Set' if os.getenv('OPENAI_API_KEY') else 'Not set'}")
print(f"   PORTKEY_API_KEY = {'Set' if os.getenv('PORTKEY_API_KEY') else 'Not set'}")
print(f"   USE_PORTKEY = {os.getenv('USE_PORTKEY', 'Not set')}")

print(f"\n5. Testing embedding configuration:")
from page_indexing_rag.config_agentsdk import (
    CHAT_BACKEND,
    EMBEDDING_BACKEND,
    COLLECTION_NAME,
    EMBEDDING_MODEL
)

print(f"   CHAT_BACKEND = {CHAT_BACKEND}")
print(f"   EMBEDDING_BACKEND = {EMBEDDING_BACKEND}")
print(f"   COLLECTION_NAME = {COLLECTION_NAME}")
print(f"   EMBEDDING_MODEL = {EMBEDDING_MODEL}")

print("\n6. Testing agent initialization...")
try:
    from page_indexing_rag.rag_agent import SemiconductorRAGAgent
    agent = SemiconductorRAGAgent(chroma_path=str(CHROMA_PATH))
    print("   ✅ Agent initialized successfully")

    print(f"\n7. Testing document processing...")

    # Find a test PDF
    test_pdf = None
    if pdf_files:
        test_pdf = pdf_files[0]
    elif legacy_pdf_dir.exists():
        legacy_pdfs = list(legacy_pdf_dir.glob("*.pdf"))
        if legacy_pdfs:
            test_pdf = legacy_pdfs[0]

    if test_pdf:
        print(f"   Testing with: {test_pdf.name}")

        async def test_ingestion():
            try:
                file_type, chunks = await agent.process_document(
                    str(test_pdf),
                    force_reingest=True
                )
                print(f"   ✅ Result: file_type={file_type}, chunks={chunks}")
                return True
            except Exception as e:
                print(f"   ❌ Error: {e}")
                import traceback
                traceback.print_exc()
                return False

        success = asyncio.run(test_ingestion())
    else:
        print("   ⚠️ No PDF files found to test")

except Exception as e:
    print(f"   ❌ Failed to initialize agent: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 70)
print("DEBUG COMPLETE")
print("=" * 70)