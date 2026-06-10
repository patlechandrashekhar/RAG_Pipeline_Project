#!/usr/bin/env python3
"""
Test script for PDF ingestion pipeline.
Run this to verify the ingestion system is working properly.
"""

import asyncio
import sys
from pathlib import Path

# Add source to path
PROJECT_ROOT = Path(__file__).resolve().parent
SRC_ROOT = PROJECT_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from page_indexing_rag.rag_agent import SemiconductorRAGAgent
from page_indexing_rag.config_agentsdk import PDF_DATA_DIR, CHROMA_PATH


async def test_ingestion():
    """Test the PDF ingestion pipeline."""
    print("=" * 60)
    print("PDF Ingestion Pipeline Test")
    print("=" * 60)

    # Initialize agent
    print("\n1. Initializing RAG Agent...")
    try:
        agent = SemiconductorRAGAgent(chroma_path=str(CHROMA_PATH))
        print("✅ Agent initialized successfully")
    except Exception as e:
        print(f"❌ Failed to initialize agent: {e}")
        return

    # Check PDF directory
    print(f"\n2. Checking PDF directory: {PDF_DATA_DIR}")
    if not PDF_DATA_DIR.exists():
        print(f"❌ Directory does not exist: {PDF_DATA_DIR}")
        return

    pdf_files = list(PDF_DATA_DIR.glob("*.pdf"))
    print(f"✅ Found {len(pdf_files)} PDF files")

    if not pdf_files:
        print("⚠️ No PDF files found in directory")
        return

    # Test ingestion on first PDF
    test_pdf = pdf_files[0]
    print(f"\n3. Testing ingestion on: {test_pdf.name}")

    # Test without force (check duplicate detection)
    print("\n   a. Testing duplicate detection...")
    try:
        file_type, chunks = await agent.process_document(
            str(test_pdf),
            force_reingest=False
        )
        if file_type == "already_ingested":
            print(f"   ℹ️ File already in database (duplicate detection working)")
        else:
            print(f"   ✅ Ingested {chunks} chunks (file type: {file_type})")
    except Exception as e:
        print(f"   ❌ Error during ingestion: {e}")

    # Test with force re-ingestion
    print("\n   b. Testing force re-ingestion...")
    try:
        file_type, chunks = await agent.process_document(
            str(test_pdf),
            force_reingest=True
        )
        if chunks > 0:
            print(f"   ✅ Force re-ingested {chunks} chunks (file type: {file_type})")
        else:
            print(f"   ⚠️ No chunks generated (file type: {file_type})")
            print(f"      This could mean the PDF has no extractable text content")
    except Exception as e:
        print(f"   ❌ Error during force re-ingestion: {e}")
        import traceback
        traceback.print_exc()

    # Check collection stats
    print("\n4. Checking knowledge base stats...")
    stats = agent.get_collection_stats()
    print(f"   📊 Total chunks in KB: {stats['total_chunks']}")
    print(f"   📄 Unique sources: {len(stats['sources'])}")
    if stats['sources']:
        print("   📚 Sample sources:")
        for source in stats['sources'][:5]:
            print(f"      - {source}")

    # Test a simple query
    print("\n5. Testing retrieval...")
    try:
        results = await agent.vector_search("test query", n_results=3)
        if results:
            print(f"   ✅ Retrieved {len(results)} chunks")
            print(f"   📄 Top result from: {results[0]['source']}")
        else:
            print("   ⚠️ No results retrieved (knowledge base might be empty)")
    except Exception as e:
        print(f"   ❌ Error during retrieval: {e}")

    print("\n" + "=" * 60)
    print("Test Complete!")
    print("=" * 60)


if __name__ == "__main__":
    asyncio.run(test_ingestion())