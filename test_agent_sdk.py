"""Smoke test for the active RAG backend configuration."""

from __future__ import annotations

import asyncio
import os
import shutil
import sys
import tempfile
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent
SRC_ROOT = PROJECT_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from page_indexing_rag.config_agentsdk import (
    CHAT_BACKEND,
    EMBEDDING_BACKEND,
    OPENAI_API_KEY,
    PORTKEY_API_KEY,
    USE_PORTKEY,
)
from page_indexing_rag.rag_agent import SemiconductorRAGAgent


def _check_credentials() -> bool:
    print("\n1. Checking credentials...")
    if CHAT_BACKEND == "openai_compatible":
        if USE_PORTKEY and not PORTKEY_API_KEY:
            print("  [ERROR] PORTKEY_API_KEY not set")
            return False
        if not USE_PORTKEY and not OPENAI_API_KEY:
            print("  [ERROR] OPENAI_API_KEY not set")
            return False
        print(f"  [OK] Chat backend: {CHAT_BACKEND} ({'Portkey' if USE_PORTKEY else 'OpenAI'})")
    else:
        if not os.getenv("ANTHROPIC_API_KEY"):
            print("  [ERROR] ANTHROPIC_API_KEY not set")
            return False
        print(f"  [OK] Chat backend: {CHAT_BACKEND}")

    print(f"  [OK] Embedding backend: {EMBEDDING_BACKEND}")
    return True


async def test_basic_functionality() -> bool:
    """Test the active agent configuration end to end."""
    print("=" * 60)
    print("Testing Agent SDK RAG System")
    print("=" * 60)

    if not _check_credentials():
        return False

    print("\n2. Initializing RAG agent...")
    try:
        agent = SemiconductorRAGAgent()
        print("  [OK] Agent initialized successfully")
    except Exception as e:
        print(f"  [ERROR] Failed to initialize agent: {e}")
        return False

    print("\n3. Checking ChromaDB connection...")
    try:
        stats = agent.get_collection_stats()
        print("  [OK] ChromaDB connected")
        print(f"    Collection: {agent.collection_name}")
        print(f"    Chunks: {stats['total_chunks']}")
    except Exception as e:
        print(f"  [ERROR] ChromaDB error: {e}")
        return False

    print("\n4. Testing embeddings...")
    try:
        embedding = await agent.get_titan_embedding("This is a test of the configured embedding model.")
        if embedding and len(embedding) == agent.embedding_dimensions:
            print(f"  [OK] Embedding successful (dimension: {len(embedding)})")
        else:
            print(f"  [ERROR] Unexpected embedding dimension: {len(embedding)}")
            return False
    except Exception as e:
        print(f"  [ERROR] Embedding error: {e}")
        return False

    print("\n5. Testing chat generation...")
    try:
        text = agent._complete_text(
            model=agent.claude_model,
            max_tokens=8,
            temperature=0,
            system="Reply with OK only.",
            user_prompt="ping",
        )
        if "OK" in text.upper():
            print("  [OK] Chat generation successful")
        else:
            print(f"  [ERROR] Unexpected chat response: {text[:80]}")
            return False
    except Exception as e:
        print(f"  [ERROR] Chat generation failed: {e}")
        return False

    print("\n6. Testing query classification...")
    from page_indexing_rag.retrieval_agentsdk import classify_query_type

    test_queries = [
        ("What is the maximum voltage for ADuCM410?", "O"),
        ("Show me the TML code for SPI testing", "P"),
        ("How does a transistor work?", "U"),
    ]
    for query, expected_type in test_queries:
        query_type = classify_query_type(query)
        if query_type == expected_type:
            print(f"  [OK] '{query[:30]}...' -> {query_type}")
        else:
            print(f"  [ERROR] '{query[:30]}...' -> {query_type} (expected {expected_type})")
            return False

    print("\n7. Testing file classification...")
    from page_indexing_rag.ingestion_agentsdk import classify_file

    test_files = [
        ("test_program.tml", "testsuite main", "tml"),
        ("datasheet_aducm410.pdf", "absolute maximum ratings", "datasheet"),
        ("cookbook.tdc", "test development cookbook", "tdc"),
        ("schematic.hib", "hardware interconnect", "schematic"),
    ]
    for filename, content, expected_type in test_files:
        file_type = classify_file(filename, content)
        if file_type == expected_type:
            print(f"  [OK] {filename} -> {file_type}")
        else:
            print(f"  [ERROR] {filename} -> {file_type} (expected {expected_type})")
            return False

    print("\n8. Testing web tool path...")
    try:
        web_results = await agent.web_search("ADuCM410 datasheet", "O")
        print(f"  [OK] Web tool path returned {len(web_results)} result(s)")
    except Exception as e:
        print(f"  [ERROR] Web tool path failed: {e}")
        return False

    print("\n" + "=" * 60)
    print("[OK] All basic tests passed!")
    print("=" * 60)
    return True


async def test_document_ingestion() -> None:
    """Test document ingestion with a small temporary file."""
    print("\n9. Testing document ingestion...")
    test_fd, test_name = tempfile.mkstemp(prefix="agent_doc_", suffix=".txt", dir=PROJECT_ROOT / "data")
    os.close(test_fd)
    test_file = Path(test_name)
    test_content = """
    The ADuCM410 is a mixed-signal microcontroller from Analog Devices.
    It includes SPI, I2C, and UART interfaces.
    Operating voltage ranges from 1.8V to 3.6V.
    """

    try:
        test_file.write_text(test_content, encoding="utf-8")
        print(f"  [OK] Created test file: {test_file}")

        test_chroma_dir = Path(tempfile.mkdtemp(prefix="agent_smoke_", dir=PROJECT_ROOT / "data"))
        agent = SemiconductorRAGAgent(chroma_path=str(test_chroma_dir))
        file_type, chunks = await agent.process_document(str(test_file))
        if chunks > 0:
            print(f"  [OK] Ingested {chunks} chunks (type: {file_type})")
        else:
            print("  [WARN] File already in database or no chunks created")
    except Exception as e:
        print(f"  [ERROR] Ingestion error: {e}")
    finally:
        if test_file.exists():
            try:
                test_file.unlink()
                print("  [OK] Cleaned up test file")
            except PermissionError as e:
                print(f"  [WARN] Could not remove temporary test file: {e}")
        if "test_chroma_dir" in locals() and test_chroma_dir.exists():
            try:
                shutil.rmtree(test_chroma_dir)
                print("  [OK] Cleaned up isolated Chroma test store")
            except PermissionError as e:
                print(f"  [WARN] Could not remove isolated Chroma test store: {e}")


async def main() -> None:
    success = await test_basic_functionality()
    if success:
        response = input("\nRun document ingestion test? (y/n): ")
        if response.lower() == "y":
            await test_document_ingestion()

    print("\n[OK] Testing complete!")
    print("\nTo run the Streamlit app with Agent SDK backend:")
    print("  streamlit run app/streamlit_app_agentsdk.py")


if __name__ == "__main__":
    asyncio.run(main())
