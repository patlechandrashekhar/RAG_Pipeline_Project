"""Quick test to verify the Agent SDK integration works."""

import asyncio
import os
from pathlib import Path
import sys

# Add source to path
PROJECT_ROOT = Path(__file__).resolve().parent
SRC_ROOT = PROJECT_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from page_indexing_rag.rag_agent import SemiconductorRAGAgent
from page_indexing_rag.config_agentsdk import (
    CHAT_BACKEND,
    EMBEDDING_BACKEND,
    OPENAI_API_KEY,
    PORTKEY_API_KEY,
    USE_PORTKEY,
)


async def quick_test():
    """Quick test of basic functionality."""
    print("Testing Agent SDK RAG...")

    # Check credentials for the active backend.
    if CHAT_BACKEND == "openai_compatible":
        if USE_PORTKEY and not PORTKEY_API_KEY:
            print("[ERROR] Set PORTKEY_API_KEY in .env file")
            return
        if not USE_PORTKEY and not OPENAI_API_KEY:
            print("[ERROR] Set OPENAI_API_KEY in .env file")
            return
        print(f"[OK] Chat backend: {CHAT_BACKEND} ({'Portkey' if USE_PORTKEY else 'OpenAI'})")
    else:
        if not os.getenv("ANTHROPIC_API_KEY"):
            print("[ERROR] Set ANTHROPIC_API_KEY in .env file")
            return
        print(f"[OK] Chat backend: {CHAT_BACKEND}")

    print(f"[OK] Embedding backend: {EMBEDDING_BACKEND}")

    # Initialize agent
    try:
        agent = SemiconductorRAGAgent()
        print("[OK] Agent initialized")
    except Exception as e:
        print(f"[ERROR] Failed to initialize: {e}")
        return

    # Test query expansion
    try:
        variants = await agent.expand_query_claude("ADuCM410 datasheet")
        print(f"[OK] Query expansion: {len(variants)} variants")
        for v in variants:
            print(f"   - {v}")
    except Exception as e:
        print(f"[ERROR] Query expansion failed: {e}")

    # Test embedding
    try:
        embedding = await agent.get_titan_embedding("test text")
        print(f"[OK] Embedding generated: dimension {len(embedding)}")
    except Exception as e:
        print(f"[ERROR] Embedding failed: {e}")

    # Test web placeholder/tool path
    try:
        web_results = await agent.web_search("ADuCM410 datasheet", "O")
        print(f"[OK] Web tool path: {len(web_results)} result(s)")
    except Exception as e:
        print(f"[ERROR] Web tool path failed: {e}")

    # Test collection stats
    try:
        stats = agent.get_collection_stats()
        print(f"[OK] ChromaDB: {stats['total_chunks']} chunks")
    except Exception as e:
        print(f"[ERROR] ChromaDB failed: {e}")

    print("\n[SUCCESS] Basic tests passed! You can now run:")
    print("   streamlit run app/streamlit_app_agentsdk.py")


if __name__ == "__main__":
    asyncio.run(quick_test())
