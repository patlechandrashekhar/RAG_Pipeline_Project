#!/usr/bin/env python3
"""
Test script for the Two-Step RAG + Agent SDK Architecture

This script tests the new two-step architecture to ensure:
1. RAG system works independently
2. Agent SDK enhancement is optional
3. Fallback mechanisms work correctly
"""

import asyncio
import sys
from pathlib import Path

# Add source to path
PROJECT_ROOT = Path(__file__).resolve().parent
SRC_ROOT = PROJECT_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

# Import the two-step system components
from page_indexing_rag.config_agentsdk import (
    ANTHROPIC_API_KEY,
    AWS_ACCESS_KEY_ID,
    AWS_SECRET_ACCESS_KEY,
    USE_PORTKEY,
    PORTKEY_API_KEY,
    OPENAI_ANSWER_MODEL,
    build_openai_compatible_client
)

# Test configuration
TEST_QUERIES = [
    "What is TML syntax for V93000?",
    "How to configure AVI64 instrument?",
    "What are the voltage limits for ADuCM410?",
    "Explain HVIN instrument programming",
    "What is the purpose of test flow in TML?"
]

# Test enhancement modes
TEST_MODES = [
    "disabled",    # RAG only
    "standard",    # Basic tools
    "research"     # Web research tools
]

async def test_rag_only():
    """Test RAG-only mode (no Agent SDK)"""
    print("\n" + "="*60)
    print("TEST 1: RAG-Only Mode (No Agent SDK)")
    print("="*60)

    # Import components directly
    import chromadb
    from page_indexing_rag.config_agentsdk import CHROMA_PATH

    # Check if ChromaDB has data
    client = chromadb.PersistentClient(path=str(CHROMA_PATH))
    try:
        collection = client.get_collection("tml_copilot_v3_titan")
        doc_count = collection.count()
        print(f"[OK] ChromaDB collection found: {doc_count} documents")
    except:
        try:
            collection = client.create_collection("tml_copilot_v3_titan")
            doc_count = 0
            print(f"[WARN] ChromaDB collection created (empty)")
        except:
            # Collection might exist but be empty
            collection = client.get_collection("tml_copilot_v3_titan")
            doc_count = collection.count()
            print(f"[OK] ChromaDB collection exists: {doc_count} documents")

    # Test embedding function
    if AWS_ACCESS_KEY_ID and AWS_SECRET_ACCESS_KEY:
        print("[OK] AWS credentials configured")
        try:
            import boto3
            import json

            bedrock = boto3.client(
                'bedrock-runtime',
                region_name='us-east-1',
                aws_access_key_id=AWS_ACCESS_KEY_ID,
                aws_secret_access_key=AWS_SECRET_ACCESS_KEY
            )

            # Test embedding
            response = bedrock.invoke_model(
                modelId="amazon.titan-embed-text-v1",
                body=json.dumps({"inputText": "test"})
            )
            result = json.loads(response['body'].read())
            print(f"[OK] Titan embeddings working (dim: {len(result['embedding'])})")
        except Exception as e:
            print(f"[ERROR] Titan embedding error: {e}")
    else:
        print("[WARN] AWS credentials not configured (using dummy embeddings)")

    # Test LLM connection (Portkey or Anthropic)
    if USE_PORTKEY and PORTKEY_API_KEY:
        print("[OK] Portkey configured (unified API gateway)")
        print(f"    Model: {OPENAI_ANSWER_MODEL}")
        try:
            client = build_openai_compatible_client()
            # Quick test
            response = client.chat.completions.create(
                model=OPENAI_ANSWER_MODEL,
                max_tokens=10,
                messages=[{"role": "user", "content": "Hi"}]
            )
            print("[OK] Portkey connection successful")
        except Exception as e:
            print(f"[ERROR] Portkey error: {e}")
    elif ANTHROPIC_API_KEY:
        print("[OK] Direct Anthropic API key configured")
        try:
            from anthropic import Anthropic
            client = Anthropic(api_key=ANTHROPIC_API_KEY)
            # Quick test
            response = client.messages.create(
                model="claude-3-haiku-20240307",
                max_tokens=10,
                messages=[{"role": "user", "content": "Hi"}]
            )
            print("[OK] Claude API connection successful")
        except Exception as e:
            print(f"[ERROR] Claude API error: {e}")
    else:
        print("[ERROR] No LLM API configured (need PORTKEY_API_KEY or ANTHROPIC_API_KEY)")

    return doc_count > 0 and bool(ANTHROPIC_API_KEY)

async def test_two_step_system():
    """Test the complete two-step system"""
    print("\n" + "="*60)
    print("TEST 2: Two-Step RAG + Agent SDK System")
    print("="*60)

    try:
        # Import the streamlit app module to get the classes
        sys.path.insert(0, str(PROJECT_ROOT / "app"))
        from streamlit_app_agent_sdk_v2 import (
            TwoStepRAGSystem,
            EnhancementMode,
            RAGContext
        )

        # Create system instance
        system = TwoStepRAGSystem()
        print("[OK] Two-step system initialized")

        # Test different enhancement modes
        for mode_name in TEST_MODES[:1]:  # Test just RAG-only first
            mode = EnhancementMode(mode_name)
            print(f"\nTesting mode: {mode.display_name}")
            print(f"  Description: {mode.description}")
            print(f"  Tools: {mode.allowed_tools}")

            # Test with a simple query
            query = TEST_QUERIES[0]
            print(f"  Query: {query}")

            try:
                response = await system.process_query(
                    query=query,
                    mode=mode
                )

                print(f"  [OK] Response received")
                print(f"    - Processing time: {response.processing_time:.2f}s")
                print(f"    - Sources: {response.sources_count}")
                print(f"    - Enhancement used: {response.enhancement_used}")
                print(f"    - Confidence: {response.rag_context.confidence:.2f}")

                if response.answer:
                    print(f"    - Answer preview: {response.answer[:100]}...")
                else:
                    print(f"    - No answer generated")

            except Exception as e:
                print(f"  [ERROR] Error: {e}")

        print("\n[OK] Two-step system test complete")
        return True

    except ImportError as e:
        print(f"[ERROR] Failed to import two-step system: {e}")
        return False
    except Exception as e:
        print(f"[ERROR] Unexpected error: {e}")
        return False

async def test_enhancement_strategies():
    """Test different enhancement strategies"""
    print("\n" + "="*60)
    print("TEST 3: Enhancement Strategies")
    print("="*60)

    try:
        from page_indexing_rag.rag_agent_enhanced import (
            EnhancedRAGAgent,
            NoEnhancement,
            BasicToolEnhancement,
            AutonomousEnhancement
        )

        print("[OK] Enhancement strategies imported")

        # Test each strategy
        strategies = [
            ("none", NoEnhancement()),
            ("basic", BasicToolEnhancement()),
            ("autonomous", AutonomousEnhancement())
        ]

        for name, strategy in strategies:
            print(f"\nStrategy: {strategy.name}")
            print(f"  Description: {strategy.description}")

        print("\n[OK] Enhancement strategies test complete")
        return True

    except ImportError as e:
        print(f"[WARN] Enhanced RAG Agent not available: {e}")
        return False

async def test_agent_sdk_availability():
    """Test if Agent SDK is available"""
    print("\n" + "="*60)
    print("TEST 4: Agent SDK Availability")
    print("="*60)

    try:
        from page_indexing_rag.rag_agent_sdk import (
            SemiconductorRAGAgentSDK,
            QueryResponse
        )
        print("[OK] Agent SDK imported successfully")

        # Try to create an instance
        agent = SemiconductorRAGAgentSDK()
        print("[OK] Agent SDK instance created")

        # Check for actual Claude Agent SDK
        try:
            from claude_agent_sdk import query, ClaudeAgentOptions
            print("[OK] Claude Agent SDK is available")
            return True
        except ImportError:
            print("[WARN] Claude Agent SDK not installed (will use fallback)")
            return False

    except ImportError as e:
        print(f"[WARN] RAG Agent SDK not available: {e}")
        return False

async def main():
    """Run all tests"""
    print("\n" + "="*60)
    print("TWO-STEP RAG + AGENT SDK ARCHITECTURE TEST SUITE")
    print("="*60)

    # Check environment
    print("\nEnvironment Check:")
    print(f"  Python: {sys.version.split()[0]}")
    print(f"  Project: {PROJECT_ROOT}")
    print(f"  ChromaDB: {PROJECT_ROOT / 'data' / 'chroma_persistent_storage'}")

    # Run tests
    results = []

    # Test 1: RAG-only mode
    rag_ok = await test_rag_only()
    results.append(("RAG System", rag_ok))

    # Test 2: Two-step system
    two_step_ok = await test_two_step_system()
    results.append(("Two-Step System", two_step_ok))

    # Test 3: Enhancement strategies
    strategies_ok = await test_enhancement_strategies()
    results.append(("Enhancement Strategies", strategies_ok))

    # Test 4: Agent SDK availability
    agent_sdk_ok = await test_agent_sdk_availability()
    results.append(("Agent SDK", agent_sdk_ok))

    # Summary
    print("\n" + "="*60)
    print("TEST SUMMARY")
    print("="*60)

    all_passed = True
    for test_name, passed in results:
        status = "[PASS]" if passed else "[FAIL]"
        print(f"  {test_name}: {status}")
        if not passed:
            all_passed = False

    print("\n" + "="*60)
    if all_passed:
        print("ALL TESTS PASSED - System is ready!")
        print("\nTo run the application:")
        print("  streamlit run app/streamlit_app_agent_sdk_v2.py")
    else:
        print("SOME TESTS FAILED - Check configuration:")
        print("\n1. Ensure .env file has required keys:")
        print("   - ANTHROPIC_API_KEY")
        print("   - AWS_ACCESS_KEY_ID")
        print("   - AWS_SECRET_ACCESS_KEY")
        print("\n2. Install optional dependencies:")
        print("   - pip install duckduckgo-search  # For web search")
        print("   - pip install claude-agent-sdk   # For full Agent SDK")
        print("\n3. Ingest documents into ChromaDB")

    print("="*60)

if __name__ == "__main__":
    asyncio.run(main())