#!/usr/bin/env python3
"""
Test Portkey integration for the Two-Step RAG System

This script verifies that Portkey is properly configured and working
for both LLM calls and embeddings.
"""

import sys
import asyncio
from pathlib import Path

# Add source to path
PROJECT_ROOT = Path(__file__).resolve().parent
SRC_ROOT = PROJECT_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

# Import configuration
from page_indexing_rag.config_agentsdk import (
    USE_PORTKEY,
    PORTKEY_API_KEY,
    PORTKEY_BASE_URL,
    OPENAI_ANSWER_MODEL,
    OPENAI_EMBEDDING_MODEL,
    build_openai_compatible_client
)

print("=" * 60)
print("PORTKEY INTEGRATION TEST")
print("=" * 60)

# Check configuration
print("\n1. Configuration Check:")
print(f"   USE_PORTKEY: {USE_PORTKEY}")
print(f"   PORTKEY_API_KEY: {'SET' if PORTKEY_API_KEY else 'NOT SET'}")
print(f"   PORTKEY_BASE_URL: {PORTKEY_BASE_URL}")
print(f"   ANSWER_MODEL: {OPENAI_ANSWER_MODEL}")
print(f"   EMBEDDING_MODEL: {OPENAI_EMBEDDING_MODEL}")

if not USE_PORTKEY or not PORTKEY_API_KEY:
    print("\n[ERROR] Portkey not configured. Set USE_PORTKEY=true and PORTKEY_API_KEY in .env")
    sys.exit(1)

# Test LLM calls through Portkey
print("\n2. Testing LLM calls through Portkey...")
try:
    client = build_openai_compatible_client()

    # Test chat completion
    response = client.chat.completions.create(
        model=OPENAI_ANSWER_MODEL,
        max_tokens=50,
        temperature=0,
        messages=[
            {"role": "system", "content": "You are a helpful assistant."},
            {"role": "user", "content": "Say 'Portkey is working' if you can see this."}
        ]
    )

    answer = response.choices[0].message.content
    print(f"   [OK] LLM Response: {answer[:100]}")

except Exception as e:
    print(f"   [ERROR] LLM call failed: {e}")

# Test embeddings through Portkey
print("\n3. Testing embeddings through Portkey...")
try:
    client = build_openai_compatible_client()

    # Test embedding
    response = client.embeddings.create(
        input="Test embedding through Portkey",
        model=OPENAI_EMBEDDING_MODEL
    )

    embedding = response.data[0].embedding
    print(f"   [OK] Embedding dimensions: {len(embedding)}")
    print(f"   [OK] First 5 values: {embedding[:5]}")

except Exception as e:
    print(f"   [ERROR] Embedding call failed: {e}")

# Test the two-step system
print("\n4. Testing Two-Step RAG System...")
try:
    # Import and test the streamlit app module
    sys.path.insert(0, str(PROJECT_ROOT / "app"))
    from streamlit_app_agent_sdk_v2 import (
        TwoStepRAGSystem,
        EnhancementMode
    )

    # Create system
    system = TwoStepRAGSystem()
    print("   [OK] Two-step system initialized")

    # Test a simple query
    async def test_query():
        response = await system.process_query(
            query="What is TML?",
            mode=EnhancementMode.DISABLED  # RAG-only mode
        )
        return response

    response = asyncio.run(test_query())
    print(f"   [OK] Query processed in {response.processing_time:.2f}s")
    print(f"   [OK] Answer preview: {response.answer[:100]}...")

except Exception as e:
    print(f"   [ERROR] Two-step system test failed: {e}")

# Summary
print("\n" + "=" * 60)
print("PORTKEY INTEGRATION SUMMARY")
print("=" * 60)
print("""
If all tests passed, your Portkey integration is working correctly!

The two-step RAG system is now using Portkey for:
- Claude LLM calls (through Bedrock)
- Embeddings (through Azure OpenAI)
- All API routing through a single gateway

To run the app:
  streamlit run app/streamlit_app_agent_sdk_v2.py
""")