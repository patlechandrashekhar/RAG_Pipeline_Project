"""
Test script to demonstrate Claude Agent SDK with built-in tools.

This shows the real power of Agent SDK - Claude can autonomously:
- Search the web
- Read and edit files
- Run commands
- Use specialized agents
"""

import asyncio
import os
from pathlib import Path
import sys

# Add source to path
PROJECT_ROOT = Path(__file__).resolve().parent
SRC_ROOT = PROJECT_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from page_indexing_rag.rag_agent_sdk import (
    SemiconductorRAGAgentSDK,
    ask_semiconductor_question,
    debug_tml_file,
    search_datasheets
)


async def test_agent_sdk_capabilities():
    """Test various Agent SDK capabilities."""
    print("=" * 60)
    print("Testing Claude Agent SDK with Built-in Tools")
    print("=" * 60)

    # Check for API key
    if not os.getenv("ANTHROPIC_API_KEY"):
        print("[ERROR] Set ANTHROPIC_API_KEY in .env file")
        return

    print("\n1. Testing Simple Question (with WebSearch)")
    print("-" * 40)
    try:
        answer = await ask_semiconductor_question(
            "What is the latest ADuCM410 datasheet specification for operating voltage?"
        )
        print(f"Answer: {answer[:500]}...")
        print("[OK] Simple question answered")
    except Exception as e:
        print(f"[ERROR] {e}")

    print("\n2. Testing Datasheet Analysis")
    print("-" * 40)
    try:
        result = await search_datasheets(
            "Find the maximum junction temperature for ADuCM410"
        )
        print(f"Result: {result[:500]}...")
        print("[OK] Datasheet analysis completed")
    except Exception as e:
        print(f"[ERROR] {e}")

    print("\n3. Testing Full Agent SDK with Multiple Tools")
    print("-" * 40)
    try:
        agent = SemiconductorRAGAgentSDK()
        response = await agent.process_with_agent_sdk(
            "Search the web for V93000 ATE platform specifications and summarize key features",
            allowed_tools=["WebSearch", "WebFetch"],
            use_subagents=False
        )
        print(f"Answer: {response.answer[:500]}...")
        if response.web_sources:
            print(f"Web sources used: {len(response.web_sources)}")
        print("[OK] Multi-tool query processed")
    except Exception as e:
        print(f"[ERROR] {e}")

    print("\n4. Testing Interactive Streaming")
    print("-" * 40)
    try:
        agent = SemiconductorRAGAgentSDK()
        print("Streaming response:")
        async for chunk in agent.interactive_query(
            "What are the key features of Analog Devices semiconductor products?"
        ):
            print(chunk, end="", flush=True)
        print("[OK] Streaming completed")
    except Exception as e:
        print(f"[ERROR] {e}")


async def test_file_operations():
    """Test file operations with Agent SDK."""
    print("\n" + "=" * 60)
    print("Testing File Operations with Agent SDK")
    print("=" * 60)

    # Create a test TML file with a bug
    test_tml_path = Path("test_buggy.tml")
    test_tml_content = """
    // Test program with intentional bug
    testsuite main {
        test spi_test {
            // Bug: Missing semicolon
            int voltage = 3.3

            // Bug: Incorrect instrument reference
            avi64.force(voltage);

            // Bug: Missing closing brace

    }
    """

    print("\n1. Creating test TML file with bugs")
    test_tml_path.write_text(test_tml_content)
    print(f"[OK] Created {test_tml_path}")

    print("\n2. Using Agent SDK to debug and fix the file")
    try:
        result = await debug_tml_file(str(test_tml_path))
        print(f"Debug result: {result[:500]}...")
        print("[OK] TML debugging completed")
    except Exception as e:
        print(f"[ERROR] {e}")
    finally:
        # Clean up
        if test_tml_path.exists():
            test_tml_path.unlink()
            print(f"[OK] Cleaned up {test_tml_path}")


async def test_session_management():
    """Test session management for context preservation."""
    print("\n" + "=" * 60)
    print("Testing Session Management")
    print("=" * 60)

    agent = SemiconductorRAGAgentSDK()

    print("\n1. First query - establishing context")
    response1 = await agent.process_with_agent_sdk(
        "Tell me about the ADuCM410 microcontroller",
        allowed_tools=["WebSearch"]
    )
    session_id = response1.session_id
    print(f"Session ID: {session_id[:8]}..." if session_id else "No session ID")
    print(f"Answer: {response1.answer[:200]}...")

    if session_id:
        print("\n2. Second query - using context from first query")
        response2 = await agent.process_with_agent_sdk(
            "What are its key specifications?",  # "its" refers to ADuCM410
            allowed_tools=["WebSearch"],
            session_id=session_id
        )
        print(f"Answer: {response2.answer[:200]}...")
        print("[OK] Session context preserved")
    else:
        print("[WARN] No session ID to test context preservation")


async def main():
    """Run all tests."""
    print("CLAUDE AGENT SDK INTEGRATION TEST")
    print("=" * 60)
    print("This test demonstrates the real Claude Agent SDK capabilities")
    print("with autonomous tool use (WebSearch, File operations, etc.)")
    print()

    # Run basic tests
    await test_agent_sdk_capabilities()

    # Test file operations (optional)
    response = input("\nTest file operations? (y/n): ")
    if response.lower() == 'y':
        await test_file_operations()

    # Test session management (optional)
    response = input("\nTest session management? (y/n): ")
    if response.lower() == 'y':
        await test_session_management()

    print("\n" + "=" * 60)
    print("Testing Complete!")
    print("\nTo use the new Streamlit app with Agent SDK:")
    print("  streamlit run app/streamlit_app_agent_sdk.py")
    print("\nKey improvements with Agent SDK:")
    print("  ✅ Real web search (not placeholders)")
    print("  ✅ Autonomous file operations")
    print("  ✅ Command execution capability")
    print("  ✅ Specialized subagents")
    print("  ✅ Session management")
    print("=" * 60)


if __name__ == "__main__":
    asyncio.run(main())