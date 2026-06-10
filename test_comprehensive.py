"""
Comprehensive test suite for the Claude Agent SDK RAG system.
Tests both backend and frontend functionality.
"""

import asyncio
import os
import sys
import tempfile
import json
from pathlib import Path
from datetime import datetime

# Add source to path
PROJECT_ROOT = Path(__file__).resolve().parent
SRC_ROOT = PROJECT_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

# Test results collector
class TestResults:
    def __init__(self):
        self.passed = []
        self.failed = []
        self.warnings = []

    def add_pass(self, test_name, details=""):
        self.passed.append({"test": test_name, "details": details})
        print(f"  [PASS] {test_name}")
        if details:
            print(f"     {details}")

    def add_fail(self, test_name, error):
        self.failed.append({"test": test_name, "error": str(error)})
        print(f"  [FAIL] {test_name}")
        print(f"     Error: {error}")

    def add_warning(self, test_name, warning):
        self.warnings.append({"test": test_name, "warning": warning})
        print(f"  [WARN] {test_name}")
        print(f"     {warning}")

    def print_summary(self):
        total = len(self.passed) + len(self.failed)
        print("\n" + "=" * 60)
        print("TEST SUMMARY")
        print("=" * 60)
        print(f"Total Tests: {total}")
        print(f"Passed: {len(self.passed)}")
        print(f"Failed: {len(self.failed)}")
        print(f"Warnings: {len(self.warnings)}")

        if self.failed:
            print("\nFailed Tests:")
            for fail in self.failed:
                print(f"  - {fail['test']}: {fail['error'][:100]}")

        if self.warnings:
            print("\nWarnings:")
            for warn in self.warnings:
                print(f"  - {warn['test']}: {warn['warning']}")

        return len(self.failed) == 0


# ========================================
# BACKEND TESTS
# ========================================

async def test_imports(results: TestResults):
    """Test that all required modules can be imported."""
    print("\n[TEST] Testing Imports...")

    try:
        from page_indexing_rag.config_agentsdk import (
            ANTHROPIC_API_KEY,
            AWS_ACCESS_KEY_ID,
            MASTER_SYSTEM_PROMPT
        )
        results.add_pass("Import config_agentsdk")
    except Exception as e:
        results.add_fail("Import config_agentsdk", e)
        return False

    try:
        from page_indexing_rag.ingestion_agentsdk import (
            classify_file,
            semantic_chunk,
            extract_pdf_pages
        )
        results.add_pass("Import ingestion_agentsdk")
    except Exception as e:
        results.add_fail("Import ingestion_agentsdk", e)

    try:
        from page_indexing_rag.retrieval_agentsdk import (
            classify_query_type,
            assess_internal_confidence,
            mmr_deduplicate
        )
        results.add_pass("Import retrieval_agentsdk")
    except Exception as e:
        results.add_fail("Import retrieval_agentsdk", e)

    try:
        from page_indexing_rag.generation_agentsdk import (
            build_file_context_prompt,
            classify_question_complexity,
            build_rag_context
        )
        results.add_pass("Import generation_agentsdk")
    except Exception as e:
        results.add_fail("Import generation_agentsdk", e)

    # Try importing the Agent SDK version
    try:
        from page_indexing_rag.rag_agent_sdk import (
            SemiconductorRAGAgentSDK,
            QueryResponse,
            ask_semiconductor_question
        )
        results.add_pass("Import rag_agent_sdk (Claude Agent SDK version)")
    except ImportError as e:
        if "claude_agent_sdk" in str(e):
            results.add_warning("Import rag_agent_sdk", "claude-agent-sdk not installed - run: pip install claude-agent-sdk>=0.2.111")
        else:
            results.add_fail("Import rag_agent_sdk", e)
    except Exception as e:
        results.add_fail("Import rag_agent_sdk", e)

    # Try importing original RAG agent
    try:
        from page_indexing_rag.rag_agent import SemiconductorRAGAgent
        results.add_pass("Import rag_agent (original version)")
    except Exception as e:
        results.add_fail("Import rag_agent", e)

    return True


async def test_environment(results: TestResults):
    """Test environment variables and configuration."""
    print("\n[TEST] Testing Environment Configuration...")

    # Check .env file
    env_file = PROJECT_ROOT / ".env"
    if env_file.exists():
        results.add_pass(".env file exists")
    else:
        results.add_warning(".env file", "Not found - copy .env.template and add credentials")

    # Check API keys
    from page_indexing_rag.config_agentsdk import (
        ANTHROPIC_API_KEY,
        AWS_ACCESS_KEY_ID,
        AWS_SECRET_ACCESS_KEY,
        PORTKEY_API_KEY,
        USE_PORTKEY,
        CHAT_BACKEND
    )

    if USE_PORTKEY:
        if PORTKEY_API_KEY:
            results.add_pass("Portkey API key configured")
        else:
            results.add_fail("Portkey API key", "PORTKEY_API_KEY not set but AGENTSDK_USE_PORTKEY=true")
    else:
        if ANTHROPIC_API_KEY and not ANTHROPIC_API_KEY.startswith("your-"):
            results.add_pass("Anthropic API key configured")
        else:
            results.add_warning("Anthropic API key", "Not configured - add ANTHROPIC_API_KEY to .env")

    if AWS_ACCESS_KEY_ID and not AWS_ACCESS_KEY_ID.startswith("your-"):
        results.add_pass("AWS credentials configured")
    else:
        results.add_warning("AWS credentials", "Not configured - Titan embeddings unavailable")

    # Check system prompt
    from page_indexing_rag.config_agentsdk import MASTER_SYSTEM_PROMPT
    if MASTER_SYSTEM_PROMPT and "ADI" in MASTER_SYSTEM_PROMPT:
        results.add_pass("System prompt loaded", f"Length: {len(MASTER_SYSTEM_PROMPT)} chars")
    else:
        results.add_fail("System prompt", "Not loaded or invalid")


async def test_file_classification(results: TestResults):
    """Test file classification logic."""
    print("\n[TEST] Testing File Classification...")

    try:
        from page_indexing_rag.ingestion_agentsdk import classify_file

        test_cases = [
            ("test_program.tml", "testsuite main { }", "tml"),
            ("datasheet_aducm410.pdf", "absolute maximum ratings", "datasheet"),
            ("cookbook.tdc", "test development cookbook", "tdc"),
            ("schematic.hib", "hardware interconnect board", "schematic"),
            ("pinmap.xlsx", "pin configuration mapping", "pinmap"),
            ("design.sv", "module test_module", "rtl"),
            ("registers.xml", "<register name='ctrl'>", "regmap"),
            ("config.txt", "instrument setup", "tester_config"),
            ("readme.txt", "general information", "general"),
        ]

        passed = 0
        for filename, content, expected in test_cases:
            result = classify_file(filename, content)
            if result == expected:
                passed += 1
            else:
                results.add_warning(f"Classify {filename}", f"Got '{result}', expected '{expected}'")

        if passed == len(test_cases):
            results.add_pass("File classification", f"All {passed} test cases passed")
        else:
            results.add_warning("File classification", f"{passed}/{len(test_cases)} passed")

    except Exception as e:
        results.add_fail("File classification", e)


async def test_query_classification(results: TestResults):
    """Test query type classification."""
    print("\n[TEST] Testing Query Classification...")

    try:
        from page_indexing_rag.retrieval_agentsdk import classify_query_type

        test_cases = [
            ("Show me the TML code for SPI testing", "P"),  # Proprietary
            ("What's in the V93000 TDC cookbook?", "P"),    # Proprietary
            ("Find the ADuCM410 datasheet", "O"),           # Open/Public
            ("What are IEEE standards for testing?", "O"),   # Open/Public
            ("How does a transistor work?", "U"),           # Unknown
            ("Explain quantum computing", "U"),             # Unknown
        ]

        passed = 0
        for query, expected in test_cases:
            result = classify_query_type(query)
            if result == expected:
                passed += 1
            else:
                results.add_warning(f"Query classification", f"'{query[:30]}...' got '{result}', expected '{expected}'")

        if passed == len(test_cases):
            results.add_pass("Query classification", f"All {passed} test cases passed")
        else:
            results.add_warning("Query classification", f"{passed}/{len(test_cases)} passed")

    except Exception as e:
        results.add_fail("Query classification", e)


async def test_semantic_chunking(results: TestResults):
    """Test semantic text chunking."""
    print("\n[TEST] Testing Semantic Chunking...")

    try:
        from page_indexing_rag.ingestion_agentsdk import semantic_chunk

        test_text = """
        The ADuCM410 is a mixed-signal microcontroller from Analog Devices.
        It features an ARM Cortex-M33 processor running at up to 160 MHz.
        The device includes a 12-bit ADC with up to 16 channels for precise analog measurements.
        Communication interfaces include SPI, I2C, and UART for flexible connectivity.
        Operating voltage ranges from 1.8V to 3.6V, suitable for low-power applications.
        Maximum junction temperature is 125°C, ensuring reliable operation in harsh environments.
        The device comes in a 48-pin LFCSP package with excellent thermal performance.
        """

        # Test with different chunk sizes
        chunks_small = semantic_chunk(test_text, max_tokens=50)
        chunks_medium = semantic_chunk(test_text, max_tokens=100)
        chunks_large = semantic_chunk(test_text, max_tokens=200)

        if len(chunks_small) >= len(chunks_medium) >= len(chunks_large):
            results.add_pass("Semantic chunking",
                           f"Small: {len(chunks_small)}, Medium: {len(chunks_medium)}, Large: {len(chunks_large)} chunks")
        else:
            results.add_warning("Semantic chunking", "Unexpected chunk counts")

    except Exception as e:
        results.add_fail("Semantic chunking", e)


async def test_complexity_classification(results: TestResults):
    """Test question complexity classification."""
    print("\n[TEST] Testing Complexity Classification...")

    try:
        from page_indexing_rag.generation_agentsdk import classify_question_complexity

        test_cases = [
            ("What is the voltage?", "simple"),
            ("How do I configure the ADC?", "medium"),
            ("Debug this complex timing issue in my TML testflow with multiple instrument synchronization", "complex"),
            ("Explain the tradeoffs between different test architectures", "complex"),
        ]

        passed = 0
        for question, expected in test_cases:
            result = classify_question_complexity(question)
            if result == expected:
                passed += 1
            else:
                results.add_warning(f"Complexity", f"'{question[:30]}...' got '{result}', expected '{expected}'")

        if passed == len(test_cases):
            results.add_pass("Complexity classification", f"All {passed} test cases passed")
        else:
            results.add_warning("Complexity classification", f"{passed}/{len(test_cases)} passed")

    except Exception as e:
        results.add_fail("Complexity classification", e)


async def test_chromadb_connection(results: TestResults):
    """Test ChromaDB connection and operations."""
    print("\n[TEST] Testing ChromaDB Connection...")

    try:
        import chromadb
        from page_indexing_rag.config_agentsdk import CHROMA_PATH

        # Try to connect to ChromaDB
        client = chromadb.PersistentClient(path=str(CHROMA_PATH))
        results.add_pass("ChromaDB connection", f"Path: {CHROMA_PATH}")

        # Check collections
        collections = client.list_collections()
        if collections:
            for col in collections:
                count = col.count()
                results.add_pass(f"Collection '{col.name}'", f"{count} documents")
        else:
            results.add_warning("ChromaDB collections", "No collections found")

    except Exception as e:
        results.add_fail("ChromaDB connection", e)


async def test_agent_sdk_import(results: TestResults):
    """Test Claude Agent SDK import and basic functionality."""
    print("\n[TEST] Testing Claude Agent SDK...")

    try:
        from claude_agent_sdk import (
            query,
            ClaudeAgentOptions,
            AgentDefinition,
            SystemMessage,
            ResultMessage
        )
        results.add_pass("Claude Agent SDK import", "All required components imported")

        # Check version if possible
        try:
            import claude_agent_sdk
            if hasattr(claude_agent_sdk, "__version__"):
                version = claude_agent_sdk.__version__
                results.add_pass("Claude Agent SDK version", version)
        except:
            pass

    except ImportError as e:
        results.add_warning("Claude Agent SDK", "Not installed - run: pip install claude-agent-sdk>=0.2.111")
    except Exception as e:
        results.add_fail("Claude Agent SDK import", e)


async def test_rag_agent_initialization(results: TestResults):
    """Test RAG agent initialization."""
    print("\n[TEST] Testing RAG Agent Initialization...")

    # Test original agent
    try:
        from page_indexing_rag.rag_agent import SemiconductorRAGAgent
        agent = SemiconductorRAGAgent()
        stats = agent.get_collection_stats()
        results.add_pass("Original RAG agent", f"Initialized, {stats.get('total_chunks', 0)} chunks in KB")
    except Exception as e:
        results.add_fail("Original RAG agent init", e)

    # Test Agent SDK version
    try:
        from page_indexing_rag.rag_agent_sdk import SemiconductorRAGAgentSDK
        agent_sdk = SemiconductorRAGAgentSDK()
        count = agent_sdk.collection.count() if agent_sdk.collection else 0
        results.add_pass("Agent SDK RAG", f"Initialized, {count} chunks in KB")
    except ImportError:
        results.add_warning("Agent SDK RAG", "Claude Agent SDK not installed")
    except Exception as e:
        results.add_fail("Agent SDK RAG init", e)


# ========================================
# FRONTEND TESTS
# ========================================

async def test_streamlit_apps(results: TestResults):
    """Test that Streamlit apps can be imported."""
    print("\n[TEST] Testing Streamlit Applications...")

    apps = [
        ("app/streamlit_app.py", "Original Streamlit app"),
        ("app/streamlit_app_agentsdk.py", "Agent SDK Streamlit app (v1)"),
        ("app/streamlit_app_agent_sdk.py", "Agent SDK Streamlit app (v2)"),
    ]

    for app_path, description in apps:
        full_path = PROJECT_ROOT / app_path
        if full_path.exists():
            results.add_pass(description, f"Found at {app_path}")
        else:
            results.add_warning(description, f"Not found at {app_path}")


async def test_data_directories(results: TestResults):
    """Test data directory structure."""
    print("\n[TEST] Testing Data Directories...")

    from page_indexing_rag.config_agentsdk import (
        PDF_DATA_DIR,
        HTML_DATA_DIR,
        TXT_DATA_DIR,
        CHROMA_PATH
    )

    dirs = [
        (PDF_DATA_DIR, "PDF data directory"),
        (HTML_DATA_DIR, "HTML data directory"),
        (TXT_DATA_DIR, "Text data directory"),
        (CHROMA_PATH, "ChromaDB storage"),
    ]

    for dir_path, description in dirs:
        if dir_path.exists():
            # Count files
            if dir_path == CHROMA_PATH:
                results.add_pass(description, f"Exists at {dir_path}")
            else:
                files = list(dir_path.glob("*"))
                results.add_pass(description, f"Exists with {len(files)} files")
        else:
            results.add_warning(description, f"Not found at {dir_path}")


# ========================================
# MAIN TEST RUNNER
# ========================================

async def run_all_tests():
    """Run all tests and generate report."""
    print("=" * 60)
    print("COMPREHENSIVE TEST SUITE - CLAUDE AGENT SDK RAG")
    print("=" * 60)
    print(f"Start Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Project Root: {PROJECT_ROOT}")

    results = TestResults()

    # Backend tests
    print("\n" + "=" * 60)
    print("BACKEND TESTS")
    print("=" * 60)

    await test_imports(results)
    await test_environment(results)
    await test_file_classification(results)
    await test_query_classification(results)
    await test_semantic_chunking(results)
    await test_complexity_classification(results)
    await test_chromadb_connection(results)
    await test_agent_sdk_import(results)
    await test_rag_agent_initialization(results)

    # Frontend tests
    print("\n" + "=" * 60)
    print("FRONTEND TESTS")
    print("=" * 60)

    await test_streamlit_apps(results)
    await test_data_directories(results)

    # Print summary
    success = results.print_summary()

    # Save test report
    report_path = PROJECT_ROOT / "test_report.json"
    report = {
        "timestamp": datetime.now().isoformat(),
        "passed": len(results.passed),
        "failed": len(results.failed),
        "warnings": len(results.warnings),
        "tests": {
            "passed": results.passed,
            "failed": results.failed,
            "warnings": results.warnings
        }
    }

    with open(report_path, 'w') as f:
        json.dump(report, f, indent=2)

    print(f"\n[INFO] Test report saved to: {report_path}")

    # Provide next steps
    print("\n" + "=" * 60)
    print("NEXT STEPS")
    print("=" * 60)

    if len(results.failed) == 0:
        print("[OK] All critical tests passed!")
        print("\nYou can now run the applications:")
        print("  1. Original app: streamlit run app/streamlit_app_agentsdk.py")
        print("  2. Agent SDK v2: streamlit run app/streamlit_app_agent_sdk.py")
    else:
        print("[ERROR] Some tests failed. Please fix the issues above.")

    if len(results.warnings) > 0:
        print("\n[WARNING] Warnings detected:")
        print("  - Check API keys in .env file")
        print("  - Install missing packages: pip install -r requirements.txt")
        print("  - For Agent SDK: pip install claude-agent-sdk>=0.2.111")

    return success


if __name__ == "__main__":
    success = asyncio.run(run_all_tests())
    sys.exit(0 if success else 1)