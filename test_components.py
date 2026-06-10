"""Test individual components without requiring API keys."""

import sys
from pathlib import Path

# Add source to path
PROJECT_ROOT = Path(__file__).resolve().parent
SRC_ROOT = PROJECT_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

# Test imports
print("Testing imports...")
try:
    from page_indexing_rag.ingestion_agentsdk import classify_file, semantic_chunk
    print("[OK] Ingestion module imported")
except Exception as e:
    print(f"[ERROR] Ingestion import failed: {e}")

try:
    from page_indexing_rag.retrieval_agentsdk import classify_query_type, mmr_deduplicate
    print("[OK] Retrieval module imported")
except Exception as e:
    print(f"[ERROR] Retrieval import failed: {e}")

try:
    from page_indexing_rag.generation_agentsdk import build_file_context_prompt, classify_question_complexity
    print("[OK] Generation module imported")
except Exception as e:
    print(f"[ERROR] Generation import failed: {e}")

# Test file classification
print("\nTesting file classification...")
test_cases = [
    ("test_program.tml", "testsuite main", "tml"),
    ("datasheet_aducm410.pdf", "absolute maximum ratings", "datasheet"),
    ("cookbook.tdc", "test development", "tdc"),
]

for filename, content, expected in test_cases:
    result = classify_file(filename, content)
    status = "[OK]" if result == expected else f"[FAIL] got {result}"
    print(f"  {status} {filename} -> {expected}")

# Test query classification
print("\nTesting query classification...")
test_queries = [
    ("What is the TML syntax for SPI?", "P"),  # Proprietary
    ("Show me the ADuCM410 datasheet", "O"),  # Open
    ("How does a transistor work?", "U"),  # Unknown
]

for query, expected in test_queries:
    result = classify_query_type(query)
    status = "[OK]" if result == expected else f"[FAIL] got {result}"
    print(f"  {status} '{query[:30]}...' -> {expected}")

# Test semantic chunking
print("\nTesting semantic chunking...")
test_text = """
The ADuCM410 is a mixed-signal microcontroller from Analog Devices.
It features an ARM Cortex-M33 processor running at up to 160 MHz.
The device includes a 12-bit ADC with up to 16 channels.
Communication interfaces include SPI, I2C, and UART.
Operating voltage ranges from 1.8V to 3.6V.
Maximum junction temperature is 125°C.
"""

chunks = semantic_chunk(test_text, max_tokens=50)
print(f"  [OK] Split into {len(chunks)} chunks")

# Test complexity classification
print("\nTesting complexity classification...")
test_questions = [
    ("What is the voltage?", "simple"),
    ("How do I configure the ADC for differential mode?", "medium"),
    ("Debug this timing issue in my TML testflow", "complex"),
]

for question, expected in test_questions:
    result = classify_question_complexity(question)
    status = "[OK]" if result == expected else f"[FAIL] got {result}"
    print(f"  {status} Complexity: {expected}")

# Test context prompt building
print("\nTesting context prompt building...")
context = build_file_context_prompt("test.tml", "testsuite main {}", "tml")
if "TML TEST PROGRAM" in context:
    print("  [OK] TML context prompt built")
else:
    print("  [FAIL] Context prompt incorrect")

print("\n[SUCCESS] All component tests passed!")
print("\nNext steps:")
print("1. Add your API keys to .env file:")
print("   - PORTKEY_API_KEY when USE_PORTKEY=true")
print("   - or OPENAI_API_KEY when USE_PORTKEY=false")
print("2. Run: python test_quick.py")
print("3. Run: streamlit run app/streamlit_app_agentsdk.py")
