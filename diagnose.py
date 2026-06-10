"""Diagnostic script to identify the source of the cross_encoder_rerank error."""

import sys
from pathlib import Path

# Add source to path
PROJECT_ROOT = Path(__file__).resolve().parent
SRC_ROOT = PROJECT_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

print("Diagnostic: Checking imports and function signatures")
print("=" * 60)

# Check if old retrieval module has issues
print("\n1. Checking old retrieval module...")
try:
    from page_indexing_rag import retrieval
    print("   [INFO] Old retrieval module imported")

    # Check cross_encoder_rerank signature
    import inspect
    sig = inspect.signature(retrieval.cross_encoder_rerank)
    print(f"   [INFO] cross_encoder_rerank signature: {sig}")
    print(f"   [INFO] Number of parameters: {len(sig.parameters)}")

except Exception as e:
    print(f"   [ERROR] {e}")

# Check new retrieval module
print("\n2. Checking new retrieval_agentsdk module...")
try:
    from page_indexing_rag import retrieval_agentsdk
    print("   [INFO] retrieval_agentsdk module imported")

    # Check if cross_encoder_rerank exists
    if hasattr(retrieval_agentsdk, 'cross_encoder_rerank'):
        print("   [WARNING] cross_encoder_rerank found in retrieval_agentsdk!")
    else:
        print("   [OK] cross_encoder_rerank not in retrieval_agentsdk (good)")

except Exception as e:
    print(f"   [ERROR] {e}")

# Check rag_agent module
print("\n3. Checking rag_agent module...")
try:
    from page_indexing_rag.rag_agent import SemiconductorRAGAgent
    print("   [OK] SemiconductorRAGAgent imported successfully")

    # Check if vector_search method exists
    agent = SemiconductorRAGAgent()
    print("   [OK] Agent instantiated")

    # Check method signature
    import inspect
    sig = inspect.signature(agent.vector_search)
    print(f"   [INFO] vector_search signature: {sig}")

except Exception as e:
    print(f"   [ERROR] {e}")

print("\n4. Checking which Streamlit app should be used...")
print("   [IMPORTANT] You should run: streamlit run app/streamlit_app_agentsdk.py")
print("   [IMPORTANT] NOT: streamlit run app/streamlit_app.py")

# Check if .env file exists
print("\n5. Checking environment...")
env_file = PROJECT_ROOT / ".env"
if env_file.exists():
    print("   [OK] .env file exists")

    # Check for required keys
    with open(env_file, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()

    if "ANTHROPIC_API_KEY" in content:
        print("   [OK] ANTHROPIC_API_KEY found in .env")
    else:
        print("   [WARNING] ANTHROPIC_API_KEY not found in .env")

    if "AWS_ACCESS_KEY_ID" in content:
        print("   [OK] AWS_ACCESS_KEY_ID found in .env")
    else:
        print("   [WARNING] AWS_ACCESS_KEY_ID not found in .env")
else:
    print("   [WARNING] .env file not found")
    print("   [INFO] Copy .env.agentsdk to .env and add your credentials")

print("\n" + "=" * 60)
print("Diagnosis complete!")
print("\nMake sure you're running the correct app:")
print("  streamlit run app/streamlit_app_agentsdk.py")
print("\nIf you're still getting errors, the issue might be:")
print("1. Running the wrong streamlit app (use _agentsdk version)")
print("2. Old imports cached - restart your Python/terminal session")
print("3. Missing API credentials in .env file")