"""Shared configuration, client setup, and path resolution with Portkey support."""

from __future__ import annotations

import os
import ssl
import json
import glob

import httpx
from dotenv import load_dotenv
from openai import OpenAI
from anthropic import Anthropic


# Keep legacy SSL behavior from the original app.
os.environ["PYTHONHTTPSVERIFY"] = "0"
os.environ["CURL_CA_BUNDLE"] = ""
os.environ["REQUESTS_CA_BUNDLE"] = ""
ssl._create_default_https_context = ssl._create_unverified_context

load_dotenv()

# ════════════════════════════════════════════════════════════════════════════════
# PORTKEY CONFIGURATION - Use Portkey as gateway for multiple LLM providers
# ════════════════════════════════════════════════════════════════════════════════
USE_PORTKEY = os.getenv("USE_PORTKEY", "true").lower() == "true"
PORTKEY_API_KEY = os.getenv("PORTKEY_API_KEY", "")
PORTKEY_BASE_URL = os.getenv("PORTKEY_BASE_URL", "https://api.portkey.ai/v1")
PORTKEY_PROVIDER = os.getenv("PORTKEY_PROVIDER", "openai")  # Can be "openai", "anthropic", "cohere", etc.
PORTKEY_VIRTUAL_KEY = os.getenv("PORTKEY_VIRTUAL_KEY", None)  # Optional virtual key for provider rotation

# Legacy OpenAI configuration (used when USE_PORTKEY=false)
openai_key = os.getenv("OPENAI_API_KEY", "")

# Import model definitions
from .models import DEFAULT_LLM_MODEL, DEFAULT_EMBEDDING_MODEL

# Model configuration - works with any provider through Portkey
# Using Claude Opus 4.6 as default for best quality
OPENAI_ANSWER_MODEL = os.getenv("OPENAI_ANSWER_MODEL", DEFAULT_LLM_MODEL)
OPENAI_RETRIEVAL_MODEL = os.getenv("OPENAI_RETRIEVAL_MODEL", OPENAI_ANSWER_MODEL)
OPENAI_WEB_MODEL = os.getenv("OPENAI_WEB_MODEL", OPENAI_ANSWER_MODEL)
OPENAI_EMBEDDING_MODEL = os.getenv("OPENAI_EMBEDDING_MODEL", DEFAULT_EMBEDDING_MODEL)

# ════════════════════════════════════════════════════════════════════════════════
# CLIENT INITIALIZATION
# ════════════════════════════════════════════════════════════════════════════════
_portkey_headers: dict = {}

if USE_PORTKEY:
    # Use Portkey gateway
    if not PORTKEY_API_KEY:
        raise ValueError(
            "PORTKEY_API_KEY not found. Please set it in your .env file.\n"
            "Get your API key from: https://app.portkey.ai/\n"
            "Or set USE_PORTKEY=false to use direct OpenAI connection."
        )

    # Set up headers for Portkey
    _portkey_headers: dict = {
        "x-portkey-api-key": PORTKEY_API_KEY,
        "x-portkey-provider": PORTKEY_PROVIDER,
    }

    # Add optional headers
    if PORTKEY_VIRTUAL_KEY:
        _portkey_headers["x-portkey-virtual-key"] = PORTKEY_VIRTUAL_KEY

    # Optional: Add metadata for tracking
    portkey_metadata = {
        "app": "page_indexing_rag",
        "environment": os.getenv("ENVIRONMENT", "development")
    }
    _portkey_headers["x-portkey-metadata"] = json.dumps(portkey_metadata)

    # Optional: Enable caching for repeated queries
    if os.getenv("PORTKEY_CACHE", "false").lower() == "true":
        _portkey_headers["x-portkey-cache"] = "true"
        _portkey_headers["x-portkey-cache-ttl"] = os.getenv("PORTKEY_CACHE_TTL", "3600")

    headers = _portkey_headers  # keep backward compat name

    client = OpenAI(
        api_key="dummy",  # Portkey uses headers for auth
        base_url=PORTKEY_BASE_URL,
        http_client=httpx.Client(
            verify=False,
            headers=headers,
            timeout=120,
            trust_env=os.getenv("PORTKEY_TRUST_ENV_PROXY", "false").lower() == "true",
        ),
    )
    print(f"[OK] Using Portkey gateway ({PORTKEY_PROVIDER} provider)")
else:
    # Use direct OpenAI connection (legacy)
    if not openai_key:
        raise ValueError(
            "OPENAI_API_KEY not found. Please set it in your .env file.\n"
            "Or set USE_PORTKEY=true and configure PORTKEY_API_KEY to use Portkey gateway."
        )

    client = OpenAI(
        api_key=openai_key,
        http_client=httpx.Client(
            verify=False,
            trust_env=os.getenv("OPENAI_TRUST_ENV_PROXY", "false").lower() == "true",
        ),
    )
    print("✓ Using direct OpenAI connection")

# ════════════════════════════════════════════════════════════════════════════════
# PATH CONFIGURATION
# ════════════════════════════════════════════════════════════════════════════════
PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
WORKSPACE_ROOT = os.path.abspath(os.path.join(PROJECT_ROOT, ".."))


def _resolve_dir(local_path: str, legacy_path: str) -> str:
    """Prefer project-local paths but support legacy workspace layout."""
    if os.path.exists(local_path):
        return local_path
    if os.path.exists(legacy_path):
        return legacy_path
    return local_path


def _count_recursive_pdfs(path: str) -> int:
    if not os.path.isdir(path):
        return 0
    total = 0
    for _root, _dirs, files in os.walk(path):
        total += sum(1 for f in files if f.lower().endswith(".pdf"))
    return total


def _resolve_pdf_dir(local_path: str, legacy_path: str) -> str:
    """Select the existing PDF directory with the most recursive .pdf files."""
    candidates = [local_path, legacy_path]
    worktree_pattern = os.path.join(
        WORKSPACE_ROOT,
        ".claude",
        "worktrees",
        "*",
        "page_indexing_RAG",
        "data",
        "pdf_data",
    )
    candidates.extend(glob.glob(worktree_pattern))

    best_path = None
    best_count = -1
    for path in candidates:
        if not os.path.isdir(path):
            continue
        count = _count_recursive_pdfs(path)
        if count > best_count:
            best_path = path
            best_count = count

    if best_path is not None:
        print(f"Using PDF directory with {best_count} PDFs: {best_path}")
        return best_path

    return local_path


PDF_DATA_DIR = _resolve_pdf_dir(
    os.path.join(PROJECT_ROOT, "data", "pdf_data"),
    os.path.join(WORKSPACE_ROOT, "pdf_data"),
)
CHROMA_PATH = _resolve_dir(
    os.path.join(PROJECT_ROOT, "data", "chroma_persistent_storage"),
    os.path.join(WORKSPACE_ROOT, "chroma_persistent_storage"),
)
_CHROMA_PATH_OVERRIDE = os.getenv("CHROMA_PATH", "").strip()
if _CHROMA_PATH_OVERRIDE:
    CHROMA_PATH = _CHROMA_PATH_OVERRIDE
TXT_DATA_DIR = _resolve_dir(
    os.path.join(PROJECT_ROOT, "data", "Data"),
    os.path.join(WORKSPACE_ROOT, "Data"),
)
HTML_DATA_DIR = _resolve_dir(
    os.path.join(PROJECT_ROOT, "data", "HTML_Data"),
    os.path.join(WORKSPACE_ROOT, "HTML_Data"),
)

os.makedirs(CHROMA_PATH, exist_ok=True)

# ════════════════════════════════════════════════════════════════════════════════
# SYSTEM PROMPT
# ════════════════════════════════════════════════════════════════════════════════
MASTER_SYSTEM_PROMPT = """
You are an expert AI assistant for Analog Devices Inc. (ADI),
specializing in semiconductor test engineering and chip validation on the Teradyne UltraFlex platform.
You also function as a general-purpose technical search engine: when real-time web results are
provided, you can answer questions on ANY engineering, programming, electronics, or general
technical topic — not just ADI/ATE-specific content.

You have deep expertise in:
- Teradyne UltraFlex ATE platform and IG-XL test development software
- UltraFlex instrument cards (HDVS, HSD1000, UltraPIN1600, DCVI, HexVS, PPMU, etc.)
- TDC (Test Development Cookbook) documentation
- ADI device families: ADuCM410, ADIN6310, ADIN7210, AMC038X
- Device characterization, validation, and new product introduction (NPI)
- HIB schematic verification and pin mapping
- HTOL, HAST, BHAST reliability test boards
- SPI, I2C, TSN communication protocols on ADI devices
- IG-XL test program development, debugging, and optimization
- Pattern generation and timing setup for digital testing
- DC parametric testing and analog measurements

You receive context from up to THREE sources:
(A) INTERNAL KNOWLEDGE BASE - private/proprietary docs, manuals, internal PDFs
(B) PUBLIC WEB SOURCES - live web search results fetched in real-time from across the internet
(C) DIRECTLY UPLOADED FILES - files the user attached in this session

WEB SEARCH CAPABILITY:
- When PUBLIC WEB SOURCES are present, treat them as current, real-time information from the web.
- You can answer questions about ANY topic using web results: general electronics, programming
  languages, protocols, standards, market data, tutorials, documentation from any vendor, etc.
- Web results are ranked by relevance and source authority. Use them confidently.
- Do NOT restrict yourself to ADI topics when web search results are available.

ANSWERING RULES:
1. Check UPLOADED FILES first — they have the highest priority.
2. Check INTERNAL KNOWLEDGE BASE — cite as: [Source: <filename>, Page <N>]
3. Check PUBLIC WEB SOURCES — cite as: [Web: <url>]
4. Clearly label which source each piece of information came from.
5. NEVER invent numbers, specs, or register values. If not found anywhere, say so explicitly.
6. For TML issues — always return corrected code in a code block.
7. For specs — always include units, ranges, and conditions.
8. Flag any potential issues or risks you notice.
9. Suggest ADI best practices where relevant to the question.
10. If no source has the answer, state:
    "This information was not found in the internal knowledge base or web search results."

Always be precise, technical, and concise.
"""

# ════════════════════════════════════════════════════════════════════════════════
# LANGCHAIN FACTORY HELPERS
# ════════════════════════════════════════════════════════════════════════════════
def get_chat_llm(model: str | None = None, temperature: float = 0.1):
    """Return a LangChain ChatOpenAI configured for Portkey or direct OpenAI."""
    from langchain_openai import ChatOpenAI

    model = model or OPENAI_ANSWER_MODEL
    if USE_PORTKEY:
        return ChatOpenAI(
            model=model,
            temperature=temperature,
            openai_api_key="dummy",
            openai_api_base=PORTKEY_BASE_URL,
            default_headers=_portkey_headers,
            http_client=httpx.Client(
                verify=False,
                headers=_portkey_headers,
                timeout=120,
            ),
        )
    return ChatOpenAI(
        model=model,
        temperature=temperature,
        openai_api_key=openai_key,
        http_client=httpx.Client(verify=False),
    )


def get_embeddings():
    """Return a LangChain OpenAIEmbeddings configured for Portkey or direct OpenAI."""
    from langchain_openai import OpenAIEmbeddings

    if USE_PORTKEY:
        return OpenAIEmbeddings(
            model=OPENAI_EMBEDDING_MODEL,
            tiktoken_model_name="text-embedding-3-large",
            openai_api_key="dummy",
            openai_api_base=PORTKEY_BASE_URL,
            default_headers=_portkey_headers,
            http_client=httpx.Client(
                verify=False,
                headers=_portkey_headers,
                timeout=120,
            ),
        )
    return OpenAIEmbeddings(
        model=OPENAI_EMBEDDING_MODEL,
        tiktoken_model_name="text-embedding-3-large",
        openai_api_key=openai_key,
        http_client=httpx.Client(verify=False),
    )
