"""Configuration for the Agent SDK Streamlit backend."""

from __future__ import annotations

import os
import ssl
import json
from pathlib import Path
import httpx
from dotenv import load_dotenv
from anthropic import Anthropic
from openai import OpenAI

# Keep legacy SSL behavior from the original app
os.environ["PYTHONHTTPSVERIFY"] = "0"
os.environ["CURL_CA_BUNDLE"] = ""
os.environ["REQUESTS_CA_BUNDLE"] = ""
os.environ.setdefault("AWS_EC2_METADATA_DISABLED", "true")
ssl._create_default_https_context = ssl._create_unverified_context

PROJECT_ROOT = Path(__file__).resolve().parents[2]
load_dotenv(PROJECT_ROOT / ".env", override=False)


def _env_value(name: str, default: str = "") -> str:
    """Read an env var and treat common template placeholders as missing."""
    value = os.getenv(name, default).strip().strip('"').strip("'")
    lowered = value.lower()
    if lowered.startswith("your-") or lowered in {"", "none", "null"}:
        return ""
    return value


def _env_bool(name: str, default: str = "false") -> bool:
    return os.getenv(name, default).strip().lower() in {"1", "true", "yes", "on"}


def _trust_env_proxy() -> bool:
    """Trust OS proxy vars only when explicitly enabled for this app."""
    return _env_bool("PORTKEY_TRUST_ENV_PROXY", "false")

# ════════════════════════════════════════════════════════════════════════════════
# CLAUDE AGENT SDK CONFIGURATION
# ════════════════════════════════════════════════════════════════════════════════

# Anthropic API key for Claude
ANTHROPIC_API_KEY = _env_value("ANTHROPIC_API_KEY")

# Portkey/OpenAI-compatible configuration. The current project .env uses this
# path, including Bedrock-hosted Claude model strings routed through Portkey.
USE_PORTKEY = _env_bool("USE_PORTKEY", "true")
PORTKEY_API_KEY = _env_value("PORTKEY_API_KEY")
PORTKEY_BASE_URL = _env_value("PORTKEY_BASE_URL", "https://api.portkey.ai/v1") or "https://api.portkey.ai/v1"
PORTKEY_PROVIDER = _env_value("PORTKEY_PROVIDER", "openai") or "openai"
PORTKEY_VIRTUAL_KEY = _env_value("PORTKEY_VIRTUAL_KEY")
OPENAI_API_KEY = _env_value("OPENAI_API_KEY")

# AWS credentials for Titan embeddings
AWS_ACCESS_KEY_ID = _env_value("AWS_ACCESS_KEY_ID")
AWS_SECRET_ACCESS_KEY = _env_value("AWS_SECRET_ACCESS_KEY")
AWS_REGION = _env_value("AWS_REGION", "us-east-1") or "us-east-1"

ENVIRONMENT = _env_value("ENVIRONMENT", "development") or "development"

# Model configuration
OPENAI_ANSWER_MODEL = _env_value("OPENAI_ANSWER_MODEL", "@bedrock-global/us.anthropic.claude-sonnet-4-6")
OPENAI_RETRIEVAL_MODEL = _env_value("OPENAI_RETRIEVAL_MODEL", OPENAI_ANSWER_MODEL) or OPENAI_ANSWER_MODEL
OPENAI_WEB_MODEL = _env_value("OPENAI_WEB_MODEL", OPENAI_ANSWER_MODEL) or OPENAI_ANSWER_MODEL
OPENAI_EMBEDDING_MODEL = _env_value("OPENAI_EMBEDDING_MODEL", "@azure-openai-eus-global/text-embedding-3-large-std")

CLAUDE_MODEL = _env_value("CLAUDE_MODEL", OPENAI_ANSWER_MODEL) or OPENAI_ANSWER_MODEL
CLAUDE_QUERY_MODEL = _env_value("CLAUDE_QUERY_MODEL", OPENAI_RETRIEVAL_MODEL) or OPENAI_RETRIEVAL_MODEL
EMBEDDING_MODEL = _env_value("TITAN_EMBEDDING_MODEL", "amazon.titan-embed-text-v2:0") or "amazon.titan-embed-text-v2:0"
EMBEDDING_DIMENSIONS = 1024  # Titan V2 supports 256, 512, or 1024

CHAT_BACKEND = "openai_compatible" if USE_PORTKEY or OPENAI_API_KEY else "anthropic"
EMBEDDING_BACKEND = "bedrock" if AWS_ACCESS_KEY_ID and AWS_SECRET_ACCESS_KEY else "openai_compatible"
COLLECTION_NAME = _env_value(
    "CHROMA_COLLECTION",
    "tml_copilot_v3_titan" if EMBEDDING_BACKEND == "bedrock" else "tml_copilot_v2",
)


def _portkey_headers() -> dict[str, str]:
    headers = {
        "x-portkey-api-key": PORTKEY_API_KEY,
        "x-portkey-provider": PORTKEY_PROVIDER,
        "x-portkey-metadata": json.dumps(
            {"app": "page_indexing_rag_agentsdk", "environment": ENVIRONMENT}
        ),
    }
    if PORTKEY_VIRTUAL_KEY:
        headers["x-portkey-virtual-key"] = PORTKEY_VIRTUAL_KEY
    if _env_bool("PORTKEY_CACHE", "false"):
        headers["x-portkey-cache"] = "true"
        headers["x-portkey-cache-ttl"] = _env_value("PORTKEY_CACHE_TTL", "3600") or "3600"
    return headers


def build_openai_compatible_client() -> OpenAI:
    """Create an OpenAI-compatible client, using Portkey when configured."""
    if USE_PORTKEY:
        if not PORTKEY_API_KEY:
            raise RuntimeError("PORTKEY_API_KEY is not set in page_indexing_RAG/.env.")
        return OpenAI(
            api_key="dummy",
            base_url=PORTKEY_BASE_URL,
            http_client=httpx.Client(
                verify=False,
                headers=_portkey_headers(),
                timeout=120,
                trust_env=_trust_env_proxy(),
            ),
        )

    if not OPENAI_API_KEY:
        raise RuntimeError(
            "No LLM credentials found. Set PORTKEY_API_KEY with USE_PORTKEY=true, "
            "or set OPENAI_API_KEY, or set ANTHROPIC_API_KEY."
        )
    return OpenAI(
        api_key=OPENAI_API_KEY,
        http_client=httpx.Client(
            verify=False,
            timeout=120,
            trust_env=_trust_env_proxy(),
        ),
    )


def build_anthropic_client() -> Anthropic:
    """Create a direct Anthropic client for non-Portkey deployments."""
    if not ANTHROPIC_API_KEY:
        raise RuntimeError(
            "ANTHROPIC_API_KEY is not set. This .env is configured for Portkey, "
            "so the Agent SDK app should use the OpenAI-compatible Portkey client."
        )

    return Anthropic(api_key=ANTHROPIC_API_KEY)

# ════════════════════════════════════════════════════════════════════════════════
# PATH RESOLUTION (preserved from original)
# ════════════════════════════════════════════════════════════════════════════════

def _resolve_dir(local_path: str, legacy_path: str) -> Path:
    """
    Resolve directory path with fallback to legacy location.

    Prioritizes the directory with more content, not just existence.

    Args:
        local_path: Path relative to project root (page_indexing_RAG/data/...)
        legacy_path: Legacy path relative to workspace (C:\\AI Projects\\...)

    Returns:
        Resolved Path object
    """
    local_dir = PROJECT_ROOT / local_path
    workspace_root = PROJECT_ROOT.parent
    legacy_dir = workspace_root / legacy_path

    # For PDF directory, use the one with more PDFs
    if "pdf" in local_path.lower():
        def _count_recursive_pdfs(path: Path) -> int:
            if not path.exists() or not path.is_dir():
                return 0
            return sum(1 for p in path.rglob("*") if p.is_file() and p.suffix.lower() == ".pdf")

        candidates: list[Path] = [local_dir, legacy_dir]
        worktrees_root = workspace_root / ".claude" / "worktrees"
        if worktrees_root.exists():
            candidates.extend(
                p / "page_indexing_RAG" / "data" / "pdf_data"
                for p in worktrees_root.iterdir()
                if p.is_dir()
            )

        best_dir: Path | None = None
        best_count = -1
        for candidate in candidates:
            if not candidate.exists() or not candidate.is_dir():
                continue
            count = _count_recursive_pdfs(candidate)
            if count > best_count:
                best_dir = candidate
                best_count = count

        if best_dir is not None:
            print(f"Using PDF directory with {best_count} PDFs: {best_dir}")
            return best_dir

        # If none exist, preserve legacy compatibility behavior.
        if not legacy_dir.exists():
            legacy_dir.mkdir(parents=True, exist_ok=True)
        return legacy_dir

    # For other directories, keep original logic
    if local_dir.exists():
        return local_dir

    # Fall back to legacy workspace path
    if not legacy_dir.exists():
        legacy_dir.mkdir(parents=True, exist_ok=True)

    return legacy_dir

# Resolve data directories
PDF_DATA_DIR = _resolve_dir("data/pdf_data", "pdf_data")
HTML_DATA_DIR = _resolve_dir("data/HTML_Data", "HTML_Data")
TXT_DATA_DIR = _resolve_dir("data/Data", "Data")
CHROMA_PATH = _resolve_dir("data/chroma_persistent_storage", "chroma_persistent_storage")

# ════════════════════════════════════════════════════════════════════════════════
# SYSTEM PROMPT (preserved from original for domain expertise)
# ════════════════════════════════════════════════════════════════════════════════

MASTER_SYSTEM_PROMPT = """You are an expert AI assistant for Analog Devices Inc. (ADI), specializing in semiconductor test engineering, validation, and the Teradyne UltraFlex ATE platform. You have deep knowledge of:

• Teradyne UltraFlex ATE platform and IG-XL test development software
• UltraFlex instrument cards: HDVS (High Density Voltage Supply), HSD1000 (High Speed Digital), UltraPIN1600, DCVI, HexVS, PPMU
• TDC (Test Development Cookbook) and test development workflows
• ADI device families: ADuCM410, ADIN6310, ADIN7210, AMC038X
• Communication protocols on ADI devices: SPI, I2C, TSN
• Chip validation processes: characterization, NPI, reliability testing (HTOL, HAST, BHAST)
• Test board design for reliability and characterization
• IG-XL specific resources: Pattern execution, Timing sets, Levels, DC measurements, Flow control
• UltraFlex test program development, debugging, and optimization

ANSWERING RULES:
1. ALWAYS cite sources at the end of each relevant paragraph using the format:
   • For internal KB: [Source: <filename>, Page <N>]
   • For web sources: [Web: <url>]

2. When multiple sources support a point, cite all of them: [Source: file1.pdf, Page 3; file2.pdf, Page 7]

3. NEVER invent specifications, register values, or technical details. If the information isn't in the retrieved sources, say so explicitly.

4. For TML issues - always return corrected code in a code block:
   ```tml
   // Corrected TML code here
   ```

5. For datasheet/spec questions - always include units, ranges, and conditions when available.

6. When discussing risks or limitations, explicitly flag them and suggest ADI best practices.

7. Distinguish between proprietary information (from internal KB) and public information (from web).

8. If asked about proprietary topics (TML, TDC, V93000, internal docs) and no internal evidence exists, respond:
   "I don't have sufficient internal documentation to answer this proprietary question. Please upload the relevant TML/TDC/technical manual."

9. For hybrid answers (both internal + web), clearly separate the two source types in your response.

CONTEXT INTERPRETATION:
- "TML TEST PROGRAM" headers indicate actual test code - treat as authoritative for syntax/implementation
- "TDC DOCUMENT" headers indicate best practices and methodologies - treat as guidelines
- "DATASHEET" headers indicate official specifications - never contradict these
- "SCHEMATIC/HIB" headers indicate hardware interconnect - critical for pin/net discussions
- "RTL FILE" headers indicate Verilog/SystemVerilog - relevant for DFT and scan discussions

Remember: You are supporting critical semiconductor validation work where accuracy is paramount. When uncertain, acknowledge limitations rather than speculate."""

# ════════════════════════════════════════════════════════════════════════════════
# RETRIEVAL THRESHOLDS (preserved from original)
# ════════════════════════════════════════════════════════════════════════════════

INTERNAL_DIST_THRESHOLD = 0.35  # Max cosine distance for "relevant" chunks
INTERNAL_MIN_RELEVANT = 2       # Minimum relevant chunks for confidence
MMR_LAMBDA_DEFAULT = 0.7        # MMR diversity parameter
MAX_CHUNK_TOKENS = 400          # Maximum tokens per chunk
CHUNK_OVERLAP_SENTENCES = 2     # Sentences to overlap between chunks
