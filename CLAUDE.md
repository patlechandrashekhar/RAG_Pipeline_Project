# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**page-indexing-rag** is a Streamlit-based RAG (Retrieval-Augmented Generation) application designed for semiconductor validation workflows. It integrates Analog Devices documentation (test programs, TDC, datasheets, schematics) with LLM backends and ChromaDB vector storage, providing hybrid internal KB + web fallback retrieval.

**Domain Context**: The app targets semiconductor test engineers working with Teradyne UltraFlex ATE platform, ADI devices, and related validation documentation.

**Dual Backend Architecture**: The application has two parallel implementations:
- **Original path**: OpenAI/Portkey-based (`streamlit_app.py`, `config.py`, `ingestion.py`, `retrieval.py`, `generation.py`)
- **Agent SDK path**: Anthropic Claude + AWS Bedrock Titan-based (`streamlit_app_agentsdk.py`, `config_agentsdk.py`, `rag_agent.py`, `*_agentsdk.py` modules)

When making changes, be explicit about which path you're modifying. Avoid mixing assumptions between the two paths.

## Architecture

### Original Path (OpenAI/Portkey)

The original pipeline follows a **4-layer modular design**:

1. **Ingestion** (`ingestion.py`): PDF/HTML/CHM/TXT parsing, semantic chunking, file classification (TML, TDC, datasheet, schematic, etc.), deduplication via MD5, ChromaDB writes
2. **Storage** (`config.py`): ChromaDB collection management (`tml_copilot_v2`), path resolution (prefers local `data/` but falls back to legacy workspace paths), OpenAI/Portkey client setup
3. **Retrieval** (`retrieval.py`): Query expansion, vector similarity search, optional DuckDuckGo web fallback, intent-aware routing
4. **Generation** (`generation.py`): Prompt assembly, complexity-based model routing, system prompt with semiconductor domain expertise

**Key Configuration**:
- Uses Portkey gateway by default (`USE_PORTKEY=true`)
- Falls back to direct OpenAI when `USE_PORTKEY=false`
- Collection name: `tml_copilot_v2`

### Agent SDK Path (Anthropic Claude + AWS Bedrock)

The Agent SDK pipeline uses async orchestration:

1. **Ingestion** (`ingestion_agentsdk.py`): Same parsing/chunking but optimized for Titan embeddings (1024 dimensions vs OpenAI's 3072)
2. **Storage** (`config_agentsdk.py`): Direct Anthropic and AWS Bedrock clients, ChromaDB collection (`tml_copilot_v3_titan`)
3. **Retrieval** (`retrieval_agentsdk.py`): MMR deduplication, confidence checks, query type classification, placeholder web search
4. **Generation** (`generation_agentsdk.py`): Context assembly helpers, Claude-specific prompting
5. **Orchestration** (`rag_agent.py`): Core async RAG agent coordinating ingestion, retrieval, and generation

**Key Configuration**:
- Uses direct Anthropic SDK (no Portkey)
- AWS Bedrock for Titan embeddings
- Collection name: `tml_copilot_v3_titan`
- Claude Haiku for query expansion, Claude Opus for main responses

### Shared Design Patterns

- **File Classification**: Classifies uploaded/ingested files by extension + content heuristics → type-specific handling (e.g., TML gets code formatting, datasheets get structured specs)
- **Hybrid Retrieval**: Tries internal vector DB first; on low confidence, queries web if available (DuckDuckGo in original, placeholder in Agent SDK)
- **Semantic Chunking**: Splits by sentences with token-aware overlap (default 400 tokens, 2-sentence overlap) to preserve context across chunk boundaries
- **Chat Persistence**: Session chat history saved to JSON at `data/chat_history/chipagent_chat_history.json`
- **Path Fallback**: Local `page_indexing_RAG/data/` paths preferred; if empty, falls back to legacy workspace root paths (`C:\AI Projects\pdf_data`, `C:\AI Projects\Data`, etc.) for backward compatibility
- **Document Generation**: `docx_builder.py` provides professional Word document generation with tables, code blocks, callouts, and styling

## Development Commands

From `page_indexing_RAG/` directory:

**Unix/Linux/Mac (Make)**:
```bash
make setup    # Install dependencies + pytest
make run      # Launch Streamlit app (original path by default)
make test     # Run pytest tests (output: quiet mode)
make check    # Validate Python syntax AST on app + src modules
```

**Windows (PowerShell)**:
```powershell
.\scripts\tasks.ps1 -Task setup
.\scripts\tasks.ps1 -Task run
.\scripts\tasks.ps1 -Task test
.\scripts\tasks.ps1 -Task check
```

**Direct Commands**:
```bash
# Original path (OpenAI/Portkey)
python -m streamlit run app/streamlit_app.py

# Agent SDK path (Claude + Bedrock Titan)
python -m streamlit run app/streamlit_app_agentsdk.py

# Testing
python -m pytest tests -q                          # Run tests quietly
python -m pytest tests -q -p no:cacheprovider     # Without pytest cache
python -m pytest tests/test_layout.py -v          # Run single test file verbosely
```

## Project Structure

```
page_indexing_RAG/
├── app/
│   ├── __init__.py
│   ├── streamlit_app.py           # Original UI (OpenAI/Portkey path)
│   └── streamlit_app_agentsdk.py  # Agent SDK UI (Claude + Titan path)
├── src/page_indexing_rag/
│   ├── __init__.py
│   │
│   ├── # Original path modules
│   ├── config.py                  # Env vars, OpenAI/Portkey client, paths, system prompt
│   ├── ingestion.py               # PDF/HTML/CHM parsing, chunking, ChromaDB writes
│   ├── retrieval.py               # Vector search, query expansion, web fallback
│   ├── generation.py              # Prompt building, LLM calls
│   │
│   ├── # Agent SDK path modules
│   ├── config_agentsdk.py         # Anthropic + AWS Bedrock config
│   ├── rag_agent.py               # Core async RAG orchestration
│   ├── ingestion_agentsdk.py      # Parsing/chunking for Titan embeddings
│   ├── retrieval_agentsdk.py      # MMR deduplication, confidence checks
│   ├── generation_agentsdk.py     # Context assembly for Claude
│   │
│   ├── # Shared utilities
│   ├── docx_builder.py            # Professional Word document generation
│   ├── models.py                  # Model definitions (Portkey-oriented)
│   └── project_workspace.py       # Path utilities
├── data/
│   ├── pdf_data/                  # Local PDF ingestion folder (preferred)
│   ├── HTML_Data/                 # Local HTML/CHM folder (preferred)
│   ├── Data/                      # Local text file folder (preferred)
│   ├── chroma_persistent_storage/ # ChromaDB vector DB (preferred local)
│   └── chat_history/              # Chat persistence (auto-created)
├── tests/
│   ├── test_layout.py             # Directory structure validation
│   ├── test_modules.py            # Module imports/syntax
│   └── test_config.py             # Configuration module tests
├── scripts/
│   ├── tasks.ps1                  # Windows PowerShell task runner
│   └── run_app.ps1                # Quick Windows app launcher
├── docs/
│   └── ARCHITECTURE.md            # Architecture notes and planned agentic split
├── Makefile                       # Default task runner
├── pyproject.toml                 # Project metadata + dependencies
├── requirements.txt               # Pip install list
├── .env.example                   # Template for original path
├── .env.agentsdk                  # Template for Agent SDK path
├── migrate_to_titan.py            # Migration script for existing ChromaDB data
└── README_AGENT_SDK.md            # Agent SDK-specific documentation
```

## Configuration & Environment

### Original Path Setup

1. Copy `.env.example` → `.env`
2. Set required variables:
   - `OPENAI_API_KEY` (required when `USE_PORTKEY=false`)
   - `PORTKEY_API_KEY` (required when `USE_PORTKEY=true`, default)
   - `PORTKEY_PROVIDER` (e.g., `openai`, `anthropic`)
   - `PORTKEY_BASE_URL` (gateway URL)
3. Optional: override model defaults
   - `OPENAI_ANSWER_MODEL` (default: `gpt-4.1`)
   - `OPENAI_RETRIEVAL_MODEL` (default: inherits answer model)
   - `OPENAI_WEB_MODEL` (default: inherits answer model)
   - `OPENAI_EMBEDDING_MODEL` (default: `text-embedding-3-large`)

**Key config (`config.py`)**:
- Uses Portkey gateway by default (`USE_PORTKEY=true`)
- SSL verification disabled (legacy behavior: `PYTHONHTTPSVERIFY=0`, `CURL_CA_BUNDLE=""`)
- `PROJECT_ROOT` = `src/../..` (resolves to `page_indexing_RAG/`)
- `WORKSPACE_ROOT` = `page_indexing_RAG/../` (resolves to `C:\AI Projects\`)
- Master system prompt hardcoded (ADI semiconductor domain focus)

### Agent SDK Path Setup

1. Copy `.env.agentsdk` → `.env`
2. Set required variables:
   - `ANTHROPIC_API_KEY` - Get from [Anthropic Console](https://console.anthropic.com/)
   - `AWS_ACCESS_KEY_ID` - AWS credentials for Bedrock
   - `AWS_SECRET_ACCESS_KEY` - AWS secret key
   - `AWS_REGION` - Default: `us-east-1`

**Key config (`config_agentsdk.py`)**:
- Direct Anthropic SDK calls (no Portkey)
- AWS Bedrock for Titan embeddings (1024 dimensions)
- SSL verification disabled (preserved from original)
- Claude Haiku for query expansion (fast & cheap)
- Claude Opus for main responses (high quality)
- Uses `tml_copilot_v3_titan` ChromaDB collection

### Security Notes

Do not print or inspect `.env` files unless explicitly asked. Treat API keys, AWS credentials, and Portkey keys as secrets. Both configuration modules preserve legacy SSL-disable behavior for corporate network compatibility - do not change casually.

## File Ingestion & Data Handling

**Supported Formats**:
- **PDF**: `pdfplumber` for text extraction, `pypdf` for metadata, page-level chunking
- **HTML/XHTML**: BeautifulSoup parsing, link preservation
- **CHM**: `subprocess` + `hh.exe` (Windows-only) extraction to temp, then HTML parsing
- **TXT**: Plain text chunking

**File Classification Heuristics** (`classify_file()`):
- **Test Programs**: Contains "tester", "testsuite", "testflow", "testmethod" keywords (supports both IG-XL and legacy formats)
- **TDC**: Filename contains "tdc" OR content has "cookbook"
- **Datasheet**: Filename contains "datasheet" OR content has "absolute max", "electrical characteristics"
- **Schematic/Pinmap**: Filename contains "hib"/"pin" OR content has "schematic"/"pinmap"
- **RTL**: `.sv`/`.v` or content has "module"
- Defaults to general if no matches

**Deduplication**: MD5 hash of content → skip re-ingestion of identical files

**Chunking**:
- **Semantic chunking**: Split on sentence boundaries, accumulate sentences until token budget (default 400 tokens) reached
- **Overlap**: Last 2 sentences of previous chunk overlap into next chunk to preserve context
- **Token counting**: Approximate (word count / 1.3) to estimate embedding tokens

## Key Modules

### Original Path

#### `ingestion.py`
- `classify_file(filename, content)` → file type (tdc, datasheet, tml, etc.)
- `semantic_chunk(text, max_tokens, overlap_sentences)` → list of chunks
- `ingest_pdf(path)` / `ingest_all_pdfs()` → Parse, chunk, embed, store in ChromaDB
- `ingest_uploaded_text_file(uploaded_file)` → Handle runtime file uploads
- `ingest_all_html_documents()` / `ingest_all_text_data()` → Batch ingestion
- `extract_pdf_pages(pdf_path)` → List (page_num, text) tuples
- `is_pdf_ingested(pdf_path)` / `is_file_ingested(filename)` → Check for duplicates
- `get_openai_embedding(text)` → Call OpenAI embedding API

#### `retrieval.py`
- `expand_query(question)` → Generate query variants (reformulations, synonyms)
- `_dedupe_keep_order(values)` → Remove duplicates while preserving order
- `hybrid_query_documents(question, top_k, use_web)` → Internal vector search + optional web fallback
- `HAS_DDGS` → Boolean flag; if False, web fallback unavailable

#### `generation.py`
- `build_file_context_prompt(filename, content, file_type)` → Type-specific wrapping (e.g., TML code block formatting)
- `classify_question_complexity(question)` → simple/medium/complex (based on length + keywords)
- `route_model(complexity)` → Select LLM (currently static: uses OPENAI_ANSWER_MODEL)
- `build_rag_context(top_chunks)` → Format chunks with source/page citations
- `generate_hybrid_response(question, internal_context, web_context, file_context)` → Final LLM call

#### `config.py`
- Global `client` (OpenAI/Portkey instance)
- Path constants: `PROJECT_ROOT`, `WORKSPACE_ROOT`, `PDF_DATA_DIR`, `CHROMA_PATH`, `TXT_DATA_DIR`, `HTML_DATA_DIR`
- `MASTER_SYSTEM_PROMPT` (hardcoded domain expertise)
- `_resolve_dir(local_path, legacy_path)` → Fallback logic for backward compatibility

### Agent SDK Path

#### `rag_agent.py`
- Core async RAG orchestration using Anthropic SDK
- `ingest_document()` → Async document ingestion with Titan embeddings
- `query()` → Main async query endpoint coordinating retrieval + generation
- Handles ChromaDB operations with `tml_copilot_v3_titan` collection

#### `config_agentsdk.py`
- `build_anthropic_client()` → Direct Anthropic SDK client
- `build_openai_compatible_client()` → Compatibility layer for Portkey if needed
- `_resolve_dir()` → Path fallback logic (same as original)
- AWS Bedrock configuration for Titan embeddings
- Claude model configuration (Haiku for expansion, Opus for responses)

#### `ingestion_agentsdk.py`, `retrieval_agentsdk.py`, `generation_agentsdk.py`
- Similar APIs to original path modules but optimized for Claude + Titan
- MMR (Maximal Marginal Relevance) deduplication in retrieval
- Async-first design throughout

### Shared Utilities

#### `docx_builder.py`
Professional Word document generation with rich formatting:

**Main API**:
- `build_docx(spec: DocumentSpec, output_path: str)` → Build and save .docx file
- `build_docx_bytes(spec: DocumentSpec)` → Return document as bytes (for Streamlit downloads)

**DocumentSpec Structure**:
- Title page with metadata (title, author, date, confidentiality)
- Auto-generated table of contents
- Headers/footers with page numbers
- Multiple section support (`SectionSpec` with heading levels 1-3)

**Content Blocks**:
- `ParagraphBlock`: Text with formatting (bold, italic, alignment)
- `BulletBlock` / `NumberedBlock`: List formatting
- `TableSpec`: Professional tables with headers, zebra striping, column widths
- `CodeBlock`: Syntax-highlighted code with optional caption
- `CalloutBlock`: Styled callouts (note, warning, tip, important)
- `ImageBlock`: Embedded images with captions

**Styling**: Predefined professional styles (Calibri font, blue color scheme), consistent spacing, page layout (letter/A4)

## Testing & Validation

**Test Structure** (`tests/`):
- `test_layout.py`: Validates directory structure exists (app, src, data, docs, scripts, tests)
- `test_modules.py`: Tests module imports + syntax
- `test_config.py`: Tests configuration module functionality

**Run**:
```bash
# Via task runners
make test                               # pytest tests -q
make check                              # AST syntax validation
.\scripts\tasks.ps1 -Task test          # Windows equivalent

# Direct pytest commands
pytest tests -v                         # Verbose output
pytest tests -q -p no:cacheprovider    # Without pytest cache
pytest tests/test_layout.py -v         # Single file
pytest tests -k "test_config"          # Match test name pattern
```

**Check Command** (`make check`):
- Parses Python AST on `app/streamlit_app.py` and all `src/page_indexing_rag/*.py` files
- Fails if syntax errors present; prints "AST_OK" on success

**Test Coverage**:
```bash
pytest tests --cov=src/page_indexing_rag --cov-report=html
```
Generates HTML coverage report in `htmlcov/`

## Streamlit App (`app/streamlit_app.py`)

**Key Features**:
- **UI Layout**: Title → Divider → Chat interface (messages + input) → Sidebar (controls)
- **Session State**:
  - `chat_history`: List of dicts (role, content, created_at); saved to JSON on each append
  - `chat_file_contexts`: Attached file metadata
- **File Upload**: Sidebar file picker (no extension whitelist); auto-classify; option to ingest into KB
- **Sidebar Controls**:
  - `Ingest All PDFs`, `Ingest HTML Data`, `Load Text Data` buttons
  - Sidebar toggles: "Auto-add attached files to KB"
  - Model config display
- **Chat Flow**: User message → retrieval → web fallback (if enabled) → file context → LLM generation → display + persist
- **Context Limits**:
  - `MAX_FILE_CONTEXT_CHARS = 12000` (file attachment context size)
  - `MAX_TOTAL_CONTEXT_CHARS = 32000` (total retrieval + file context)
- **Chat History**: Persisted to `data/chat_history/chipagent_chat_history.json`; format includes `updated_at`, `message_count`, `messages` array

## Data Migration Between Paths

The two paths use **separate ChromaDB collections** with different embedding dimensions:
- Original path: `tml_copilot_v2` (OpenAI embeddings, 3072 dimensions)
- Agent SDK path: `tml_copilot_v3_titan` (Titan embeddings, 1024 dimensions)

**Migration Script** (`migrate_to_titan.py`):
```bash
python migrate_to_titan.py
```
Reads documents from `tml_copilot_v2`, re-embeds with Titan, writes to `tml_copilot_v3_titan`. Batch size: 10 documents.

Collections are **not** automatically synced. If you ingest data in one path, it will not appear in the other unless you run the migration script.

## Domain-Specific Notes

- **Semiconductor Focus**: App is tailored for ADI customers/engineers working with Teradyne UltraFlex ATE, IG-XL test programs, TDC, device datasheets, pin mappings, schematics
- **System Prompt**: Hardcoded in both `config.py` and `config_agentsdk.py`, specifies ADI expertise, UltraFlex platform, IG-XL test development, device families (ADuCM410, ADIN6310, etc.), communication protocols
- **Answer Rules**: System prompt instructs LLM to cite sources [Source: filename, Page N] for internal KB, [Web: url] for web results; forbids invented specs; returns corrected test code in blocks
- **Fallback Behavior**: 
  - Original path: Tries web search (DuckDuckGo) if internal confidence is low
  - Agent SDK path: Has placeholder web search (not fully wired)
  - Both report "not found in sources" if neither internal nor web retrieval succeeds

## Important Caveats

1. **Dual Backend Isolation**: The two paths (Original vs Agent SDK) have separate collections and do not share data automatically. Use `migrate_to_titan.py` to sync.

2. **Model Mixing**: `models.py` contains Portkey-oriented model definitions still used by original path. Check imports before removing.

3. **Web Search Status**: 
   - Original path: DuckDuckGo fully wired
   - Agent SDK path: Placeholder implementation only

4. **Encoding Issues**: Some UI strings in `streamlit_app_agentsdk.py` may display mojibake from previous encoding issues. Don't propagate to new files.

5. **Git Repository**: This folder may not be initialized as a Git repository. Don't rely on Git commands.

6. **Data Files**: Avoid committing:
   - ChromaDB files (`data/chroma_persistent_storage/`, `data/chroma_runtime_storage/`)
   - Chat history JSON (`data/chat_history/`)
   - `__pycache__`, pytest cache folders

## Legacy Compatibility

- **Path Fallback**: Both config modules prefer local `data/` paths but fall back to legacy workspace paths (`C:\AI Projects\pdf_data`, `C:\AI Projects\Data`, `C:\AI Projects\chroma_persistent_storage`) when local paths are empty
- This allows old scripts/workflows to continue working without modification
- Preserve this compatibility unless explicitly removing it
