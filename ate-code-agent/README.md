# ATE Code Agent

Cursor-style AI agent for Teradyne UltraFLEX / IG-XL test programs.
Built with Claude Agent SDK.

## What it does

- Reads IG-XL Excel sheets (Test Instances, Limits, Flow, Levels, Timing)
- Reads and writes VBA modules (Windows + Excel)
- Plans changes, supports human-approval checkpoints, and applies edits
- Runs IG-XL programs in engineering mode
- Parses STDF logs and summarizes pass/fail outcomes
- Creates git backup commits before edits

## Project structure

- main.py: CLI entrypoint
- agent.py: Claude Agent SDK loop and tool registration
- system_prompt.py: ATE domain prompt
- safety/rules.py: hard safety guardrails
- tools/: excel/vba/run/log/git tool modules
- api/: optional FastAPI wrapper
- tests/: unit tests for core tools

## Quick start

1. Install dependencies:

```powershell
pip install -r requirements.txt
```

2. Set your API key in .env:

```env
ANTHROPIC_API_KEY=your_api_key_here
```

3. Run tests:

```powershell
pytest tests -v
```

4. Start interactive agent:

```powershell
python main.py C:\path\to\YourProgram.xlsx
```

## Safety

- Pin Map and Spec sheets are read-only
- Voltage writes are bounded by safety limits
- Allowed run modes: engineering, characterization
- Git backup enabled before edits
- Max iterations configured in safety rules
- Human approval gate: edit/run style tasks require explicit engineer approval

## Approval workflow

- Interactive mode: when a task implies edits or run, the CLI asks for `APPROVE`.
- API mode: send `approved=true` in request body for edit/run tasks.

## Standalone MCP server for Claude Code

This repo now includes a standalone stdio MCP server:

- [page_indexing_RAG/ate-code-agent/rag_mcp_server.py](page_indexing_RAG/ate-code-agent/rag_mcp_server.py)
- [page_indexing_RAG/ate-code-agent/testprog_mcp_server.py](page_indexing_RAG/ate-code-agent/testprog_mcp_server.py)

It exposes tools:

- `query_ultraflex_docs(question, web_mode="fallback", n_per_query=5)`
- `ping()`

The test-program MCP exposes tools:

- `build_pinmap_package(channel_assignment, project_name="unnamed_dut", default_instrument="UltraPin1600", write_files=False, output_dir="")`
- `generate_igxl_seed_package(channel_assignment, project_name="unnamed_dut", default_instrument="UltraPin1600", default_test_name="DC_SANITY", default_levelset_name="LVL_DEFAULT", default_timingset_name="TIM_DEFAULT", default_period_ns="100", protocol_preset="", write_files=True, output_dir="")`
- `ping()`

`build_pinmap_package` accepts flexible channel assignment formats (JSON, CSV, or tokenized lines) and returns:

- pinmap CSV content
- channel map CSV content
- datasheet Markdown content
- parser/validation warnings and assumptions

`generate_igxl_seed_package` generates IG-XL-ready seed CSV artifacts:

- PinMap seed CSV
- ChannelMap seed CSV
- Limits seed CSV
- Levels seed CSV
- Timing seed CSV

`protocol_preset` can auto-fill timing/level defaults for common buses:

- `SPI`
- `I2C`
- `UART`
- `JTAG`

### Register in Claude Code

1. Use the template in [page_indexing_RAG/ate-code-agent/claude_code_mcp_config.example.json](page_indexing_RAG/ate-code-agent/claude_code_mcp_config.example.json).
2. Copy the `mcpServers.ultraflex-rag` block into your Claude Code MCP config.
3. Restart Claude Code so it discovers the server and tools.

### Quick verification

Ask Claude Code to call `ping` and then `query_ultraflex_docs` with a simple question.

## Chipagent on Linux: ready-to-run setup

Use these files for Linux deployment:

- `ate-code-agent/rag_mcp_server.py` (MCP stdio server)
- `ate-code-agent/chipagent_mcp_config.linux.example.json` (Chipagent config template)
- `ate-code-agent/scripts/start_rag_mcp_linux.sh` (local launcher)
- `ate-code-agent/rag_http_api.py` (optional REST wrapper)
- `ate-code-agent/scripts/start_rag_http_linux.sh` (REST launcher)

### 1) Install runtime dependencies

```bash
cd /opt/page_indexing_RAG
python3 -m venv .venv
source .venv/bin/activate
pip install -U pip
pip install -r requirements.txt
pip install mcp fastapi uvicorn
```

### 2) Configure shared Chroma DB path

Set a common path for all systems that should read the same vector DB:

```bash
export CHROMA_PATH=/mnt/shared/chroma/chroma_persistent_storage
```

`src/page_indexing_rag/config.py` now honors `CHROMA_PATH` when set.

### 3) Force vector-DB-only retrieval

To ensure retrieval is from internal chunks only (no web search), set:

```bash
export ULTRAFLEX_MCP_VECTOR_ONLY=true
```

This forces `web_mode=off` inside the MCP and HTTP wrappers.

### 4) Chipagent MCP configuration

Copy and adapt:

- `ate-code-agent/chipagent_mcp_config.linux.example.json`

### 5) Optional REST API wrapper

Start API server:

```bash
cd /opt/page_indexing_RAG/ate-code-agent
./scripts/start_rag_http_linux.sh
```

Endpoints:

- `GET /health`
- `POST /query`

`POST /query` supports:

- `mode="answer"` for generated final answer
- `mode="retrieve_only"` for chunk/source context only (Chipagent generates final answer)

### Local smoke test script

Run this before opening Claude Code to verify MCP startup quickly:

```powershell
Set-Location "c:\AI Projects\page_indexing_RAG\ate-code-agent"
.\scripts\smoke_test_mcp.ps1
```
