# AGENTS.md

Code-first guidance for agents working in page_indexing_RAG.

## Purpose

This file defines a pure coding-agent mode:

- Focus on code, tests, and file changes.
- Prefer direct implementation over long planning.
- Keep edits minimal, explicit, and verifiable.
- Avoid product chatter and non-coding workflows.

## Scope

Only modify files in page_indexing_RAG unless the task explicitly asks for cross-app changes.

Do not mix logic or dependencies across workspace apps:

- page_indexing_RAG (this app)
- chm_rag_pipeline
- root utility scripts

## Default Working Style

1. Understand the request and find the exact target files.
2. Make the smallest correct code change.
3. Run focused validation (tests, lint, or direct run path).
4. Report what changed, what was verified, and any remaining risk.

When unsure between broad refactor and targeted fix, choose targeted fix.

## Code Change Rules

- Preserve existing architecture unless the task asks for redesign.
- Do not rename public APIs or move modules without clear need.
- Do not edit .env or expose secrets.
- Keep new text ASCII unless a file already requires Unicode.
- Avoid touching generated/runtime data.

Prefer small, reviewable diffs over wide cleanup.

## Architecture Guardrails

This repo has two backend paths. Do not cross-wire them accidentally.

- Original path:
  - app/streamlit_app.py
  - src/page_indexing_rag/config.py
  - src/page_indexing_rag/ingestion.py
  - src/page_indexing_rag/retrieval.py
  - src/page_indexing_rag/generation.py
- Agent SDK path:
  - app/streamlit_app_agentsdk.py
  - src/page_indexing_rag/config_agentsdk.py
  - src/page_indexing_rag/rag_agent.py
  - src/page_indexing_rag/ingestion_agentsdk.py
  - src/page_indexing_rag/retrieval_agentsdk.py
  - src/page_indexing_rag/generation_agentsdk.py

Be explicit about which path you are editing.

## Run and Validate

Run from page_indexing_RAG.

PowerShell:

```powershell
.\scripts\tasks.ps1 -Task run
.\scripts\tasks.ps1 -Task test
.\scripts\tasks.ps1 -Task check
```

Direct commands:

```powershell
python -m streamlit run app\streamlit_app.py
python -m streamlit run app\streamlit_app_agentsdk.py
python -m pytest tests -q
python -m pytest tests -q -p no:cacheprovider
```

Validate only what is relevant to the change first, then broaden if needed.

## Environment and Security

- Treat API keys and credentials as secrets.
- Do not print .env contents.
- Preserve current SSL-disable behavior used in this environment unless explicitly asked to change it.

For any new Python script in this workspace, include the required SSL compatibility block:

```python
import os, ssl, httpx
os.environ["PYTHONHTTPSVERIFY"] = "0"
os.environ["CURL_CA_BUNDLE"] = ""
os.environ["REQUESTS_CA_BUNDLE"] = ""
ssl._create_default_https_context = ssl._create_unverified_context
```

And pass http_client=httpx.Client(verify=False) to OpenAI clients.

## Data and Persistence

Prefer project-local data paths under data/.

Avoid modifying runtime artifacts unless task-relevant:

- data/chroma_persistent_storage
- data/chroma_runtime_storage
- data/chat_history
- __pycache__ and test cache folders

## Known Constraints

- pyproject.toml should use requires-python >=3.10, not >=3.14.
- rag_agent.py web search is placeholder behavior unless explicitly reworked.
- Keep new UI text ASCII to avoid mojibake spread.

## Non-Goals

- Do not add process-heavy "agent mode" instructions.
- Do not document speculative features as implemented.
- Do not perform broad cleanup unrelated to the user request.
