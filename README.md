# Page Indexing RAG

Streamlit-based RAG assistant for semiconductor validation workflows (TML, TDC, datasheets, and mixed internal/public retrieval).

## Project Layout

```text
page_indexing_RAG/
  app/
    __init__.py
    streamlit_app.py         # Main Streamlit application
  src/
    page_indexing_rag/
      __init__.py
      config.py              # Env/client/path configuration
      ingestion.py           # Parsing/chunking/embedding/Chroma writes
      retrieval.py           # Query expansion + retrieval + web fallback
      generation.py          # Prompt assembly + response generation
  data/
    pdf_data/                # Preferred local PDF source folder
    HTML_Data/               # Preferred local HTML/CHM source folder
    Data/                    # Preferred local TXT source folder
    chroma_persistent_storage/  # Preferred local Chroma DB path
  docs/
  scripts/
    tasks.ps1               # One-command setup/run/test/check
    run_app.ps1             # Shortcut to run app
  tests/
  Makefile
  pageIndexing_RAG.py        # Backward-compatible legacy entrypoint
  pyproject.toml
  requirements.txt
  .env.example
```

## Run

From `C:\AI Projects\page_indexing_RAG`:

```powershell
streamlit run app/streamlit_app.py
```

Or use task runner:

```powershell
.\scripts\tasks.ps1 -Task run
```

Legacy command still works:

```powershell
streamlit run pageIndexing_RAG.py
```

## Environment

1. Copy `.env.example` to `.env`.
2. Set `OPENAI_API_KEY`.
3. Optional quality/cost knobs:
   - `OPENAI_ANSWER_MODEL` (default: `gpt-4.1`)
   - `OPENAI_RETRIEVAL_MODEL` (default: inherits answer model)
   - `OPENAI_WEB_MODEL` (default: inherits answer model)
   - `OPENAI_EMBEDDING_MODEL` (default: `text-embedding-3-large`)

## Chat Attachments

- File attachment is integrated into the chat input (ChatGPT-style).
- The chat picker accepts any file type (no extension whitelist).
- Attach files and ask in the same message, or attach first and ask later.
- Sidebar toggle `Auto-add attached files to KB` controls whether attached files are ingested into Chroma automatically.

## Local Knowledge Sources

- `data/pdf_data`: use sidebar button `Ingest All PDFs`
- `data/HTML_Data`: use sidebar button `Ingest HTML Data` (supports `.html/.htm/.xhtml` and `.chm`)
- `data/Data`: use sidebar button `Load Text Data`

## Developer Tasks

PowerShell:

```powershell
.\scripts\tasks.ps1 -Task setup
.\scripts\tasks.ps1 -Task run
.\scripts\tasks.ps1 -Task test
.\scripts\tasks.ps1 -Task check
```

Make:

```bash
make setup
make run
make test
make check
```

## Data Path Behavior

The app prefers local project folders under `data/`.
If those are empty but legacy workspace folders exist (`C:\AI Projects\pdf_data`, `C:\AI Projects\Data`, `C:\AI Projects\chroma_persistent_storage`), it automatically falls back to them.
