from __future__ import annotations

# ── SSL bypass (corporate self-signed certs) ────────────────────────────────
# Must be done BEFORE any network-touching import (tiktoken, httpx, requests…)
import os
import ssl

os.environ["PYTHONHTTPSVERIFY"] = "0"
os.environ["CURL_CA_BUNDLE"] = ""
os.environ["REQUESTS_CA_BUNDLE"] = ""
os.environ["SSL_CERT_FILE"] = ""
ssl._create_default_https_context = ssl._create_unverified_context

# Patch requests.get/Session.send so tiktoken BPE downloads skip SSL verify
import requests
import requests.adapters

_orig_send = requests.adapters.HTTPAdapter.send


def _ssl_free_send(self, request, **kwargs):
    kwargs["verify"] = False
    return _orig_send(self, request, **kwargs)


requests.adapters.HTTPAdapter.send = _ssl_free_send
# ────────────────────────────────────────────────────────────────────────────

import json
import inspect
import sys
import tempfile
import asyncio
import re
import io
from pathlib import Path
import importlib
from collections import Counter
from datetime import datetime, timedelta
import uuid
from types import SimpleNamespace
from typing import Any

import streamlit as st

chat_input_fileupload = None
HAS_CHAT_INPUT_FILEUPLOAD = False
try:
    _chat_upload_module = importlib.import_module("streamlit_chat_input_fileupload")
    chat_input_fileupload = getattr(_chat_upload_module, "chat_input_fileupload", None)
    HAS_CHAT_INPUT_FILEUPLOAD = callable(chat_input_fileupload)
except Exception:
    chat_input_fileupload = None
    HAS_CHAT_INPUT_FILEUPLOAD = False

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = PROJECT_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

# Force reload of modules to pick up changes (config first, then dependents)
if 'page_indexing_rag.config' in sys.modules:
    import page_indexing_rag.config
    importlib.reload(page_indexing_rag.config)
if 'page_indexing_rag.ingestion' in sys.modules:
    import page_indexing_rag.ingestion
    importlib.reload(page_indexing_rag.ingestion)
if 'page_indexing_rag.retrieval' in sys.modules:
    import page_indexing_rag.retrieval
    importlib.reload(page_indexing_rag.retrieval)
if 'page_indexing_rag.generation' in sys.modules:
    import page_indexing_rag.generation
    importlib.reload(page_indexing_rag.generation)
if 'page_indexing_rag.rag_agent_sdk' in sys.modules:
    import page_indexing_rag.rag_agent_sdk
    importlib.reload(page_indexing_rag.rag_agent_sdk)

_SESSIONS_PATH = PROJECT_ROOT / "data" / "chat_history" / "sessions.json"


def load_sessions() -> dict:
    """Load all chat sessions from disk; return empty dict on failure."""
    try:
        if _SESSIONS_PATH.exists():
            data = json.loads(_SESSIONS_PATH.read_text(encoding="utf-8"))
            return data
    except Exception:
        pass
    return {"sessions": {}}


def save_sessions(sessions: dict) -> None:
    """Save all chat sessions to disk."""
    try:
        _SESSIONS_PATH.parent.mkdir(parents=True, exist_ok=True)
        _SESSIONS_PATH.write_text(
            json.dumps(sessions, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
    except Exception:
        pass


def get_current_messages() -> list:
    """Return messages for the current session."""
    if "current_session_id" not in st.session_state:
        return []

    session_id = st.session_state.current_session_id
    sessions = st.session_state.sessions_index.get("sessions", {})

    if session_id not in sessions:
        return []

    return sessions[session_id].get("messages", [])


def new_session() -> str:
    """Create a new chat session and return its ID."""
    session_id = str(uuid.uuid4())
    now = datetime.now().isoformat()

    new_session_data = {
        "id": session_id,
        "title": "New Chat",
        "created_at": now,
        "updated_at": now,
        "messages": []
    }

    if "sessions" not in st.session_state.sessions_index:
        st.session_state.sessions_index["sessions"] = {}

    st.session_state.sessions_index["sessions"][session_id] = new_session_data
    save_sessions(st.session_state.sessions_index)

    return session_id


def delete_session(session_id: str) -> None:
    """Remove a session from the sessions dict and save."""
    sessions = st.session_state.sessions_index.get("sessions", {})

    if session_id in sessions:
        del sessions[session_id]
        st.session_state.sessions_index["sessions"] = sessions
        save_sessions(st.session_state.sessions_index)

        # If we deleted the current session, switch to another or create new
        if st.session_state.current_session_id == session_id:
            if sessions:
                # Switch to the most recently updated session
                st.session_state.current_session_id = max(
                    sessions.keys(),
                    key=lambda k: sessions[k].get("updated_at", "")
                )
            else:
                # No sessions left, create a new one
                st.session_state.current_session_id = new_session()


def auto_title_session(session_id: str, first_user_msg: str) -> None:
    """Set session title from first user message (max 40 chars)."""
    sessions = st.session_state.sessions_index.get("sessions", {})

    if session_id in sessions:
        # Only auto-title if it's still "New Chat"
        if sessions[session_id]["title"] == "New Chat":
            title = first_user_msg[:40]
            if len(first_user_msg) > 40:
                title += "..."
            sessions[session_id]["title"] = title
            st.session_state.sessions_index["sessions"] = sessions
            save_sessions(st.session_state.sessions_index)

from page_indexing_rag.config import PDF_DATA_DIR, TXT_DATA_DIR
from page_indexing_rag.models import LLM_MODELS, DEFAULT_LLM_MODEL
from page_indexing_rag.document_generator import DocumentGenerator
from page_indexing_rag.generation import (
    build_file_context_prompt,
    generate_code_agent_hybrid_response,
    generate_hybrid_response,
)
from page_indexing_rag.agentic_loop import run_agentic_retrieve_loop
from page_indexing_rag.ingestion import (
    classify_file,
    collection,
    extract_pdf_pages,
    ingest_assets_from_directory,
    get_pdf_folders,
    ingest_all_pdfs,
    ingest_pdf,
    ingest_pdfs_from_folder,
    ingest_uploaded_text_file,
    is_pdf_ingested,
)
from page_indexing_rag.retrieval import HAS_DDGS, hybrid_query_documents
from page_indexing_rag.rag_agent_sdk import SemiconductorRAGAgentSDK


_CODE_AGENT_EXTENSIONS = {
    ".bas", ".cls", ".vb", ".vbs", ".tml", ".txt", ".csv", ".json", ".xml",
    ".yaml", ".yml", ".ini", ".cfg", ".log", ".py", ".md",
}


def _collect_code_agent_workspace(paths: list[str], max_files: int = 120, max_chars_per_file: int = 3500) -> tuple[str, dict]:
    """Scan attached directories and build compact context for code-agent debugging."""
    all_files: list[Path] = []
    missing: list[str] = []

    for raw in paths:
        if not raw.strip():
            continue
        p = Path(raw.strip())
        if not p.exists() or not p.is_dir():
            missing.append(str(p))
            continue

        for fp in p.rglob("*"):
            if not fp.is_file():
                continue
            if fp.suffix.lower() not in _CODE_AGENT_EXTENSIONS:
                continue
            all_files.append(fp)

    all_files = sorted(all_files)[:max_files]
    parts: list[str] = []
    files_meta: list[dict] = []

    for fp in all_files:
        try:
            raw = fp.read_bytes()
            text = raw.decode("utf-8", errors="ignore")
            snippet = text[:max_chars_per_file]
            files_meta.append({
                "path": str(fp),
                "size": len(raw),
                "chars": len(text),
            })
            parts.append(
                f"=== FILE: {fp} ===\n"
                f"SIZE_BYTES: {len(raw)}\n"
                f"CONTENT:\n{snippet}\n"
            )
        except Exception as exc:
            parts.append(f"=== FILE: {fp} ===\n[Read error: {exc}]\n")

    summary = {
        "scanned_files": len(all_files),
        "missing_paths": missing,
        "max_files": max_files,
        "max_chars_per_file": max_chars_per_file,
        "files": files_meta,
    }

    return "\n\n".join(parts), summary


def _extract_directory_paths_from_text(text: str) -> list[str]:
    """Extract potential directory paths from free-form user text."""
    if not text:
        return []

    candidates: list[str] = []

    # Quoted absolute Windows paths (supports spaces)
    for m in re.findall(r"[\"']([A-Za-z]:\\[^\"']+)[\"']", text):
        candidates.append(m.strip())

    # Unquoted absolute Windows paths up to common delimiters
    for m in re.findall(r"([A-Za-z]:\\[^\n,;]+)", text):
        candidates.append(m.strip())

    # Keep existing order while de-duplicating and validating directories
    deduped: list[str] = []
    seen: set[str] = set()
    for raw in candidates:
        p = Path(raw)
        if p.exists() and p.is_dir():
            norm = str(p)
            if norm not in seen:
                seen.add(norm)
                deduped.append(norm)

    return deduped


def _default_code_agent_project() -> dict:
    """Create a blank code-agent project workspace."""
    return {
        "paths": "",
        "workspace_context": "",
        "workspace_summary": None,
        "workspace_sig": "",
        "messages": [],
        "notes": "",
        "attached_file_paths": [],
    }


def _ensure_code_agent_projects() -> None:
    """Ensure multi-project state exists and migrate legacy single-project fields."""
    if "code_agent_projects" not in st.session_state or not isinstance(st.session_state.code_agent_projects, dict):
        st.session_state.code_agent_projects = {}

    projects = st.session_state.code_agent_projects
    if not projects:
        projects["Project 1"] = {
            "paths": st.session_state.get("code_agent_paths", ""),
            "workspace_context": st.session_state.get("code_agent_workspace_context", ""),
            "workspace_summary": st.session_state.get("code_agent_workspace_summary"),
            "workspace_sig": st.session_state.get("code_agent_workspace_sig", ""),
            "messages": st.session_state.get("code_agent_messages", []),
            "notes": "",
        }

    for name, project in projects.items():
        if not isinstance(project, dict):
            projects[name] = _default_code_agent_project()
            continue
        project.setdefault("paths", "")
        project.setdefault("workspace_context", "")
        project.setdefault("workspace_summary", None)
        project.setdefault("workspace_sig", "")
        project.setdefault("messages", [])
        project.setdefault("notes", "")
        project.setdefault("attached_file_paths", [])

    if (
        "code_agent_active_project" not in st.session_state
        or st.session_state.code_agent_active_project not in projects
    ):
        st.session_state.code_agent_active_project = next(iter(projects.keys()))


def _pick_directory_via_dialog(initial_dir: str | None = None) -> str | None:
    """Open a native directory picker and return selected folder path."""
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        selected = filedialog.askdirectory(
            initialdir=initial_dir or str(PROJECT_ROOT.parent),
            title="Select directory to attach",
        )
        root.destroy()
        return selected or None
    except Exception:
        return None


def _project_slug(name: str) -> str:
    return re.sub(r"[^A-Za-z0-9_.-]+", "_", name).strip("_") or "project"


def _persist_code_agent_uploaded_files(
    uploaded_files: list[Any],
    session_id: str,
    project_name: str,
) -> list[str]:
    """Save uploaded project files and return saved absolute file paths."""
    if not uploaded_files:
        return []

    out_dir = PROJECT_ROOT / "data" / "code_agent_uploads" / session_id / _project_slug(project_name)
    out_dir.mkdir(parents=True, exist_ok=True)

    saved_paths: list[str] = []
    for uf in uploaded_files:
        target = out_dir / uf.name
        target.write_bytes(uf.getvalue())
        saved_paths.append(str(target))

    return saved_paths


def _apply_claude_code_command(raw_query: str, active_project: dict) -> tuple[str, str | None]:
    """Apply Claude Code style slash command behavior to a user query."""
    q = (raw_query or "").strip()
    if not q.startswith("/"):
        return q, None

    command, _, remainder = q.partition(" ")
    payload = remainder.strip()
    cmd = command.lower()

    project_paths = [p.strip() for p in active_project.get("paths", "").splitlines() if p.strip()]
    attached_files = active_project.get("attached_file_paths", [])

    base_context = (
        "Project context:\n"
        f"- Directories attached: {len(project_paths)}\n"
        f"- Files attached: {len(attached_files)}\n"
    )

    if cmd == "/init":
        transformed = (
            "Initialize this project like a coding agent. "
            "Map directory structure, identify entry points, highlight risks, and propose a step-by-step plan.\n\n"
            + base_context
            + (f"User focus: {payload}\n" if payload else "")
        )
        return transformed, "init"

    if cmd == "/plan":
        transformed = (
            "Create an implementation plan with ordered steps, assumptions, and validation checks.\n\n"
            + base_context
            + (f"Task: {payload}\n" if payload else "")
        )
        return transformed, "plan"

    if cmd == "/fix":
        transformed = (
            "Debug and fix the issue. Start with probable root cause, then provide concrete edits and validation steps.\n\n"
            + base_context
            + (f"Issue: {payload}\n" if payload else "")
        )
        return transformed, "fix"

    if cmd == "/review":
        transformed = (
            "Perform a code review focused on bugs, risks, regressions, and missing tests.\n\n"
            + base_context
            + (f"Scope: {payload}\n" if payload else "")
        )
        return transformed, "review"

    if cmd == "/edit":
        transformed = (
            "Apply requested edits to project files. If ambiguous, state assumptions and provide exact patch-ready changes.\n\n"
            + base_context
            + (f"Edit request: {payload}\n" if payload else "")
        )
        return transformed, "edit"

    return raw_query, None


@st.cache_resource
def _get_code_agent_sdk() -> SemiconductorRAGAgentSDK:
    """Create one SDK agent instance per Streamlit session process."""
    return SemiconductorRAGAgentSDK()


def _run_ate_code_agent_query(program_path: str, task: str, approved: bool = False) -> tuple[str, str, str]:
    """Execute a single query through the standalone ATE code-agent runtime."""
    agent_path = PROJECT_ROOT / "ate-code-agent" / "agent.py"
    if not agent_path.exists():
        return (
            f"ATE agent not found at {agent_path}. Build it under page_indexing_RAG/ate-code-agent first.",
            "ate-agent",
            "error",
        )

    ate_root = str(agent_path.parent)
    if ate_root not in sys.path:
        sys.path.insert(0, ate_root)

    try:
        spec = importlib.util.spec_from_file_location("ate_code_agent_runtime", str(agent_path))
        if spec is None or spec.loader is None:
            return ("Failed to load ATE agent module specification.", "ate-agent", "error")

        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)

        if not hasattr(module, "run_single_query"):
            return ("ATE agent module is missing run_single_query().", "ate-agent", "error")

        try:
            answer = _run_async_blocking(module.run_single_query(program_path, task, approved=approved))
        except TypeError:
            # Backward-compatible call for older module signatures.
            answer = _run_async_blocking(module.run_single_query(program_path, task))

        return str(answer), "claude-sonnet-4-6", "ate-agent"
    except Exception as exc:
        return (f"ATE agent execution error: {exc}", "ate-agent", "error")


def _build_ate_init_markdown(project_dir: str, notes: str = "") -> tuple[str, str]:
    """Create and persist a local ATE /init report markdown file for an attached directory."""
    root = Path(project_dir)
    if not root.exists() or not root.is_dir():
        return "", f"Invalid directory: {project_dir}"

    files: list[Path] = []
    for p in root.rglob("*"):
        if not p.is_file():
            continue
        if any(part in {".git", ".venv", "__pycache__", "node_modules"} for part in p.parts):
            continue
        files.append(p)
        if len(files) >= 400:
            break

    rel_files = [str(f.relative_to(root)) for f in files]
    ext_counts = Counter((f.suffix.lower() or "<no_ext>") for f in files)

    entrypoint_markers = {
        "main.py",
        "app.py",
        "streamlit_app.py",
        "program.xlsx",
        "program.xlsm",
        "readme.md",
        "makefile",
        "pyproject.toml",
    }
    entrypoints = [
        rf for rf in rel_files
        if Path(rf).name.lower() in entrypoint_markers
        or rf.lower().endswith((".xlsx", ".xlsm", ".bas", ".cls", ".vb"))
    ][:30]

    top_files = rel_files[:60]
    ext_summary = sorted(ext_counts.items(), key=lambda kv: kv[1], reverse=True)[:12]

    md_lines = [
        "# ATE Code Agent Init Report",
        "",
        f"Generated: {datetime.now().isoformat(timespec='seconds')}",
        f"Project root: {root}",
        "",
        "## Workspace Summary",
        f"- Files scanned: {len(files)}",
        f"- Distinct extensions: {len(ext_counts)}",
        "",
        "## Extension Distribution",
    ]

    if ext_summary:
        md_lines.extend([f"- {ext}: {count}" for ext, count in ext_summary])
    else:
        md_lines.append("- No files found")

    md_lines.extend([
        "",
        "## Likely Entrypoints / Key Artifacts",
    ])
    if entrypoints:
        md_lines.extend([f"- {ep}" for ep in entrypoints])
    else:
        md_lines.append("- No obvious entrypoints detected")

    md_lines.extend([
        "",
        "## File Inventory (Sample)",
    ])
    if top_files:
        md_lines.extend([f"- {rf}" for rf in top_files])
    else:
        md_lines.append("- No files found")

    md_lines.extend([
        "",
        "## Initial Risks",
        "- Missing environment credentials and toolchain dependencies",
        "- Safety-sensitive sheets/modules changed without review",
        "- Run-mode misconfiguration (must remain engineering/characterization)",
        "",
        "## Step-by-Step Plan",
        "1. Validate target workbook/module paths and available tools.",
        "2. Read current Excel/VBA state and identify required edits.",
        "3. Present exact plan and wait for human approval.",
        "4. Apply minimal edits with git backup.",
        "5. Run and parse logs, then iterate until stable or escalate.",
    ])

    if notes.strip():
        md_lines.extend([
            "",
            "## Project Notes",
            notes.strip()[:4000],
        ])

    report = "\n".join(md_lines) + "\n"
    out_path = root / "ATE_INIT_REPORT.md"
    out_path.write_text(report, encoding="utf-8")
    return report, str(out_path)


def _looks_like_ate_query(query: str) -> bool:
    """Heuristic intent routing for ATE-specialized workflow."""
    q = (query or "").lower()
    ate_tokens = [
        "ig-xl",
        "igxl",
        "ultraflex",
        "ate",
        "stdf",
        "vba",
        "test instance",
        "test suite",
        "pin map",
        "levels sheet",
        "limits sheet",
        ".xlsx",
        ".xlsm",
        ".bas",
        ".cls",
        "dut",
    ]
    return any(token in q for token in ate_tokens)


def _run_async_blocking(coro):
    """Run async coroutine from Streamlit sync code."""
    try:
        return asyncio.run(coro)
    except RuntimeError:
        loop = asyncio.new_event_loop()
        try:
            return loop.run_until_complete(coro)
        finally:
            loop.close()


st.set_page_config(page_title="ADI TestAgent", page_icon="ADI", layout="wide")

# Initialize session state FIRST, before any UI elements
if "sessions_index" not in st.session_state:
    st.session_state.sessions_index = load_sessions()

if "current_session_id" not in st.session_state:
    sessions = st.session_state.sessions_index.get("sessions", {})
    if sessions:
        # Pick most recently updated session
        st.session_state.current_session_id = max(
            sessions.keys(),
            key=lambda k: sessions[k].get("updated_at", "")
        )
    else:
        # Create first session
        st.session_state.current_session_id = new_session()

if "document_cache" not in st.session_state:
    st.session_state.document_cache = {}
if "uploaded_files_cache" not in st.session_state:
    st.session_state.uploaded_files_cache = {}
if "active_files" not in st.session_state:
    st.session_state.active_files = []
if "code_agent_messages" not in st.session_state:
    st.session_state.code_agent_messages = []
if "code_agent_paths" not in st.session_state:
    st.session_state.code_agent_paths = ""
if "code_agent_workspace_context" not in st.session_state:
    st.session_state.code_agent_workspace_context = ""
if "code_agent_workspace_summary" not in st.session_state:
    st.session_state.code_agent_workspace_summary = None
if "code_agent_workspace_sig" not in st.session_state:
    st.session_state.code_agent_workspace_sig = ""
if "code_agent_sdk_mode" not in st.session_state:
    st.session_state.code_agent_sdk_mode = True
if "code_agent_sdk_use_internal_kb" not in st.session_state:
    st.session_state.code_agent_sdk_use_internal_kb = False
if "code_agent_allowed_tools" not in st.session_state:
    st.session_state.code_agent_allowed_tools = ["Read", "Grep", "Glob", "Edit", "Write"]

st.title("ADI TestAgent")
st.caption("Analog Devices Internal Tool")
st.divider()

# Sidebar
with st.sidebar:
    # New Chat button at the top
    if st.button("➕ New Chat", use_container_width=True, type="primary"):
        st.session_state.current_session_id = new_session()
        # Clear all caches for new session
        st.session_state.document_cache = {}
        st.session_state.uploaded_files_cache = {}
        st.session_state.active_files = []
        st.rerun()

    st.divider()

    # Recent Chats section
    st.subheader("Recent Chats")
    sessions = st.session_state.sessions_index.get("sessions", {})

    if sessions:
        # Sort sessions by updated_at in descending order
        sorted_sessions = sorted(
            sessions.items(),
            key=lambda x: x[1].get("updated_at", ""),
            reverse=True
        )[:30]  # Show max 30 sessions

        for session_id, session_data in sorted_sessions:
            col1, col2 = st.columns([10, 1])

            with col1:
                # Highlight current session
                is_current = session_id == st.session_state.current_session_id
                button_label = f"{'● ' if is_current else ''}{session_data['title']}"

                if st.button(
                    button_label,
                    key=f"sess_{session_id}",
                    use_container_width=True,
                    type="secondary" if is_current else "secondary"
                ):
                    if not is_current:  # Only switch if not already current
                        st.session_state.current_session_id = session_id
                        # Clear caches when switching sessions
                        st.session_state.document_cache = {}
                        st.session_state.uploaded_files_cache = {}
                        st.session_state.active_files = []
                        st.rerun()

            with col2:
                # Don't allow deleting the last session
                if len(sessions) > 1 or session_id != st.session_state.current_session_id:
                    if st.button("🗑", key=f"del_{session_id}", help="Delete this chat"):
                        delete_session(session_id)
                        st.rerun()
    else:
        st.info("No chats yet. Start a new chat!")

    st.divider()

    st.header("Settings")

    # ── Model selector ──────────────────────────────────────────────────
    st.subheader("LLM Models")
    _model_options = []
    for vendor, models in LLM_MODELS.items():
        for display_name, model_str, _ in models:
            _model_options.append((f"{display_name} ({vendor})", model_str))

    _model_labels = [label for label, _ in _model_options]
    _model_strings = [ms for _, ms in _model_options]

    # Find index of current / default model
    _current_model = st.session_state.get("selected_model", DEFAULT_LLM_MODEL)
    _default_idx = next(
        (i for i, ms in enumerate(_model_strings) if ms == _current_model), 0
    )

    _chosen_idx = st.selectbox(
        "Select model",
        options=range(len(_model_labels)),
        format_func=lambda i: _model_labels[i],
        index=_default_idx,
        key="model_selectbox",
        label_visibility="collapsed",
    )
    st.session_state["selected_model"] = _model_strings[_chosen_idx]

    st.divider()

    db_count = collection.count()
    if db_count > 0:
        st.success(f"Knowledge Base: {db_count} chunks")
    else:
        st.warning("Knowledge Base is empty")

    st.divider()

    st.subheader("PDF Knowledge Base")
    st.caption(f"PDF folder: `{PDF_DATA_DIR}`")

    # Get all folders containing PDFs
    pdf_folders = get_pdf_folders()

    if not pdf_folders:
        st.error(f"No PDF files found in `{PDF_DATA_DIR}` or its subdirectories")
    else:
        # Calculate total PDFs across all folders
        total_pdfs = sum(folder_info["count"] for folder_info in pdf_folders.values())
        st.info(f"Found **{total_pdfs}** PDFs in **{len(pdf_folders)}** folder(s)")

        # Button to ingest ALL PDFs from all folders
        if st.button("🔄 Ingest All PDFs", use_container_width=True, type="primary"):
            progress_bar = st.progress(0.0, text="Starting ingestion...")
            status_text = st.empty()

            def update_progress(frac: float, fname: str) -> None:
                progress_bar.progress(frac, text=f"Processing: {fname}")
                status_text.text(f"Embedding: {fname}")

            stats = ingest_all_pdfs(progress_cb=update_progress)
            progress_bar.progress(1.0, text="Done")

            if "error" in stats:
                st.error(stats["error"])
            else:
                st.success(
                    f"Ingestion complete\n\n"
                    f"- New: {stats['new']} PDFs\n"
                    f"- Skipped (already in KB): {stats['skipped']} PDFs\n"
                    f"- New chunks added: {stats['total_chunks']}"
                )
            st.rerun()

        # # Show folders with individual ingestion buttons
        # with st.expander("📁 Folders & Manual Ingestion", expanded=True):
        #     for folder_name, folder_info in sorted(pdf_folders.items()):
        #         col1, col2, col3 = st.columns([3, 1, 2])

        #         with col1:
        #             if folder_name == "(root)":
        #                 st.markdown(f"📂 **Root Directory**")
        #             else:
        #                 st.markdown(f"📁 **{folder_name}**")

        #         with col2:
        #             st.markdown(f"📄 {folder_info['count']} PDFs")

        #         with col3:
        #             button_key = f"ingest_{folder_name.replace('/', '_').replace('\\', '_')}"
        #             if st.button("Ingest", key=button_key, use_container_width=True):
        #                 with st.spinner(f"Ingesting PDFs from {folder_name}..."):
        #                     progress_bar = st.progress(0.0, text="Starting...")
        #                     status_text = st.empty()

        #                     def update_progress(frac: float, fname: str) -> None:
        #                         progress_bar.progress(frac, text=f"Processing: {fname}")
        #                         status_text.text(f"Embedding: {fname}")

        #                     stats = ingest_pdfs_from_folder(
        #                         folder_info["path"],
        #                         progress_cb=update_progress
        #                     )
        #                     progress_bar.progress(1.0, text="Done")

        #                     if "error" in stats:
        #                         st.error(stats["error"])
        #                     else:
        #                         st.success(
        #                             f"Folder ingestion complete\n\n"
        #                             f"- New: {stats['new']} PDFs\n"
        #                             f"- Skipped: {stats['skipped']} PDFs\n"
        #                             f"- Chunks added: {stats['total_chunks']}"
        #                         )
        #                 st.rerun()

        # # Show individual PDF status (optional, collapsed by default)
        # with st.expander("📋 Individual PDF Status", expanded=False):
        #     for folder_name, folder_info in sorted(pdf_folders.items()):
        #         if folder_name != "(root)":
        #             st.markdown(f"**{folder_name}/**")

        #         folder_path = folder_info["path"]
        #         pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith(".pdf")]

        #         for fname in sorted(pdf_files):
        #             path = os.path.join(folder_path, fname)
        #             ingested = is_pdf_ingested(path)
        #             icon = "✅" if ingested else "⏳"
        #             if folder_name == "(root)":
        #                 st.markdown(f"{icon} `{fname}`")
        #             else:
        #                 st.markdown(f"  {icon} `{fname}`")

    st.divider()

    st.subheader("Load Text, Program, and Image Assets")
    st.caption(f"Assets folder: `{TXT_DATA_DIR}`")
    if st.button("Load Text/Program/Image Data", use_container_width=True):
        if os.path.exists(TXT_DATA_DIR):
            with st.spinner("Loading and embedding text/program/image assets..."):
                progress_bar = st.progress(0.0, text="Scanning assets...")

                def update_progress(frac: float, fname: str) -> None:
                    progress_bar.progress(frac, text=f"Processing: {fname}")

                stats = ingest_assets_from_directory(TXT_DATA_DIR, progress_cb=update_progress)
                progress_bar.progress(1.0, text="Done")

            if "error" in stats:
                st.error(stats["error"])
            else:
                st.success(
                    f"Asset ingestion complete\n\n"
                    f"- New files: {stats['new']}\n"
                    f"- Skipped (already in KB): {stats['skipped']}\n"
                    f"- Failed: {stats['failed']}\n"
                    f"- New chunks added: {stats['total_chunks']}"
                )
            st.rerun()
        else:
            st.error(f"{TXT_DATA_DIR} directory not found")

    st.divider()

    st.subheader("Retrieval Settings")
    n_results = st.slider("Chunks per query", 3, 10, 5)
    mmr_lambda = st.slider("MMR diversity (0=diverse, 1=relevant)", 0.0, 1.0, 0.7, step=0.05)
    agentic_mode = st.toggle(
        "Enable agentic retrieval loop",
        value=False,
        help="Planner + iterative retrieve-only loop + grep/read evidence judge.",
    )
    agentic_tool_budget = st.slider(
        "Agentic tool budget",
        0,
        20,
        6,
        disabled=not agentic_mode,
        help="Maximum grep/read tool calls for agentic retrieval.",
    )
    show_sources = st.toggle("Show retrieved sources", value=True)

    _WEB_MODE_OPTIONS = ["always"]
    _WEB_MODE_LABELS = {
        "always":   "Always search (web + internal KB)",
    }
    web_mode = st.selectbox(
        "Web Search Mode",
        options=_WEB_MODE_OPTIONS,
        index=0,
        disabled=True,
        format_func=lambda x: _WEB_MODE_LABELS[x],
        help=(
            "**Always** — Always fetch web results AND search internal KB; merges both.\n\n"
            "This mode is locked to ensure every query uses combined web + internal KB context."
        ),
    )
    if not HAS_DDGS:
        st.warning("`duckduckgo-search` not installed. Run: `pip install duckduckgo-search`")
    if web_mode == "always":
        st.info("Always mode: combines internal KB + real-time web results.")
    if agentic_mode:
        st.info("Agentic mode: planner + iterative retrieval + bounded grep/read evidence collection.")

    st.divider()

    if st.button("Clear Knowledge Base", use_container_width=True):
        collection.delete(where={"source": {"$ne": ""}})
        st.success("Knowledge base cleared")
        st.rerun()

    if st.button("Clear Current Chat", use_container_width=True):
        # Clear messages in current session only
        session_id = st.session_state.current_session_id
        sessions = st.session_state.sessions_index.get("sessions", {})
        if session_id in sessions:
            sessions[session_id]["messages"] = []
            sessions[session_id]["updated_at"] = datetime.now().isoformat()
            st.session_state.sessions_index["sessions"] = sessions
            save_sessions(st.session_state.sessions_index)
        # Clear document cache for current session
        st.session_state.document_cache = {}
        st.rerun()

    st.divider()

    st.subheader("💡 Quick Tip")
    st.info(
        "**Document Generation:**\n\n"
        "Ask for a document and get instant downloads:\n\n"
        "- *\"Generate a report on...\"*\n"
        "- *\"Create a document about...\"*\n"
        "- *\"Export this as a doc\"*\n\n"
        "📥 DOCX — Editable Word format\n"
        "📄 PDF — Shareable format"
    )

    st.divider()
    st.caption("ADI Internal Tool - Not for external use")


def _render_doc_buttons(key_prefix: str, uq: str, content: str) -> None:
    """Render DOCX + PDF download buttons, using session cache to avoid re-generation."""
    session_id = st.session_state.get("current_session_id", "default")
    ck_docx = f"{session_id}_{key_prefix}_docx"
    ck_pdf = f"{session_id}_{key_prefix}_pdf"
    b1, b2 = st.columns(2)
    with b1:
        try:
            if ck_docx not in st.session_state.document_cache:
                st.session_state.document_cache[ck_docx] = (
                    DocumentGenerator.create_from_llm_response(uq, content, format="docx")
                )
            st.download_button(
                label="📥 Download DOCX",
                data=st.session_state.document_cache[ck_docx],
                file_name=f"chipagent_{key_prefix}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key=f"dl_docx_{key_prefix}",
                use_container_width=True,
            )
        except Exception as exc:
            st.error(f"DOCX error: {exc}")
    # with b2:
    #     try:
    #         if ck_pdf not in st.session_state.document_cache:
    #             st.session_state.document_cache[ck_pdf] = (
    #                 DocumentGenerator.create_from_llm_response(uq, content, format="pdf")
    #             )
    #         st.download_button(
    #             label="📄 Download PDF",
    #             data=st.session_state.document_cache[ck_pdf],
    #             file_name=f"chipagent_{key_prefix}.pdf",
    #             mime="application/pdf",
    #             key=f"dl_pdf_{key_prefix}",
    #             use_container_width=True,
    #         )
    #     except Exception as exc:
    #         st.error(f"PDF error: {exc}")


def _ingest_uploaded_file(uploaded_file) -> None:
    """Parse uploaded file and cache it for chat-context use."""
    filename = uploaded_file.name
    if filename in st.session_state.uploaded_files_cache:
        return

    raw_bytes = uploaded_file.read()
    file_ext = filename.lower().rsplit(".", 1)[-1] if "." in filename else ""

    if file_ext == "pdf":
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            tmp.write(raw_bytes)
            tmp_path = tmp.name
        content = " ".join(p["text"] for p in extract_pdf_pages(tmp_path))
        os.unlink(tmp_path)
    elif file_ext in ["jpg", "jpeg", "png", "gif", "bmp"]:
        content = f"[Image file: {filename}]"
    else:
        try:
            content = raw_bytes.decode("utf-8")
        except UnicodeDecodeError:
            content = raw_bytes.decode("latin-1", errors="ignore")

    file_type = classify_file(filename, content)
    st.session_state.uploaded_files_cache[filename] = {
        "content": content,
        "file_type": file_type,
        "raw_bytes": raw_bytes,
        "size": len(raw_bytes),
    }
    if filename not in st.session_state.active_files:
        st.session_state.active_files.append(filename)


def _parse_ask_uploaded_file(uploaded_file: Any) -> tuple[str, str]:
    """Parse supported ask-section files and return filename + extracted text."""
    filename = uploaded_file.name
    ext = filename.lower().rsplit(".", 1)[-1] if "." in filename else ""
    raw_bytes = uploaded_file.getvalue()

    if ext == "pdf":
        try:
            import pdfplumber

            with pdfplumber.open(io.BytesIO(raw_bytes)) as pdf:
                parts = [(page.extract_text() or "") for page in pdf.pages]
            text = "\n".join(parts).strip()
            return filename, text
        except Exception as exc:
            raise RuntimeError(f"PDF parsing requires pdfplumber: {exc}") from exc

    if ext == "docx":
        try:
            from docx import Document
        except Exception as exc:
            raise RuntimeError(f"DOCX parsing requires python-docx: {exc}") from exc

        document = Document(io.BytesIO(raw_bytes))
        text = "\n".join([p.text for p in document.paragraphs]).strip()
        return filename, text

    if ext == "txt":
        try:
            text = raw_bytes.decode("utf-8")
        except UnicodeDecodeError:
            text = raw_bytes.decode("latin-1", errors="ignore")
        return filename, text.strip()

    if ext == "csv":
        try:
            text = raw_bytes.decode("utf-8")
        except UnicodeDecodeError:
            text = raw_bytes.decode("latin-1", errors="ignore")
        return filename, text.strip()

    raise RuntimeError("Unsupported file type. Use PDF, DOCX, TXT, or CSV.")


def _safe_render_package_chat_input(widget_key: str):
    """Best-effort wrapper for optional package-based chat input uploader."""
    if not HAS_CHAT_INPUT_FILEUPLOAD or chat_input_fileupload is None:
        return "", None, False

    try:
        payload = chat_input_fileupload(
            key=widget_key,
            placeholder="Ask about TML, TDC specs, instrument cards, ADI devices...",
            file_types=["pdf", "docx", "txt", "csv"],
        )
    except Exception:
        return "", None, False

    question_text = ""
    uploaded_file = None
    submitted = False

    if isinstance(payload, dict):
        question_text = str(payload.get("text") or payload.get("message") or "")
        uploaded_file = payload.get("file") or payload.get("uploaded_file")
        submitted = bool(payload.get("submitted") or payload.get("send") or payload.get("ask"))
    elif isinstance(payload, tuple):
        if len(payload) >= 1:
            question_text = str(payload[0] or "")
        if len(payload) >= 2:
            uploaded_file = payload[1]
        submitted = len(question_text.strip()) > 0
    elif isinstance(payload, str):
        question_text = payload
        submitted = len(question_text.strip()) > 0

    return question_text, uploaded_file, submitted


def _build_followup_options(question: str, answer: str, is_doc_request: bool = False) -> list[str]:
    """Build concise follow-up suggestions for the latest assistant response."""
    q = (question or "").lower()
    options: list[str] = []

    if "weather" in q or "forecast" in q:
        options = [
            "Show hourly forecast for today",
            "Show 7-day forecast summary",
            "Add rain and wind details",
        ]
    elif is_doc_request:
        options = [
            "Give a shorter executive summary",
            "Expand with a step-by-step checklist",
            "Add risks and assumptions section",
        ]
    else:
        options = [
            "Summarize this in 5 bullets",
            "Give step-by-step action plan",
            "List assumptions, risks, and checks",
        ]

    out: list[str] = []
    seen: set[str] = set()
    for item in options:
        label = item.strip()
        key = label.lower()
        if label and key not in seen:
            seen.add(key)
            out.append(label)
    return out[:3]


def _resolve_followup_selection(user_text: str, options: list[str]) -> str | None:
    """Resolve yes/number/text input to one of the provided follow-up options."""
    txt = (user_text or "").strip().lower()
    if not txt or not options:
        return None

    if txt in {"yes", "y", "ok", "okay", "sure", "go ahead", "continue", "proceed"}:
        return options[0]

    if txt.isdigit():
        idx = int(txt)
        if 1 <= idx <= len(options):
            return options[idx - 1]

    for opt in options:
        if txt == opt.lower() or txt in opt.lower():
            return opt

    return None


def _compose_followup_prompt(option: str, base_question: str, base_answer: str) -> str:
    """Compose a follow-up prompt grounded in the previous answer."""
    short_answer = (base_answer or "")[:4000]
    return (
        "Follow-up request based on the previous answer.\n\n"
        f"Original question: {base_question}\n\n"
        "Previous answer context:\n"
        f"{short_answer}\n\n"
        f"Please do this next: {option}"
    )


def _render_chat_tab() -> None:
    st.subheader("Ask")
    st.caption("RAG chat over knowledge base, optional web search, and attached files.")

    # Compact input controls and improve chat readability.
    st.markdown(
        """
        <style>
        div[data-testid="stChatMessage"] p {
            line-height: 1.55;
        }
        div[data-testid="stForm"] {
            padding-top: 0.25rem;
            padding-bottom: 0.25rem;
            margin-bottom: 0.2rem;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    # Initialize session state for file upload
    if "chat_ask_uploaded_name" not in st.session_state:
        st.session_state.chat_ask_uploaded_name = ""
    if "chat_ask_uploaded_text" not in st.session_state:
        st.session_state.chat_ask_uploaded_text = ""
    if "chat_ask_uploaded_sig" not in st.session_state:
        st.session_state.chat_ask_uploaded_sig = ""
    if "chat_ask_uploader_nonce" not in st.session_state:
        st.session_state.chat_ask_uploader_nonce = 0
    if "chat_show_ask_uploader" not in st.session_state:
        st.session_state.chat_show_ask_uploader = False
    if "chat_pending_followups" not in st.session_state:
        st.session_state.chat_pending_followups = {}
    if "chat_followup_clicked_payload" not in st.session_state:
        st.session_state.chat_followup_clicked_payload = None

    # Chat message area: keep generous height so responses stay readable.
    chat_area_height = 680 if not st.session_state.get("chat_show_ask_uploader") else 600
    chat_area = st.container(height=chat_area_height, border=False)
    question = None
    question_for_llm = ""

    if st.session_state.chat_ask_uploaded_name:
        name_col, clear_col = st.columns([10, 2])
        with name_col:
            st.caption(f"+ {st.session_state.chat_ask_uploaded_name}")
        with clear_col:
            if st.button("Clear", key="chat_ask_remove_file", use_container_width=True):
                st.session_state.chat_ask_uploaded_name = ""
                st.session_state.chat_ask_uploaded_text = ""
                st.session_state.chat_ask_uploaded_sig = ""
                st.session_state.chat_ask_uploader_nonce += 1
                st.rerun()

    # File uploader (conditionally shown)
    if st.session_state.chat_show_ask_uploader:
        uploader_key = f"chat_ask_section_uploader_{st.session_state.chat_ask_uploader_nonce}"
        ask_uploaded = st.file_uploader(
            "Add source file",
            type=["pdf", "docx", "txt", "csv"],
            accept_multiple_files=False,
            key=uploader_key,
            help="Attach one PDF, DOCX, TXT, or CSV file for this ask.",
        )

        if ask_uploaded is not None:
            file_sig = f"{ask_uploaded.name}:{ask_uploaded.size}"
            if file_sig != st.session_state.chat_ask_uploaded_sig:
                try:
                    fname, ftext = _parse_ask_uploaded_file(ask_uploaded)
                    st.session_state.chat_ask_uploaded_name = fname
                    st.session_state.chat_ask_uploaded_text = ftext
                    st.session_state.chat_ask_uploaded_sig = file_sig
                    st.session_state.chat_show_ask_uploader = False
                    st.rerun()
                except Exception as exc:
                    st.error(f"Unable to parse file: {exc}")

    # Ask form with compact input controls
    with st.form("chat_tab_ask_form", clear_on_submit=True, enter_to_submit=True):
        package_question, package_uploaded, package_submitted = _safe_render_package_chat_input("ask_chat_input_package")
        if package_uploaded is not None:
            file_sig = f"{package_uploaded.name}:{getattr(package_uploaded, 'size', 0)}"
            if file_sig != st.session_state.chat_ask_uploaded_sig:
                try:
                    fname, ftext = _parse_ask_uploaded_file(package_uploaded)
                    st.session_state.chat_ask_uploaded_name = fname
                    st.session_state.chat_ask_uploaded_text = ftext
                    st.session_state.chat_ask_uploaded_sig = file_sig
                except Exception as exc:
                    st.error(f"Unable to parse file: {exc}")

        if package_submitted and package_question.strip():
            question_text = package_question
            ask_clicked = True
            attach_clicked = False
        else:
            ask_col_attach, ask_col_input, ask_col_submit = st.columns([1, 10, 1])
            # Keep Ask as the first submit button in code order so Enter submits Ask.
            with ask_col_submit:
                ask_clicked = st.form_submit_button("Ask", use_container_width=True)
            with ask_col_input:
                question_text = st.text_input(
                    "Ask Chat",
                    placeholder="Ask about TML, TDC specs, instrument cards, ADI devices...",
                    key="chat_tab_text_query",
                    label_visibility="collapsed",
                )
            with ask_col_attach:
                attach_clicked = st.form_submit_button("📎", use_container_width=True)

        # If form submit happened (Ask/Enter) but text was typed in the package input,
        # treat that package text as the active question.
        if ask_clicked and not question_text.strip() and package_question.strip():
            question_text = package_question

    if attach_clicked:
        st.session_state.chat_show_ask_uploader = not st.session_state.chat_show_ask_uploader
        st.rerun()

    # Process question and append file content if attached
    if ask_clicked and question_text.strip():
        question = question_text.strip()
        question_for_llm = question
        if st.session_state.chat_ask_uploaded_text:
            # Append file content to the question for LLM
            file_block = (
                f"\n\n=== ATTACHED FILE: {st.session_state.chat_ask_uploaded_name} ===\n"
                f"{st.session_state.chat_ask_uploaded_text[:20000]}"  # Limit to 20k chars
            )
            question_for_llm += file_block

    # Resolve textual follow-up selection: yes / number / option text.
    if question:
        pending_map = st.session_state.get("chat_pending_followups", {})
        pending = pending_map.get(st.session_state.current_session_id, {}) if isinstance(pending_map, dict) else {}
        selected_option = _resolve_followup_selection(question, pending.get("options", []))
        if selected_option:
            question_for_llm = _compose_followup_prompt(
                selected_option,
                str(pending.get("base_question", "")),
                str(pending.get("base_answer", "")),
            )
            question = selected_option

    with chat_area:
        current_messages = get_current_messages()
        if not current_messages:
            st.info(
                "Welcome to ADI ChipAgent. Ask about UltraFlex/TML, docs, device specs, or attached files."
            )

        for idx, msg in enumerate(current_messages):
            with st.chat_message(msg["role"]):
                st.markdown(msg.get("content", ""))
                if msg["role"] == "user" and msg.get("attached_file_name"):
                    st.caption(f"+ {msg.get('attached_file_name')}")
                if msg["role"] == "assistant" and msg.get("content"):
                    uq = msg.get("user_query", "")
                    is_doc = msg.get("is_doc_request", False)
                    if is_doc:
                        st.success("Document ready — click below to download")
                        _render_doc_buttons(f"h{idx}", uq, msg["content"])
                    else:
                        with st.expander("Download this response as a document"):
                            _render_doc_buttons(f"h{idx}", uq, msg["content"])

                if msg["role"] == "assistant" and show_sources:
                    _web = msg.get("web_sources", [])
                    if _web:
                        with st.expander(f"Web Sources ({len(_web)})"):
                            for wr in _web:
                                st.markdown(f"**[{wr.get('title', 'Untitled')}]({wr['url']})**")
                                st.caption(wr.get("snippet", "")[:200])

                if msg["role"] == "assistant":
                    followup_options = msg.get("followup_options", [])
                    if followup_options:
                        st.caption("Follow-up options:")
                        for oi, opt in enumerate(followup_options, start=1):
                            if st.button(
                                f"{oi}. {opt}",
                                key=f"chat_fu_{st.session_state.current_session_id}_{idx}_{oi}",
                                use_container_width=True,
                            ):
                                st.session_state.chat_followup_clicked_payload = {
                                    "session_id": st.session_state.current_session_id,
                                    "message_index": idx,
                                    "option": opt,
                                }
                                st.rerun()

        if not question:
            payload = st.session_state.get("chat_followup_clicked_payload")
            if isinstance(payload, dict) and payload.get("session_id") == st.session_state.current_session_id:
                midx = payload.get("message_index", -1)
                selected = str(payload.get("option", "")).strip()
                base_q = ""
                base_a = ""
                if isinstance(midx, int) and 0 <= midx < len(current_messages):
                    base_msg = current_messages[midx]
                    base_q = str(base_msg.get("user_query", ""))
                    base_a = str(base_msg.get("content", ""))
                question_for_llm = _compose_followup_prompt(selected, base_q, base_a)
                question = selected
                st.session_state.chat_followup_clicked_payload = None

        if not question:
            return

        is_doc_request, _ = DocumentGenerator.detect_document_request(question)
        with st.chat_message("user"):
            st.markdown(question)
            if st.session_state.chat_ask_uploaded_name:
                st.caption(f"+ {st.session_state.chat_ask_uploaded_name}")

        session_id = st.session_state.current_session_id
        sessions = st.session_state.sessions_index.get("sessions", {})
        if session_id in sessions:
            sessions[session_id]["messages"].append(
                {
                    "role": "user",
                    "content": question,
                    "attached_file_name": st.session_state.chat_ask_uploaded_name if st.session_state.chat_ask_uploaded_name else "",
                }
            )
            sessions[session_id]["updated_at"] = datetime.now().isoformat()
            if len(sessions[session_id]["messages"]) == 1:
                auto_title_session(session_id, question)
            st.session_state.sessions_index["sessions"] = sessions
            save_sessions(st.session_state.sessions_index)

        uploaded_context = ""
        has_attached_ask_file = bool((st.session_state.get("chat_ask_uploaded_text") or "").strip())

        with st.chat_message("assistant"):
            answer = ""
            model_used = "-"
            complexity = "unknown"
            _web_spinner = {
                "off": "Searching knowledge base...",
                "fallback": "Searching knowledge base...",
                "always": "Searching knowledge base + web...",
                "web_only": "Searching the web...",
            }.get(web_mode, "Searching...")
            if agentic_mode:
                _web_spinner = "Running agentic retrieval loop..."
            spinner_msg = "Generating document..." if is_doc_request else _web_spinner
            try:
                with st.spinner(spinner_msg):
                    if agentic_mode:
                        hybrid = run_agentic_retrieve_loop(
                            question,
                            n_per_query=n_results,
                            max_iterations=3,
                            tool_budget=max(0, int(agentic_tool_budget)),
                        )
                    else:
                        hybrid = hybrid_query_documents(
                            question,  # Use original question for search, not the one with file content
                            n_per_query=n_results,
                            web_mode=web_mode,
                            mmr_lambda=mmr_lambda,
                        )
            except Exception as exc:
                hybrid = {
                    "query_type": "U",
                    "internal_chunks": [],
                    "web_results": [],
                    "internal_confident": False,
                    "internal_sufficient": False,
                    "web_confident": False,
                    "abstain": True,
                    "abstain_reason": f"Retrieval failed: {exc}",
                }

            # If the user attached a file in Ask, do not block on KB abstain.
            # We still keep retrieved context when available, but always allow
            # answer generation from the uploaded file content.
            if has_attached_ask_file and hybrid.get("abstain"):
                hybrid["abstain"] = False
                hybrid["abstain_reason"] = ""
                hybrid["query_type"] = "U"
                st.caption("📄 Using attached file content for this answer.")

            _qtype_label = {"P": "Proprietary", "O": "Public", "U": "Unknown"}
            _qtype = _qtype_label.get(hybrid["query_type"], "Unknown")

            if hybrid["abstain"]:
                reason = (hybrid.get("abstain_reason") or "").strip()
                if not reason:
                    reason = "I could not retrieve enough context for this query. Please try again."
                st.warning(reason)
                answer = reason
                complexity = "abstain"
            else:
                try:
                    with st.spinner("Composing answer..."):
                        answer, model_used, complexity = generate_hybrid_response(
                            question_for_llm,  # This includes the file content if attached
                            hybrid,
                            uploaded_context,
                            document_mode=is_doc_request,
                        )
                except Exception as exc:
                    answer = f"Response generation failed: {exc}"
                    model_used = "-"
                    complexity = "error"

                # Guard against occasional empty model outputs.
                if not (answer or "").strip():
                    try:
                        retry_answer, retry_model, retry_complexity = generate_hybrid_response(
                            question_for_llm,
                            hybrid,
                            uploaded_context,
                            document_mode=is_doc_request,
                            concise_mode=True,
                        )
                        if (retry_answer or "").strip():
                            answer = retry_answer
                            model_used = retry_model
                            complexity = f"{retry_complexity}-retry"
                    except Exception:
                        pass

                if not (answer or "").strip():
                    answer = (
                        "I could not generate a response for this query. "
                        "Please retry or simplify the question."
                    )
                    model_used = "-"
                    complexity = "empty-fallback"

                st.markdown(answer)

            meta = (
                f"Model: `{model_used}` | Complexity: `{complexity}` "
                f"| Query: {_qtype} "
                f"| DB: `{len(hybrid['internal_chunks'])}` chunks "
                f"| Web: `{len(hybrid['web_results'])}` results "
                f"| Mode: `{'agentic' if agentic_mode else web_mode}`"
                + (" | Document mode" if is_doc_request else "")
                + (f" | 📎 File attached: {st.session_state.chat_ask_uploaded_name}" if has_attached_ask_file else "")
            )

            new_key = f"n{len(get_current_messages())}"
            if is_doc_request and answer:
                st.success("Document ready — click below to download")
                _render_doc_buttons(new_key, question, answer)
            elif answer:
                with st.expander("Download this response as a document"):
                    _render_doc_buttons(new_key, question, answer)

            if show_sources:
                if hybrid["internal_chunks"]:
                    with st.expander(f"Internal Sources ({len(hybrid['internal_chunks'])})"):
                        for src in hybrid["internal_chunks"]:
                            st.markdown(
                                f"**{src['source']}** - Page {src['page_number']}/{src['total_pages']} "
                                f"| Type: `{src['file_type']}` | Cosine dist: `{src['distance']:.3f}`"
                            )
                            st.code(
                                src["text"][:300] + ("..." if len(src["text"]) > 300 else ""),
                                language="text",
                            )
                if hybrid["web_results"]:
                    with st.expander(f"Web Sources ({len(hybrid['web_results'])})"):
                        for wr in hybrid["web_results"]:
                            st.markdown(f"**[{wr.get('title', 'Untitled')}]({wr['url']})**")
                            st.caption(wr.get("snippet", "")[:200])
                if hybrid.get("agentic_trace"):
                    with st.expander("Agentic Retrieval Trace"):
                        trace = hybrid.get("agentic_trace") or {}
                        plan = trace.get("plan", [])
                        steps = trace.get("steps", [])
                        judge = trace.get("judge", {})
                        st.markdown("**Plan**")
                        for item in plan:
                            st.markdown(f"- {item}")
                        st.markdown("**Steps**")
                        st.json(steps)
                        st.markdown("**Judge**")
                        st.json(judge)

        sessions = st.session_state.sessions_index.get("sessions", {})
        if session_id in sessions:
            followup_options = _build_followup_options(question, answer, is_doc_request=is_doc_request)
            sessions[session_id]["messages"].append(
                {
                    "role": "assistant",
                    "content": answer,
                    "meta": meta,
                    "sources": hybrid["internal_chunks"],
                    "web_sources": hybrid["web_results"],
                    "query_type": hybrid["query_type"],
                    "is_doc_request": is_doc_request,
                    "user_query": question,
                    "followup_options": followup_options,
                }
            )
            sessions[session_id]["updated_at"] = datetime.now().isoformat()
            st.session_state.sessions_index["sessions"] = sessions
            save_sessions(st.session_state.sessions_index)

            pending_map = st.session_state.get("chat_pending_followups", {})
            if not isinstance(pending_map, dict):
                pending_map = {}
            if followup_options:
                pending_map[session_id] = {
                    "base_question": question,
                    "base_answer": answer,
                    "options": followup_options,
                }
            else:
                pending_map.pop(session_id, None)
            st.session_state.chat_pending_followups = pending_map

            # Clear attached file after processing
            if has_attached_ask_file:
                st.session_state.chat_ask_uploaded_name = ""
                st.session_state.chat_ask_uploaded_text = ""
                st.session_state.chat_ask_uploaded_sig = ""
                st.session_state.chat_ask_uploader_nonce += 1


def _render_code_agent_tab() -> None:
    st.subheader("Code Agent")
    st.caption("Minimal Claude Code style workflow: attach folders/files and chat with a code-first agent.")

    _ensure_code_agent_projects()
    projects = st.session_state.code_agent_projects

    if "Main Project" not in projects:
        seed_name = st.session_state.get("code_agent_active_project")
        seed_project = projects.get(seed_name) if isinstance(projects, dict) else None
        if not seed_project and projects:
            seed_project = next(iter(projects.values()))
        projects = {"Main Project": seed_project or _default_code_agent_project()}
        st.session_state.code_agent_projects = projects

    active_project_name = "Main Project"
    st.session_state.code_agent_active_project = active_project_name
    active_project = projects[active_project_name]

    st.markdown("### Workspace")
    st.caption("Commands supported: /init, /plan, /fix, /review, /edit")

    paths_widget_key = f"code_agent_paths_{active_project_name}"
    if st.session_state.get(paths_widget_key) != active_project.get("paths", ""):
        st.session_state[paths_widget_key] = active_project.get("paths", "")

    st.text_area(
        "Attached directories (one per line)",
        key=paths_widget_key,
        placeholder=(
            "C:\\Projects\\UltraFlexProgram\\\n"
            "C:\\Projects\\DeviceA\\IGXL\\"
        ),
        height=120,
    )
    active_project["paths"] = st.session_state.get(paths_widget_key, "")

    bp1, bp2 = st.columns(2)
    with bp1:
        if st.button("Attach directory via Explorer", use_container_width=True, key=f"browse_dir_{active_project_name}"):
            seed_paths = [p.strip() for p in active_project.get("paths", "").splitlines() if p.strip()]
            initial_dir = seed_paths[0] if seed_paths else str(PROJECT_ROOT.parent)
            selected_dir = _pick_directory_via_dialog(initial_dir=initial_dir)
            if selected_dir:
                merged = list(dict.fromkeys(seed_paths + [selected_dir]))
                active_project["paths"] = "\n".join(merged)
                st.session_state.code_agent_projects = projects
                st.success(f"Attached: {selected_dir}")
                st.rerun()
            else:
                st.warning("No directory selected.")
    with bp2:
        if st.button("Clear attached directories", use_container_width=True, key=f"clear_dirs_{active_project_name}"):
            active_project["paths"] = ""
            active_project["workspace_context"] = ""
            active_project["workspace_summary"] = None
            active_project["workspace_sig"] = ""
            st.session_state.code_agent_projects = projects
            st.rerun()

    uploaded_project_files = st.file_uploader(
        "Attach files to active project",
        accept_multiple_files=True,
        key=f"code_agent_project_files_{active_project_name}",
    )
    if uploaded_project_files:
        saved_files = _persist_code_agent_uploaded_files(
            uploaded_project_files,
            st.session_state.current_session_id,
            active_project_name,
        )
        if saved_files:
            existing_files = active_project.get("attached_file_paths", [])
            merged_files = list(dict.fromkeys(existing_files + saved_files))
            active_project["attached_file_paths"] = merged_files

            seed_paths = [p.strip() for p in active_project.get("paths", "").splitlines() if p.strip()]
            file_dirs = [str(Path(fp).parent) for fp in saved_files]
            active_project["paths"] = "\n".join(list(dict.fromkeys(seed_paths + file_dirs)))
            st.session_state.code_agent_projects = projects
            st.success(f"Attached {len(saved_files)} file(s) to project.")

    if active_project.get("attached_file_paths"):
        with st.expander("Attached project files", expanded=False):
            for fp in active_project.get("attached_file_paths", [])[:80]:
                st.markdown(f"- {fp}")

    notes_widget_key = f"code_agent_notes_{active_project_name}"
    if st.session_state.get(notes_widget_key) != active_project.get("notes", ""):
        st.session_state[notes_widget_key] = active_project.get("notes", "")

    st.text_area(
        "Project notes / code snippets",
        key=notes_widget_key,
        placeholder="Add reference code, TODOs, assumptions, or debugging notes for this project.",
        height=120,
    )
    active_project["notes"] = st.session_state.get(notes_widget_key, "")

    max_files = 120
    max_chars = 3500
    use_rag_context = False
    sdk_mode = True
    sdk_use_internal_kb = False

    base_tools = ["Read", "Grep", "Glob", "Edit", "Write"]
    enable_web_tools = st.toggle(
        "Enable web tools (optional)",
        value=False,
        key="code_agent_enable_web_tools",
        help="Off by default for code-first behavior. Turn on only when external lookup is needed.",
    )
    st.session_state.code_agent_allowed_tools = (
        base_tools + ["WebSearch", "WebFetch"] if enable_web_tools else base_tools
    )

    approve_file_actions = st.toggle(
        "Approve file actions for this prompt",
        value=False,
        key="code_agent_approve_actions",
        help="Required to execute Edit/Write actions; otherwise the agent returns a structured runbook only.",
    )

    use_ate_agent = st.toggle(
        "Use ATE tools workflow",
        value=False,
        key="code_agent_use_ate_agent",
        help="Routes this prompt to the standalone ATE code-agent (Excel/VBA/run/log safety workflow).",
    )
    ate_approved = st.toggle(
        "Approve edit/run actions for this prompt",
        value=False,
        key="code_agent_ate_approved",
        help="Required by ATE safety gate for tasks that imply edits or program execution.",
    )

    st.markdown("### UltraFLEX RAG")
    with st.form("code_agent_ultraflex_rag_form", clear_on_submit=True):
        rag_question = st.text_input(
            "Ask UltraFLEX RAG",
            placeholder="Ask a tester/IG-XL question and retrieve from RAG directly...",
        )
        rag_submit = st.form_submit_button("Ask UltraFLEX RAG", use_container_width=True)

    if rag_submit and rag_question.strip():
        q = rag_question.strip()
        with st.spinner("Querying UltraFLEX RAG..."):
            rag_hybrid = hybrid_query_documents(
                q,
                n_per_query=5,
                web_mode="always",
                mmr_lambda=0.7,
            )
            if rag_hybrid.get("abstain"):
                rag_answer = rag_hybrid.get("abstain_reason", "No answer available from RAG.")
                rag_model = "-"
                rag_complexity = "abstain"
            else:
                rag_answer, rag_model, rag_complexity = generate_hybrid_response(q, rag_hybrid)

        active_project.setdefault("messages", []).append(
            {"role": "user", "content": f"[UltraFLEX RAG] {q}"}
        )
        active_project.setdefault("messages", []).append(
            {
                "role": "assistant",
                "content": (
                    f"{rag_answer}\n\n"
                    f"Model: {rag_model} | Complexity: {rag_complexity} | "
                    f"KB chunks: {len(rag_hybrid.get('internal_chunks', []))} | "
                    f"Web: {len(rag_hybrid.get('web_results', []))}"
                ),
            }
        )
        st.session_state.code_agent_projects = projects
        st.rerun()

    if st.button("Scan attached directories", use_container_width=True, key="scan_code_agent_dirs"):
        paths = [p.strip() for p in active_project.get("paths", "").splitlines() if p.strip()]
        with st.spinner("Scanning directories and building debug context..."):
            workspace_context, workspace_summary = _collect_code_agent_workspace(
                paths,
                max_files=max_files,
                max_chars_per_file=max_chars,
            )
        active_project["workspace_context"] = workspace_context
        active_project["workspace_summary"] = workspace_summary
        active_project["workspace_sig"] = "|".join(paths) + f"::{max_files}::{max_chars}"
        st.session_state.code_agent_projects = projects
        st.success(f"Workspace loaded: {workspace_summary['scanned_files']} files")

    summary = active_project.get("workspace_summary")
    if summary:
        st.info(
            f"Scanned files: {summary['scanned_files']} | "
            f"Max files: {summary['max_files']} | "
            f"Max chars/file: {summary['max_chars_per_file']}"
        )
        if summary.get("missing_paths"):
            st.warning("Missing/invalid paths: " + ", ".join(summary["missing_paths"]))
        with st.expander("Scanned file list", expanded=False):
            for row in summary.get("files", [])[:80]:
                st.markdown(f"- {row['path']} ({row['size']} bytes)")

    for msg in active_project.get("messages", []):
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    code_question = st.chat_input(
        "Ask the code agent, or use /init /plan /fix /review /edit ...",
        key="code_agent_chat_input",
    )
    if not code_question:
        return

    transformed_question, command_used = _apply_claude_code_command(code_question, active_project)

    active_project.setdefault("messages", []).append({"role": "user", "content": code_question})
    st.session_state.code_agent_projects = projects
    with st.chat_message("user"):
        st.markdown(code_question)
        if command_used:
            st.caption(f"Command mode: {command_used}")

    workspace_context = active_project.get("workspace_context", "")
    attached_dirs = [p.strip() for p in active_project.get("paths", "").splitlines() if p.strip()]
    project_notes = (active_project.get("notes") or "").strip()

    # Allow users to include directories directly in the chat request.
    inferred_dirs = _extract_directory_paths_from_text(code_question)
    if inferred_dirs:
        merged = list(dict.fromkeys(attached_dirs + inferred_dirs))
        active_project["paths"] = "\n".join(merged)
        attached_dirs = merged
        st.info(f"Added directories from prompt: {len(inferred_dirs)}")

    hybrid = {
        "query_type": "U",
        "internal_chunks": [],
        "web_results": [],
        "internal_confident": False,
        "internal_sufficient": False,
        "web_confident": False,
        "abstain": False,
        "abstain_reason": "",
    }

    if use_rag_context and not sdk_mode:
        with st.spinner("Retrieving KB/web context for code-agent reasoning..."):
            hybrid = hybrid_query_documents(
                transformed_question,
                n_per_query=n_results,
                web_mode=web_mode,
                mmr_lambda=mmr_lambda,
            )

    if (not sdk_mode) and attached_dirs:
        current_sig = "|".join(attached_dirs) + f"::{max_files}::{max_chars}"
        if current_sig != active_project.get("workspace_sig", ""):
            with st.spinner("Refreshing directory context from attached paths..."):
                workspace_context, workspace_summary = _collect_code_agent_workspace(
                    attached_dirs,
                    max_files=max_files,
                    max_chars_per_file=max_chars,
                )
            active_project["workspace_context"] = workspace_context
            active_project["workspace_summary"] = workspace_summary
            active_project["workspace_sig"] = current_sig
            st.session_state.code_agent_projects = projects
            st.caption(f"Directory context refreshed: {workspace_summary['scanned_files']} files")

    workspace_context = active_project.get("workspace_context", "")
    if project_notes:
        workspace_context = (
            "=== PROJECT NOTES / CODE SNIPPETS ===\n"
            f"{project_notes[:12000]}\n\n"
            + workspace_context
        )

    if not workspace_context and not sdk_mode:
        st.info("No directory context loaded. Proceeding with retrieval context only.")

    with st.chat_message("assistant"):
        telemetry = {}
        if sdk_mode:
            if use_ate_agent:
                with st.spinner("ATE code agent executing workflow..."):
                    if not attached_dirs:
                        answer = "Attach at least one project directory to use ATE tools workflow."
                        model_used = "ate-agent"
                        complexity = "blocked"
                        tools_used = []
                    elif command_used == "init":
                        ate_user_query = transformed_question if command_used else code_question
                        if project_notes:
                            ate_user_query += f"\n\nProject notes/code:\n{project_notes[:4000]}"

                        llm_answer, llm_model, llm_complexity = _run_ate_code_agent_query(
                            attached_dirs[0],
                            ate_user_query,
                            approved=ate_approved,
                        )

                        if "ATE agent execution error:" in llm_answer:
                            report_md, report_path = _build_ate_init_markdown(
                                attached_dirs[0],
                                project_notes,
                            )
                            if not report_md:
                                answer = f"Failed to generate init report: {report_path}"
                                model_used = "local-init"
                                complexity = "error"
                                tools_used = ["ATE-INIT"]
                            else:
                                preview = report_md[:3000]
                                answer = (
                                    f"{llm_answer}\n\n"
                                    f"Fallback generated init report: {report_path}\n\n"
                                    f"Preview:\n\n```markdown\n{preview}\n```"
                                )
                                model_used = "local-init"
                                complexity = "init-fallback"
                                tools_used = ["ATE-MCP", "ATE-INIT"]
                        else:
                            answer = llm_answer
                            model_used = llm_model
                            complexity = llm_complexity
                            tools_used = ["ATE-MCP"]
                    else:
                        ate_user_query = transformed_question if command_used else code_question
                        if project_notes:
                            ate_user_query += f"\n\nProject notes/code:\n{project_notes[:4000]}"
                        if _looks_like_ate_query(ate_user_query):
                            answer, model_used, complexity = _run_ate_code_agent_query(
                                attached_dirs[0],
                                ate_user_query,
                                approved=ate_approved,
                            )
                            tools_used = ["ATE-MCP"]
                        else:
                            # Keep ATE toggle enabled but route general prompts to the standard code-agent path.
                            sdk_agent = _get_code_agent_sdk()
                            project_contexts = [SimpleNamespace(path=p) for p in attached_dirs]
                            sdk_kwargs = {
                                "user_query": ate_user_query,
                                "allowed_tools": st.session_state.code_agent_allowed_tools,
                                "session_id": st.session_state.current_session_id,
                                "project_contexts": project_contexts,
                            }
                            if "use_internal_kb" in inspect.signature(sdk_agent.process_with_agent_sdk).parameters:
                                sdk_kwargs["use_internal_kb"] = sdk_use_internal_kb
                            sdk_response = _run_async_blocking(sdk_agent.process_with_agent_sdk(**sdk_kwargs))
                            answer = (
                                "Routed to general code-agent workflow because this prompt does not look ATE-specific.\n\n"
                                + sdk_response.answer
                            )
                            model_used = sdk_response.model_used
                            complexity = f"{sdk_response.complexity}-routed"
                            tools_used = sdk_response.tools_used
            else:
                with st.spinner("Code agent executing tool-enabled workflow..."):
                    sdk_agent = _get_code_agent_sdk()
                    project_contexts = [SimpleNamespace(path=p) for p in attached_dirs]
                    sdk_user_query = code_question
                    if command_used:
                        sdk_user_query = transformed_question
                    if project_notes:
                        sdk_user_query += f"\n\nProject notes/code:\n{project_notes[:4000]}"
                    if active_project.get("attached_file_paths"):
                        attached_file_block = "\n".join(active_project.get("attached_file_paths", [])[:40])
                        sdk_user_query += f"\n\nAttached project files:\n{attached_file_block}"
                    sdk_kwargs = {
                        "user_query": sdk_user_query,
                        "allowed_tools": st.session_state.code_agent_allowed_tools,
                        "session_id": st.session_state.current_session_id,
                        "project_contexts": project_contexts,
                        "approved_actions": approve_file_actions,
                    }
                    # Backward-compatible call for older loaded SDK versions.
                    if "use_internal_kb" in inspect.signature(sdk_agent.process_with_agent_sdk).parameters:
                        sdk_kwargs["use_internal_kb"] = sdk_use_internal_kb
                    sdk_response = _run_async_blocking(
                        sdk_agent.process_with_agent_sdk(**sdk_kwargs)
                    )
                    answer = sdk_response.answer
                    model_used = sdk_response.model_used
                    complexity = sdk_response.complexity
                    tools_used = sdk_response.tools_used
                    telemetry = getattr(sdk_response, "telemetry", {}) or {}
        else:
            with st.spinner("Code agent analyzing workspace + retrieval context..."):
                answer, model_used, complexity = generate_code_agent_hybrid_response(
                    transformed_question,
                    workspace_context=workspace_context,
                    hybrid=hybrid,
                    debug_goal="Diagnose and fix UltraFlex/IG-XL test-program issues using attached directory files.",
                )
            tools_used = []
            telemetry = {}
        st.markdown(answer)
        if sdk_mode:
            telemetry_bits = ""
            if telemetry:
                retries = telemetry.get("retries", 0)
                conf = telemetry.get("confidence_score", "n/a")
                failures = telemetry.get("failure_reasons", [])
                telemetry_bits = (
                    f" | Planner conf: {conf} | Retries: {retries}"
                    + (f" | Failures: {', '.join(failures)}" if failures else "")
                )
            st.caption(
                f"Model: {model_used} | Complexity: {complexity} | "
                f"Tools used: {', '.join(tools_used) if tools_used else 'None'} | "
                f"Attached dirs: {len(attached_dirs)}"
                + (f" | ATE approval: {'yes' if ate_approved else 'no'}" if use_ate_agent else "")
                + (f" | File actions approved: {'yes' if approve_file_actions else 'no'}")
                + telemetry_bits
            )
        else:
            st.caption(
                f"Model: {model_used} | Complexity: {complexity} | "
                f"KB chunks: {len(hybrid['internal_chunks'])} | Web: {len(hybrid['web_results'])} | "
                f"Mode: {web_mode if use_rag_context else 'code-only'}"
            )

    active_project.setdefault("messages", []).append({"role": "assistant", "content": answer})
    st.session_state.code_agent_projects = projects


def _render_search_tab() -> None:
    st.subheader("Search")
    st.caption("Semantic search over internal knowledge base and optional web results.")

    with st.form("search_tab_form", clear_on_submit=False):
        search_query = st.text_input(
            "Search query",
            placeholder="Search tester docs, program notes, and instrument details...",
            key="search_tab_query",
            label_visibility="collapsed",
        )
        search_clicked = st.form_submit_button("Search", use_container_width=True)

    if not search_clicked or not search_query.strip():
        st.info("Enter a query and click Search.")
        return

    with st.spinner("Searching..."):
        hybrid = hybrid_query_documents(
            search_query.strip(),
            n_per_query=n_results,
            web_mode=web_mode,
            mmr_lambda=mmr_lambda,
        )

    st.markdown(f"Internal chunks: **{len(hybrid.get('internal_chunks', []))}** | Web results: **{len(hybrid.get('web_results', []))}**")
    if hybrid.get("abstain"):
        st.warning(hybrid.get("abstain_reason", "No strong match found."))

    if hybrid.get("internal_chunks"):
        with st.expander(f"Internal Matches ({len(hybrid['internal_chunks'])})", expanded=True):
            for src in hybrid["internal_chunks"][:20]:
                st.markdown(f"**{src.get('source', 'Unknown')}** | Type: `{src.get('file_type', '')}`")
                st.code((src.get("text", "") or "")[:280], language="text")

    if hybrid.get("web_results"):
        with st.expander(f"Web Matches ({len(hybrid['web_results'])})", expanded=False):
            for row in hybrid["web_results"][:20]:
                st.markdown(f"**[{row.get('title', 'Untitled')}]({row.get('url', '#')})**")
                st.caption((row.get("snippet", "") or "")[:240])


def _render_settings_tab() -> None:
    st.subheader("Settings")
    st.caption("Primary controls remain in the sidebar for a clean, dark-mode compatible layout.")
    st.markdown(
        "\n".join(
            [
                f"- Active model: `{st.session_state.get('selected_model', DEFAULT_LLM_MODEL)}`",
                f"- Web mode: `{web_mode}`",
                f"- Chunks per query: `{n_results}`",
                f"- Show sources: `{show_sources}`",
                f"- KB chunks: `{collection.count()}`",
            ]
        )
    )


chat_tab, code_agent_tab = st.tabs(["Chat", "Code Agent"])
with chat_tab:
    _render_chat_tab()
with code_agent_tab:
    _render_code_agent_tab()