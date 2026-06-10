"""
Core ATE Code Agent wiring based on Claude Agent SDK.
"""

import os
import ssl
import httpx
import anyio
import json
import sys
from pathlib import Path

from system_prompt import ATE_SYSTEM_PROMPT
from safety.rules import SAFETY_RULES
from tools.excel_tools import read_excel_sheets, write_excel_cell
from tools.vba_tools import read_vba_module, write_vba_module, get_vba_module_names
from tools.run_tools import run_igxl_program
from tools.log_tools import parse_stdf_log, summarize_results
from tools.git_tools import git_commit_backup

os.environ["PYTHONHTTPSVERIFY"] = "0"
os.environ["CURL_CA_BUNDLE"] = ""
os.environ["REQUESTS_CA_BUNDLE"] = ""
ssl._create_default_https_context = ssl._create_unverified_context

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = PROJECT_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

try:
    from claude_agent_sdk import (
        tool,
        create_sdk_mcp_server,
        ClaudeAgentOptions,
        ClaudeSDKClient,
        AssistantMessage,
        TextBlock,
        ResultMessage,
    )
except Exception as exc:  # pragma: no cover
    raise RuntimeError(
        "claude_agent_sdk is required. Install dependencies from requirements.txt"
    ) from exc


@tool(
    name="read_igxl_excel",
    description="Read IG-XL Excel sheets from test program workbook.",
    input_schema={"file_path": str},
)
async def tool_read_excel(args: dict) -> dict:
    try:
        data = read_excel_sheets(args["file_path"])
        return {"content": [{"type": "text", "text": json.dumps(data, indent=2, default=str)}]}
    except Exception as exc:
        return {"content": [{"type": "text", "text": f"ERROR: {exc}"}], "isError": True}


@tool(
    name="write_igxl_cell",
    description="Write one IG-XL sheet cell with safety checks and git backup.",
    input_schema={"file_path": str, "sheet_name": str, "row": int, "col": int, "value": str},
)
async def tool_write_excel(args: dict) -> dict:
    try:
        value = args["value"]
        try:
            value = float(value) if "." in str(value) else int(value)
        except (ValueError, TypeError):
            pass

        result = write_excel_cell(
            args["file_path"],
            args["sheet_name"],
            args["row"],
            args["col"],
            value,
        )
        return {"content": [{"type": "text", "text": json.dumps(result)}]}
    except Exception as exc:
        return {"content": [{"type": "text", "text": f"ERROR: {exc}"}], "isError": True}


@tool(
    name="list_vba_modules",
    description="List VBA module names from workbook.",
    input_schema={"file_path": str},
)
async def tool_list_vba_modules(args: dict) -> dict:
    try:
        names = get_vba_module_names(args["file_path"])
        return {"content": [{"type": "text", "text": json.dumps(names)}]}
    except Exception as exc:
        return {"content": [{"type": "text", "text": f"ERROR: {exc}"}], "isError": True}


@tool(
    name="read_vba_module",
    description="Read a full VBA module source code.",
    input_schema={"file_path": str, "module_name": str},
)
async def tool_read_vba(args: dict) -> dict:
    try:
        code = read_vba_module(args["file_path"], args["module_name"])
        return {"content": [{"type": "text", "text": code}]}
    except Exception as exc:
        return {"content": [{"type": "text", "text": f"ERROR: {exc}"}], "isError": True}


@tool(
    name="write_vba_module",
    description="Overwrite one VBA module with full source text.",
    input_schema={"file_path": str, "module_name": str, "new_code": str},
)
async def tool_write_vba(args: dict) -> dict:
    try:
        result = write_vba_module(args["file_path"], args["module_name"], args["new_code"])
        return {"content": [{"type": "text", "text": json.dumps(result)}]}
    except Exception as exc:
        return {"content": [{"type": "text", "text": f"ERROR: {exc}"}], "isError": True}


@tool(
    name="run_igxl_program",
    description="Execute IG-XL test program in engineering mode.",
    input_schema={"program_path": str, "log_output_path": str, "timeout_seconds": int},
)
async def tool_run_program(args: dict) -> dict:
    try:
        result = run_igxl_program(
            args["program_path"],
            args["log_output_path"],
            timeout_seconds=args.get("timeout_seconds", 600),
        )
        return {"content": [{"type": "text", "text": json.dumps(result)}]}
    except Exception as exc:
        return {"content": [{"type": "text", "text": f"ERROR: {exc}"}], "isError": True}


@tool(
    name="parse_test_log",
    description="Parse STDF log and summarize pass/fail metrics.",
    input_schema={"log_path": str},
)
async def tool_parse_log(args: dict) -> dict:
    try:
        results = parse_stdf_log(args["log_path"])
        summary = summarize_results(results)
        return {
            "content": [
                {"type": "text", "text": json.dumps({"summary": summary, "results": results}, indent=2, default=str)}
            ]
        }
    except Exception as exc:
        return {"content": [{"type": "text", "text": f"ERROR: {exc}"}], "isError": True}


@tool(
    name="git_backup",
    description="Create a git backup commit for a file.",
    input_schema={"file_path": str, "message": str},
)
async def tool_git_backup(args: dict) -> dict:
    try:
        commit_hash = git_commit_backup(args["file_path"], args["message"])
        return {"content": [{"type": "text", "text": f"Backup committed: {commit_hash}"}]}
    except Exception as exc:
        return {"content": [{"type": "text", "text": f"ERROR: {exc}"}], "isError": True}


@tool(
    name="query_ultraflex_docs",
    description=(
        "Query the existing UltraFLEX/IG-XL RAG knowledge base and return answer with sources. "
        "Use this for domain documentation lookups before making edits."
    ),
    input_schema={"question": str, "web_mode": str},
)
async def tool_query_ultraflex_docs(args: dict) -> dict:
    """Expose existing page_indexing_rag retrieval+generation as an MCP tool."""
    try:
        # Lazy import to keep agent startup resilient when optional deps are missing.
        from page_indexing_rag.retrieval import hybrid_query_documents
        from page_indexing_rag.generation import generate_hybrid_response

        question = str(args.get("question", "")).strip()
        if not question:
            return {
                "content": [{"type": "text", "text": "ERROR: question is required"}],
                "isError": True,
            }

        web_mode = str(args.get("web_mode", "fallback") or "fallback").strip()
        hybrid = hybrid_query_documents(question, n_per_query=5, web_mode=web_mode, mmr_lambda=0.7)
        answer, model_used, complexity = generate_hybrid_response(question, hybrid)

        # Keep payload compact for tool context efficiency.
        internal_sources = [
            {
                "source": src.get("source", "Unknown"),
                "page": src.get("page_number", 0),
                "file_type": src.get("file_type", ""),
            }
            for src in hybrid.get("internal_chunks", [])[:8]
        ]
        web_sources = [
            {
                "title": w.get("title", "Untitled"),
                "url": w.get("url", ""),
            }
            for w in hybrid.get("web_results", [])[:8]
        ]

        payload = {
            "answer": answer,
            "model_used": model_used,
            "complexity": complexity,
            "query_type": hybrid.get("query_type", "U"),
            "web_mode": hybrid.get("web_mode", web_mode),
            "abstain": hybrid.get("abstain", False),
            "abstain_reason": hybrid.get("abstain_reason", ""),
            "internal_sources": internal_sources,
            "web_sources": web_sources,
        }
        return {"content": [{"type": "text", "text": json.dumps(payload, indent=2, default=str)}]}
    except Exception as exc:
        return {"content": [{"type": "text", "text": f"ERROR: {exc}"}], "isError": True}


ate_tool_server = create_sdk_mcp_server(
    name="ate-tools",
    version="1.0.0",
    tools=[
        tool_read_excel,
        tool_write_excel,
        tool_list_vba_modules,
        tool_read_vba,
        tool_write_vba,
        tool_run_program,
        tool_parse_log,
        tool_git_backup,
        tool_query_ultraflex_docs,
    ],
)


def get_agent_options(program_path: str, approved: bool = False) -> ClaudeAgentOptions:
    """Build agent options with ATE tools and approval-aware permissions."""
    allowed = [
        "mcp__ate-tools__read_igxl_excel",
        "mcp__ate-tools__list_vba_modules",
        "mcp__ate-tools__read_vba_module",
        "mcp__ate-tools__parse_test_log",
        "mcp__ate-tools__git_backup",
        "mcp__ate-tools__query_ultraflex_docs",
    ]

    if approved:
        allowed.extend([
            "mcp__ate-tools__write_igxl_cell",
            "mcp__ate-tools__write_vba_module",
            "mcp__ate-tools__run_igxl_program",
        ])

    return ClaudeAgentOptions(
        system_prompt=ATE_SYSTEM_PROMPT,
        cwd=str(program_path),
        mcp_servers={"ate-tools": ate_tool_server},
        allowed_tools=allowed,
        tools=["Read", "Glob", "Grep"],
        max_turns=50,
        model="claude-sonnet-4-6",
    )


def _requires_human_approval(task: str) -> bool:
    """Detect write/run style requests that require explicit approval."""
    text = (task or "").lower()
    risky_tokens = [
        "write",
        "edit",
        "modify",
        "change",
        "update",
        "overwrite",
        "set limit",
        "run",
        "execute",
        "retest",
    ]
    return any(token in text for token in risky_tokens)


async def run_interactive_agent(program_path: str) -> None:
    """Run an interactive ATE agent session."""
    print("\n" + "=" * 60)
    print("  ATE CODE AGENT - UltraFLEX / IG-XL")
    print("  Powered by Claude Agent SDK")
    print("=" * 60)
    print(f"  Program: {program_path}")
    print("  Type your task. Type 'quit' to exit.")
    print("=" * 60 + "\n")

    while True:
        try:
            task = input("You: ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\nExiting ATE Agent.")
            break

        if task.lower() in {"quit", "exit", "q"}:
            print("Exiting ATE Agent.")
            break

        if not task:
            continue

        approved = False
        if SAFETY_RULES.get("require_human_approval", True) and _requires_human_approval(task):
            print(
                "Approval required: this task may edit files or run hardware. "
                "Type APPROVE to continue, or anything else to cancel."
            )
            approval = input("Approval: ").strip()
            if approval.upper() != "APPROVE":
                print("Task canceled. No edits or runs were executed.\n")
                continue
            approved = True
            task = (
                task
                + "\n\n"
                + "Approval granted by engineer for this task. "
                + "You may execute required write/run tools while respecting safety rules."
            )

        print("\nAgent: ", end="", flush=True)
        options = get_agent_options(program_path, approved=approved)
        async with ClaudeSDKClient(options=options) as client:
            async for message in client.query(task):
                if isinstance(message, AssistantMessage):
                    for block in message.content:
                        if isinstance(block, TextBlock):
                            print(block.text, end="", flush=True)
                elif isinstance(message, ResultMessage) and message.is_error:
                    print(f"\n[Agent stopped with error: {message.result}]")
        print("\n")


async def run_single_query(program_path: str, task: str, approved: bool = False) -> str:
    """Run one task and return full text response."""
    if SAFETY_RULES.get("require_human_approval", True) and _requires_human_approval(task) and not approved:
        return (
            "Approval required before edit/run actions. "
            "Resubmit with approved=true after engineer review."
        )

    options = get_agent_options(program_path, approved=approved)
    full_response: list[str] = []

    async with ClaudeSDKClient(options=options) as client:
        async for message in client.query(task):
            if isinstance(message, AssistantMessage):
                for block in message.content:
                    if isinstance(block, TextBlock):
                        full_response.append(block.text)
            elif isinstance(message, ResultMessage) and message.is_error:
                full_response.append(f"[ERROR] {message.result}")

    return "\n".join(full_response)
