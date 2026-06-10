"""
Standalone MCP stdio server exposing page_indexing_RAG RAG as callable tools.
Use this server from Claude Code MCP config.
"""

from __future__ import annotations

import os
import ssl
import sys
import json
import time
from pathlib import Path
from typing import Any

import httpx
from mcp.server.fastmcp import FastMCP

os.environ["PYTHONHTTPSVERIFY"] = "0"
os.environ["CURL_CA_BUNDLE"] = ""
os.environ["REQUESTS_CA_BUNDLE"] = ""
ssl._create_default_https_context = ssl._create_unverified_context

# Ensure the page_indexing_RAG src modules are importable when server is launched externally.
PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = PROJECT_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))


mcp = FastMCP("ultraflex-rag")
_CACHE_TTL_SECONDS = int(os.getenv("ULTRAFLEX_MCP_CACHE_TTL", "300"))
_RESULT_CACHE: dict[str, tuple[float, str]] = {}
_FAST_QUERY_DEFAULT_K = int(os.getenv("ULTRAFLEX_MCP_FAST_N_PER_QUERY", "2"))
_VECTOR_ONLY_MODE = os.getenv("ULTRAFLEX_MCP_VECTOR_ONLY", "false").strip().lower() == "true"


def _normalize_question_for_cache(question: str) -> str:
    return " ".join(question.strip().lower().split())


def _cache_key(
    question: str,
    web_mode: str,
    n_per_query: int,
    fast_mode: bool,
    concise_answer: bool,
    mode: str,
    agentic_mode: bool,
    tool_budget: int,
) -> str:
    return json.dumps(
        {
            "q": _normalize_question_for_cache(question),
            "web_mode": (web_mode or "off").strip().lower(),
            "n_per_query": max(1, int(n_per_query)),
            "fast_mode": bool(fast_mode),
            "concise_answer": bool(concise_answer),
            "mode": (mode or "answer").strip().lower(),
            "agentic_mode": bool(agentic_mode),
            "tool_budget": max(0, int(tool_budget)),
        },
        sort_keys=True,
    )


def _cache_get(key: str) -> str | None:
    row = _RESULT_CACHE.get(key)
    if not row:
        return None

    ts, value = row
    if time.monotonic() - ts > _CACHE_TTL_SECONDS:
        _RESULT_CACHE.pop(key, None)
        return None
    return value


def _cache_set(key: str, value: str) -> None:
    _RESULT_CACHE[key] = (time.monotonic(), value)


def _compact_sources(hybrid: dict[str, Any]) -> dict[str, Any]:
    internal_sources = [
        {
            "source": src.get("source", "Unknown"),
            "page": src.get("page_number", 0),
            "file_type": src.get("file_type", ""),
        }
        for src in hybrid.get("internal_chunks", [])[:10]
    ]
    web_sources = [
        {
            "title": row.get("title", "Untitled"),
            "url": row.get("url", ""),
        }
        for row in hybrid.get("web_results", [])[:10]
    ]
    return {"internal_sources": internal_sources, "web_sources": web_sources}


def _compact_internal_context(hybrid: dict[str, Any], max_items: int = 6, max_chars: int = 400) -> list[dict[str, Any]]:
    """Return concise chunk payload for external answer generation."""
    items: list[dict[str, Any]] = []
    for src in hybrid.get("internal_chunks", [])[:max_items]:
        text = str(src.get("text", "")).strip()
        snippet = text[:max_chars] + ("..." if len(text) > max_chars else "")
        items.append(
            {
                "source": src.get("source", "Unknown"),
                "page": src.get("page_number", 0),
                "file_type": src.get("file_type", ""),
                "distance": src.get("distance", None),
                "snippet": snippet,
            }
        )
    return items


@mcp.tool()
def query_ultraflex_docs(
    question: str,
    web_mode: str = "off",
    n_per_query: int = 2,
    fast_mode: bool = True,
    concise_answer: bool = True,
    deep_mode: bool = False,
    mode: str = "answer",
    agentic_mode: bool = False,
    tool_budget: int = 6,
) -> str:
    """
    Query the existing UltraFLEX RAG and return answer + source metadata.

    Args:
        question: User question.
        web_mode: off|fallback|always|web_only.
        n_per_query: Top-k chunks per query expansion.
        fast_mode: If true, use low-latency retrieval path (recommended for MCP).
        concise_answer: If true, generate a short direct answer for lower latency.
        deep_mode: If true, disable fast/concise behavior for higher-recall responses.
        mode: answer|retrieve_only. retrieve_only returns chunks/sources only.
        agentic_mode: If true, use planner+judge iterative retrieval loop.
        tool_budget: Budget for grep/read tool calls in agentic mode.
    """
    if not question or not question.strip():
        return json.dumps({"error": "question is required"}, indent=2)

    normalized_question = question.strip()
    normalized_mode = (web_mode or "off").strip().lower()
    if normalized_mode not in {"off", "fallback", "always", "web_only"}:
        normalized_mode = "off"
    normalized_response_mode = (mode or "answer").strip().lower()
    if normalized_response_mode not in {"answer", "retrieve_only"}:
        normalized_response_mode = "answer"
    safe_tool_budget = max(0, int(tool_budget))
    if _VECTOR_ONLY_MODE:
        normalized_mode = "off"

    if deep_mode:
        effective_fast_mode = False
        effective_concise = False
        safe_k = max(1, int(n_per_query))
    else:
        effective_fast_mode = bool(fast_mode)
        effective_concise = bool(concise_answer)
        requested_k = max(1, int(n_per_query))
        fast_default_k = max(1, int(_FAST_QUERY_DEFAULT_K))
        safe_k = min(requested_k, fast_default_k) if effective_fast_mode else requested_k

    key = _cache_key(
        normalized_question,
        normalized_mode,
        safe_k,
        effective_fast_mode,
        effective_concise,
        normalized_response_mode,
        bool(agentic_mode),
        safe_tool_budget,
    )
    cached = _cache_get(key)
    if cached is not None:
        return cached

    try:
        # Lazy import so MCP server can start even if RAG provider env vars are not set.
        from page_indexing_rag.retrieval import hybrid_query_documents
    except Exception as exc:
        return json.dumps(
            {
                "error": "RAG modules could not be loaded",
                "detail": str(exc),
                "hint": "Set required provider environment variables (Portkey/OpenAI) before calling this tool.",
            },
            indent=2,
        )

    if agentic_mode:
        from page_indexing_rag.agentic_loop import run_agentic_retrieve_loop

        hybrid = run_agentic_retrieve_loop(
            normalized_question,
            n_per_query=safe_k,
            max_iterations=3,
            tool_budget=safe_tool_budget,
        )
    else:
        hybrid = hybrid_query_documents(
            normalized_question,
            n_per_query=safe_k,
            web_mode=normalized_mode,
            mmr_lambda=0.7,
            fast_mode=effective_fast_mode,
        )

    if normalized_response_mode == "retrieve_only":
        payload: dict[str, Any] = {
            "mode": normalized_response_mode,
            "retrieval_style": "agentic" if agentic_mode else "semantic",
            "vector_only": _VECTOR_ONLY_MODE,
            "query_type": hybrid.get("query_type", "U"),
            "web_mode": hybrid.get("web_mode", normalized_mode),
            "abstain": hybrid.get("abstain", False),
            "abstain_reason": hybrid.get("abstain_reason", ""),
            "internal_context": _compact_internal_context(hybrid),
            "web_context": [
                {
                    "title": row.get("title", "Untitled"),
                    "url": row.get("url", ""),
                    "snippet": str(row.get("snippet", ""))[:400],
                }
                for row in hybrid.get("web_results", [])[:5]
            ],
        }
        if hybrid.get("agentic_trace"):
            payload["agentic_trace"] = hybrid.get("agentic_trace")
        payload.update(_compact_sources(hybrid))
        serialized = json.dumps(payload, indent=2, default=str)
        _cache_set(key, serialized)
        return serialized

    try:
        from page_indexing_rag.generation import generate_hybrid_response
    except Exception as exc:
        return json.dumps(
            {
                "error": "Generation module could not be loaded",
                "detail": str(exc),
                "hint": "Use mode=retrieve_only or set required provider environment variables.",
            },
            indent=2,
        )

    answer, model_used, complexity = generate_hybrid_response(
        normalized_question,
        hybrid,
        concise_mode=effective_concise,
    )

    payload: dict[str, Any] = {
        "mode": normalized_response_mode,
        "retrieval_style": "agentic" if agentic_mode else "semantic",
        "vector_only": _VECTOR_ONLY_MODE,
        "answer": answer,
        "model_used": model_used,
        "complexity": complexity,
        "query_type": hybrid.get("query_type", "U"),
        "web_mode": hybrid.get("web_mode", web_mode),
        "abstain": hybrid.get("abstain", False),
        "abstain_reason": hybrid.get("abstain_reason", ""),
    }
    if hybrid.get("agentic_trace"):
        payload["agentic_trace"] = hybrid.get("agentic_trace")
    payload.update(_compact_sources(hybrid))
    serialized = json.dumps(payload, indent=2, default=str)
    _cache_set(key, serialized)
    return serialized


@mcp.tool()
def ping() -> str:
    """Health check tool for MCP connectivity checks."""
    return "ultraflex-rag MCP server is running"


if __name__ == "__main__":
    # stdio transport by default for Claude Code MCP.
    mcp.run()
