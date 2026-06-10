"""FastAPI wrapper for UltraFLEX RAG retrieval and answer generation."""

from __future__ import annotations

import os
import ssl
import sys
from pathlib import Path
from typing import Any

import httpx
from fastapi import FastAPI
from pydantic import BaseModel, Field

os.environ["PYTHONHTTPSVERIFY"] = "0"
os.environ["CURL_CA_BUNDLE"] = ""
os.environ["REQUESTS_CA_BUNDLE"] = ""
ssl._create_default_https_context = ssl._create_unverified_context

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = PROJECT_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))


class RagQueryRequest(BaseModel):
    question: str = Field(..., description="Question to ask the RAG system")
    web_mode: str = Field("off", description="off|fallback|always|web_only")
    n_per_query: int = Field(2, ge=1, description="Top-k chunks per query expansion")
    fast_mode: bool = Field(True, description="Low-latency retrieval path")
    concise_answer: bool = Field(True, description="Short answer mode")
    deep_mode: bool = Field(False, description="Disable fast/concise for higher recall")
    mode: str = Field("answer", description="answer|retrieve_only")
    agentic_mode: bool = Field(False, description="Enable planner+judge iterative retrieval")
    tool_budget: int = Field(6, ge=0, description="Grep/read tool budget for agentic retrieval")


app = FastAPI(title="UltraFLEX RAG API", version="1.0.0")
_VECTOR_ONLY_MODE = os.getenv("ULTRAFLEX_MCP_VECTOR_ONLY", "false").strip().lower() == "true"


@app.get("/health")
def health() -> dict[str, Any]:
    return {
        "status": "ok",
        "service": "ultraflex-rag-http",
        "vector_only": _VECTOR_ONLY_MODE,
    }


@app.post("/query")
def query(req: RagQueryRequest) -> dict[str, Any]:
    question = (req.question or "").strip()
    if not question:
        return {"error": "question is required"}

    normalized_web_mode = (req.web_mode or "off").strip().lower()
    if normalized_web_mode not in {"off", "fallback", "always", "web_only"}:
        normalized_web_mode = "off"
    if _VECTOR_ONLY_MODE:
        normalized_web_mode = "off"

    normalized_mode = (req.mode or "answer").strip().lower()
    if normalized_mode not in {"answer", "retrieve_only"}:
        normalized_mode = "answer"

    if req.deep_mode:
        effective_fast_mode = False
        effective_concise = False
        safe_k = max(1, int(req.n_per_query))
    else:
        effective_fast_mode = bool(req.fast_mode)
        effective_concise = bool(req.concise_answer)
        requested_k = max(1, int(req.n_per_query))
        fast_default_k = max(1, int(os.getenv("ULTRAFLEX_MCP_FAST_N_PER_QUERY", "2")))
        safe_k = min(requested_k, fast_default_k) if effective_fast_mode else requested_k

    if req.agentic_mode:
        from page_indexing_rag.agentic_loop import run_agentic_retrieve_loop

        hybrid = run_agentic_retrieve_loop(
            question,
            n_per_query=safe_k,
            max_iterations=3,
            tool_budget=max(0, int(req.tool_budget)),
        )
    else:
        from page_indexing_rag.retrieval import hybrid_query_documents

        hybrid = hybrid_query_documents(
            question,
            n_per_query=safe_k,
            web_mode=normalized_web_mode,
            mmr_lambda=0.7,
            fast_mode=effective_fast_mode,
        )

    base_payload: dict[str, Any] = {
        "mode": normalized_mode,
        "retrieval_style": "agentic" if req.agentic_mode else "semantic",
        "vector_only": _VECTOR_ONLY_MODE,
        "query_type": hybrid.get("query_type", "U"),
        "web_mode": hybrid.get("web_mode", normalized_web_mode),
        "abstain": hybrid.get("abstain", False),
        "abstain_reason": hybrid.get("abstain_reason", ""),
        "internal_sources": [
            {
                "source": src.get("source", "Unknown"),
                "page": src.get("page_number", 0),
                "file_type": src.get("file_type", ""),
            }
            for src in hybrid.get("internal_chunks", [])[:10]
        ],
        "web_sources": [
            {
                "title": row.get("title", "Untitled"),
                "url": row.get("url", ""),
            }
            for row in hybrid.get("web_results", [])[:10]
        ],
    }
    if hybrid.get("agentic_trace"):
        base_payload["agentic_trace"] = hybrid.get("agentic_trace")

    if normalized_mode == "retrieve_only":
        base_payload["internal_context"] = [
            {
                "source": src.get("source", "Unknown"),
                "page": src.get("page_number", 0),
                "file_type": src.get("file_type", ""),
                "distance": src.get("distance", None),
                "snippet": (str(src.get("text", ""))[:400] + "...")
                if len(str(src.get("text", ""))) > 400
                else str(src.get("text", "")),
            }
            for src in hybrid.get("internal_chunks", [])[:6]
        ]
        base_payload["web_context"] = [
            {
                "title": row.get("title", "Untitled"),
                "url": row.get("url", ""),
                "snippet": str(row.get("snippet", ""))[:400],
            }
            for row in hybrid.get("web_results", [])[:5]
        ]
        return base_payload

    from page_indexing_rag.generation import generate_hybrid_response

    answer, model_used, complexity = generate_hybrid_response(
        question,
        hybrid,
        concise_mode=effective_concise,
    )
    base_payload.update(
        {
            "answer": answer,
            "model_used": model_used,
            "complexity": complexity,
        }
    )
    return base_payload
