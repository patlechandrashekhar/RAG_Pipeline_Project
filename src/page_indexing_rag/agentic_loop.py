"""Cookbook-style agentic retrieval loop for hybrid RAG systems.

This module adds an iterative retrieval workflow on top of the existing
vector retrieval stack:
1. Planner node proposes retrieval sub-queries.
2. Retrieve-only loop gathers vector evidence.
3. Grep/read-file tools collect exact evidence from local text corpora.
4. Evidence judge decides whether evidence is sufficient.
"""

from __future__ import annotations

import os
import re
import ssl
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import httpx

os.environ["PYTHONHTTPSVERIFY"] = "0"
os.environ["CURL_CA_BUNDLE"] = ""
os.environ["REQUESTS_CA_BUNDLE"] = ""
ssl._create_default_https_context = ssl._create_unverified_context

from .config import HTML_DATA_DIR, TXT_DATA_DIR
from .retrieval import classify_query_type, query_documents

_TEXT_EXTS = {
    ".txt",
    ".md",
    ".rst",
    ".csv",
    ".json",
    ".yaml",
    ".yml",
    ".xml",
    ".html",
    ".htm",
    ".xhtml",
    ".log",
    ".bas",
    ".cls",
    ".vb",
    ".vbs",
    ".py",
    ".ini",
    ".cfg",
}

_STOPWORDS = {
    "the",
    "and",
    "for",
    "with",
    "that",
    "this",
    "from",
    "what",
    "when",
    "where",
    "which",
    "who",
    "your",
    "about",
    "into",
    "need",
    "show",
    "tell",
    "spec",
    "specs",
}


@dataclass
class PlannerNode:
    """Simple deterministic planner for iterative retrieval queries."""

    max_iterations: int = 3

    def build_plan(self, question: str) -> list[str]:
        q = (question or "").strip()
        if not q:
            return []

        tokens = [
            t.lower()
            for t in re.findall(r"[A-Za-z0-9_\-]{4,}", q)
            if t.lower() not in _STOPWORDS
        ]

        unique_tokens: list[str] = []
        for t in tokens:
            if t not in unique_tokens:
                unique_tokens.append(t)

        plan = [q]
        if unique_tokens:
            plan.append(" ".join(unique_tokens[:4]))
        if len(unique_tokens) > 2:
            plan.append(f"{unique_tokens[0]} {unique_tokens[-1]} troubleshooting")

        deduped: list[str] = []
        for p in plan:
            p = p.strip()
            if p and p.lower() not in {d.lower() for d in deduped}:
                deduped.append(p)

        return deduped[: max(1, self.max_iterations)]


@dataclass
class GrepReadTools:
    """Local grep/read tools with bounded file and read budgets."""

    search_roots: list[Path]
    max_files_scan: int = 200
    max_chars_per_read: int = 4000

    def _iter_text_files(self):
        scanned = 0
        for root in self.search_roots:
            if not root.exists() or not root.is_dir():
                continue
            for p in root.rglob("*"):
                if scanned >= self.max_files_scan:
                    return
                if p.is_file() and p.suffix.lower() in _TEXT_EXTS:
                    scanned += 1
                    yield p

    def grep(self, pattern: str, max_hits: int = 40) -> list[dict[str, Any]]:
        if not pattern.strip():
            return []
        try:
            rx = re.compile(pattern, re.IGNORECASE)
        except re.error:
            rx = re.compile(re.escape(pattern), re.IGNORECASE)

        hits: list[dict[str, Any]] = []
        for p in self._iter_text_files():
            try:
                lines = p.read_text(encoding="utf-8", errors="ignore").splitlines()
            except OSError:
                continue
            for idx, line in enumerate(lines, start=1):
                if rx.search(line):
                    hits.append(
                        {
                            "file": str(p),
                            "line": idx,
                            "text": line[:300],
                        }
                    )
                    if len(hits) >= max_hits:
                        return hits
        return hits

    def read_file(self, file_path: str) -> dict[str, Any] | None:
        target = Path(file_path)
        if not target.exists() or not target.is_file():
            return None

        allowed = False
        for root in self.search_roots:
            try:
                if target.resolve().is_relative_to(root.resolve()):
                    allowed = True
                    break
            except Exception:
                continue
        if not allowed:
            return None

        try:
            content = target.read_text(encoding="utf-8", errors="ignore")
        except OSError:
            return None

        trimmed = content[: self.max_chars_per_read]
        return {
            "file": str(target),
            "content": trimmed,
            "truncated": len(content) > len(trimmed),
        }


@dataclass
class EvidenceJudgeNode:
    """Judge whether gathered evidence is sufficient to answer reliably."""

    min_relevant_chunks: int = 2
    max_good_distance: float = 0.42

    def judge(
        self,
        question: str,
        chunks: list[dict[str, Any]],
        grep_hits: list[dict[str, Any]],
        reads: list[dict[str, Any]],
    ) -> dict[str, Any]:
        _ = question
        strong_chunks = [c for c in chunks if float(c.get("distance", 1.0)) <= self.max_good_distance]
        sufficient = len(strong_chunks) >= self.min_relevant_chunks or (
            len(chunks) >= self.min_relevant_chunks and len(grep_hits) >= 2 and len(reads) >= 1
        )
        return {
            "sufficient": sufficient,
            "strong_chunk_count": len(strong_chunks),
            "chunk_count": len(chunks),
            "grep_hits": len(grep_hits),
            "read_count": len(reads),
        }


def _default_search_roots() -> list[Path]:
    roots = [Path(TXT_DATA_DIR), Path(HTML_DATA_DIR)]
    project_root = Path(__file__).resolve().parents[2]
    local_data = project_root / "data" / "Data"
    if local_data.exists():
        roots.append(local_data)
    return roots


def _merge_chunks(existing: dict[str, dict[str, Any]], incoming: list[dict[str, Any]]) -> None:
    for chunk in incoming:
        key = "|".join(
            [
                str(chunk.get("source", "")),
                str(chunk.get("page_number", 0)),
                str(chunk.get("chunk_index", 0)),
            ]
        )
        if key not in existing or chunk.get("distance", 1.0) < existing[key].get("distance", 1.0):
            existing[key] = chunk


def run_agentic_retrieve_loop(
    question: str,
    n_per_query: int = 3,
    max_iterations: int = 3,
    tool_budget: int = 6,
) -> dict[str, Any]:
    """Run planner -> retrieve loop -> grep/read -> evidence judge.

    Returns a hybrid-like payload compatible with generation functions.
    """
    planner = PlannerNode(max_iterations=max_iterations)
    search_plan = planner.build_plan(question)

    chunk_map: dict[str, dict[str, Any]] = {}
    step_log: list[dict[str, Any]] = []

    for step_idx, subquery in enumerate(search_plan, start=1):
        retrieved = query_documents(
            subquery,
            n_per_query=max(1, int(n_per_query)),
            use_query_expansion=False,
            use_rerank=False,
            max_results=max(2, int(n_per_query)),
        )
        _merge_chunks(chunk_map, retrieved)
        step_log.append(
            {
                "step": step_idx,
                "type": "vector_query",
                "query": subquery,
                "chunks": len(retrieved),
            }
        )

    chunks = sorted(chunk_map.values(), key=lambda c: float(c.get("distance", 1.0)))[:8]

    keywords = [
        t
        for t in re.findall(r"[A-Za-z0-9_\-]{4,}", question)
        if t.lower() not in _STOPWORDS
    ]
    pattern_terms = keywords[: max(1, min(4, len(keywords)))]
    grep_pattern = "|".join(re.escape(t) for t in pattern_terms) if pattern_terms else re.escape(question.strip())

    tools = GrepReadTools(search_roots=_default_search_roots())
    grep_budget = max(0, tool_budget // 2)
    read_budget = max(0, tool_budget - grep_budget)

    grep_hits = tools.grep(grep_pattern, max_hits=max(10, grep_budget * 10)) if grep_budget > 0 else []
    if grep_hits:
        step_log.append(
            {
                "step": len(step_log) + 1,
                "type": "grep",
                "pattern": grep_pattern,
                "hits": len(grep_hits),
            }
        )

    reads: list[dict[str, Any]] = []
    seen_files: set[str] = set()
    for hit in grep_hits:
        if len(reads) >= read_budget:
            break
        file_path = str(hit.get("file", ""))
        if not file_path or file_path in seen_files:
            continue
        seen_files.add(file_path)
        doc = tools.read_file(file_path)
        if doc:
            reads.append(doc)

    if reads:
        step_log.append(
            {
                "step": len(step_log) + 1,
                "type": "read_file",
                "files": [r["file"] for r in reads],
            }
        )

    judge = EvidenceJudgeNode().judge(question, chunks, grep_hits, reads)

    return {
        "query_type": classify_query_type(question),
        "web_mode": "off",
        "internal_chunks": chunks,
        "web_results": [],
        "internal_confident": judge["strong_chunk_count"] >= 2,
        "internal_sufficient": bool(judge["sufficient"]),
        "web_confident": False,
        "abstain": not bool(judge["sufficient"]),
        "abstain_reason": "Insufficient evidence from agentic retrieval loop." if not judge["sufficient"] else "",
        "agentic_trace": {
            "plan": search_plan,
            "steps": step_log,
            "judge": judge,
            "grep_pattern": grep_pattern,
            "tool_budget": int(tool_budget),
        },
    }
