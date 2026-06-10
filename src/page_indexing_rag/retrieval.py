"""Retrieval pipeline: internal vector retrieval plus optional web fallback."""

from __future__ import annotations

import json
import logging
import re
from html import unescape
from urllib.parse import parse_qsl, urlencode, urlparse, urlunparse

import httpx

from langchain_core.output_parsers import StrOutputParser
from langchain_core.prompts import ChatPromptTemplate

from .config import OPENAI_RETRIEVAL_MODEL, OPENAI_WEB_MODEL, get_chat_llm, get_embeddings
from .ingestion import collection

_lc_embeddings = get_embeddings()

_expand_prompt = ChatPromptTemplate.from_messages([
    ("system", (
        "You optimize retrieval queries for semiconductor and ATE documentation. "
        "Return up to 4 short query rewrites, each on a new line, with no numbering."
    )),
    ("user", "{question}"),
])

_rerank_prompt = ChatPromptTemplate.from_messages([
    ("system", "You are a strict relevance ranker. Return valid JSON only."),
    ("user", "{prompt}"),
])

_web_query_prompt = ChatPromptTemplate.from_messages([
    ("system", (
        "You generate high-precision web search queries for electronics/ATE questions. "
        "Return up to 5 concise queries, one per line, no numbering."
    )),
    ("user", "Question: {question}\nType: {query_type} (O=public, U=unknown)\n"
             "Prioritize authoritative official pages, standards, and datasheets."),
])

_web_rerank_prompt = ChatPromptTemplate.from_messages([
    ("system", "You are a strict ranking engine for authoritative technical web sources. JSON only."),
    ("user", "{prompt}"),
])

try:
    from ddgs import DDGS

    HAS_DDGS = True
except ImportError:
    DDGS = None  # type: ignore[assignment]
    HAS_DDGS = False

_logger = logging.getLogger(__name__)
_WEB_OPEN_LOGGED_BLOCKS: set[tuple[str, int]] = set()


def _parse_json_object(raw: str) -> dict | None:
    """Best-effort JSON object parse for LLM outputs that may include wrappers."""
    if not raw:
        return None

    text = raw.strip()
    if not text:
        return None

    # Common case: markdown fenced JSON output.
    fenced = re.fullmatch(r"```(?:json)?\s*(.*?)\s*```", text, flags=re.DOTALL | re.IGNORECASE)
    if fenced:
        text = fenced.group(1).strip()

    try:
        payload = json.loads(text)
        return payload if isinstance(payload, dict) else None
    except json.JSONDecodeError:
        pass

    # Fallback: extract first JSON object from mixed text.
    match = re.search(r"\{[\s\S]*\}", text)
    if not match:
        return None

    try:
        payload = json.loads(match.group(0))
        return payload if isinstance(payload, dict) else None
    except json.JSONDecodeError:
        return None


def _dedupe_keep_order(values: list[str]) -> list[str]:
    out: list[str] = []
    seen: set[str] = set()
    for value in values:
        norm = value.strip()
        if not norm:
            continue
        key = norm.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(norm)
    return out


def expand_query(question: str, enable_llm: bool = True, max_variants: int = 4) -> list[str]:
    """Generate semantically diverse internal-retrieval query variants."""
    fallback = [
        question,
        f"{question} troubleshooting",
        f"{question} register settings",
        f"{question} testflow best practice",
    ]
    if not enable_llm:
        return _dedupe_keep_order([question])[: max(1, max_variants)]

    try:
        chain = _expand_prompt | get_chat_llm(OPENAI_RETRIEVAL_MODEL, temperature=0) | StrOutputParser()
        result = chain.invoke({"question": question})
        variants = [line.strip() for line in result.splitlines() if line.strip()]
        return _dedupe_keep_order([question, *variants])[: max(1, max_variants)]
    except Exception as exc:
        _logger.warning("expand_query failed, using fallback: %s", exc)
        return _dedupe_keep_order(fallback)[: max(1, max_variants)]


def mmr_deduplicate(candidates: list[dict], lambda_param: float = 0.7, top_k: int = 6) -> list[dict]:
    """MMR selection balancing relevance and diversity."""
    if len(candidates) <= top_k:
        return candidates

    selected: list[dict] = []
    remaining = candidates.copy()

    best = min(remaining, key=lambda x: x["distance"])
    selected.append(best)
    remaining.remove(best)

    while len(selected) < top_k and remaining:
        mmr_scores = []
        for cand in remaining:
            relevance = 1 - cand["distance"]
            redundancy = max(
                len(set(cand["text"].split()) & set(s["text"].split())) / max(len(set(cand["text"].split())), 1)
                for s in selected
            )
            mmr_scores.append(lambda_param * relevance - (1 - lambda_param) * redundancy)

        best_idx = mmr_scores.index(max(mmr_scores))
        selected.append(remaining[best_idx])
        remaining.pop(best_idx)

    return selected


def cross_encoder_rerank(question: str, candidates: list[dict], enable_llm: bool = True) -> list[dict]:
    """LLM reranker (0-10 relevance score per chunk)."""
    if len(candidates) <= 2 or not enable_llm:
        return candidates

    chunk_list = "\n\n".join(
        f"[CHUNK {i + 1}] (Source: {c['source']}, Page {c['page_number']})\n{c['text'][:600]}"
        for i, c in enumerate(candidates)
    )
    prompt = (
        f"Question: {question}\n\n"
        "Score each chunk relevance from 0-10 and return only JSON in this shape:\n"
        '{"scores": [{"chunk": 1, "score": 8}]}\n\n'
        f"{chunk_list}"
    )

    try:
        chain = _rerank_prompt | get_chat_llm(OPENAI_RETRIEVAL_MODEL, temperature=0) | StrOutputParser()
        raw = chain.invoke({"prompt": prompt})
        payload = _parse_json_object(raw)
        score_items = payload.get("scores", []) if payload else []

        idx_to_score: dict[int, int] = {}
        for row in score_items:
            if not isinstance(row, dict):
                continue
            idx = row.get("chunk")
            score = row.get("score")
            if isinstance(idx, int) and isinstance(score, int) and 1 <= idx <= len(candidates):
                idx_to_score[idx - 1] = max(0, min(10, score))

        if len(idx_to_score) >= max(2, len(candidates) // 2):
            scored = []
            for i, cand in enumerate(candidates):
                scored.append({**cand, "rerank_score": idx_to_score.get(i, 0)})
            return sorted(scored, key=lambda x: (x["rerank_score"], -x["distance"]), reverse=True)
    except Exception as exc:
        _logger.warning("cross_encoder_rerank fallback: %s", exc)

    return candidates


def query_documents(
    question: str,
    n_per_query: int = 5,
    mmr_lambda: float = 0.7,
    use_query_expansion: bool = True,
    use_rerank: bool = True,
    max_results: int = 5,
) -> list[dict]:
    """
    Retrieval steps:
    1. Query expansion
    2. Chroma vector search
    3. MMR dedupe
    4. LLM rerank
    """
    queries = expand_query(question, enable_llm=use_query_expansion, max_variants=4)
    candidate_map: dict[str, dict] = {}

    for query in queries:
        query_embedding = _lc_embeddings.embed_query(query)
        results = collection.query(
            query_embeddings=[query_embedding],
            n_results=min(n_per_query, max(collection.count(), 1)),
            include=["documents", "metadatas", "distances"],
        )
        for i, doc in enumerate(results["documents"][0]):
            chunk_id = results["ids"][0][i]
            distance = results["distances"][0][i]
            meta = results["metadatas"][0][i] or {}

            if chunk_id not in candidate_map or distance < candidate_map[chunk_id]["distance"]:
                candidate_map[chunk_id] = {
                    "text": doc,
                    "distance": distance,
                    "source": meta.get("source", "unknown"),
                    "file_type": meta.get("file_type", "general"),
                    "page_number": meta.get("page_number", 0),
                    "total_pages": meta.get("total_pages", 0),
                    "has_tables": meta.get("has_tables", "False") == "True",
                    "chunk_index": meta.get("chunk_index", 0),
                }

    candidates = [c for c in candidate_map.values() if c["distance"] < 0.65]
    candidates.sort(key=lambda x: x["distance"])
    diverse = mmr_deduplicate(candidates, lambda_param=mmr_lambda, top_k=8)
    reranked = cross_encoder_rerank(question, diverse, enable_llm=use_rerank)
    return reranked[: max(1, max_results)]


_PROPRIETARY_KW = (
    "teradyne",
    "ultraflex",
    "ig-xl",
    "igxl",
    "tester manual",
    "internal manual",
    "customer doc",
    "confidential",
    "restricted",
    "nda",
    "training doc",
    "proprietary",
    "advantest",
    "v93000",
    "tml",
    "tdc",
    "cookbook",
    "hib",
    "testsuite",
    "testflow",
    "test program",
    "application note internal",
)
_PUBLIC_KW = (
    "datasheet",
    "spec",
    "specification",
    "ieee",
    "standard",
    "typical",
    "features",
    "product brief",
    "public",
    "analog.com",
    "adin",
    "aduc",
    "amc",
    "voltage range",
    "pinout",
    "package",
    "ordering",
    "block diagram",
    "eval board",
    "weather",
    "temperature",
    "forecast",
)

_REALTIME_WEB_KW = (
    "current weather",
    "weather",
    "temperature now",
    "forecast",
    "right now",
    "currently",
    "today",
    "tomorrow",
    "live",
    "latest",
    "news today",
)

INTERNAL_DIST_THRESHOLD = 0.35
INTERNAL_MIN_RELEVANT = 2

_AUTHORITATIVE_DOMAINS = (
    "analog.com",
    "teradyne.com",
    "ieee.org",
    "ietf.org",
    "advantest.com",
    "ti.com",
    "nxp.com",
    "microchip.com",
    "onsemi.com",
    "infineon.com",
)


def classify_query_type(question: str) -> str:
    """Classify as P (proprietary), O (open/public), or U (unknown)."""
    q = question.lower()
    if any(kw in q for kw in _REALTIME_WEB_KW):
        return "O"
    p_hits = sum(1 for kw in _PROPRIETARY_KW if kw in q)
    o_hits = sum(1 for kw in _PUBLIC_KW if kw in q)
    if p_hits > o_hits and p_hits >= 1:
        return "P"
    if o_hits > p_hits and o_hits >= 1:
        return "O"
    return "U"


def requires_realtime_web(question: str) -> bool:
    """Return True for prompts that require fresh public information."""
    q = question.lower()
    return any(kw in q for kw in _REALTIME_WEB_KW)


def assess_internal_confidence(chunks: list[dict]) -> tuple[bool, bool]:
    """Returns (has_confidence, has_sufficiency)."""
    relevant = [c for c in chunks if c["distance"] < INTERNAL_DIST_THRESHOLD]
    has_confidence = len(relevant) >= INTERNAL_MIN_RELEVANT
    total_chars = sum(len(c["text"]) for c in relevant)
    has_sufficiency = has_confidence and total_chars > 200
    return has_confidence, has_sufficiency


def _tokenize(text: str) -> set[str]:
    return set(re.findall(r"[a-zA-Z0-9]{3,}", text.lower()))


def _normalize_url(url: str) -> str:
    parsed = urlparse(url.strip())
    scheme = parsed.scheme or "https"
    netloc = parsed.netloc.lower().removeprefix("www.")
    path = parsed.path.rstrip("/")
    clean_query_pairs = [(k, v) for k, v in parse_qsl(parsed.query, keep_blank_values=False) if not k.startswith("utm_")]
    query = urlencode(clean_query_pairs)
    return urlunparse((scheme, netloc, path, "", query, ""))


def _domain(url: str) -> str:
    return urlparse(url).netloc.lower().removeprefix("www.")


def _domain_boost(url: str) -> float:
    host = _domain(url)
    if any(host == dom or host.endswith(f".{dom}") for dom in _AUTHORITATIVE_DOMAINS):
        return 2.5
    if host.endswith(".edu") or host.endswith(".gov"):
        return 1.5
    return 0.0


def _lexical_relevance(query: str, title: str, snippet: str) -> float:
    q_tokens = _tokenize(query)
    if not q_tokens:
        return 0.0
    result_tokens = _tokenize(f"{title} {snippet}")
    overlap = len(q_tokens & result_tokens)
    return overlap / max(len(q_tokens), 1)


def _clean_html(html: str) -> str:
    text = re.sub(r"<script[^>]*>.*?</script>", " ", html, flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r"<style[^>]*>.*?</style>", " ", text, flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r"<noscript[^>]*>.*?</noscript>", " ", text, flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r"<[^>]+>", " ", text)
    text = unescape(text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def web_open(url: str, max_chars: int = 5000) -> str:
    """Fetch URL and convert HTML to plain text."""
    try:
        resp = httpx.get(
            url,
            timeout=12,
            follow_redirects=True,
            verify=False,
            headers={
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
                )
            },
        )
        resp.raise_for_status()
        text = _clean_html(resp.text)
        return text[:max_chars]
    except httpx.HTTPStatusError as exc:
        status = exc.response.status_code if exc.response is not None else 0
        host = _domain(url)
        key = (host, status)
        # 401/403/404 are common anti-bot or not-found outcomes; keep logs quiet.
        if status in {401, 403, 404}:
            if key not in _WEB_OPEN_LOGGED_BLOCKS:
                _WEB_OPEN_LOGGED_BLOCKS.add(key)
                _logger.info("web_open blocked for host=%s status=%s", host, status)
        else:
            _logger.warning("web_open failed for %s: %s", url, exc)
        return ""
    except Exception as exc:
        _logger.warning("web_open failed for %s: %s", url, exc)
        return ""


def build_web_queries(question: str, query_type: str) -> list[str]:
    """Generate focused web queries for public information retrieval."""
    fallback = [
        question,
        f"{question} official datasheet",
        f"{question} product brief",
        f"{question} application note",
    ]
    if query_type == "O":
        fallback.append(f"{question} site:analog.com OR site:ti.com OR site:nxp.com")

    try:
        chain = _web_query_prompt | get_chat_llm(OPENAI_WEB_MODEL, temperature=0) | StrOutputParser()
        result = chain.invoke({"question": question, "query_type": query_type})
        generated = [line.strip() for line in result.splitlines() if line.strip()]
        return _dedupe_keep_order([*generated, *fallback])[:5]
    except Exception as exc:
        _logger.warning("build_web_queries fallback: %s", exc)
        return _dedupe_keep_order(fallback)[:5]


def web_search(queries: list[str] | str, top_k: int = 8, per_query: int = 8) -> list[dict]:
    """Run multi-query DDG search and return deduped, pre-scored candidates."""
    if not HAS_DDGS:
        _logger.warning("duckduckgo-search not installed - web search disabled")
        return []

    query_list = [queries] if isinstance(queries, str) else queries
    merged: dict[str, dict] = {}

    try:
        with DDGS() as ddgs:
            for query in query_list:
                try:
                    raw = ddgs.text(query, max_results=per_query)
                except TypeError:
                    raw = ddgs.text(query)

                for rank, row in enumerate(raw or [], start=1):
                    url = str(row.get("href") or "").strip()
                    if not url:
                        continue

                    norm_url = _normalize_url(url)
                    title = str(row.get("title") or "").strip()
                    snippet = str(row.get("body") or "").strip()
                    lex_score = _lexical_relevance(query, title, snippet)
                    domain_score = _domain_boost(norm_url)
                    score = domain_score + (2.2 * lex_score) + (0.6 / rank)

                    existing = merged.get(norm_url)
                    if existing is None:
                        merged[norm_url] = {
                            "title": title,
                            "url": norm_url,
                            "snippet": snippet,
                            "content": "",
                            "score": score,
                            "lexical_score": lex_score,
                            "domain_score": domain_score,
                            "query_hits": 1,
                        }
                    else:
                        existing["score"] = max(existing["score"], score)
                        existing["query_hits"] += 1
                        existing["lexical_score"] = max(existing["lexical_score"], lex_score)
                        if len(snippet) > len(existing.get("snippet", "")):
                            existing["snippet"] = snippet
                        if len(title) > len(existing.get("title", "")):
                            existing["title"] = title
    except Exception as exc:
        _logger.warning("Web search failed: %s", exc)
        return []

    ranked = sorted(
        merged.values(),
        key=lambda x: (x["score"], x["query_hits"], x["domain_score"], len(x.get("snippet", ""))),
        reverse=True,
    )
    return ranked[: max(top_k, 1)]


def _llm_rerank_web_results(question: str, results: list[dict], top_k: int = 6) -> list[dict]:
    if len(results) <= 2:
        return results[:top_k]

    lines = []
    for i, row in enumerate(results, start=1):
        title = row.get("title", "")[:180]
        snippet = row.get("snippet", "")[:260]
        lines.append(f"[{i}] {title}\nURL: {row.get('url', '')}\nSnippet: {snippet}")

    prompt = (
        f"Question: {question}\n\n"
        "Rank relevance and trustworthiness for these web results. "
        "Return only JSON in this exact shape: "
        '{"scores": [{"idx": 1, "score": 9}]} where score is integer 0-10.\n\n'
        + "\n\n".join(lines)
    )

    try:
        chain = _web_rerank_prompt | get_chat_llm(OPENAI_WEB_MODEL, temperature=0) | StrOutputParser()
        raw = chain.invoke({"prompt": prompt})
        payload = _parse_json_object(raw)
        score_items = payload.get("scores", []) if payload else []
        scores: dict[int, int] = {}
        for row in score_items:
            if not isinstance(row, dict):
                continue
            idx = row.get("idx")
            score = row.get("score")
            if isinstance(idx, int) and isinstance(score, int) and 1 <= idx <= len(results):
                scores[idx - 1] = max(0, min(10, score))

        reranked = []
        for i, row in enumerate(results):
            reranked.append({**row, "llm_score": scores.get(i, 0)})
        reranked.sort(
            key=lambda x: (x.get("llm_score", 0), x.get("domain_score", 0.0), x.get("score", 0.0)),
            reverse=True,
        )
        return reranked[:top_k]
    except Exception as exc:
        _logger.warning("_llm_rerank_web_results fallback: %s", exc)
        return sorted(results, key=lambda x: x.get("score", 0.0), reverse=True)[:top_k]


def assess_web_confidence(results: list[dict]) -> bool:
    """Confidence when sources are authoritative or strongly relevant by rerank score."""
    if not results:
        return False

    authoritative = 0
    strong = 0
    for row in results:
        url = str(row.get("url", "")).lower()
        if any(dom in url for dom in _AUTHORITATIVE_DOMAINS):
            authoritative += 1
        if row.get("llm_score", 0) >= 6:
            strong += 1

    return authoritative >= 1 or strong >= 2


def _web_retrieve(question: str, query_type: str, top_k: int = 5) -> list[dict]:
    queries = build_web_queries(question, query_type)
    candidates = web_search(queries, top_k=max(top_k * 3, 10), per_query=8)
    if not candidates:
        return []

    reranked = _llm_rerank_web_results(question, candidates, top_k=max(top_k, 6))

    for row in reranked[:top_k]:
        row["content"] = web_open(row["url"])
        if not row["content"]:
            row["content"] = row.get("snippet", "")

    return reranked[:top_k]


def hybrid_query_documents(
    question: str,
    n_per_query: int = 5,
    enable_web: bool = False,  # legacy — prefer web_mode
    mmr_lambda: float = 0.7,
    web_mode: str = "fallback",
    fast_mode: bool = False,
) -> dict:
    """
    Retrieval decision policy controlled by *web_mode*:

    - "off"      : Internal KB only. Abstain if insufficient.
    - "fallback" : Internal first; web only when internal is insufficient
                   and topic is public/unknown (default behaviour).
    - "always"   : Run BOTH internal KB and web search regardless of confidence.
                   Best answers use both sources together.
    - "web_only" : Skip vector DB entirely — pure real-time web search for
                   any topic. Acts like a general-purpose search engine.
    """
    # Legacy shim: enable_web=True maps to "fallback" behaviour
    if enable_web and web_mode == "fallback":
        pass  # nothing to change, fallback is the web-enabled mode
    elif not enable_web and web_mode == "fallback":
        web_mode = "off"

    query_type = classify_query_type(question)
    needs_realtime_web = requires_realtime_web(question)
    # web_only treats everything as publicly searchable
    if web_mode == "web_only":
        query_type = "O"

    _logger.info(
        "[HYBRID] query_type=%s  web_mode=%s  realtime_web=%s  q=%r",
        query_type,
        web_mode,
        needs_realtime_web,
        question[:60],
    )

    # ── Internal KB retrieval (skipped in web_only) ──────────────────────────
    internal_chunks: list[dict] = []
    effective_n_per_query = max(1, int(n_per_query))
    if fast_mode:
        # Fast mode reduces LLM-heavy retrieval work while preserving RAG grounding.
        effective_n_per_query = min(effective_n_per_query, 3)

    if web_mode != "web_only":
        internal_chunks = query_documents(
            question,
            n_per_query=effective_n_per_query,
            mmr_lambda=mmr_lambda,
            use_query_expansion=not fast_mode,
            use_rerank=not fast_mode,
            max_results=4 if fast_mode else 5,
        )

    int_conf, int_suff = assess_internal_confidence(internal_chunks)
    _logger.info(
        "[HYBRID] internal: %d chunks  confident=%s  sufficient=%s",
        len(internal_chunks),
        int_conf,
        int_suff,
    )

    web_results: list[dict] = []
    web_conf = False
    abstain = False
    abstain_reason = ""

    # ── web_only: search the open web for anything ───────────────────────────
    if web_mode == "web_only":
        web_results = _web_retrieve(question, query_type="O", top_k=4 if fast_mode else 6)
        web_conf = assess_web_confidence(web_results)
        if not web_results:
            abstain = True
            abstain_reason = (
                "Web search returned no results for this query. "
                "Try rephrasing or switching to Fallback mode."
            )

    # ── always: internal + web in parallel ───────────────────────────────────
    elif web_mode == "always":
        eff_qtype = "O" if query_type == "P" else query_type
        web_results = _web_retrieve(question, query_type=eff_qtype, top_k=4 if fast_mode else 5)
        web_conf = assess_web_confidence(web_results)
        if not int_suff and not web_conf and not web_results:
            if query_type == "P":
                abstain = True
                abstain_reason = (
                    "This appears to be proprietary/internal content and "
                    "no useful public web sources were found. "
                    "Please upload the relevant document to the knowledge base."
                )
            else:
                abstain = True
                abstain_reason = (
                    "No sufficient information found in the internal KB or web sources."
                )

    # ── off: internal only ────────────────────────────────────────────────────
    elif web_mode == "off":
        if query_type == "P" and not int_conf:
            abstain = True
            abstain_reason = (
                "This appears to be proprietary/internal content. "
                "My internal knowledge base doesn't contain enough detail to answer reliably. "
                "If you have the relevant manual or document, please upload it so I can index it."
            )
        elif not int_suff:
            abstain = True
            abstain_reason = (
                "I couldn't find sufficient information in the internal knowledge base. "
                "Try enabling web search in the sidebar for public topics."
            )

    # ── fallback: original behaviour ─────────────────────────────────────────
    else:
        if query_type == "P":
            if not int_conf:
                abstain = True
                abstain_reason = (
                    "This appears to be proprietary/internal content. "
                    "My internal knowledge base doesn't contain enough detail to answer reliably. "
                    "If you have the relevant manual or document, please upload it so I can index it."
                )
        elif query_type == "O":
            if needs_realtime_web or not int_suff:
                web_results = _web_retrieve(question, query_type=query_type, top_k=4 if fast_mode else 5)
                web_conf = assess_web_confidence(web_results)
            if not int_suff and not web_conf:
                abstain = True
                abstain_reason = (
                    "I couldn't find verified information in the internal DB or public web sources."
                )
        else:
            if not int_suff:
                has_proprietary_hint = any(kw in question.lower() for kw in _PROPRIETARY_KW)
                if not has_proprietary_hint:
                    web_results = _web_retrieve(question, query_type=query_type, top_k=3 if fast_mode else 4)
                    web_conf = assess_web_confidence(web_results)

    _logger.info("[HYBRID] web: %d results  confident=%s  abstain=%s", len(web_results), web_conf, abstain)
    return {
        "query_type": query_type,
        "web_mode": web_mode,
        "internal_chunks": internal_chunks,
        "web_results": web_results,
        "internal_confident": int_conf,
        "internal_sufficient": int_suff,
        "web_confident": web_conf,
        "abstain": abstain,
        "abstain_reason": abstain_reason,
    }
