"""Prompt-building and final answer generation."""

from __future__ import annotations

from langchain_core.output_parsers import StrOutputParser
from langchain_core.prompts import ChatPromptTemplate

from .config import MASTER_SYSTEM_PROMPT, OPENAI_ANSWER_MODEL, get_chat_llm

_rag_prompt = ChatPromptTemplate.from_messages([
    ("system", "{system_prompt}"),
    ("user", "{user_prompt}"),
])

QUICK_ANSWER_ADDENDUM = """

--- QUICK ANSWER MODE ---
Respond with a concise technical answer optimized for IDE workflows.
- Keep it short and direct by default.
- Use compact bullet points for specs or steps.
- Include only the most relevant citations.
- Avoid long narrative unless the user explicitly asks for depth.
"""

DOCUMENT_MODE_ADDENDUM = """

--- DOCUMENT GENERATION MODE ---
The user explicitly requested a formal document. Your output is parsed into DOCX/PDF,
so write in strict, clean markdown with high readability and engineering rigor.

MANDATORY OUTPUT SHAPE:
1. First line must be exactly one H1 title: `# Document Title`
2. Then use `##` for major sections and `###` for subsections.
3. Do not use bold text as fake headings.
4. Do not include preamble text like "Here is the report".

CLAUDE-CODE STYLE WRITING:
- Be precise, structured, and implementation-focused.
- Prefer short paragraphs, then lists, then tables where useful.
- Separate facts, assumptions, and recommendations clearly.
- For procedures, provide deterministic numbered steps.

REQUIRED SECTIONS (in order):
- `## Executive Summary` (2-4 sentences)
- `## Technical Details`
- `## Recommendations`
- `## Key Considerations`
- `## Summary`

LISTING RULES:
- Use bullet lists for capabilities, constraints, and checks.
- Use numbered lists for procedures or ordered actions.
- Keep list items specific and non-overlapping.

TABLE RULES:
- Insert a markdown table whenever the content has repeated fields across 2+ items,
    or when comparing options, specs, parameters, limits, error codes, or trade-offs.
- Use this exact markdown structure:
    | Column A | Column B | Column C |
    |---|---|---|
    | value | value | value |
- Do not describe tabular data only in prose when a table would be clearer.

FORMAT RULES:
- Use `inline code` for pins, signals, registers, commands, literals, and config keys.
- Use fenced code blocks with language tags for multi-line code.
- Use callouts for critical guidance:
    - `> WARNING: ...`
    - `> NOTE: ...`
    - `> TIP: ...`

QUALITY BAR:
- No placeholders.
- No duplicated sections.
- Keep output directly usable as an engineering document.
"""

CODE_AGENT_MODE_ADDENDUM = """

--- CODE AGENT MODE ---
You are operating as a code-debug agent for Teradyne UltraFlex and IG-XL test programs.
The user has attached one or more project directories and expects concrete debugging help.

OUTPUT RULES:
1. Start with a short diagnosis summary (2-4 bullets).
2. List likely root causes with confidence labels: High / Medium / Low.
3. Reference file paths exactly as provided in the workspace context.
4. When proposing fixes, provide code edits in fenced code blocks.
5. If details are missing, state assumptions clearly before giving a patch.
6. End with a focused validation checklist (what to run/check in IG-XL/TML).

IMPORTANT:
- Prefer actionable fixes over theory.
- Do not invent file contents not present in the provided workspace snippets.
- If the provided context is insufficient, ask for exactly what additional file(s) are needed.
"""


def build_file_context_prompt(filename: str, content: str, file_type: str) -> str:
    headers = {
        "test_program": f"=== TEST PROGRAM: {filename} ===\nType: UltraFlex IG-XL Test Program\n",
        "tml": f"=== TEST PROGRAM: {filename} ===\nType: Test Program\n",  # Legacy support
        "rtl": f"=== RTL FILE: {filename} ===\nType: SystemVerilog/Verilog\n",
        "regmap": f"=== REGISTER MAP: {filename} ===\nType: Device Register Config\n",
        "pinmap": f"=== PIN MAP: {filename} ===\nType: Device Pin Mapping\n",
        "tester_config": f"=== TESTER CONFIG: {filename} ===\nType: UltraFlex Instrument Config\n",
        "datasheet": f"=== DATASHEET: {filename} ===\nType: Device Specification\n",
        "tdc": f"=== TDC DOCUMENT: {filename} ===\nType: Test Development Cookbook\n",
        "general": f"=== DOCUMENT: {filename} ===\n",
    }
    header = headers.get(file_type, headers["general"])
    return header + f"CONTENT:\n{content}\n"


def classify_question_complexity(question: str) -> str:
    q = question.lower()
    complexity_tokens = (
        "why",
        "debug",
        "optimize",
        "tradeoff",
        "root cause",
        "architecture",
        "timing",
        "reliability",
        "equation",
        "derive",
    )
    medium_tokens = (
        "configure",
        "setup",
        "how do i",
        "how to",
        "steps",
        "procedure",
    )
    if len(question) > 200 or any(token in q for token in complexity_tokens):
        return "complex"
    if any(token in q for token in medium_tokens):
        return "medium"
    if len(question) > 80:
        return "medium"
    return "simple"


def route_model(complexity: str) -> str:
    """Get the model to use, checking session state first."""
    _ = complexity
    # Check if user selected a model in the UI
    try:
        import streamlit as st
        if "selected_model" in st.session_state:
            return st.session_state["selected_model"]
    except Exception:
        pass
    # Fall back to default from config
    return OPENAI_ANSWER_MODEL


def build_rag_context(top_chunks: list[dict]) -> str:
    """Format retrieved chunks with source/page provenance."""
    parts = []
    for rank, chunk in enumerate(top_chunks, start=1):
        page_info = f"Page {chunk['page_number']}/{chunk['total_pages']}" if chunk["page_number"] else "N/A"
        table_tag = " [CONTAINS TABLE]" if chunk["has_tables"] else ""
        header = (
            f"--- [Rank {rank}] Source: {chunk['source']} | {page_info}{table_tag} "
            f"| Type: {chunk['file_type']} ---"
        )
        parts.append(f"{header}\n{chunk['text']}")
    return "\n\n".join(parts)


def generate_response(question: str, top_chunks: list[dict], uploaded_file_context: str = "") -> tuple[str, str, str]:
    rag_context = build_rag_context(top_chunks)
    user_prompt = f"QUESTION: {question}\n\n"
    if uploaded_file_context:
        user_prompt += f"DIRECTLY UPLOADED FILE CONTEXT:\n{uploaded_file_context}\n\n"
    user_prompt += f"RETRIEVED KNOWLEDGE BASE CONTEXT:\n{rag_context}"

    complexity = classify_question_complexity(question)
    model = route_model(complexity)
    chain = _rag_prompt | get_chat_llm(model) | StrOutputParser()
    result = chain.invoke({"system_prompt": MASTER_SYSTEM_PROMPT, "user_prompt": user_prompt})
    return result, model, complexity


def _build_web_context(web_results: list[dict]) -> str:
    parts = []
    for i, r in enumerate(web_results, start=1):
        snippet = r.get("content") or r.get("snippet", "")
        if not snippet:
            continue
        parts.append(f"--- [Web {i}] {r.get('title', 'Untitled')} | URL: {r['url']} ---\n{snippet[:1500]}")
    return "\n\n".join(parts)


def generate_hybrid_response(
    question: str,
    hybrid: dict,
    uploaded_file_context: str = "",
    document_mode: bool = False,
    concise_mode: bool = False,
) -> tuple[str, str, str]:
    """Generate answer from hybrid retrieval output."""
    if hybrid["abstain"]:
        return hybrid["abstain_reason"], "-", "abstain"

    internal_ctx = build_rag_context(hybrid["internal_chunks"])
    web_ctx = _build_web_context(hybrid["web_results"])

    user_prompt = f"QUESTION: {question}\n\n"
    if uploaded_file_context:
        user_prompt += f"DIRECTLY UPLOADED FILE CONTEXT:\n{uploaded_file_context}\n\n"
    if internal_ctx:
        user_prompt += f"=== INTERNAL KNOWLEDGE BASE CONTEXT ===\n{internal_ctx}\n\n"
    if web_ctx:
        user_prompt += f"=== PUBLIC WEB SOURCES ===\n{web_ctx}\n\n"
    if not internal_ctx and not web_ctx:
        user_prompt += "(No retrieved context available.)\n\n"

    complexity = classify_question_complexity(question)
    model = route_model(complexity)
    system = MASTER_SYSTEM_PROMPT + DOCUMENT_MODE_ADDENDUM if document_mode else MASTER_SYSTEM_PROMPT
    if concise_mode and not document_mode:
        system += QUICK_ANSWER_ADDENDUM
    chain = _rag_prompt | get_chat_llm(model) | StrOutputParser()
    result = chain.invoke({"system_prompt": system, "user_prompt": user_prompt})
    return result, model, complexity


def generate_code_agent_response(
    question: str,
    workspace_context: str,
    debug_goal: str = "Debug and improve UltraFlex test program code.",
) -> tuple[str, str, str]:
    """Generate a coding-agent style response using attached directory context."""
    user_prompt = (
        f"DEBUG GOAL: {debug_goal}\n\n"
        f"QUESTION: {question}\n\n"
        f"=== ATTACHED WORKSPACE CONTEXT ===\n{workspace_context}\n"
    )
    complexity = classify_question_complexity(question)
    model = route_model(complexity)
    system = MASTER_SYSTEM_PROMPT + CODE_AGENT_MODE_ADDENDUM
    chain = _rag_prompt | get_chat_llm(model) | StrOutputParser()
    result = chain.invoke({"system_prompt": system, "user_prompt": user_prompt})
    return result, model, complexity


def generate_code_agent_hybrid_response(
    question: str,
    workspace_context: str,
    hybrid: dict,
    debug_goal: str = "Debug and improve UltraFlex test program code.",
) -> tuple[str, str, str]:
    """Generate code-agent response using workspace + internal KB/web retrieval."""
    if hybrid.get("abstain") and not workspace_context:
        return hybrid.get("abstain_reason", "No context available."), "-", "abstain"

    internal_ctx = build_rag_context(hybrid.get("internal_chunks", []))
    web_ctx = _build_web_context(hybrid.get("web_results", []))

    user_prompt = (
        f"DEBUG GOAL: {debug_goal}\n\n"
        f"QUESTION: {question}\n\n"
    )
    if workspace_context:
        user_prompt += f"=== ATTACHED WORKSPACE CONTEXT ===\n{workspace_context}\n\n"
    if internal_ctx:
        user_prompt += f"=== INTERNAL KNOWLEDGE BASE CONTEXT ===\n{internal_ctx}\n\n"
    if web_ctx:
        user_prompt += f"=== PUBLIC WEB SOURCES ===\n{web_ctx}\n\n"
    if not workspace_context and not internal_ctx and not web_ctx:
        user_prompt += "(No workspace or retrieval context available.)\n"

    complexity = classify_question_complexity(question)
    model = route_model(complexity)
    system = MASTER_SYSTEM_PROMPT + CODE_AGENT_MODE_ADDENDUM
    chain = _rag_prompt | get_chat_llm(model) | StrOutputParser()
    result = chain.invoke({"system_prompt": system, "user_prompt": user_prompt})
    return result, model, complexity
