"""
Generation functions for Agent SDK version.

Standalone generation functions that don't depend on the old config.
"""

from typing import List, Dict, Any


def build_file_context_prompt(filename: str, content: str, file_type: str) -> str:
    """
    Build a context prompt for an uploaded file.

    Args:
        filename: Name of the file
        content: File content
        file_type: Classification of the file

    Returns:
        Formatted context prompt
    """
    headers = {
        "tml": f"=== TML TEST PROGRAM: {filename} ===\nType: Advantest TML\n",
        "rtl": f"=== RTL FILE: {filename} ===\nType: SystemVerilog/Verilog\n",
        "regmap": f"=== REGISTER MAP: {filename} ===\nType: Device Register Config\n",
        "pinmap": f"=== PIN MAP: {filename} ===\nType: Device Pin Mapping\n",
        "tester_config": f"=== TESTER CONFIG: {filename} ===\nType: V93000 Instrument Config\n",
        "datasheet": f"=== DATASHEET: {filename} ===\nType: Device Specification\n",
        "tdc": f"=== TDC DOCUMENT: {filename} ===\nType: Test Development Cookbook\n",
        "schematic": f"=== SCHEMATIC: {filename} ===\nType: Hardware Interconnect\n",
        "general": f"=== DOCUMENT: {filename} ===\n",
    }
    header = headers.get(file_type, headers["general"])
    return header + f"CONTENT:\n{content}\n"


def classify_question_complexity(question: str) -> str:
    """
    Classify question complexity for model routing.

    Args:
        question: User's question

    Returns:
        Complexity level: "simple", "medium", or "complex"
    """
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


def build_rag_context(top_chunks: List[Dict[str, Any]]) -> str:
    """
    Format retrieved chunks with source/page provenance.

    Args:
        top_chunks: List of retrieved chunks with metadata

    Returns:
        Formatted RAG context string
    """
    if not top_chunks:
        return "No relevant context found in knowledge base."

    parts = []
    for rank, chunk in enumerate(top_chunks, start=1):
        # Get metadata
        source = chunk.get('source', 'Unknown')
        file_type = chunk.get('file_type', 'general')
        page_number = chunk.get('page_number', 0)
        total_pages = chunk.get('total_pages', 0)
        has_tables = chunk.get('has_tables', False)
        text = chunk.get('text', '')[:600]  # Limit chunk text

        # Format page info
        if page_number and total_pages:
            page_info = f"Page {page_number}/{total_pages}"
        elif page_number:
            page_info = f"Page {page_number}"
        else:
            page_info = "N/A"

        # Add table indicator
        table_tag = " [CONTAINS TABLE]" if has_tables else ""

        # Build header
        header = (
            f"--- [Rank {rank}] Source: {source} | {page_info}{table_tag} "
            f"| Type: {file_type} ---"
        )

        parts.append(f"{header}\n{text}")

    return "\n\n".join(parts)
