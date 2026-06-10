"""
Retrieval functions for Agent SDK version.

Standalone retrieval functions that don't depend on the old config.
"""

from typing import List, Dict, Any, Tuple
import numpy as np


# Proprietary keywords
_PROPRIETARY_KW = [
    "teradyne", "ultraflex", "igxl", "advantest", "v93000", "tml", "tdc",
    "cookbook", "hib", "testsuite", "testflow", "test program", "test flow",
    "nda", "confidential", "restricted", "internal manual", "internal only",
    "application note internal", "customer confidential", "customer doc",
    "customer document", "training doc", "training document", "proprietary"
]

# Public keywords
_PUBLIC_KW = [
    "datasheet", "data sheet", "spec", "specification", "ieee", "ietf",
    "standard", "features", "product brief", "eval board", "evaluation",
    "analog.com", "adin", "aduc", "amc", "pinout", "pin out", "package",
    "voltage range", "operating range", "thermal", "ordering", "block diagram"
]


def classify_query_type(question: str) -> str:
    """
    Classify query as Proprietary (P), Open (O), or Unknown (U).

    Args:
        question: User's query

    Returns:
        Query type: "P", "O", or "U"
    """
    q_lower = question.lower()

    # Check for proprietary keywords
    proprietary_count = sum(1 for kw in _PROPRIETARY_KW if kw in q_lower)

    # Check for public keywords
    public_count = sum(1 for kw in _PUBLIC_KW if kw in q_lower)

    # Decision logic
    if proprietary_count > public_count:
        return "P"  # Proprietary
    elif public_count > proprietary_count:
        return "O"  # Open/Public
    else:
        return "U"  # Unknown


def assess_internal_confidence(chunks: List[Dict[str, Any]]) -> Dict[str, bool]:
    """
    Assess confidence in internal retrieval results.

    Args:
        chunks: Retrieved chunks with distance scores

    Returns:
        Dictionary with confidence indicators
    """
    INTERNAL_DIST_THRESHOLD = 0.35
    INTERNAL_MIN_RELEVANT = 2

    if not chunks:
        return {"has_confidence": False, "has_sufficiency": False}

    # Count relevant chunks (distance < threshold)
    relevant_chunks = [c for c in chunks if c.get('distance', 1.0) < INTERNAL_DIST_THRESHOLD]
    has_confidence = len(relevant_chunks) >= INTERNAL_MIN_RELEVANT

    # Check total content
    total_chars = sum(len(c.get('text', '')) for c in chunks)
    has_sufficiency = has_confidence and total_chars > 200

    return {
        "has_confidence": has_confidence,
        "has_sufficiency": has_sufficiency
    }


def mmr_deduplicate(
    candidates: List[Dict[str, Any]],
    lambda_param: float = 0.7,
    top_k: int = 8
) -> List[Dict[str, Any]]:
    """
    Maximal Marginal Relevance deduplication.

    Args:
        candidates: List of candidate chunks
        lambda_param: Balance between relevance and diversity (0=diverse, 1=relevant)
        top_k: Number of results to return

    Returns:
        Deduplicated list of chunks
    """
    if not candidates:
        return []

    # If we have fewer candidates than requested, return all
    if len(candidates) <= top_k:
        return candidates

    # Sort by distance (lower is better)
    candidates = sorted(candidates, key=lambda x: x.get('distance', 1.0))

    # Initialize with best candidate
    selected = [candidates[0]]
    remaining = candidates[1:]

    # MMR selection
    while len(selected) < top_k and remaining:
        mmr_scores = []

        for candidate in remaining:
            # Relevance score (inverse of distance)
            relevance = 1.0 - candidate.get('distance', 1.0)

            # Diversity score (minimum similarity to selected items)
            # Using token overlap as similarity metric
            cand_tokens = set(candidate.get('text', '').lower().split())
            max_sim = 0.0

            for selected_item in selected:
                sel_tokens = set(selected_item.get('text', '').lower().split())
                if cand_tokens and sel_tokens:
                    overlap = len(cand_tokens & sel_tokens)
                    union = len(cand_tokens | sel_tokens)
                    similarity = overlap / union if union > 0 else 0.0
                    max_sim = max(max_sim, similarity)

            diversity = 1.0 - max_sim

            # MMR score
            mmr = lambda_param * relevance + (1 - lambda_param) * diversity
            mmr_scores.append((candidate, mmr))

        # Select best MMR score
        if mmr_scores:
            mmr_scores.sort(key=lambda x: x[1], reverse=True)
            best_candidate = mmr_scores[0][0]
            selected.append(best_candidate)
            remaining.remove(best_candidate)

    return selected