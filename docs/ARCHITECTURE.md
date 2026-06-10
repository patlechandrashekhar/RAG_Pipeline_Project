# Architecture Notes

## Current

- `app/streamlit_app.py` handles Streamlit UI and orchestration only.
- `src/page_indexing_rag/config.py` owns environment, OpenAI client, and data path resolution.
- `src/page_indexing_rag/ingestion.py` owns file classification, PDF/HTML(CHM) extraction, chunking, and Chroma ingestion.
- `src/page_indexing_rag/retrieval.py` owns query expansion, vector retrieval, reranking, and web fallback.
- `src/page_indexing_rag/generation.py` owns context assembly and answer generation.

## Planned Agentic Split

1. `planner` agent for task decomposition.
2. `retriever` agent for query expansion, vector search, and reranking.
3. `web` agent for external retrieval when allowed.
4. `evidence_judge` for confidence and grounding checks.
5. `answer_writer` for final response formatting/citations.
