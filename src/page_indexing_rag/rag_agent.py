"""
RAG Agent using Claude Agent SDK for semiconductor test engineering.

This module provides the core agent that integrates with Claude Agent SDK
while preserving all domain-specific features for ADI semiconductor workflows.
"""

from __future__ import annotations

import hashlib
import asyncio
from pathlib import Path
from typing import List, Dict, Optional, Tuple, Any
from dataclasses import dataclass

import chromadb
import boto3
import numpy as np
from openai import APIConnectionError, APIError, APITimeoutError

# Import existing modules to reuse logic
from .ingestion_agentsdk import (
    classify_file,
    semantic_chunk,
    extract_pdf_pages,
    extract_html_document,
    extract_chm_documents
)
from .retrieval_agentsdk import (
    mmr_deduplicate,
    assess_internal_confidence,
    classify_query_type,
    _PROPRIETARY_KW,
    _PUBLIC_KW
)
from .generation_agentsdk import (
    build_file_context_prompt,
    classify_question_complexity,
    build_rag_context
)
from .config_agentsdk import (
    AWS_ACCESS_KEY_ID,
    AWS_REGION,
    AWS_SECRET_ACCESS_KEY,
    CHAT_BACKEND,
    COLLECTION_NAME,
    CLAUDE_MODEL,
    CLAUDE_QUERY_MODEL,
    EMBEDDING_BACKEND,
    EMBEDDING_MODEL,
    MASTER_SYSTEM_PROMPT,
    OPENAI_ANSWER_MODEL,
    OPENAI_EMBEDDING_MODEL,
    OPENAI_RETRIEVAL_MODEL,
    _resolve_dir,
    build_anthropic_client,
    build_openai_compatible_client,
)

AGENT_CONFIG_VERSION = 4


@dataclass
class QueryResponse:
    """Response from RAG query containing answer and metadata."""
    answer: str
    sources: List[Dict[str, Any]]
    web_sources: Optional[List[Dict[str, Any]]] = None
    query_type: str = "U"  # P=Proprietary, O=Open, U=Unknown
    complexity: str = "medium"
    model_used: str = OPENAI_ANSWER_MODEL
    confidence: Dict[str, bool] = None


class SemiconductorRAGAgent:
    """
    RAG Agent for semiconductor test engineering using Claude Agent SDK.

    This agent maintains all domain-specific features from the original
    Streamlit application while leveraging Claude Agent SDK capabilities.
    """

    def __init__(self, chroma_path: Optional[str] = None):
        """
        Initialize the RAG agent.

        Args:
            chroma_path: Path to ChromaDB storage. Defaults to data/chroma_persistent_storage
        """
        self.chat_backend = CHAT_BACKEND
        self.embedding_backend = EMBEDDING_BACKEND
        self.collection_name = COLLECTION_NAME

        if self.chat_backend == "openai_compatible":
            self.openai_client = build_openai_compatible_client()
            self.anthropic = None
            self.claude_model = OPENAI_ANSWER_MODEL
            self.query_model = OPENAI_RETRIEVAL_MODEL
        else:
            self.openai_client = None
            self.anthropic = build_anthropic_client()
            self.claude_model = CLAUDE_MODEL
            self.query_model = CLAUDE_QUERY_MODEL

        self.embedding_model = (
            EMBEDDING_MODEL if self.embedding_backend == "bedrock" else OPENAI_EMBEDDING_MODEL
        )
        self.embedding_dimensions = self._resolve_embedding_dimensions(self.embedding_model)
        self.config_version = AGENT_CONFIG_VERSION

        # Initialize AWS Bedrock only when Titan embeddings are actually configured.
        self.bedrock = None
        if self.embedding_backend == "bedrock":
            bedrock_kwargs = {"region_name": AWS_REGION}
            if AWS_ACCESS_KEY_ID and AWS_SECRET_ACCESS_KEY:
                bedrock_kwargs["aws_access_key_id"] = AWS_ACCESS_KEY_ID
                bedrock_kwargs["aws_secret_access_key"] = AWS_SECRET_ACCESS_KEY
            try:
                self.bedrock = boto3.client('bedrock-runtime', **bedrock_kwargs)
            except Exception as e:
                print(f"Warning: Could not initialize Bedrock client: {e}")
                self.bedrock = None

        # Initialize ChromaDB
        self.chroma_path = chroma_path or _resolve_dir(
            "data/chroma_persistent_storage",
            "chroma_persistent_storage"
        )
        self._init_chromadb()

        # System prompt for semiconductor domain
        self.system_prompt = MASTER_SYSTEM_PROMPT

        # File classifier categories
        self.file_types = [
            "tml", "tdc", "datasheet", "schematic",
            "pinmap", "rtl", "regmap", "tester_config", "general"
        ]

    def _init_chromadb(self):
        """Initialize ChromaDB collection with Titan embeddings."""
        try:
            self.chroma_client = chromadb.PersistentClient(path=str(self.chroma_path))
            # New collection for Titan embeddings
            self.collection = self.chroma_client.get_or_create_collection(
                name=self.collection_name,
                metadata={"hnsw:space": "cosine"}
            )
        except Exception as e:
            print(f"Warning: Could not initialize persistent ChromaDB: {e}")
            # Fallback to ephemeral client
            self.chroma_client = chromadb.EphemeralClient()
            self.collection = self.chroma_client.get_or_create_collection(
                name=self.collection_name,
                metadata={"hnsw:space": "cosine"}
            )

    async def expand_query_claude(self, query: str) -> List[str]:
        """
        Generate query variants using Claude for better retrieval.

        Args:
            query: Original query

        Returns:
            List of query variants
        """
        # Fallback if Claude call fails
        fallback = [
            query,
            f"{query} semiconductor",
            f"{query} ADI Analog Devices",
            f"{query} test validation"
        ]

        try:
            text = self._complete_text(
                model=self.query_model,  # Use a smaller model for speed/cost
                max_tokens=200,
                temperature=0,
                system="You optimize retrieval queries for semiconductor and ATE documentation. Return up to 4 short query rewrites, each on a new line, with no numbering.",
                user_prompt=query,
            )

            variants = [line.strip() for line in text.splitlines() if line.strip()]
            # Add original query if not present
            if query not in variants:
                variants.insert(0, query)
            return variants[:4]
        except Exception as e:
            print(f"Query expansion failed, using fallback: {e}")
            return fallback[:4]

    async def get_titan_embedding(self, text: str) -> List[float]:
        """
        Get embedding from Amazon Titan Embed V2 model with fallback to OpenAI.

        Args:
            text: Text to embed

        Returns:
            Embedding vector
        """
        import json

        # Validate input
        if not text or not text.strip():
            print("Warning: Empty text provided for embedding, using placeholder")
            text = "placeholder"

        # Try OpenAI first if configured for OpenAI backend
        if self.embedding_backend == "openai_compatible":
            try:
                return self._get_openai_embedding(text)
            except Exception as e:
                print(f"OpenAI embedding failed: {e}")
                # Return zero vector as fallback
                return [0.0] * self.embedding_dimensions

        # Try Titan embeddings
        if self.bedrock is not None:
            try:
                body = json.dumps({
                    "inputText": text[:8000],  # Titan has 8K token limit
                    "dimensions": 1024,  # Titan V2 supports 256, 512, or 1024
                    "normalize": True
                })

                response = self.bedrock.invoke_model(
                    modelId=self.embedding_model,
                    body=body,
                    contentType="application/json",
                    accept="application/json"
                )

                result = json.loads(response['body'].read())
                embedding = result.get('embedding')

                if embedding and len(embedding) > 0:
                    return embedding
            except Exception as e:
                print(f"Titan embedding failed: {e}")

        # Try OpenAI as fallback if available
        if self.openai_client:
            try:
                print("Falling back to OpenAI embeddings...")
                return self._get_openai_embedding(text)
            except Exception as e:
                print(f"OpenAI fallback also failed: {e}")

        # Last resort: return a random vector
        print("Warning: All embedding methods failed, using random vector")
        import random
        random.seed(hash(text) % 2**32)
        return [random.random() for _ in range(self.embedding_dimensions)]

    async def process_document(self, file_path: str, source_name: Optional[str] = None,
                              force_reingest: bool = False) -> Tuple[str, int]:
        """
        Process and ingest a document into the vector store.

        Args:
            file_path: Path to the document
            source_name: Optional override for source name
            force_reingest: If True, bypass duplicate check and re-ingest

        Returns:
            Tuple of (file_type, chunks_added)
        """
        file_path = Path(file_path)
        source_name = source_name or file_path.name

        # Check if already ingested (unless force_reingest is True)
        fingerprint = self._get_file_fingerprint(file_path)
        if not force_reingest:
            existing = self.collection.get(
                where={"fingerprint": fingerprint[:16]}
            )
            if existing and existing['ids']:
                print(f"Skipping {source_name}: Already in database (fingerprint: {fingerprint[:16]})")
                return "already_ingested", 0
        else:
            # If force re-ingestion, delete existing chunks first
            existing = self.collection.get(
                where={"source": source_name}
            )
            if existing and existing['ids']:
                print(f"Removing {len(existing['ids'])} existing chunks for {source_name}")
                self.collection.delete(ids=existing['ids'])

        # Read file content
        try:
            with open(file_path, 'rb') as f:
                content = f.read()
                if file_path.suffix.lower() != '.pdf':
                    content = content.decode('utf-8', errors='ignore')
        except Exception as e:
            print(f"Error reading file: {e}")
            return "error", 0

        # Classify file
        file_type = classify_file(source_name, content if isinstance(content, str) else "")

        # Extract chunks based on file type
        chunks = []
        metadata_list = []

        if file_path.suffix.lower() == '.pdf':
            pages = extract_pdf_pages(str(file_path))

            # Check if PDF has any content
            if not pages:
                print(f"Warning: {source_name} - No pages extracted from PDF")
                return "no_pages", 0

            total_text_length = sum(len(page.get('text', '')) for page in pages)
            if total_text_length < 50:
                print(f"Warning: {source_name} - PDF has minimal text content ({total_text_length} chars)")
                return "minimal_content", 0

            # Process each page
            for page in pages:
                page_text = page.get('text', '').strip()
                if not page_text:
                    print(f"Skipping empty page {page.get('page_number', '?')} in {source_name}")
                    continue

                # Try to extract meaningful chunks
                page_chunks = semantic_chunk(page_text)

                # If no chunks generated but page has text, create at least one chunk
                if not page_chunks and len(page_text) > 20:
                    page_chunks = [page_text[:2000]]  # Take first 2000 chars
                    print(f"Created fallback chunk for page {page.get('page_number', '?')}")

                for i, chunk in enumerate(page_chunks):
                    if len(chunk.strip()) < 10:  # Skip very small chunks
                        continue

                    chunks.append(chunk)
                    metadata_list.append({
                        "source": source_name,
                        "file_type": file_type,
                        "page_number": page['page_number'],
                        "total_pages": page['total_pages'],
                        "has_tables": str(page.get('has_tables', False)),
                        "chunk_index": i,
                        "fingerprint": fingerprint[:16]
                    })

        elif file_path.suffix.lower() in ['.html', '.xhtml']:
            doc_text = extract_html_document(str(file_path), source_name)
            doc_chunks = semantic_chunk(doc_text)
            for i, chunk in enumerate(doc_chunks):
                chunks.append(chunk)
                metadata_list.append({
                    "source": source_name,
                    "file_type": file_type,
                    "page_number": 0,
                    "total_pages": 1,
                    "has_tables": "False",
                    "chunk_index": i,
                    "fingerprint": fingerprint[:16]
                })

        elif file_path.suffix.lower() == '.chm':
            documents = extract_chm_documents(str(file_path), source_name)
            for doc in documents:
                doc_chunks = semantic_chunk(doc['text'])
                for i, chunk in enumerate(doc_chunks):
                    chunks.append(chunk)
                    metadata_list.append({
                        "source": doc['source'],
                        "file_type": file_type,
                        "page_number": 0,
                        "total_pages": 1,
                        "has_tables": "False",
                        "chunk_index": i,
                        "fingerprint": fingerprint[:16]
                    })

        else:
            # Plain text file
            text_chunks = semantic_chunk(content)
            for i, chunk in enumerate(text_chunks):
                chunks.append(chunk)
                metadata_list.append({
                    "source": source_name,
                    "file_type": file_type,
                    "page_number": 0,
                    "total_pages": 1,
                    "has_tables": "False",
                    "chunk_index": i,
                    "fingerprint": fingerprint[:16]
                })

        # Get embeddings and store in ChromaDB
        if chunks:
            print(f"Processing {len(chunks)} chunks for {source_name}")
            embeddings = []

            # Get embeddings with error handling
            for i, chunk in enumerate(chunks):
                try:
                    embedding = await self.get_titan_embedding(chunk)
                    if embedding is None or len(embedding) == 0:
                        print(f"Warning: Empty embedding for chunk {i} of {source_name}")
                        continue
                    embeddings.append(embedding)
                except Exception as e:
                    print(f"Error getting embedding for chunk {i} of {source_name}: {e}")
                    # Try fallback to OpenAI if available
                    if self.embedding_backend != "bedrock" and self.openai_client:
                        try:
                            embedding = self._get_openai_embedding(chunk)
                            embeddings.append(embedding)
                        except Exception as e2:
                            print(f"Fallback embedding also failed: {e2}")
                            continue

            if len(embeddings) != len(chunks):
                print(f"Warning: Only got {len(embeddings)} embeddings for {len(chunks)} chunks")
                # Adjust chunks and metadata to match embeddings
                chunks = chunks[:len(embeddings)]
                metadata_list = metadata_list[:len(embeddings)]

            if not embeddings:
                print(f"Error: No embeddings generated for {source_name}")
                return "embedding_failed", 0

            # Generate unique IDs
            ids = [f"{fingerprint[:8]}_{i:04d}" for i in range(len(chunks))]

            # Add to collection with error handling
            try:
                self.collection.add(
                    ids=ids,
                    embeddings=embeddings,
                    documents=chunks,
                    metadatas=metadata_list
                )
                print(f"Successfully added {len(chunks)} chunks for {source_name}")
                return file_type, len(chunks)
            except Exception as e:
                print(f"Error adding to ChromaDB: {e}")
                return "storage_failed", 0

        else:
            print(f"No chunks generated for {source_name}")
            return "no_chunks", 0

    def _get_file_fingerprint(self, file_path: Path) -> str:
        """Generate SHA-256 fingerprint of file."""
        sha256_hash = hashlib.sha256()
        with open(file_path, "rb") as f:
            # Read first 64KB for fingerprint
            chunk = f.read(65536)
            sha256_hash.update(chunk)
        return sha256_hash.hexdigest()

    async def vector_search(self, query: str, n_results: int = 5) -> List[Dict[str, Any]]:
        """
        Perform vector similarity search on the knowledge base.

        Args:
            query: Search query
            n_results: Number of results per query variant

        Returns:
            List of relevant chunks with metadata
        """
        # Expand query for better retrieval
        query_variants = await self.expand_query_claude(query)

        # Search for each variant
        all_results = []
        for variant in query_variants:
            # Get embedding for query
            query_embedding = await self.get_titan_embedding(variant)

            # Search in ChromaDB
            try:
                results = self.collection.query(
                    query_embeddings=[query_embedding],
                    n_results=n_results
                )
            except Exception as e:
                print(f"Vector search failed for query variant '{variant}': {e}")
                continue

            # Process results
            if not results.get('ids') or not results['ids'] or not results['ids'][0]:
                continue
            for i in range(len(results['ids'][0])):
                chunk_data = {
                    'text': results['documents'][0][i],
                    'distance': results['distances'][0][i],
                    'metadata': results['metadatas'][0][i],
                    'source': results['metadatas'][0][i]['source'],
                    'file_type': results['metadatas'][0][i]['file_type'],
                    'page_number': results['metadatas'][0][i].get('page_number', 0),
                    'total_pages': results['metadatas'][0][i].get('total_pages', 1),
                    'has_tables': results['metadatas'][0][i].get('has_tables', 'False')
                }
                all_results.append(chunk_data)

        # Deduplicate and rerank
        deduped = mmr_deduplicate(all_results)

        # Note: cross_encoder_rerank uses OpenAI/Portkey from config
        # For now, skip reranking and just use MMR results
        # TODO: Implement Claude-based reranking
        reranked = sorted(deduped, key=lambda x: x['distance'])[:n_results]

        return reranked

    async def web_search(self, query: str, query_type: str) -> List[Dict[str, Any]]:
        """
        Perform web search for public/open queries.

        Since we're using Agent SDK as a library (not as a full agent),
        we'll use a simple placeholder for web search.

        Args:
            query: Search query
            query_type: P=Proprietary, O=Open, U=Unknown

        Returns:
            List of web search results
        """
        # Only search for public/open queries
        if query_type == "P":
            return []

        web_results = []

        # Simulate web search results for common queries
        # This is a placeholder - in production, integrate with real web search API
        # or use Claude Agent SDK's WebSearch tool when running as full agent

        if "datasheet" in query.lower() or "specification" in query.lower():
            web_results.append({
                "title": "ADI Product Documentation",
                "url": "https://analog.com/documentation",
                "content": "Analog Devices provides comprehensive datasheets and specifications for all products.",
                "domain": "analog.com"
            })

        if "ieee" in query.lower() or "standard" in query.lower():
            web_results.append({
                "title": "IEEE Standards",
                "url": "https://ieee.org/standards",
                "content": "IEEE standards define protocols and specifications for electronic systems.",
                "domain": "ieee.org"
            })

        if "v93000" in query.lower() or "advantest" in query.lower():
            web_results.append({
                "title": "Advantest V93000 Platform",
                "url": "https://advantest.com/v93000",
                "content": "The V93000 is a scalable ATE platform for semiconductor test.",
                "domain": "advantest.com"
            })

        return web_results

    async def process_query(
        self,
        query: str,
        n_results: int = 5,
        mmr_lambda: float = 0.7,
        enable_web: bool = True,
        uploaded_file_context: Optional[str] = None
    ) -> QueryResponse:
        """
        Process a user query through the RAG pipeline.

        Args:
            query: User's question
            n_results: Number of chunks to retrieve
            mmr_lambda: MMR diversity parameter
            enable_web: Whether to enable web fallback
            uploaded_file_context: Optional context from uploaded file

        Returns:
            QueryResponse with answer and metadata
        """
        # Classify query type
        query_type = classify_query_type(query)

        # Vector search
        internal_results = await self.vector_search(query, n_results)

        # Assess confidence
        confidence = assess_internal_confidence(internal_results)

        # Determine if web search is needed
        web_results = []
        if enable_web and query_type in ["O", "U"] and not confidence["has_sufficiency"]:
            web_results = await self.web_search(query, query_type)

        # Check for abstention
        if query_type == "P" and not confidence["has_confidence"]:
            return QueryResponse(
                answer="I cannot answer this proprietary question without sufficient internal documentation. Please upload relevant TML, TDC, or technical documentation.",
                sources=[],
                query_type=query_type,
                confidence=confidence
            )

        # Build context
        rag_context = build_rag_context(internal_results)

        # Build prompt
        user_prompt = f"QUESTION: {query}\n\n"
        if uploaded_file_context:
            user_prompt += f"DIRECTLY UPLOADED FILE CONTEXT:\n{uploaded_file_context}\n\n"
        user_prompt += f"RETRIEVED KNOWLEDGE BASE CONTEXT:\n{rag_context}"

        if web_results:
            web_context = "\n\n".join([
                f"[Web: {r['url']}]\n{r['content'][:1500]}"
                for r in web_results[:5]
            ])
            user_prompt += f"\n\nPUBLIC WEB SOURCES:\n{web_context}"

        # Classify complexity
        complexity = classify_question_complexity(query)

        # Generate response using Claude
        answer = self._complete_text(
            model=self.claude_model,
            max_tokens=2000,
            temperature=0.1,
            system=self.system_prompt,
            user_prompt=user_prompt,
        )

        return QueryResponse(
            answer=answer,
            sources=internal_results,
            web_sources=web_results if web_results else None,
            query_type=query_type,
            complexity=complexity,
            model_used=self.claude_model,
            confidence=confidence
        )

    def _complete_text(
        self,
        *,
        model: str,
        max_tokens: int,
        temperature: float,
        system: str,
        user_prompt: str,
    ) -> str:
        """Generate text using either Portkey/OpenAI-compatible chat or Anthropic."""
        if self.chat_backend == "openai_compatible":
            try:
                response = self.openai_client.chat.completions.create(
                    model=model,
                    messages=[
                        {"role": "system", "content": system},
                        {"role": "user", "content": user_prompt},
                    ],
                    max_tokens=max_tokens,
                    temperature=temperature,
                )
                return response.choices[0].message.content or ""
            except (APIConnectionError, APITimeoutError) as e:
                raise RuntimeError(
                    "Could not reach the Portkey/OpenAI-compatible backend. "
                    "The app now ignores the broken OS proxy settings by default; "
                    "if this still fails, check VPN/network access to https://api.portkey.ai "
                    "or set PORTKEY_BASE_URL in .env."
                ) from e
            except APIError as e:
                raise RuntimeError(f"Portkey/OpenAI-compatible backend error: {e}") from e

        response = self.anthropic.messages.create(
            model=model,
            max_tokens=max_tokens,
            temperature=temperature,
            system=system,
            messages=[{"role": "user", "content": user_prompt}],
            extra_headers=self._anthropic_auth_headers(),
        )
        return response.content[0].text

    def _get_openai_embedding(self, text: str) -> List[float]:
        """Get an embedding through Portkey/OpenAI-compatible embeddings."""
        try:
            response = self.openai_client.embeddings.create(
                model=self.embedding_model,
                input=text[:8000],
            )
            return list(response.data[0].embedding)
        except Exception as e:
            print(f"Error getting OpenAI-compatible embedding: {e}")
            return [0.0] * self.embedding_dimensions

    def _resolve_embedding_dimensions(self, model: str) -> int:
        model_lower = model.lower()
        if self.embedding_backend == "bedrock" or "titan" in model_lower:
            return 1024
        if "small" in model_lower or "ada-002" in model_lower:
            return 1536
        if "gemini" in model_lower:
            return 768
        return 3072

    def _anthropic_auth_headers(self) -> Dict[str, str]:
        """
        Ensure Anthropic SDK request validation sees auth on every request.

        Streamlit can keep a previously constructed client in session state
        across reruns, so we defensively include the active client's auth header
        at call time as well as at client construction time.
        """
        api_key = getattr(self.anthropic, "api_key", None)
        auth_token = getattr(self.anthropic, "auth_token", None)
        if api_key:
            return {"X-Api-Key": api_key}
        if auth_token:
            return {"Authorization": f"Bearer {auth_token}"}
        return {}

    def get_collection_stats(self) -> Dict[str, Any]:
        """Get statistics about the knowledge base."""
        try:
            count = self.collection.count()
            # Get sample to analyze sources
            if count > 0:
                sample = self.collection.get(limit=min(100, count))
                sources = set()
                file_types = []
                for metadata in sample['metadatas']:
                    sources.add(metadata.get('source', 'unknown'))
                    file_types.append(metadata.get('file_type', 'general'))

                return {
                    'total_chunks': count,
                    'sources': list(sources)[:10],  # Top 10 sources
                    'file_types': dict(zip(*np.unique(file_types, return_counts=True)))
                }
            return {'total_chunks': 0, 'sources': [], 'file_types': {}}
        except Exception as e:
            print(f"Error getting collection stats: {e}")
            return {'total_chunks': 0, 'sources': [], 'file_types': {}}
