"""
Enhanced RAG Agent with Optional Agent SDK Layer

This module implements a two-step architecture:
1. Core RAG retrieval (always active)
2. Optional Agent SDK enhancement

The key innovation is that Agent SDK is not a replacement but an enhancement
that takes RAG context as input and enriches the response with additional
tool capabilities.
"""

import os
import asyncio
import logging
from typing import Optional, List, Dict, Any, Tuple
from dataclasses import dataclass, field
from enum import Enum
import hashlib
import json
from datetime import datetime

import chromadb
from anthropic import AsyncAnthropic, Anthropic
import boto3

# Try importing Agent SDK components
try:
    from claude_agent_sdk import (
        query as agent_query,
        ClaudeAgentOptions,
        AgentDefinition,
        AssistantMessage,
        ResultMessage,
        SystemMessage,
    )
    HAS_AGENT_SDK = True
except ImportError:
    HAS_AGENT_SDK = False
    # Define mock classes for type hints
    class ClaudeAgentOptions:
        pass
    class AgentDefinition:
        pass

logger = logging.getLogger(__name__)

# ============================================================================
# ENHANCEMENT STRATEGIES
# ============================================================================

class EnhancementStrategy:
    """Base class for enhancement strategies"""

    def __init__(self, name: str, description: str):
        self.name = name
        self.description = description

    async def enhance(
        self,
        query: str,
        rag_context: Dict[str, Any],
        **kwargs
    ) -> Tuple[str, Dict[str, Any]]:
        """
        Enhance the response using this strategy

        Args:
            query: Original user query
            rag_context: Context retrieved from RAG
            **kwargs: Additional parameters

        Returns:
            Tuple of (enhanced_answer, metadata)
        """
        raise NotImplementedError


class NoEnhancement(EnhancementStrategy):
    """No enhancement - pure RAG response"""

    def __init__(self):
        super().__init__(
            name="no_enhancement",
            description="Use RAG context only, no additional enhancement"
        )

    async def enhance(
        self,
        query: str,
        rag_context: Dict[str, Any],
        **kwargs
    ) -> Tuple[str, Dict[str, Any]]:
        """Generate response using only RAG context"""

        client = AsyncAnthropic()

        system_prompt = """You are an expert semiconductor test engineer assistant.
Answer based on the provided context from our knowledge base.
Cite sources using [Source: filename, Page N] format.
If the context doesn't contain enough information, say so clearly."""

        # Build context string
        context_parts = []

        # Add internal KB chunks
        if rag_context.get("chunks"):
            context_parts.append("=== INTERNAL KNOWLEDGE BASE ===")
            for i, chunk in enumerate(rag_context["chunks"], 1):
                context_parts.append(f"\n[{i}] {chunk.get('source', 'Unknown')}, "
                                   f"Page {chunk.get('page_number', 'N/A')}")
                context_parts.append(chunk.get("text", ""))

        # Add web results
        if rag_context.get("web_results"):
            context_parts.append("\n=== WEB SEARCH RESULTS ===")
            for i, result in enumerate(rag_context["web_results"], 1):
                context_parts.append(f"\n[{i}] {result.get('title', '')}")
                context_parts.append(f"URL: {result.get('url', '')}")
                context_parts.append(result.get("snippet", ""))

        context_str = "\n".join(context_parts)

        user_prompt = f"""Question: {query}

Context from knowledge base:
{context_str}

Please provide a comprehensive answer based on this context."""

        response = await client.messages.create(
            model="claude-3-opus-20240229",
            max_tokens=2000,
            temperature=0.1,
            system=system_prompt,
            messages=[{"role": "user", "content": user_prompt}]
        )

        return response.content[0].text, {"strategy": "no_enhancement"}


class BasicToolEnhancement(EnhancementStrategy):
    """Enhancement with basic read-only tools"""

    def __init__(self):
        super().__init__(
            name="basic_tools",
            description="Enhance with Read, Grep, Glob tools"
        )
        self.allowed_tools = ["Read", "Grep", "Glob"]

    async def enhance(
        self,
        query: str,
        rag_context: Dict[str, Any],
        **kwargs
    ) -> Tuple[str, Dict[str, Any]]:
        """Enhance using basic tools"""

        if not HAS_AGENT_SDK:
            # Fallback to no enhancement
            strategy = NoEnhancement()
            return await strategy.enhance(query, rag_context, **kwargs)

        # Build enhanced prompt with RAG context
        enhanced_prompt = self._build_enhanced_prompt(query, rag_context)

        # Create agent definition with limited tools
        agent_def = AgentDefinition(
            name="basic_enhancement_agent",
            description="Agent with basic read-only tools",
            system_prompt="""You are an expert semiconductor test engineer.
You have been given context from our knowledge base.
Use the available tools to verify or expand on this information.""",
            tools=self.allowed_tools
        )

        # Run agent
        options = ClaudeAgentOptions(
            model="claude-3-opus-20240229",
            max_thinking_tokens=2000,
            max_messages=10
        )

        result = await agent_query(
            prompt=enhanced_prompt,
            agent_definition=agent_def,
            options=options
        )

        # Extract answer and metadata
        answer = result.content if hasattr(result, 'content') else str(result)
        metadata = {
            "strategy": "basic_tools",
            "tools_used": self._extract_tools_used(result)
        }

        return answer, metadata

    def _build_enhanced_prompt(self, query: str, rag_context: Dict) -> str:
        """Build prompt that includes RAG context"""
        context_summary = self._summarize_rag_context(rag_context)

        return f"""User Query: {query}

I've already retrieved the following context from our knowledge base:

{context_summary}

You have access to these tools to verify or expand on this information:
{', '.join(self.allowed_tools)}

Please provide a comprehensive answer, using the tools if needed to:
1. Verify specific details mentioned in the context
2. Look up additional related information
3. Check for more recent updates

Always cite your sources."""

    def _summarize_rag_context(self, rag_context: Dict) -> str:
        """Summarize RAG context for the prompt"""
        parts = []

        if rag_context.get("chunks"):
            parts.append(f"Found {len(rag_context['chunks'])} relevant documents:")
            for chunk in rag_context["chunks"][:3]:  # First 3 chunks
                parts.append(f"- {chunk.get('source', 'Unknown')}: "
                           f"{chunk.get('text', '')[:200]}...")

        if rag_context.get("web_results"):
            parts.append(f"\nAlso found {len(rag_context['web_results'])} web results")

        return "\n".join(parts)

    def _extract_tools_used(self, result: Any) -> List[str]:
        """Extract list of tools used from agent result"""
        tools = []
        if hasattr(result, 'tool_calls'):
            for call in result.tool_calls:
                if hasattr(call, 'tool_name'):
                    tools.append(call.tool_name)
        return tools


class AutonomousEnhancement(EnhancementStrategy):
    """Full autonomous enhancement with all tools"""

    def __init__(self):
        super().__init__(
            name="autonomous",
            description="Full autonomous mode with all available tools"
        )
        self.allowed_tools = [
            "Read", "Grep", "Glob", "Edit", "Write",
            "WebSearch", "WebFetch", "Bash", "Agent"
        ]

    async def enhance(
        self,
        query: str,
        rag_context: Dict[str, Any],
        **kwargs
    ) -> Tuple[str, Dict[str, Any]]:
        """Enhance with full autonomy"""

        if not HAS_AGENT_SDK:
            strategy = NoEnhancement()
            return await strategy.enhance(query, rag_context, **kwargs)

        # Build comprehensive prompt
        prompt = f"""Query: {query}

Initial Context from Knowledge Base:
{self._format_context(rag_context)}

You have full autonomy with these tools: {', '.join(self.allowed_tools)}

Instructions:
1. Start with the provided context
2. Use tools to investigate further if needed
3. Can spawn subagents for complex subtasks
4. Provide comprehensive, well-researched answer
5. Cite all sources (KB, web, files)"""

        agent_def = AgentDefinition(
            name="autonomous_agent",
            description="Fully autonomous agent with all tools",
            system_prompt="""You are an expert semiconductor test engineer with
full autonomy to research and solve problems. Use all available tools
to provide the most comprehensive and accurate answer possible.""",
            tools=self.allowed_tools,
            allow_subagents=True
        )

        options = ClaudeAgentOptions(
            model="claude-3-opus-20240229",
            max_thinking_tokens=5000,
            max_messages=20,
            auto_approve_tools=True
        )

        result = await agent_query(
            prompt=prompt,
            agent_definition=agent_def,
            options=options
        )

        answer = result.content if hasattr(result, 'content') else str(result)
        metadata = {
            "strategy": "autonomous",
            "tools_used": self._extract_tools_used(result),
            "subagents_spawned": self._count_subagents(result)
        }

        return answer, metadata

    def _format_context(self, rag_context: Dict) -> str:
        """Format RAG context for autonomous agent"""
        lines = []

        if rag_context.get("chunks"):
            lines.append("=== Internal Knowledge Base ===")
            for i, chunk in enumerate(rag_context["chunks"], 1):
                lines.append(f"{i}. {chunk.get('source', 'Unknown')} "
                           f"(p.{chunk.get('page_number', 'N/A')})")
                lines.append(f"   {chunk.get('text', '')[:300]}...")
                lines.append("")

        if rag_context.get("confidence"):
            lines.append(f"Confidence Score: {rag_context['confidence']:.2f}")

        return "\n".join(lines)

    def _extract_tools_used(self, result: Any) -> List[str]:
        """Extract tools used from result"""
        tools = []
        if hasattr(result, 'metadata') and 'tools_used' in result.metadata:
            tools = result.metadata['tools_used']
        return tools

    def _count_subagents(self, result: Any) -> int:
        """Count number of subagents spawned"""
        if hasattr(result, 'metadata') and 'subagents' in result.metadata:
            return len(result.metadata['subagents'])
        return 0


# ============================================================================
# ENHANCED RAG AGENT
# ============================================================================

class EnhancedRAGAgent:
    """
    Two-step RAG agent with optional enhancement

    Architecture:
        1. Always retrieve from RAG (ChromaDB + web)
        2. Optionally enhance with Agent SDK tools

    This preserves the existing RAG system while adding
    powerful Agent SDK capabilities when needed.
    """

    def __init__(
        self,
        chroma_path: str = "chroma_persistent_storage",
        collection_name: str = "tml_copilot_v3_titan"
    ):
        """Initialize enhanced RAG agent"""
        # Core RAG components
        self.chroma_client = chromadb.PersistentClient(path=chroma_path)
        self.collection = self._get_or_create_collection(collection_name)

        # AWS Bedrock for embeddings
        self.bedrock_client = boto3.client(
            'bedrock-runtime',
            region_name=os.environ.get('AWS_REGION', 'us-east-1')
        )

        # Anthropic for LLM
        self.anthropic_client = AsyncAnthropic()

        # Enhancement strategies
        self.strategies = {
            "none": NoEnhancement(),
            "basic": BasicToolEnhancement(),
            "autonomous": AutonomousEnhancement()
        }

        logger.info(f"Enhanced RAG Agent initialized with {len(self.strategies)} strategies")

    def _get_or_create_collection(self, name: str):
        """Get or create ChromaDB collection"""
        try:
            return self.chroma_client.get_collection(name)
        except:
            return self.chroma_client.create_collection(
                name=name,
                metadata={"hnsw:space": "cosine"}
            )

    async def query(
        self,
        query: str,
        enhancement: str = "none",
        top_k: int = 5,
        confidence_threshold: float = 0.7,
        session_id: Optional[str] = None,
        **kwargs
    ) -> Dict[str, Any]:
        """
        Process query with two-step architecture

        Args:
            query: User query
            enhancement: Enhancement strategy ("none", "basic", "autonomous")
            top_k: Number of RAG chunks to retrieve
            confidence_threshold: Threshold for web search fallback
            session_id: Optional session ID for conversation continuity

        Returns:
            Dict containing answer, sources, metadata
        """
        import time
        start_time = time.time()

        # ====================================================================
        # STEP 1: CORE RAG RETRIEVAL (Always Active)
        # ====================================================================

        logger.info(f"Step 1: RAG retrieval for query: {query[:100]}...")

        # Get embeddings for query
        query_embedding = await self._get_embedding(query)

        # Search ChromaDB
        rag_results = self.collection.query(
            query_embeddings=[query_embedding],
            n_results=top_k,
            include=["documents", "metadatas", "distances"]
        )

        # Process RAG chunks
        rag_chunks = []
        for i in range(len(rag_results["documents"][0])):
            rag_chunks.append({
                "text": rag_results["documents"][0][i],
                "metadata": rag_results["metadatas"][0][i],
                "distance": rag_results["distances"][0][i],
                "source": rag_results["metadatas"][0][i].get("source", "Unknown"),
                "page_number": rag_results["metadatas"][0][i].get("page_number", 0)
            })

        # Calculate confidence
        confidence = self._calculate_confidence(rag_chunks)
        logger.info(f"RAG confidence: {confidence:.2f}")

        # Web search fallback if low confidence
        web_results = []
        if confidence < confidence_threshold:
            logger.info("Low confidence, searching web...")
            web_results = await self._search_web(query)

        # Build RAG context
        rag_context = {
            "chunks": rag_chunks,
            "web_results": web_results,
            "confidence": confidence,
            "query": query
        }

        # ====================================================================
        # STEP 2: OPTIONAL ENHANCEMENT
        # ====================================================================

        strategy = self.strategies.get(enhancement, self.strategies["none"])
        logger.info(f"Step 2: Applying enhancement strategy: {strategy.name}")

        # Enhance the response
        answer, enhancement_metadata = await strategy.enhance(
            query=query,
            rag_context=rag_context,
            session_id=session_id,
            **kwargs
        )

        # Build final response
        processing_time = time.time() - start_time

        return {
            "answer": answer,
            "sources": rag_chunks,
            "web_sources": web_results,
            "confidence": confidence,
            "enhancement": enhancement,
            "enhancement_metadata": enhancement_metadata,
            "processing_time": processing_time,
            "session_id": session_id,
            "timestamp": datetime.now().isoformat()
        }

    async def _get_embedding(self, text: str) -> List[float]:
        """Get embedding using AWS Bedrock Titan"""
        response = self.bedrock_client.invoke_model(
            modelId="amazon.titan-embed-text-v1",
            body=json.dumps({"inputText": text[:8000]})
        )
        result = json.loads(response['body'].read())
        return result['embedding']

    def _calculate_confidence(self, chunks: List[Dict]) -> float:
        """Calculate confidence score for RAG results"""
        if not chunks:
            return 0.0

        # Average distance (lower is better)
        avg_distance = sum(c["distance"] for c in chunks) / len(chunks)

        # Convert to confidence (0-1 scale)
        confidence = max(0, 1 - avg_distance)

        # Boost if multiple high-quality chunks
        high_quality = [c for c in chunks if c["distance"] < 0.3]
        if len(high_quality) >= 2:
            confidence = min(1.0, confidence * 1.2)

        return confidence

    async def _search_web(self, query: str, max_results: int = 3) -> List[Dict]:
        """Search web for additional context"""
        try:
            from duckduckgo_search import DDGS

            results = []
            with DDGS() as ddgs:
                for r in ddgs.text(query, max_results=max_results):
                    results.append({
                        "title": r.get("title", ""),
                        "url": r.get("link", ""),
                        "snippet": r.get("snippet", "")
                    })
            return results
        except Exception as e:
            logger.warning(f"Web search failed: {e}")
            return []

    async def ingest_document(
        self,
        content: str,
        metadata: Dict[str, Any],
        chunk_size: int = 400,
        chunk_overlap: int = 50
    ) -> int:
        """
        Ingest document into RAG system

        Args:
            content: Document content
            metadata: Document metadata
            chunk_size: Size of chunks in tokens
            chunk_overlap: Overlap between chunks

        Returns:
            Number of chunks created
        """
        # Check if already ingested
        doc_hash = hashlib.md5(content.encode()).hexdigest()
        existing = self.collection.get(where={"doc_hash": doc_hash})
        if existing["ids"]:
            logger.info(f"Document already ingested: {metadata.get('source', 'Unknown')}")
            return 0

        # Chunk the document
        chunks = self._chunk_text(content, chunk_size, chunk_overlap)

        # Generate embeddings and add to collection
        ids = []
        embeddings = []
        documents = []
        metadatas = []

        for i, chunk in enumerate(chunks):
            chunk_id = f"{doc_hash}_{i}"
            embedding = await self._get_embedding(chunk)

            ids.append(chunk_id)
            embeddings.append(embedding)
            documents.append(chunk)
            metadatas.append({
                **metadata,
                "doc_hash": doc_hash,
                "chunk_index": i,
                "chunk_total": len(chunks)
            })

        # Add to collection
        self.collection.add(
            ids=ids,
            embeddings=embeddings,
            documents=documents,
            metadatas=metadatas
        )

        logger.info(f"Ingested {len(chunks)} chunks from {metadata.get('source', 'Unknown')}")
        return len(chunks)

    def _chunk_text(
        self,
        text: str,
        chunk_size: int,
        chunk_overlap: int
    ) -> List[str]:
        """Chunk text with overlap"""
        # Simple word-based chunking
        words = text.split()
        chunks = []

        start = 0
        while start < len(words):
            end = min(start + chunk_size, len(words))
            chunk = " ".join(words[start:end])
            chunks.append(chunk)
            start = end - chunk_overlap

            if start >= len(words):
                break

        return chunks


# ============================================================================
# CONVENIENCE FUNCTIONS
# ============================================================================

async def create_enhanced_agent(
    enhancement_mode: str = "none"
) -> EnhancedRAGAgent:
    """Create an enhanced RAG agent with specified mode"""
    agent = EnhancedRAGAgent()
    logger.info(f"Created enhanced RAG agent with mode: {enhancement_mode}")
    return agent


async def query_with_enhancement(
    query: str,
    enhancement: str = "none",
    **kwargs
) -> Dict[str, Any]:
    """Query with specified enhancement level"""
    agent = await create_enhanced_agent(enhancement)
    return await agent.query(query, enhancement=enhancement, **kwargs)