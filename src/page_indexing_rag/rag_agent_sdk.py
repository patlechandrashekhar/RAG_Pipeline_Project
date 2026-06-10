"""
RAG Agent with actual file and web access capabilities.

This version provides a working implementation with real file access
and web search functionality.
"""

from __future__ import annotations

import asyncio
import json
import re
from html import unescape
from pathlib import Path
from typing import List, Dict, Optional, Any, Tuple
from dataclasses import dataclass, field
from datetime import datetime
from urllib.parse import urlparse

import chromadb
import boto3
from anthropic import AsyncAnthropic

# Try importing web search capabilities
try:
    from duckduckgo_search import DDGS
    HAS_WEB_SEARCH = True
except ImportError:
    HAS_WEB_SEARCH = False
    print("Warning: duckduckgo-search not installed. Web search disabled.")

try:
    import aiohttp
    HAS_WEB_FETCH = True
except ImportError:
    HAS_WEB_FETCH = False
    print("Warning: aiohttp not installed. Web fetch disabled.")

from .ingestion_agentsdk import classify_file, semantic_chunk
from .retrieval_agentsdk import classify_query_type, assess_internal_confidence
from .generation_agentsdk import build_file_context_prompt, build_rag_context
from .config_agentsdk import (
    AWS_ACCESS_KEY_ID,
    AWS_REGION,
    AWS_SECRET_ACCESS_KEY,
    MASTER_SYSTEM_PROMPT,
    ANTHROPIC_API_KEY,
    _resolve_dir,
)


@dataclass
class QueryResponse:
    """Response from RAG query containing answer and metadata."""
    answer: str
    sources: List[Dict[str, Any]]
    web_sources: Optional[List[Dict[str, Any]]] = None
    query_type: str = "U"
    complexity: str = "medium"
    model_used: str = "claude-3-opus-20240229"
    confidence: Dict[str, bool] = None
    session_id: Optional[str] = None
    tools_used: List[str] = field(default_factory=list)
    telemetry: Dict[str, Any] = field(default_factory=dict)


class SemiconductorRAGAgentSDK:
    """
    Simplified RAG Agent with simulated tool capabilities.

    This version provides basic RAG functionality with mock tool support
    without requiring external Agent SDK packages.
    """

    def __init__(self, chroma_path: Optional[str] = None):
        """Initialize the RAG system."""
        # ChromaDB setup
        self.chroma_path = chroma_path or _resolve_dir(
            "data/chroma_persistent_storage",
            "chroma_persistent_storage"
        )
        self._init_chromadb()

        # AWS Bedrock for embeddings
        self.bedrock = None
        if AWS_ACCESS_KEY_ID and AWS_SECRET_ACCESS_KEY:
            self.bedrock = boto3.client(
                'bedrock-runtime',
                region_name=AWS_REGION,
                aws_access_key_id=AWS_ACCESS_KEY_ID,
                aws_secret_access_key=AWS_SECRET_ACCESS_KEY
            )

        # Anthropic client for LLM
        self.anthropic_client = None
        if ANTHROPIC_API_KEY:
            self.anthropic_client = AsyncAnthropic(api_key=ANTHROPIC_API_KEY)

        # System prompt for semiconductor domain
        self.system_prompt = MASTER_SYSTEM_PROMPT

        # Active sessions for context persistence
        self.sessions: Dict[str, List[Dict[str, str]]] = {}

    def _init_chromadb(self):
        """Initialize ChromaDB client and collection."""
        self.chroma_client = chromadb.PersistentClient(path=str(self.chroma_path))
        try:
            self.collection = self.chroma_client.get_collection("tml_copilot_v3_titan")
        except:
            self.collection = self.chroma_client.create_collection(
                name="tml_copilot_v3_titan",
                metadata={"hnsw:space": "cosine"}
            )

    async def get_titan_embedding(self, text: str) -> List[float]:
        """Get embedding using AWS Bedrock Titan."""
        if not self.bedrock:
            # Return dummy embedding if no Bedrock client
            return [0.0] * 1024

        try:
            response = self.bedrock.invoke_model(
                modelId="amazon.titan-embed-text-v1",
                body=json.dumps({"inputText": text[:8000]})
            )
            result = json.loads(response['body'].read())
            return result['embedding']
        except Exception as e:
            print(f"Embedding error: {e}")
            return [0.0] * 1024

    async def retrieve_context(
        self,
        query: str,
        top_k: int = 5
    ) -> Tuple[List[Dict[str, Any]], float]:
        """Retrieve relevant context from ChromaDB."""
        # Get embedding for query
        embedding = await self.get_titan_embedding(query)

        # Search ChromaDB
        if self.collection.count() == 0:
            return [], 0.0

        results = self.collection.query(
            query_embeddings=[embedding],
            n_results=min(top_k, self.collection.count()),
            include=["documents", "metadatas", "distances"]
        )

        # Format results
        chunks = []
        for i in range(len(results["documents"][0])):
            chunks.append({
                "text": results["documents"][0][i],
                "source": results["metadatas"][0][i].get("source", "Unknown"),
                "page_number": results["metadatas"][0][i].get("page_number", 0),
                "distance": results["distances"][0][i]
            })

        # Calculate confidence
        avg_distance = sum(c["distance"] for c in chunks) / len(chunks) if chunks else 1.0
        confidence = max(0, 1 - avg_distance)

        return chunks, confidence

    async def process_with_agent_sdk(
        self,
        user_query: str,
        allowed_tools: List[str] = None,
        use_subagents: bool = False,
        session_id: Optional[str] = None,
        project_contexts: Optional[List[Any]] = None,
        use_internal_kb: bool = True,
        approved_actions: bool = False,
    ) -> QueryResponse:
        """
        Process query with actual file access capabilities.

        This method implements real tool usage for file operations
        within the boundaries of attached project folders.
        """
        allowed_tools = allowed_tools or []
        intents = self._infer_tool_intents(user_query)
        plan = self._build_agentic_plan(
            user_query=user_query,
            allowed_tools=allowed_tools,
            intents=intents,
            use_internal_kb=use_internal_kb,
            approved_actions=approved_actions,
        )

        telemetry: Dict[str, Any] = {
            "phases_enabled": [
                "phase1_planner_confidence_routing",
                "phase2_critic_retry",
                "phase3_approval_runbooks",
                "phase4_telemetry",
            ],
            "planner": plan,
            "critic": [],
            "retries": 0,
            "failure_reasons": [],
            "tool_choice": [],
        }

        # Retrieve context from ChromaDB based on planner decision.
        chunks: List[Dict[str, Any]] = []
        confidence = 0.0
        context = ""
        if plan["use_internal_kb"]:
            chunks, confidence = await self.retrieve_context(user_query)
            context = build_rag_context(chunks) if chunks else ""

        # Execute actual tools based on query and allowed tools
        tools_used = []
        tool_results = []
        runbook_steps: List[Dict[str, str]] = []

        # Parse allowed directories from project contexts
        allowed_dirs = []
        if project_contexts:
            for project in project_contexts:
                if hasattr(project, 'path'):
                    p = Path(project.path)
                    if p.exists() and p.is_dir():
                        allowed_dirs.append(p)

        # Also allow directories provided directly in the user request.
        for raw_dir in self._extract_directory_paths(user_query):
            p = Path(raw_dir)
            if p.exists() and p.is_dir():
                allowed_dirs.append(p)

        # De-duplicate while preserving order
        allowed_dirs = list(dict.fromkeys(allowed_dirs))

        if allowed_tools and allowed_dirs:
            # Implement actual tool execution
            # READ tool - Read files from allowed directories
            if "Read" in allowed_tools:
                file_paths = self._extract_file_paths(user_query)
                for file_path in file_paths:
                    result = await self._read_file(file_path, allowed_dirs)
                    if result:
                        tools_used.append("Read")
                        tool_results.append(result)

            # GREP tool - Search for patterns in allowed directories
            if "Grep" in allowed_tools and intents.get("grep"):
                pattern = self._extract_search_pattern(user_query)
                if pattern:
                    results = await self._grep_directories(pattern, allowed_dirs)
                    if results:
                        tools_used.append("Grep")
                        tool_results.append(f"Found {len(results)} matches for '{pattern}':\n" + "\n".join(results[:10]))

            # GLOB tool - List files matching patterns
            if "Glob" in allowed_tools and intents.get("glob"):
                results = await self._list_files(allowed_dirs)
                if results:
                    tools_used.append("Glob")
                    tool_results.append(f"Files in attached directories:\n" + "\n".join(results[:20]))

            # EDIT tool - Modify files (with confirmation)
            if "Edit" in allowed_tools and intents.get("edit"):
                if not approved_actions:
                    runbook_steps.append({
                        "action": "Edit",
                        "purpose": "Apply an approved targeted replacement in a file",
                        "status": "pending_human_approval",
                    })
                    tool_results.append("[Approval required] Edit was requested but not executed.")
                    telemetry["failure_reasons"].append("approval_required_edit")
                else:
                    edit_request = self._extract_edit_request(user_query)
                    if edit_request:
                        file_path, old_text, new_text = edit_request
                        result = await self._edit_file(file_path, old_text, new_text, allowed_dirs)
                        tools_used.append("Edit")
                        tool_results.append(result)
                    else:
                        tools_used.append("Edit")
                        tool_results.append(
                            "[Edit tool: provide edits like: edit \"path/file.ext\" replace \"old text\" with \"new text\"]"
                        )

            # WRITE tool - Create new files
            if "Write" in allowed_tools and intents.get("write"):
                if not approved_actions:
                    runbook_steps.append({
                        "action": "Write",
                        "purpose": "Create/update a file with generated content",
                        "status": "pending_human_approval",
                    })
                    tool_results.append("[Approval required] Write was requested but not executed.")
                    telemetry["failure_reasons"].append("approval_required_write")
                else:
                    write_request = self._extract_write_request(user_query)
                    if not write_request:
                        write_request = self._extract_write_request_relaxed(user_query)
                    if write_request:
                        file_path, content = write_request
                        result = await self._write_file(file_path, content, allowed_dirs)
                        tools_used.append("Write")
                        tool_results.append(result)
                    else:
                        tools_used.append("Write")
                        tool_results.append(
                            "[Write tool: provide path and content, for example: write \"path/file.txt\" ```text\\ncontent\\n```]"
                        )

        # Web search tools (work independently of local directories)
        if allowed_tools:
            query_lower = user_query.lower()

            # WEBSEARCH tool - Search the web for any information
            if "WebSearch" in allowed_tools and HAS_WEB_SEARCH and plan["allow_web_tools"]:
                # Trigger web search for various keywords or if explicitly requested
                web_triggers = ["latest", "recent", "current", "news", "web", "search online",
                               "what is", "who is", "when", "where", "how to", "price",
                               "weather", "stock", "update", "2024", "2025", "2026"]

                if any(trigger in query_lower for trigger in web_triggers):
                    # Extract the search query
                    search_query = self._extract_web_search_query(user_query)
                    if search_query:
                        web_results = await self._web_search(search_query)
                        if web_results:
                            tools_used.append("WebSearch")
                            tool_results.append(f"Web Search Results for '{search_query}':\n{web_results}")

            # WEBFETCH tool - Fetch specific web pages
            if "WebFetch" in allowed_tools and HAS_WEB_FETCH and plan["allow_web_tools"]:
                urls = self._extract_urls(user_query)
                for url in urls:
                    content = await self._fetch_webpage(url)
                    if content:
                        tools_used.append("WebFetch")
                        tool_results.append(f"Content from {url}:\n{content[:2000]}...")

        # Confidence-based routing: escalate to web search when KB confidence is low.
        if plan["promote_web_on_low_confidence"] and confidence < plan["kb_confidence_threshold"]:
            if "WebSearch" in allowed_tools and HAS_WEB_SEARCH:
                web_results = await self._web_search(user_query)
                if web_results:
                    tools_used.append("WebSearch")
                    tool_results.append(f"Planner escalation web results:\n{web_results}")
                    telemetry["planner"]["escalated_to_web"] = True

        if runbook_steps:
            tool_results.append(self._format_runbook(runbook_steps))

        # Build full prompt with context and tool results
        system_prompt = self.system_prompt
        if allowed_tools:
            system_prompt += f"\n\nYou have access to these tools: {', '.join(allowed_tools)}"
            if allowed_dirs:
                system_prompt += f"\n\nYou can access files in these directories:\n"
                for dir_path in allowed_dirs:
                    system_prompt += f"- {dir_path}\n"

        user_prompt = (
            f"Query: {user_query}\n\n"
            f"Execution plan (phase 1 planner): {json.dumps(plan, default=str)}\n\n"
        )

        if context:
            user_prompt += f"Retrieved Context from Knowledge Base:\n{context}\n\n"

        if tool_results:
            user_prompt += "Tool Execution Results:\n"
            for result in tool_results:
                user_prompt += f"{result}\n\n"

        if allowed_dirs and not tool_results:
            user_prompt += "Note: You have access to the attached directories. You can ask me to read, search, or modify files within them.\n\n"

        user_prompt += "Please provide a comprehensive answer based on the available context and tool results."

        # Generate response using Anthropic
        answer = await self._generate_response(system_prompt, user_prompt)
        critic = self._critic_pass(answer=answer, chunks=chunks, tool_results=tool_results)
        telemetry["critic"].append(critic)

        # Phase 2: critic + retry policy (single safe retry).
        if not critic["pass"]:
            telemetry["retries"] = 1
            retry_prompt = (
                user_prompt
                + "\n\nCritic feedback: "
                + "; ".join(critic["issues"])
                + "\nImprove answer with explicit assumptions and missing-data callouts."
            )
            answer_retry = await self._generate_response(system_prompt, retry_prompt)
            critic_retry = self._critic_pass(answer=answer_retry, chunks=chunks, tool_results=tool_results)
            telemetry["critic"].append(critic_retry)
            if critic_retry["pass"]:
                answer = answer_retry
            else:
                telemetry["failure_reasons"].append("critic_failed_after_retry")
        if self._is_generation_error(answer):
            answer = await self._build_local_fallback_response(
                user_query=user_query,
                allowed_dirs=allowed_dirs,
                tool_results=tool_results,
                generation_error=answer,
            )

        unique_tools = list(dict.fromkeys(tools_used))
        telemetry["tool_choice"] = unique_tools
        telemetry["confidence_score"] = round(float(confidence), 3)
        telemetry["internal_kb_used"] = bool(plan["use_internal_kb"])

        # Create response object
        return QueryResponse(
            answer=answer,
            sources=[{"text": c["text"][:200], "source": c["source"], "page": c["page_number"]}
                    for c in chunks[:3]],
            query_type=classify_query_type(user_query),
            complexity="medium",
            model_used="claude-3-opus-20240229",
            confidence={"internal": (confidence > 0.5) if use_internal_kb else False, "web": False},
            session_id=session_id,
            tools_used=unique_tools,
            telemetry=telemetry,
        )

    def _build_agentic_plan(
        self,
        user_query: str,
        allowed_tools: List[str],
        intents: Dict[str, bool],
        use_internal_kb: bool,
        approved_actions: bool,
    ) -> Dict[str, Any]:
        """Phase 1 planner that decides tool strategy and confidence routing policy."""
        query_type = classify_query_type(user_query)
        action_requested = bool(intents.get("edit") or intents.get("write"))
        return {
            "query_type": query_type,
            "use_internal_kb": bool(use_internal_kb),
            "allow_web_tools": "WebSearch" in allowed_tools or "WebFetch" in allowed_tools,
            "promote_web_on_low_confidence": True,
            "kb_confidence_threshold": 0.45,
            "action_requested": action_requested,
            "action_execution_allowed": bool(approved_actions and action_requested),
            "strategy": "tool_first" if any(intents.values()) else "kb_first",
        }

    def _critic_pass(
        self,
        answer: str,
        chunks: List[Dict[str, Any]],
        tool_results: List[str],
    ) -> Dict[str, Any]:
        """Phase 2 critic verdict over generated answer quality."""
        issues: List[str] = []
        text = (answer or "").strip()
        if not text:
            issues.append("empty_answer")
        if self._is_generation_error(text):
            issues.append("generation_error")
        if len(text) < 80 and not chunks and not tool_results:
            issues.append("insufficient_contextual_answer")
        return {"pass": len(issues) == 0, "issues": issues}

    def _format_runbook(self, steps: List[Dict[str, str]]) -> str:
        """Phase 3 structured runbook for approval-gated actions."""
        lines = [
            "Action runbook (approval-gated):",
            "1) Review proposed file targets and safety impact.",
            "2) Confirm approval to execute actions.",
            "3) Execute one action at a time and verify outcome.",
            "4) Re-run checks/tests and summarize deltas.",
            "",
            "Planned actions:",
        ]
        for i, step in enumerate(steps, 1):
            lines.append(
                f"- Step {i}: action={step.get('action')} | purpose={step.get('purpose')} | status={step.get('status')}"
            )
        return "\n".join(lines)

    def _is_generation_error(self, answer: str) -> bool:
        """Detect known response-generation failures."""
        if not answer:
            return True
        msg = answer.strip().lower()
        return (
            msg.startswith("error generating response:")
            or msg.startswith("error: anthropic api key not configured")
        )

    def _looks_like_init_request(self, user_query: str) -> bool:
        """Detect init-style onboarding prompts, including transformed slash commands."""
        q = (user_query or "").lower()
        init_markers = [
            "/init",
            "initialize this project like a coding agent",
            "map directory structure",
            "identify entry points",
            "step-by-step plan",
            "bootstrap",
        ]
        return any(marker in q for marker in init_markers)

    async def _build_local_fallback_response(
        self,
        user_query: str,
        allowed_dirs: List[Path],
        tool_results: List[str],
        generation_error: str,
    ) -> str:
        """Return a useful local-only response when cloud generation is unavailable."""
        if not allowed_dirs:
            return (
                f"{generation_error}\n\n"
                "Local fallback could not run because no attached directory was detected. "
                "Attach at least one folder and retry /init."
            )

        listed_files = await self._list_files(allowed_dirs, max_files=60)
        files = [f for f in listed_files if f.startswith("📄 ")]

        # Basic language/entrypoint heuristics for a code-agent init summary.
        entry_hints = []
        for line in files:
            rel = line.replace("📄 ", "", 1)
            low = rel.lower()
            if low.endswith(("main.py", "app.py", "streamlit_app.py", "index.ts", "index.js", "package.json", "pyproject.toml", "makefile", "readme.md")):
                entry_hints.append(rel)

        top_files = files[:20]
        if self._looks_like_init_request(user_query):
            sections = [
                "Connection to LLM failed, so I ran local code-agent init from attached directories.",
                "",
                "Project map (sample):",
            ]
            sections.extend([f"- {f.replace('📄 ', '', 1)}" for f in top_files] or ["- No files found in attached directories"])
            sections.extend([
                "",
                "Likely entry points:",
            ])
            sections.extend([f"- {p}" for p in entry_hints[:12]] or ["- No obvious entrypoints detected from filename heuristics"])
            sections.extend([
                "",
                "Risks to check first:",
                "- Missing environment variables or API keys",
                "- Broken imports or stale module references",
                "- Misconfigured run/test commands",
                "",
                "Suggested plan:",
                "1. Confirm entrypoint and run command from README/pyproject.",
                "2. Run focused tests or smoke run for the selected path.",
                "3. Fix first failing error with minimal patch.",
                "4. Re-run validation and capture remaining issues.",
                "",
                "Note: cloud response generation is currently unavailable; retry after API/network recovery for deeper semantic analysis.",
            ])
            return "\n".join(sections)

        tool_section = "\n".join(tool_results[:6]) if tool_results else "No tool output captured before fallback."
        return (
            f"{generation_error}\n\n"
            "Cloud response generation is unavailable. "
            "I continued with local tool fallback over attached directories.\n\n"
            f"Local findings:\n{tool_section}"
        )

    def _resolve_allowed_file(
        self,
        file_path: str,
        allowed_dirs: List[Path],
        allow_missing: bool = False,
    ) -> Optional[Path]:
        """Resolve a file path and ensure it stays inside one of the allowed directories."""
        if not file_path:
            return None

        raw_path = Path(file_path.strip().strip('"').strip("'"))

        for allowed_dir in allowed_dirs:
            allowed_root = allowed_dir.resolve()
            candidate = raw_path if raw_path.is_absolute() else allowed_root / raw_path
            try:
                normalized = candidate.resolve(strict=False)
            except Exception:
                continue

            if not str(normalized).startswith(str(allowed_root)):
                continue

            if allow_missing:
                return normalized
            if normalized.exists() and normalized.is_file():
                return normalized

        return None

    async def _read_file(self, file_path: str, allowed_dirs: List[Path]) -> Optional[str]:
        """Read a file if it's within allowed directories."""
        try:
            resolved = self._resolve_allowed_file(file_path, allowed_dirs, allow_missing=False)
            if not resolved:
                return f"File '{file_path}' not found in allowed directories"

            content = resolved.read_text(encoding='utf-8', errors='ignore')
            return f"Contents of {resolved}:\n```\n{content[:3000]}\n```"
        except Exception as e:
            return f"Error reading file: {str(e)}"

    async def _write_file(self, file_path: str, content: str, allowed_dirs: List[Path]) -> str:
        """Write content to a file within allowed directories."""
        try:
            resolved = self._resolve_allowed_file(file_path, allowed_dirs, allow_missing=True)
            if not resolved:
                return f"Write denied: '{file_path}' is outside allowed directories"

            resolved.parent.mkdir(parents=True, exist_ok=True)
            resolved.write_text(content, encoding="utf-8")
            return f"Wrote {len(content)} chars to {resolved}"
        except Exception as e:
            return f"Error writing file: {str(e)}"

    async def _edit_file(self, file_path: str, old_text: str, new_text: str, allowed_dirs: List[Path]) -> str:
        """Replace exact text in a file inside allowed directories."""
        try:
            resolved = self._resolve_allowed_file(file_path, allowed_dirs, allow_missing=False)
            if not resolved:
                return f"Edit denied: '{file_path}' is outside allowed directories or does not exist"

            content = resolved.read_text(encoding="utf-8", errors="ignore")
            if old_text not in content:
                return f"Edit not applied: target text not found in {resolved}"

            updated = content.replace(old_text, new_text, 1)
            resolved.write_text(updated, encoding="utf-8")
            return f"Edited {resolved}: replaced one matching block"
        except Exception as e:
            return f"Error editing file: {str(e)}"

    async def _grep_directories(self, pattern: str, allowed_dirs: List[Path]) -> List[str]:
        """Search for pattern in files within allowed directories."""
        results = []
        try:
            for allowed_dir in allowed_dirs:
                # Search in common text file types
                for ext in ['*.py', '*.txt', '*.md', '*.json', '*.yml', '*.yaml', '*.js', '*.ts', '*.html', '*.css']:
                    for file_path in allowed_dir.rglob(ext):
                        try:
                            with open(file_path, 'r', encoding='utf-8') as f:
                                lines = f.readlines()
                                for i, line in enumerate(lines, 1):
                                    if pattern.lower() in line.lower():
                                        results.append(f"{file_path.relative_to(allowed_dir)}:{i}: {line.strip()}")
                                        if len(results) >= 20:  # Limit results
                                            return results
                        except Exception:
                            continue
        except Exception as e:
            results.append(f"Error searching: {str(e)}")
        return results

    async def _list_files(self, allowed_dirs: List[Path], max_files: int = 50) -> List[str]:
        """List files in allowed directories."""
        files = []
        try:
            for allowed_dir in allowed_dirs:
                for file_path in allowed_dir.rglob('*'):
                    if file_path.is_file():
                        relative_path = file_path.relative_to(allowed_dir)
                        files.append(f"📄 {relative_path}")
                        if len(files) >= max_files:
                            files.append(f"... and more files in {allowed_dir}")
                            return files
        except Exception as e:
            files.append(f"Error listing files: {str(e)}")
        return files

    def _extract_file_paths(self, query: str) -> List[str]:
        """Extract potential file paths from query."""
        import re
        # Look for quoted paths or common file patterns
        paths = []

        # Find quoted strings
        quoted = re.findall(r'"([^"]+)"', query) + re.findall(r"'([^']+)'", query)
        for item in quoted:
            if '.' in item or '/' in item or '\\' in item:
                paths.append(item)

        # Find common file patterns
        patterns = re.findall(r'[\w\\/.-]+\.\w+', query)
        paths.extend(patterns)

        return list(set(paths))  # Remove duplicates

    def _extract_search_pattern(self, query: str) -> Optional[str]:
        """Extract search pattern from query."""
        import re
        # Look for quoted strings or after keywords like "search for", "find"
        quoted = re.findall(r'"([^"]+)"', query) + re.findall(r"'([^']+)'", query)
        if quoted:
            return quoted[0]

        # Look for patterns after keywords
        patterns = re.findall(r'(?:search for|find|grep|look for|where is|contains)\s+([^\n\r\.]+)', query, re.IGNORECASE)
        if patterns:
            return patterns[0].strip()

        return None

    def _extract_write_request(self, query: str) -> Optional[Tuple[str, str]]:
        """Extract write/create request from natural language query."""
        fenced = re.search(
            r"(?:write|create)\s+(?:file\s+)?[\"']([^\"']+)[\"'][^`]*```[a-zA-Z0-9_+-]*\n(.*?)```",
            query,
            flags=re.IGNORECASE | re.DOTALL,
        )
        if fenced:
            return fenced.group(1).strip(), fenced.group(2)

        inline = re.search(
            r"(?:write|create)\s+(?:file\s+)?[\"']([^\"']+)[\"']\s*:\s*(.+)$",
            query,
            flags=re.IGNORECASE | re.DOTALL,
        )
        if inline:
            return inline.group(1).strip(), inline.group(2).strip()

        # Unquoted file path with fenced code block.
        unquoted_fenced = re.search(
            r"(?:write|create|add|generate)\s+(?:file\s+)?([A-Za-z0-9_./\\-]+\.[A-Za-z0-9_+-]+)[^`]*```[a-zA-Z0-9_+-]*\n(.*?)```",
            query,
            flags=re.IGNORECASE | re.DOTALL,
        )
        if unquoted_fenced:
            return unquoted_fenced.group(1).strip(), unquoted_fenced.group(2)

        # Unquoted file path with inline content.
        unquoted_inline = re.search(
            r"(?:write|create|add|generate)\s+(?:file\s+)?([A-Za-z0-9_./\\-]+\.[A-Za-z0-9_+-]+)\s*:\s*(.+)$",
            query,
            flags=re.IGNORECASE | re.DOTALL,
        )
        if unquoted_inline:
            return unquoted_inline.group(1).strip(), unquoted_inline.group(2).strip()

        return None

    def _extract_fenced_code_block(self, query: str) -> Optional[str]:
        """Extract first fenced code block body from query."""
        m = re.search(r"```[a-zA-Z0-9_+-]*\n(.*?)```", query, flags=re.DOTALL)
        if not m:
            return None
        return m.group(1)

    def _extract_write_request_relaxed(self, query: str) -> Optional[Tuple[str, str]]:
        """Fallback write extraction for natural requests that don't follow strict template."""
        file_paths = self._extract_file_paths(query)
        if not file_paths:
            return None

        code_block = self._extract_fenced_code_block(query)
        if not code_block:
            return None

        # Prefer source-like files first if multiple are mentioned.
        preferred = sorted(
            file_paths,
            key=lambda p: 0 if Path(p).suffix.lower() in {
                ".py", ".js", ".ts", ".tsx", ".jsx", ".java", ".cs", ".cpp", ".c", ".h", ".hpp",
                ".go", ".rs", ".rb", ".php", ".swift", ".kt", ".scala", ".vb", ".bas", ".md", ".txt"
            } else 1,
        )
        return preferred[0], code_block

    def _extract_edit_request(self, query: str) -> Optional[Tuple[str, str, str]]:
        """Extract edit request in the format: edit "path" replace "old" with "new"."""
        match = re.search(
            r"(?:edit|modify|change)\s+[\"']([^\"']+)[\"']\s+replace\s+[\"']([^\"']+)[\"']\s+with\s+[\"']([^\"']+)[\"']",
            query,
            flags=re.IGNORECASE | re.DOTALL,
        )
        if not match:
            return None
        return match.group(1).strip(), match.group(2), match.group(3)

    def _extract_web_search_query(self, query: str) -> str:
        """Extract search phrase from query; fallback to original text."""
        q = query.strip()
        m = re.search(r"(?:search|look up|find)\s+(?:for\s+)?(.+)$", q, re.IGNORECASE)
        if m:
            q = m.group(1).strip()
        return q[:240]

    def _extract_directory_paths(self, query: str) -> List[str]:
        """Extract absolute Windows directory paths from free-form query text."""
        candidates = []

        quoted = re.findall(r"[\"']([A-Za-z]:\\[^\"']+)[\"']", query)
        candidates.extend(quoted)

        raw = re.findall(r"([A-Za-z]:\\[^\n,;]+)", query)
        candidates.extend(raw)

        clean = []
        seen = set()
        for c in candidates:
            p = c.strip().rstrip(". )]")
            if p and p not in seen:
                seen.add(p)
                clean.append(p)
        return clean

    def _infer_tool_intents(self, query: str) -> Dict[str, bool]:
        """Infer which local tools are requested from natural language."""
        q = query.lower()

        grep_tokens = ["search", "find", "grep", "where", "contains", "match", "matches", "occurrence"]
        glob_tokens = ["list", "show files", "files", "tree", "folder", "directory", "structure", "scan"]
        edit_tokens = ["edit", "modify", "change", "update", "replace", "fix", "refactor"]
        write_tokens = ["create", "write", "add file", "new file", "generate file"]

        return {
            "grep": any(t in q for t in grep_tokens),
            "glob": any(t in q for t in glob_tokens),
            "edit": any(t in q for t in edit_tokens),
            "write": any(t in q for t in write_tokens),
        }

    async def _web_search(self, query: str, max_results: int = 5) -> str:
        """Run DuckDuckGo search and format concise results."""
        if not HAS_WEB_SEARCH:
            return "Web search unavailable: duckduckgo-search not installed"
        try:
            def _run() -> List[Dict[str, Any]]:
                with DDGS() as ddgs:
                    return list(ddgs.text(query, max_results=max_results))

            rows = await asyncio.to_thread(_run)
            lines = []
            for i, row in enumerate(rows, 1):
                title = row.get("title", "Untitled")
                href = row.get("href", "")
                body = (row.get("body", "") or "")[:300]
                lines.append(f"{i}. {title}\\nURL: {href}\\n{body}")
            return "\n\n".join(lines) if lines else "No web results found"
        except Exception as e:
            return f"Web search error: {str(e)}"

    def _extract_urls(self, query: str) -> List[str]:
        """Extract http/https URLs from query."""
        urls = re.findall(r"https?://[^\s)\]>]+", query)
        return list(dict.fromkeys(urls))

    async def _fetch_webpage(self, url: str) -> str:
        """Fetch a web page and return plain text excerpt."""
        parsed = urlparse(url)
        if parsed.scheme not in {"http", "https"}:
            return f"Unsupported URL scheme: {url}"

        if not HAS_WEB_FETCH:
            return "Web fetch unavailable: aiohttp not installed"

        try:
            timeout = aiohttp.ClientTimeout(total=12)
            async with aiohttp.ClientSession(timeout=timeout) as session:
                async with session.get(url, ssl=False) as resp:
                    body = await resp.text(errors="ignore")
            text = re.sub(r"<script[\s\S]*?</script>", " ", body, flags=re.IGNORECASE)
            text = re.sub(r"<style[\s\S]*?</style>", " ", text, flags=re.IGNORECASE)
            text = re.sub(r"<[^>]+>", " ", text)
            text = re.sub(r"\s+", " ", unescape(text)).strip()
            return text[:4000]
        except Exception as e:
            return f"Web fetch error for {url}: {str(e)}"

    async def _generate_response(self, system_prompt: str, user_prompt: str) -> str:
        """Generate response using Anthropic Claude."""
        if not self.anthropic_client:
            return "Error: Anthropic API key not configured. Please set ANTHROPIC_API_KEY in your .env file."

        try:
            response = await self.anthropic_client.messages.create(
                model="claude-3-opus-20240229",
                max_tokens=2000,
                temperature=0.1,
                system=system_prompt,
                messages=[{"role": "user", "content": user_prompt}]
            )
            return response.content[0].text
        except Exception as e:
            return f"Error generating response: {str(e)}"

    def _get_file_fingerprint(self, file_path: str) -> str:
        """Generate SHA-256 fingerprint of file for duplicate detection."""
        import hashlib
        sha256_hash = hashlib.sha256()
        try:
            with open(file_path, "rb") as f:
                # Read first 64KB for fingerprint
                chunk = f.read(65536)
                sha256_hash.update(chunk)
        except Exception as e:
            print(f"Error generating fingerprint: {e}")
            # Fallback to filename-based fingerprint
            sha256_hash.update(Path(file_path).name.encode())
        return sha256_hash.hexdigest()

    async def ingest_document(
        self,
        file_path: str,
        file_content: Optional[str] = None
    ) -> bool:
        """Ingest a document into the vector database."""
        try:
            file_path_obj = Path(file_path)
            source_name = file_path_obj.name

            # Generate fingerprint for duplicate detection
            fingerprint = self._get_file_fingerprint(file_path)

            # Check if already ingested
            existing = self.collection.get(
                where={"fingerprint": fingerprint[:16]}
            )
            if existing and existing['ids']:
                print(f"Document {source_name} already in database (fingerprint: {fingerprint[:16]})")
                return False  # Already exists

            # Handle different file types
            chunks = []
            metadata_list = []

            if file_path_obj.suffix.lower() == '.pdf':
                # Import PDF processing function
                from .ingestion_agentsdk import extract_pdf_pages

                try:
                    pages = extract_pdf_pages(str(file_path))
                    for page in pages:
                        page_chunks = semantic_chunk(page['text'], max_tokens=400, overlap_sentences=2)
                        for i, chunk in enumerate(page_chunks):
                            chunks.append(chunk)
                            metadata_list.append({
                                "source": source_name,
                                "file_type": classify_file(source_name, page['text']),
                                "page_number": page['page_number'],
                                "total_pages": page['total_pages'],
                                "chunk_index": i,
                                "fingerprint": fingerprint[:16]
                            })
                except Exception as e:
                    print(f"PDF extraction error for {source_name}: {e}")
                    return False

            elif file_path_obj.suffix.lower() in ['.html', '.xhtml', '.htm']:
                # Import HTML processing function
                from .ingestion_agentsdk import extract_html_document

                try:
                    doc_text = extract_html_document(str(file_path), source_name)
                    doc_chunks = semantic_chunk(doc_text, max_tokens=400, overlap_sentences=2)
                    for i, chunk in enumerate(doc_chunks):
                        chunks.append(chunk)
                        metadata_list.append({
                            "source": source_name,
                            "file_type": classify_file(source_name, doc_text),
                            "page_number": 0,
                            "chunk_index": i,
                            "fingerprint": fingerprint[:16]
                        })
                except Exception as e:
                    print(f"HTML extraction error for {source_name}: {e}")
                    return False

            else:
                # Plain text file or use provided content
                if file_content is None:
                    try:
                        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                            file_content = f.read()
                    except Exception as e:
                        print(f"Error reading file {source_name}: {e}")
                        return False

                # Classify and chunk the content
                file_type = classify_file(source_name, file_content)
                text_chunks = semantic_chunk(file_content, max_tokens=400, overlap_sentences=2)

                for i, chunk in enumerate(text_chunks):
                    chunks.append(chunk)
                    metadata_list.append({
                        "source": source_name,
                        "file_type": file_type,
                        "page_number": 0,
                        "chunk_index": i,
                        "fingerprint": fingerprint[:16]
                    })

            # Skip if no chunks were created
            if not chunks:
                print(f"No content extracted from {source_name}")
                return False

            # Generate unique IDs using fingerprint to avoid collisions
            ids = [f"{fingerprint[:16]}_{i}" for i in range(len(chunks))]

            # Embed and store all chunks
            embeddings = []
            for chunk in chunks:
                embedding = await self.get_titan_embedding(chunk)
                embeddings.append(embedding)

            # Add to ChromaDB in batch
            self.collection.add(
                documents=chunks,
                embeddings=embeddings,
                metadatas=metadata_list,
                ids=ids
            )

            print(f"Successfully ingested {source_name}: {len(chunks)} chunks")
            return True

        except Exception as e:
            print(f"Ingestion error for {file_path}: {e}")
            import traceback
            traceback.print_exc()
            return False

    async def query(
        self,
        question: str,
        session_id: Optional[str] = None,
        use_web: bool = False
    ) -> QueryResponse:
        """
        Main query endpoint for the RAG system.

        This is a simplified version that just retrieves context and generates a response.
        """
        return await self.process_with_agent_sdk(
            user_query=question,
            allowed_tools=["Read", "Grep", "Glob"],  # Default tools
            session_id=session_id
        )