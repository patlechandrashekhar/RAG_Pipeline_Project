# Claude Agent SDK Integration Setup

## Installation

```bash
# 1. Install the claude-agent-sdk package
pip install claude-agent-sdk>=0.2.111

# 2. Install other requirements  
pip install -r requirements.txt
```

## Environment Variables

Create a `.env` file with:

```bash
# Required for Agent SDK
ANTHROPIC_API_KEY=your_anthropic_api_key_here

# Required for Titan embeddings
AWS_ACCESS_KEY_ID=your_aws_key
AWS_SECRET_ACCESS_KEY=your_aws_secret
AWS_REGION=us-east-1
```

## Running the Agent SDK App

```bash
# From the page_indexing_RAG directory
streamlit run app/streamlit_app_agent_sdk.py
```

## Available Tools

The Claude Agent SDK provides these built-in tools:

- **📖 Read** - Read any file in the working directory
- **✏️ Edit** - Make precise edits to existing files  
- **📝 Write** - Create new files
- **🔍 Grep** - Search file contents with regex
- **📁 Glob** - Find files by pattern (e.g., **/*.py)
- **🌐 WebSearch** - Search the web for current information
- **🌏 WebFetch** - Fetch and parse web page content
- **💻 Bash** - Run terminal commands
- **🤖 Agent** - Spawn specialized subagents
- **❓ AskUserQuestion** - Ask clarifying questions
- **📊 Monitor** - Watch background scripts

## Query Modes

1. **Standard RAG** - Uses knowledge base + web search
2. **TML Debugger** - Specialized for debugging TML test programs
3. **Datasheet Analyzer** - Searches and analyzes semiconductor datasheets  
4. **Full Autonomous** - Claude decides which tools to use

## Tool Permissions

You can control which tools are available via the sidebar:
- Allow file editing (Edit, Write)
- Allow web search (WebSearch, WebFetch)
- Allow command execution (Bash)
- Allow subagents (Agent)

## Custom Agents

The system includes specialized agents:
- **tml-debugger** - Expert at debugging TML code for V93000 ATE
- **datasheet-analyzer** - Analyzes semiconductor datasheets
- **knowledge-searcher** - Searches internal KB and web

## Troubleshooting

### "Module not found: claude_agent_sdk"
```bash
pip install claude-agent-sdk
```

### "ANTHROPIC_API_KEY not set"
Add your Anthropic API key to the .env file

### Async errors in Streamlit
The app handles async operations automatically. If you see async errors, restart Streamlit.

## Features

- ✅ Autonomous tool use - Claude decides which tools to use
- ✅ Session persistence - Continue conversations across runs
- ✅ Streaming responses - See tool use in real-time
- ✅ Hybrid retrieval - ChromaDB + Web search
- ✅ File analysis - Upload and analyze TML, PDF, etc.
- ✅ Custom agents - Specialized agents for different tasks