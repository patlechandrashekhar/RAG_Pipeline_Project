# Two-Step RAG + Agent SDK Architecture

## Overview

This is an improved architecture that implements a **two-step process** for the ADI ChipAgent:

1. **Step 1: Core RAG** (Always Active)
   - ChromaDB vector search
   - Query expansion
   - Confidence assessment  
   - Web search fallback

2. **Step 2: Agent SDK Enhancement** (Optional)
   - Takes RAG context as input
   - Applies additional tools
   - Returns enriched answer

## Key Features

### ✅ Your RAG System Stays Intact
- All existing RAG functions remain unchanged
- ChromaDB, retrieval, and generation modules still work
- Agent SDK is purely additive, not a replacement

### 🎯 Multiple Enhancement Modes

| Mode | Description | Available Tools |
|------|-------------|-----------------|
| **Disabled** | RAG only, no Agent SDK | None |
| **Standard** | RAG + basic read-only tools | Read, Grep, Glob |
| **Autonomous** | RAG + all tools with auto-approval | All tools + subagents |
| **Coding** | RAG + code editing capabilities | Read, Edit, Write, Bash |
| **Research** | RAG + web research tools | WebSearch, WebFetch, Agent |

### 🔄 Two-Step Process Flow

```
User Query
    ↓
┌─────────────────────────┐
│  STEP 1: Core RAG       │
│  (Always Active)        │
├─────────────────────────┤
│ • Query expansion       │
│ • ChromaDB search       │
│ • Confidence check      │
│ • Web fallback if < 0.7 │
└─────────────────────────┘
    ↓
RAG Context
    ↓
┌─────────────────────────┐
│  STEP 2: Enhancement    │
│  (Optional)             │
├─────────────────────────┤
│ • Takes RAG context     │
│ • Applies tools         │
│ • Enriches answer       │
└─────────────────────────┘
    ↓
Final Answer
```

## Installation

### 1. Install Required Dependencies

```bash
# Core dependencies (already installed)
pip install streamlit chromadb anthropic boto3

# Optional enhancements
pip install duckduckgo-search     # For web search
pip install claude-agent-sdk      # For full Agent SDK capabilities
```

### 2. Configure Environment Variables

Create or update `.env` file:

```bash
# Core RAG requirements
ANTHROPIC_API_KEY=your_anthropic_key
AWS_ACCESS_KEY_ID=your_aws_key
AWS_SECRET_ACCESS_KEY=your_aws_secret
AWS_REGION=us-east-1

# Optional for enhanced features
OPENAI_API_KEY=your_openai_key  # If using OpenAI embeddings
```

### 3. Test the System

Run the test script to verify everything works:

```bash
python test_two_step_architecture.py
```

Expected output:
- ✓ RAG System operational
- ✓ Two-Step System initialized
- ✓ Enhancement strategies available
- ⚠ Agent SDK (optional, may show warning if not installed)

## Running the Application

### Start the Streamlit App

```bash
streamlit run app/streamlit_app_agent_sdk_v2.py
```

### Using the Interface

1. **Select Enhancement Mode** (right sidebar)
   - Start with "Disabled" to test RAG-only
   - Try "Standard" for basic tool enhancement
   - Use "Autonomous" for full capabilities

2. **Ask Questions**
   - Semiconductor test engineering
   - TML/TDC specifications
   - ADI device information
   - V93000 ATE platform

3. **Monitor the Process**
   - Enable "Show process details" to see two-step flow
   - Check confidence scores
   - View sources from RAG
   - See which tools were used (if any)

## Architecture Benefits

### 1. **Graceful Degradation**
- If Agent SDK isn't available → falls back to RAG-only
- If web search fails → uses internal KB only
- If enhancement fails → returns RAG answer

### 2. **Flexible Enhancement**
- Choose enhancement level per query
- Switch modes without restarting
- Tools are applied contextually

### 3. **Performance Optimization**
- RAG provides fast initial context
- Agent SDK only activates when needed
- Caching at both levels

### 4. **Clear Separation of Concerns**
- **RAG Layer**: Fast retrieval, always available
- **Agent Layer**: Advanced capabilities, optional

## File Structure

```
page_indexing_RAG/
├── app/
│   ├── streamlit_app_agent_sdk_v2.py   # Two-step UI
│   └── streamlit_app_agent_sdk.py      # Original (for comparison)
│
├── src/page_indexing_rag/
│   ├── config_agentsdk.py              # Configuration
│   ├── rag_agent_sdk.py               # Original Agent SDK integration
│   └── rag_agent_enhanced.py          # Enhanced two-step agent
│
├── test_two_step_architecture.py       # Test script
└── README_TWO_STEP.md                  # This file
```

## Troubleshooting

### Issue: "Agent SDK not available"
**Solution**: This is normal if claude-agent-sdk isn't installed. The system will use RAG-only mode.

### Issue: "No ChromaDB documents"
**Solution**: Ingest documents first using the original ingestion scripts or UI.

### Issue: "AWS credentials error"
**Solution**: Check your AWS credentials in .env file. System will use dummy embeddings if not configured.

### Issue: "Anthropic API error"
**Solution**: Verify your ANTHROPIC_API_KEY is valid and has sufficient credits.

## Best Practices

1. **Start Simple**: Begin with RAG-only mode to establish baseline
2. **Gradual Enhancement**: Move from Standard → Coding/Research → Autonomous
3. **Monitor Confidence**: Low confidence triggers web search automatically
4. **Use Appropriate Mode**: 
   - Questions about specs → RAG-only or Standard
   - Code debugging → Coding mode
   - Research tasks → Research mode
   - Complex multi-step → Autonomous

## Migration from Original System

The two-step system is **fully backward compatible**:

1. Existing ChromaDB collections work as-is
2. Original RAG functions are preserved
3. Can run both UIs side-by-side
4. No data migration required

To switch between systems:
```bash
# Original system
streamlit run app/streamlit_app_agent_sdk.py

# Two-step system
streamlit run app/streamlit_app_agent_sdk_v2.py
```

## Future Enhancements

Potential improvements to the two-step architecture:

1. **Custom Enhancement Strategies**
   - Domain-specific tool combinations
   - Task-based mode selection

2. **Confidence Learning**
   - Track which queries need enhancement
   - Auto-select appropriate mode

3. **Parallel Processing**
   - Run RAG and initial tools simultaneously
   - Merge results intelligently

4. **Result Caching**
   - Cache both RAG and enhanced results
   - Invalidate based on KB updates

## Support

For issues or questions:
1. Run test script for diagnostics
2. Check logs in streamlit terminal
3. Verify environment variables
4. Ensure ChromaDB has data

The two-step architecture provides the best of both worlds: fast RAG retrieval with optional Agent SDK superpowers!