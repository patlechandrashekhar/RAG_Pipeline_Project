# page_indexing_RAG - Claude Agent SDK Version

This is the migrated version of the page_indexing_RAG application using Claude Agent SDK and Amazon Titan embeddings.

## 🚀 Quick Start

### 1. Set up credentials

Copy `.env.agentsdk` to `.env` and add your credentials:

```bash
cp .env.agentsdk .env
```

Edit `.env` and add:
- `ANTHROPIC_API_KEY` - Get from [Anthropic Console](https://console.anthropic.com/)
- `AWS_ACCESS_KEY_ID` - AWS credentials for Bedrock
- `AWS_SECRET_ACCESS_KEY` - AWS secret key
- `AWS_REGION` - Default: us-east-1

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Test the setup

```bash
python test_quick.py
```

### 4. Run the application

```bash
streamlit run app/streamlit_app_agentsdk.py
```

## 📁 Key Files

- **`app/streamlit_app_agentsdk.py`** - Streamlit UI with Agent SDK backend
- **`src/page_indexing_rag/rag_agent.py`** - Core RAG agent using Claude SDK
- **`src/page_indexing_rag/config_agentsdk.py`** - Simplified configuration
- **`migrate_to_titan.py`** - Migration script for existing ChromaDB data

## 🔄 Migration from Original Version

If you have existing data in the old ChromaDB collection:

```bash
python migrate_to_titan.py
```

This will:
1. Read documents from the old `tml_copilot_v2` collection
2. Re-embed them using Amazon Titan
3. Store in new `tml_copilot_v3_titan` collection

## 🆕 What's Changed

### From Original → Agent SDK Version

| Feature | Original | Agent SDK Version |
|---------|----------|-------------------|
| LLM | OpenAI/Portkey | Claude (native) |
| Embeddings | OpenAI | Amazon Titan |
| Web Search | DuckDuckGo | Placeholder (extensible) |
| Config | Complex Portkey | Simple Claude SDK |
| Collection | tml_copilot_v2 | tml_copilot_v3_titan |

### Preserved Features

✅ All semiconductor domain expertise  
✅ File classification (TML, TDC, datasheet, etc.)  
✅ Streamlit UI interface  
✅ ChromaDB vector storage  
✅ Semantic chunking algorithm  
✅ MMR deduplication  
✅ Query expansion  
✅ Hybrid retrieval logic  

## 🐛 Troubleshooting

### "TypeError: cross_encoder_rerank() takes 2 positional arguments"
Fixed in the latest version. The reranking now uses MMR results directly.

### "ANTHROPIC_API_KEY not set"
Add your Anthropic API key to the `.env` file.

### "AWS credentials not set"
Add AWS credentials for Bedrock/Titan access to the `.env` file.

### "ChromaDB collection empty"
Either:
- Ingest new documents using the UI
- Migrate existing data: `python migrate_to_titan.py`

## 📊 Performance Notes

- **Titan embeddings**: 1024 dimensions (vs 3072 for OpenAI)
- **Claude Haiku**: Used for query expansion (fast & cheap)
- **Claude Opus**: Used for main responses (high quality)
- **Batch size**: 10 documents for migration

## 🔒 Security

- SSL verification disabled (legacy behavior from original)
- Credentials stored in `.env` (never commit this file)
- ChromaDB stored locally in `data/chroma_persistent_storage/`

## 📝 License

Internal tool for Analog Devices Inc.