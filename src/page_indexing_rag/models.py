"""Model definitions for Portkey integration with various LLM providers."""

from typing import Dict, List, Tuple

# ════════════════════════════════════════════════════════════════════════════════
# AVAILABLE MODELS - Organized by vendor and capability
# ════════════════════════════════════════════════════════════════════════════════

# Large Language Models
LLM_MODELS: Dict[str, List[Tuple[str, str, str]]] = {
    "Anthropic (Claude)": [
        ("Claude Opus 4.6 (Latest)", "@bedrock-global/us.anthropic.claude-opus-4-6-v1", "top"),
        ("Claude Opus 4.5", "@bedrock-global/us.anthropic.claude-opus-4-5-20251101-v1:0", "top"),
        ("Claude Opus 4.1", "@bedrock-global/us.anthropic.claude-opus-4-1-20250805-v1:0", "top"),
        ("Claude Opus 4", "@bedrock-global/us.anthropic.claude-opus-4-20250514-v1:0", "top"),
        ("Claude Sonnet 4.6 (Latest)", "@bedrock-global/us.anthropic.claude-sonnet-4-6", "high"),
        ("Claude Sonnet 4.5", "@bedrock-global/us.anthropic.claude-sonnet-4-5-20250929-v1:0", "high"),
        ("Claude Sonnet 4", "@bedrock-global/us.anthropic.claude-sonnet-4-20250514-v1:0", "high"),
        ("Claude Haiku 4.5", "@bedrock-global/us.anthropic.claude-haiku-4-5-20251001-v1:0", "medium"),
        ("Claude Haiku 3.5", "@bedrock-global/us.anthropic.claude-3-5-haiku-20241022-v1:0", "medium"),
    ],
    "OpenAI (GPT)": [
        ("GPT-5.2 (Latest)", "@azure-openai-eus2-global/gpt-5.2-dzs", "top"),
        ("GPT-5.1", "@azure-openai-eus2-global/gpt-5.1-dzs", "top"),
        ("GPT-5", "@azure-openai-eus2-global/gpt-5-dzs", "top"),
        ("GPT-4.1", "@azure-openai-eus-global/gpt-4.1-dzs", "high"),
        ("GPT-4o", "@azure-openai-eus-global/gpt-4o-dzs", "high"),
        ("GPT-5 Mini", "@azure-openai-eus2-global/gpt-5-mini-dzs", "medium"),
        ("GPT-5 Nano", "@azure-openai-eus2-global/gpt-5-nano-dzs", "low"),
        ("GPT-4.1 Nano", "@azure-openai-eus-global/gpt-4.1-nano-dzs", "low"),
        ("GPT-4o Mini", "@azure-openai-eus-global/gpt-4o-mini-dzs", "low"),
    ],
    "OpenAI (o-series)": [
        ("o3 (Latest Reasoning)", "@azure-openai-eus-global/o3-dzs", "top"),
        ("o3 Mini", "@azure-openai-eus-global/o3-mini-dzs", "high"),
        ("o4 Mini", "@azure-openai-eus-global/o4-mini-dzs", "medium"),
        ("o1", "@azure-openai-eus-global/o1-dzs", "high"),
    ],
    "Google (Gemini)": [
        ("Gemini 2.5 Pro", "@vertexai-global/gemini-2.5-pro", "top"),
        ("Gemini 2.5 Flash", "@vertexai-global/gemini-2.5-flash", "medium"),
        ("Gemini 2.5 Flash Lite", "@vertexai-global/gemini-2.5-flash-lite", "low"),
    ],
}

# Embedding Models
EMBEDDING_MODELS: Dict[str, Tuple[str, int]] = {
    "OpenAI Text-Embedding-3-Large (Best)": ("@azure-openai-eus-global/text-embedding-3-large-std", 3072),
    "OpenAI Text-Embedding-3-Small": ("@azure-openai-eus-global/text-embedding-3-small-std", 1536),
    "OpenAI Text-Embedding-Ada-002": ("@azure-openai-eus-global/text-embedding-ada-002", 1536),
    "Amazon Titan Embed V2": ("@bedrock-global/amazon.titan-embed-text-v2", 1024),
    "Google Gemini Embedding": ("@vertexai-global/gemini-embedding-001", 768),
}

# ════════════════════════════════════════════════════════════════════════════════
# DEFAULT CONFIGURATIONS
# ════════════════════════════════════════════════════════════════════════════════

# Default model selections (can be overridden by environment variables)
DEFAULT_LLM_MODEL = "@bedrock-global/us.anthropic.claude-sonnet-4-6"  # Claude Sonnet 4.6
DEFAULT_EMBEDDING_MODEL = "@azure-openai-eus-global/text-embedding-3-large-std"  # Best embedding

# Model presets for different use cases
MODEL_PRESETS = {
    "Best Quality (Claude Opus 4.6)": {
        "model": "@bedrock-global/us.anthropic.claude-opus-4-6-v1",
        "description": "Highest quality responses, best for complex technical questions",
        "icon": "🏆"
    },
    "Fast & Smart (GPT-5.2)": {
        "model": "@azure-openai-eus2-global/gpt-5.2-dzs",
        "description": "Latest GPT model, excellent for general purpose",
        "icon": "🚀"
    },
    "Reasoning (o3)": {
        "model": "@azure-openai-eus-global/o3-dzs",
        "description": "Best for complex reasoning and problem-solving",
        "icon": "🧠"
    },
    "Balanced (Claude Sonnet 4.6)": {
        "model": "@bedrock-global/us.anthropic.claude-sonnet-4-6",
        "description": "Good balance of quality and speed",
        "icon": "⚖️"
    },
    "Fast (Gemini 2.5 Flash)": {
        "model": "@vertexai-global/gemini-2.5-flash",
        "description": "Quick responses for simple queries",
        "icon": "⚡"
    },
    "Economy (Claude Haiku 4.5)": {
        "model": "@bedrock-global/us.anthropic.claude-haiku-4-5-20251001-v1:0",
        "description": "Cost-effective for high volume",
        "icon": "💰"
    },
}

# ════════════════════════════════════════════════════════════════════════════════
# HELPER FUNCTIONS
# ════════════════════════════════════════════════════════════════════════════════

def get_model_display_name(model_string: str) -> str:
    """Get a user-friendly display name for a model string."""
    # Search through all models to find the display name
    for vendor, models in LLM_MODELS.items():
        for display_name, model_str, _ in models:
            if model_str == model_string:
                return display_name

    # Check presets
    for preset_name, preset_info in MODEL_PRESETS.items():
        if preset_info["model"] == model_string:
            return preset_name

    # Check embedding models
    for display_name, (model_str, _) in EMBEDDING_MODELS.items():
        if model_str == model_string:
            return display_name

    # Return the model string if no display name found
    return model_string

def get_model_vendor(model_string: str) -> str:
    """Get the vendor name for a model string."""
    if "anthropic" in model_string or "claude" in model_string:
        return "Anthropic"
    elif "gpt" in model_string or "o1" in model_string or "o3" in model_string or "o4" in model_string:
        return "OpenAI"
    elif "gemini" in model_string:
        return "Google"
    elif "titan" in model_string:
        return "Amazon"
    return "Unknown"

def get_model_capability(model_string: str) -> str:
    """Get the capability level of a model (top/high/medium/low)."""
    for vendor, models in LLM_MODELS.items():
        for _, model_str, capability in models:
            if model_str == model_string:
                return capability
    return "unknown"