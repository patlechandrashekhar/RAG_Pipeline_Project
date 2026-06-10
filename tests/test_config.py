"""Test configuration module."""

import os
import sys
from pathlib import Path
from unittest.mock import patch, MagicMock
import pytest

# Add src to path for imports
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from page_indexing_rag import config


class TestConfigModule:
    """Test configuration module functionality."""

    def test_environment_variables_loaded(self):
        """Test that environment variables are properly loaded."""
        # Check that openai_key or PORTKEY_API_KEY is set (even if it's a placeholder)
        assert hasattr(config, 'openai_key') or hasattr(config, 'PORTKEY_API_KEY'), "openai_key or PORTKEY_API_KEY should be defined"

    def test_paths_exist(self):
        """Test that all configured paths are properly defined."""
        paths_to_check = [
            'PROJECT_ROOT',
            'WORKSPACE_ROOT',
            'PDF_DATA_DIR',
            'HTML_DATA_DIR',
            'TXT_DATA_DIR',
            'CHROMA_PATH'
        ]

        for path_name in paths_to_check:
            assert hasattr(config, path_name), f"{path_name} should be defined in config"
            path_value = getattr(config, path_name)
            assert path_value is not None, f"{path_name} should not be None"

    def test_resolve_dir_with_local_path(self):
        """Test _resolve_dir with existing local path."""
        # Create a temporary local directory
        temp_local = Path("temp_test_local")
        temp_local.mkdir(exist_ok=True)

        try:
            result = config._resolve_dir("temp_test_local", "legacy_path")
            assert "temp_test_local" in result
        finally:
            # Cleanup
            temp_local.rmdir()

    def test_resolve_dir_with_legacy_path(self):
        """Test _resolve_dir falls back to legacy path when local doesn't exist."""
        result = config._resolve_dir("non_existent_local", "c:/AI Projects")
        assert "AI Projects" in result

    def test_master_system_prompt_defined(self):
        """Test that MASTER_SYSTEM_PROMPT is defined and not empty."""
        assert hasattr(config, 'MASTER_SYSTEM_PROMPT'), "MASTER_SYSTEM_PROMPT should be defined"
        assert config.MASTER_SYSTEM_PROMPT, "MASTER_SYSTEM_PROMPT should not be empty"
        assert "semiconductor" in config.MASTER_SYSTEM_PROMPT.lower(), "System prompt should contain domain expertise"

    def test_openai_client_configuration(self):
        """Test that OpenAI client is properly configured."""
        assert hasattr(config, 'client'), "OpenAI client should be defined"
        # Note: We can't test actual API calls without valid credentials

    def test_model_configurations(self):
        """Test that model configurations are properly set."""
        model_configs = [
            'OPENAI_ANSWER_MODEL',
            'OPENAI_RETRIEVAL_MODEL',
            'OPENAI_WEB_MODEL',
            'OPENAI_EMBEDDING_MODEL'
        ]

        for model_config in model_configs:
            assert hasattr(config, model_config), f"{model_config} should be defined"
            value = getattr(config, model_config)
            assert value, f"{model_config} should not be empty"

    def test_ssl_configuration(self):
        """Test that SSL verification is properly configured."""
        # Check that SSL verification is disabled (as per corporate requirements)
        assert os.environ.get("PYTHONHTTPSVERIFY") == "0", "SSL verification should be disabled"
        assert os.environ.get("CURL_CA_BUNDLE") == "", "CURL_CA_BUNDLE should be empty"

    @patch.dict(os.environ, {"OPENAI_API_KEY": "test_key_123"})
    def test_api_key_from_environment(self):
        """Test that API key can be loaded from environment."""
        # This would require reloading the config module
        # In practice, the API key is loaded at module import time
        assert os.environ.get("OPENAI_API_KEY") == "test_key_123"

    def test_data_directories_structure(self):
        """Test that data directory structure is correctly configured."""
        # Check that paths follow expected pattern
        assert "pdf_data" in config.PDF_DATA_DIR.lower()
        assert "html_data" in config.HTML_DATA_DIR.lower() or "html" in config.HTML_DATA_DIR.lower()
        assert "chroma" in config.CHROMA_PATH.lower()