"""Unit tests for VBA tools."""

import os
import ssl
import httpx

import pytest

os.environ["PYTHONHTTPSVERIFY"] = "0"
os.environ["CURL_CA_BUNDLE"] = ""
os.environ["REQUESTS_CA_BUNDLE"] = ""
ssl._create_default_https_context = ssl._create_unverified_context


def test_vba_tools_require_windows_excel() -> None:
    from tools.vba_tools import get_vba_module_names

    if os.name != "nt":
        with pytest.raises(RuntimeError, match="Windows"):
            get_vba_module_names("dummy.xlsx")
