"""Unit tests for log tools."""

import os
import ssl
import httpx
from pathlib import Path
import tempfile

os.environ["PYTHONHTTPSVERIFY"] = "0"
os.environ["CURL_CA_BUNDLE"] = ""
os.environ["REQUESTS_CA_BUNDLE"] = ""
ssl._create_default_https_context = ssl._create_unverified_context


ROOT = Path(__file__).resolve().parents[1]
import sys
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


def test_parse_text_log_fallback() -> None:
    from tools.log_tools import parse_stdf_log, summarize_results

    with tempfile.NamedTemporaryFile(suffix=".stdf", delete=False, mode="w", encoding="utf-8") as handle:
        handle.write("IDDQ_TEST FAIL measured=152 limit_hi=150\n")
        handle.write("CONTINUITY_TEST PASS measured=0.4 limit_hi=1.0\n")
        path = handle.name

    try:
        results = parse_stdf_log(path)
        summary = summarize_results(results)
        assert summary["total"] >= 2
        assert summary["failed"] >= 1
    finally:
        os.unlink(path)
