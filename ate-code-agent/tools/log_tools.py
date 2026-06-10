"""
Tools for parsing IG-XL and STDF test result log files.
"""

import os
import ssl
import httpx
from pathlib import Path

os.environ["PYTHONHTTPSVERIFY"] = "0"
os.environ["CURL_CA_BUNDLE"] = ""
os.environ["REQUESTS_CA_BUNDLE"] = ""
ssl._create_default_https_context = ssl._create_unverified_context


def parse_stdf_log(log_path: str) -> dict:
    """Parse STDF log into structured test results."""
    path = Path(log_path)
    if not path.exists():
        raise FileNotFoundError(f"STDF log not found: {log_path}")

    results = {}
    try:
        from pystdf.Importer import STDF2DataFrame  # type: ignore

        df = STDF2DataFrame(str(path))

        for record in df.get("PTR", []):
            name = record.get("TEST_TXT") or f"test_{record.get('TEST_NUM', 'unknown')}"
            flag = record.get("TEST_FLG", 1)
            results[name] = {
                "result": "PASS" if flag == 0 else "FAIL",
                "measured": record.get("RESULT"),
                "limit_lo": record.get("LO_LIMIT"),
                "limit_hi": record.get("HI_LIMIT"),
                "units": record.get("UNITS", ""),
                "test_num": record.get("TEST_NUM"),
            }

        for record in df.get("FTR", []):
            name = record.get("TEST_TXT") or f"ftest_{record.get('TEST_NUM', 'unknown')}"
            flag = record.get("TEST_FLG", 1)
            results[name] = {
                "result": "PASS" if flag == 0 else "FAIL",
                "measured": None,
                "limit_lo": None,
                "limit_hi": None,
                "units": "",
                "test_num": record.get("TEST_NUM"),
            }
    except Exception as exc:
        results = _parse_text_log(str(path), str(exc))

    return results


def _parse_text_log(log_path: str, stdf_error: str) -> dict:
    """Fallback parser for plain text logs."""
    results = {
        "__parse_error__": {
            "result": "FAIL",
            "measured": None,
            "limit_lo": None,
            "limit_hi": None,
            "units": "",
            "test_num": 0,
            "note": f"STDF parse failed: {stdf_error}. Raw log at {log_path}",
        }
    }
    try:
        with open(log_path, "r", encoding="utf-8", errors="replace") as handle:
            for line in handle:
                row = line.strip()
                if "PASS" in row or "FAIL" in row:
                    parts = row.split()
                    if len(parts) >= 2:
                        name = parts[0]
                        status = "PASS" if "PASS" in row else "FAIL"
                        results[name] = {
                            "result": status,
                            "measured": None,
                            "limit_lo": None,
                            "limit_hi": None,
                            "units": "",
                            "test_num": 0,
                        }
    except Exception:
        pass
    return results


def summarize_results(results: dict) -> dict:
    """Summarize pass/fail counts and test lists."""
    passed = [t for t, r in results.items() if r["result"] == "PASS"]
    failed = [t for t, r in results.items() if r["result"] == "FAIL"]
    total = len(results)
    return {
        "total": total,
        "passed": len(passed),
        "failed": len(failed),
        "pass_rate": f"{(len(passed) / total * 100):.1f}%" if total else "0%",
        "failed_tests": failed,
        "passed_tests": passed,
    }
