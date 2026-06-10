"""
Tools for triggering IG-XL test execution on UltraFLEX.
"""

import os
import ssl
import httpx
import subprocess
from pathlib import Path

from safety.rules import SAFETY_RULES

os.environ["PYTHONHTTPSVERIFY"] = "0"
os.environ["CURL_CA_BUNDLE"] = ""
os.environ["REQUESTS_CA_BUNDLE"] = ""
ssl._create_default_https_context = ssl._create_unverified_context


def run_igxl_program(
    program_path: str,
    log_output_path: str,
    timeout_seconds: int = 600,
    mode: str = "engineering",
) -> dict:
    """Execute an IG-XL program and wait for completion."""
    if mode not in SAFETY_RULES["allowed_run_modes"]:
        raise ValueError(
            f"SAFETY BLOCK: Run mode '{mode}' not allowed. "
            f"Allowed modes: {SAFETY_RULES['allowed_run_modes']}"
        )

    cmd = [
        "igxl_run.exe",
        "--program",
        str(Path(program_path).resolve()),
        "--output",
        str(Path(log_output_path).resolve()),
        "--format",
        "stdf",
        "--mode",
        mode,
    ]

    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            timeout=timeout_seconds,
            cwd=str(Path(program_path).parent),
        )
        return {
            "return_code": result.returncode,
            "stdout": result.stdout.decode(errors="replace"),
            "stderr": result.stderr.decode(errors="replace"),
            "log_path": log_output_path,
            "status": "complete" if result.returncode == 0 else "error",
        }
    except subprocess.TimeoutExpired:
        return {
            "return_code": -1,
            "stdout": "",
            "stderr": f"Timed out after {timeout_seconds}s",
            "log_path": log_output_path,
            "status": "timeout",
        }
    except FileNotFoundError:
        return {
            "return_code": -1,
            "stdout": "",
            "stderr": "igxl_run.exe not found - is IG-XL installed?",
            "log_path": log_output_path,
            "status": "error",
        }


def mock_run_igxl_program(program_path: str, log_output_path: str) -> dict:
    """Testing helper when no tester/IG-XL runtime is available."""
    out = Path(log_output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(
        "IDDQ_TEST FAIL measured=152.3 limit_hi=150\n"
        "CONTINUITY_TEST PASS measured=0.5 limit_hi=1.0\n",
        encoding="utf-8",
    )
    return {
        "return_code": 0,
        "stdout": "mock run complete",
        "stderr": "",
        "log_path": str(out),
        "status": "complete",
    }
