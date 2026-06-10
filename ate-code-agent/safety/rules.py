"""
Safety rules for the ATE Code Agent.
These rules are enforced in tool implementations and cannot be overridden
without editing this file.
"""

import os
import ssl
import httpx
from typing import Any

os.environ["PYTHONHTTPSVERIFY"] = "0"
os.environ["CURL_CA_BUNDLE"] = ""
os.environ["REQUESTS_CA_BUNDLE"] = ""
ssl._create_default_https_context = ssl._create_unverified_context


SAFETY_RULES = {
    "max_iterations": 5,
    "require_human_approval": True,
    "readonly_sheets": ["Pin Map", "Spec"],
    "voltage_limits": {
        "VCC_max": 5.0,
        "VDD_max": 3.6,
        "VDDIO_max": 3.6,
        "VDDA_max": 5.0,
    },
    "allowed_run_modes": ["engineering", "characterization"],
    "git_backup_enabled": True,
}


def validate_plan(plan: dict[str, Any]) -> dict[str, Any]:
    """Validate a proposed plan against mandatory safety constraints."""
    violations: list[str] = []

    run_mode = plan.get("run_mode")
    if run_mode and run_mode not in SAFETY_RULES["allowed_run_modes"]:
        violations.append(
            f"Run mode '{run_mode}' is not allowed; use one of {SAFETY_RULES['allowed_run_modes']}"
        )

    edits = plan.get("edits", [])
    for edit in edits:
        sheet = str(edit.get("sheet_name", ""))
        if sheet in SAFETY_RULES["readonly_sheets"]:
            violations.append(f"Edit targets read-only sheet: {sheet}")

        value = edit.get("value")
        if sheet == "Levels" and isinstance(value, (int, float)):
            if float(value) > float(SAFETY_RULES["voltage_limits"]["VCC_max"]):
                violations.append(
                    f"Unsafe Levels voltage {value}V exceeds VCC_max={SAFETY_RULES['voltage_limits']['VCC_max']}V"
                )

    return {
        "ok": len(violations) == 0,
        "violations": violations,
    }


def get_http_client() -> httpx.Client:
    """Shared SSL-disabled httpx client for this environment."""
    return httpx.Client(verify=False)
