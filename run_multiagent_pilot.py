"""CLI pilot runner for UltraFlex multi-agent generation."""

from __future__ import annotations

import argparse
import inspect
import json
import os
import ssl
from pathlib import Path

import httpx  # Required by workspace SSL policy for all new scripts.

os.environ["PYTHONHTTPSVERIFY"] = "0"
os.environ["CURL_CA_BUNDLE"] = ""
os.environ["REQUESTS_CA_BUNDLE"] = ""
ssl._create_default_https_context = ssl._create_unverified_context

from src.page_indexing_rag.multiagent.orchestrator import build_from_reference


def _call_build_from_reference_safe(**kwargs):
    """Call build_from_reference with only parameters supported at runtime."""
    sig = inspect.signature(build_from_reference)
    supported = {k: v for k, v in kwargs.items() if k in sig.parameters}
    return build_from_reference(**supported)


def _prompt_required(label: str, default: str | None = None) -> str:
    while True:
        suffix = f" [{default}]" if default else ""
        value = input(f"{label}{suffix}: ").strip()
        if value:
            return value
        if default:
            return default
        print("This value is required.")


def _prompt_optional(label: str, default: str | None = None) -> str | None:
    suffix = f" [{default}]" if default else ""
    value = input(f"{label}{suffix} (optional): ").strip()
    if value:
        return value
    return default


def _prompt_int(label: str, default: int | None = None) -> int:
    while True:
        suffix = f" [{default}]" if default is not None else ""
        value = input(f"{label}{suffix}: ").strip()
        if not value and default is not None:
            return default
        try:
            return int(value)
        except ValueError:
            print("Please enter a valid integer.")


def _validate_base_program(base_program: str) -> None:
    base = Path(base_program)
    if not base.exists() or not base.is_dir():
        raise ValueError(f"Base program path does not exist or is not a directory: {base}")
    exports = base / "Exports"
    datalogs = base / "Datalogs"
    if not exports.exists() or not datalogs.exists():
        raise ValueError(
            f"Base program must contain Exports and Datalogs folders: {base}"
        )


def _validate_optional_file(path_value: str | None, label: str) -> None:
    if not path_value:
        return
    p = Path(path_value)
    if not p.exists():
        raise ValueError(f"{label} does not exist: {p}")
    if p.is_dir():
        raise ValueError(f"{label} must be a file, not a directory: {p}")


def _resolve_optional_input(path_value: str | None, label: str, extensions: tuple[str, ...]) -> str | None:
    """Resolve optional file input; if a directory is provided, auto-pick first matching file."""
    if not path_value:
        return None

    p = Path(path_value)
    if not p.exists():
        raise ValueError(f"{label} does not exist: {p}")

    if p.is_file():
        return str(p)

    candidates: list[Path] = []
    for ext in extensions:
        candidates.extend(sorted(p.glob(f"*{ext}")))

    if not candidates:
        ext_label = ", ".join(extensions)
        raise ValueError(f"{label} directory has no supported files ({ext_label}): {p}")

    chosen = candidates[0]
    print(f"{label}: directory provided, auto-selected file -> {chosen}")
    return str(chosen)


def main() -> None:
    parser = argparse.ArgumentParser(description="Run UltraFlex multi-agent generation pilot")
    parser.add_argument("--device", required=False, help="Target device name, e.g. ADIN6210")
    parser.add_argument("--package-pins", required=False, type=int, help="Package pin count: 32/40/64")
    parser.add_argument("--sites", required=False, type=int, help="Number of test sites")
    parser.add_argument("--base-program", required=False, help="Reference UltraFlex program path (contains Exports and Datalogs)")
    parser.add_argument("--output", required=False, help="Output directory for generated program bundle")
    parser.add_argument("--pattern-base", default=None, help="Optional base path to validate .PAT files")
    parser.add_argument("--pin-map", default=None, help="Optional pin mapping file (.xlsx/.csv)")
    parser.add_argument("--datasheet", default=None, help="Optional datasheet/spec file (.pdf/.txt)")
    parser.add_argument(
        "--required-tests",
        default=None,
        help="Optional comma-separated test names for custom BAS generation",
    )
    parser.add_argument(
        "--no-web-vbt",
        action="store_true",
        help="Disable web knowledge augmentation for BAS generation",
    )

    args = parser.parse_args()

    if not args.device:
        print("\nNo CLI arguments detected. Interactive pilot mode enabled.\n")
        args.device = _prompt_required("Device name", "ADIN6210")
        args.package_pins = _prompt_int("Package pins", 40)
        args.sites = _prompt_int("Number of sites", 2)
        args.base_program = _prompt_required("Base program path (must contain Exports and Datalogs)")
        args.output = _prompt_required("Output directory")
        args.pin_map = _prompt_optional("Pin mapping file path (.xlsx/.csv)")
        args.datasheet = _prompt_optional("Datasheet/spec file path (.pdf/.txt)")
        args.pattern_base = _prompt_optional("Pattern base path")
        args.required_tests = _prompt_optional("Required tests (comma-separated)")

    _validate_base_program(args.base_program)
    args.pin_map = _resolve_optional_input(args.pin_map, "Pin mapping file", (".xlsx", ".xlsm", ".csv"))
    args.datasheet = _resolve_optional_input(args.datasheet, "Datasheet/spec file", (".pdf", ".txt"))
    required_tests = [s.strip() for s in (args.required_tests or "").split(",") if s.strip()]

    try:
        report = _call_build_from_reference_safe(
            device_name=args.device,
            package_pins=args.package_pins,
            num_sites=args.sites,
            base_program_path=args.base_program,
            output_dir=args.output,
            pattern_base_path=args.pattern_base,
            pin_mapping_file=args.pin_map,
            datasheet_file=args.datasheet,
            required_tests=required_tests,
            use_web_vbt_knowledge=not args.no_web_vbt,
        )
    except Exception as exc:
        print(f"Generation failed: {exc}")
        print("Tip: ensure Exports contains Exec_IP_Module.bas, RunVBT.bas, and at least one MDIO module.")
        raise SystemExit(1)

    out_dir = Path(args.output)
    out_dir.mkdir(parents=True, exist_ok=True)
    report_path = out_dir / "pilot_report.json"
    report_path.write_text(json.dumps(report, indent=2), encoding="utf-8")

    print(f"Generated files: {len(report.get('generated_files', []))}")
    print(f"Output dir: {out_dir}")
    print(f"Report: {report_path}")
    print(f"Pin rows parsed: {report.get('input_pin_rows', 0)}")
    print(f"Missing patterns: {len(report.get('missing_patterns', []))}")
    print(f"Hardware failure flags: {len(report.get('hw_failures', []))}")


if __name__ == "__main__":
    main()
