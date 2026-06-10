"""
Standalone MCP stdio server for test-program development scaffolding.
User provides channel assignment text/JSON, and tools generate:
- Pin map
- Channel map
- Datasheet draft
"""

from __future__ import annotations

import csv
import io
import json
import os
import re
import ssl
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any

import httpx
from mcp.server.fastmcp import FastMCP

os.environ["PYTHONHTTPSVERIFY"] = "0"
os.environ["CURL_CA_BUNDLE"] = ""
os.environ["REQUESTS_CA_BUNDLE"] = ""
ssl._create_default_https_context = ssl._create_unverified_context

# Keep compatibility with existing project import behavior.
PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = PROJECT_ROOT / "src"
if str(SRC_ROOT) not in os.sys.path:
    os.sys.path.insert(0, str(SRC_ROOT))

# No external HTTP calls are required, but keep a local client object for
# consistency with the repository SSL/network policy.
_HTTP_CLIENT = httpx.Client(verify=False, timeout=10)

mcp = FastMCP("ultraflex-testprog")

_KEY_ALIASES = {
    "pin": "pin_name",
    "pin_name": "pin_name",
    "signal": "pin_name",
    "net": "pin_name",
    "channel": "channel",
    "ch": "channel",
    "channel_id": "channel",
    "resource": "channel",
    "site": "site",
    "site_id": "site",
    "direction": "direction",
    "dir": "direction",
    "io": "direction",
    "instrument": "instrument",
    "card": "instrument",
    "comment": "comment",
    "notes": "comment",
}

_DEFAULT_COLUMNS = ["pin_name", "channel", "site", "direction", "instrument", "comment"]

_PROTOCOL_PRESETS: dict[str, dict[str, str]] = {
    "SPI": {
        "period_ns": "40",  # 25 MHz default
        "format": "NRZ",
        "vih": "1.8",
        "vil": "0.0",
        "voh": "1.8",
        "vol": "0.0",
        "vt": "0.9",
        "drive_edge_ns": "5",
        "compare_edge_ns": "20",
    },
    "I2C": {
        "period_ns": "2500",  # 400 kHz default
        "format": "NRZ",
        "vih": "3.3",
        "vil": "0.0",
        "voh": "3.3",
        "vol": "0.0",
        "vt": "1.65",
        "drive_edge_ns": "300",
        "compare_edge_ns": "1250",
    },
    "UART": {
        "period_ns": "8680",  # 115200 baud default bit period
        "format": "NRZ",
        "vih": "3.3",
        "vil": "0.0",
        "voh": "3.3",
        "vol": "0.0",
        "vt": "1.65",
        "drive_edge_ns": "500",
        "compare_edge_ns": "4340",
    },
    "JTAG": {
        "period_ns": "100",  # 10 MHz default
        "format": "NRZ",
        "vih": "1.8",
        "vil": "0.0",
        "voh": "1.8",
        "vol": "0.0",
        "vt": "0.9",
        "drive_edge_ns": "10",
        "compare_edge_ns": "50",
    },
}

_DEFAULT_EXPORTS_DIR = Path(r"C:\AI Projects\ultraflex test program\ADIN1300_CHAR_1\Exports")


def _read_text_with_fallback(path: Path) -> str:
    raw = path.read_bytes()
    for enc in ("utf-8", "utf-16", "utf-16-le", "utf-16-be", "latin-1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def _load_datasheet_text(datasheet_text: str, datasheet_path: str) -> tuple[str, list[str]]:
    warnings: list[str] = []
    direct_text = (datasheet_text or "").strip()
    if direct_text:
        return direct_text, warnings

    path_raw = (datasheet_path or "").strip()
    if not path_raw:
        warnings.append("No datasheet_text or datasheet_path provided; generation will use generic defaults.")
        return "", warnings

    path = Path(path_raw).expanduser()
    if not path.exists() or not path.is_file():
        warnings.append(f"datasheet_path not found: {path}")
        return "", warnings

    try:
        return _read_text_with_fallback(path), warnings
    except Exception as exc:
        warnings.append(f"Failed to read datasheet_path ({path}): {exc}")
        return "", warnings


def _parse_exports_structure(exports_dir: str) -> dict[str, Any]:
    target = Path(exports_dir).expanduser()
    if not target.exists() or not target.is_dir():
        return {"error": f"Exports directory not found: {target}"}

    files = sorted([p for p in target.iterdir() if p.is_file()], key=lambda p: p.name.lower())
    by_ext: Counter[str] = Counter(p.suffix.lower() or "<no_ext>" for p in files)

    key_files = {
        "pinmap": next((f.name for f in files if f.name.lower() == "pinmap.txt"), ""),
        "flow_table": next((f.name for f in files if f.name.lower() == "flow table.txt"), ""),
        "test_instances": [f.name for f in files if f.name.lower().startswith("test_instances")],
        "vbt_modules": [f.name for f in files if f.suffix.lower() == ".bas"],
        "class_modules": [f.name for f in files if f.suffix.lower() == ".cls"],
    }

    return {
        "exports_dir": str(target),
        "file_count": len(files),
        "extensions": dict(sorted(by_ext.items())),
        "key_files": key_files,
        "files": [f.name for f in files],
    }


def _extract_template_tests(exports_dir: Path) -> list[str]:
    names: set[str] = set()
    for pattern in ("Test_Instances*.txt", "Flow Table*.txt"):
        for path in exports_dir.glob(pattern):
            try:
                text = _read_text_with_fallback(path)
            except Exception:
                continue
            for line in text.splitlines():
                if not line.strip() or line.startswith("DT"):
                    continue
                cols = [c.strip() for c in line.split("\t")]
                if not cols:
                    continue
                token = next((c for c in cols if c), "")
                if not token or token.lower() in {
                    "label",
                    "test name",
                    "group name",
                    "test procedure",
                    "gate",
                    "enable",
                }:
                    continue
                if re.match(r"^[A-Za-z_][A-Za-z0-9_]{2,}$", token):
                    names.add(token)
    return sorted(names)


def _datasheet_agent(datasheet_text: str) -> dict[str, Any]:
    text = datasheet_text or ""
    tl = text.lower()

    interfaces: list[str] = []
    for iface in ("mdio", "rgmii", "rmii", "mii", "jtag", "spi", "i2c", "uart"):
        if iface in tl:
            interfaces.append(iface.upper())

    supply_matches = sorted(
        {
            m.group(0)
            for m in re.finditer(
                r"\b(?:[AV]?DD(?:IO)?\w*|VDD\w*|AVDD\w*)\s*(?:=|:)?\s*([0-9]+(?:\.[0-9]+)?)\s*V\b",
                text,
                flags=re.IGNORECASE,
            )
        }
    )

    recommended: list[str] = [
        "Continuity_Pos",
        "Continuity_Neg",
        "Powersupply_Shorts",
        "Leakage_IIL_Hi",
        "Leakage_IIH_Hi",
    ]
    if "MDIO" in interfaces:
        recommended += ["MDIO_Read", "ABIST"]
    if "RGMII" in interfaces or "RMII" in interfaces or "MII" in interfaces:
        recommended += ["Clock_Frequency", "MII_TXCLK_Freq"]
    if "JTAG" in interfaces:
        recommended += ["VIH_level", "VIL_level"]

    return {
        "interfaces": sorted(set(interfaces)),
        "supplies": supply_matches,
        "recommended_tests": sorted(set(recommended)),
    }


def _channel_agent(rows: list[dict[str, str]]) -> dict[str, Any]:
    by_direction: Counter[str] = Counter(r["direction"] for r in rows)
    digital_pins = [r["pin_name"] for r in rows if r["direction"] in {"IO", "IN", "OUT", "I/O"}]
    power_pins = [r["pin_name"] for r in rows if "PWR" in r["direction"] or "POWER" in r["direction"]]

    pin_groups = {
        "all_pins": sorted({r["pin_name"] for r in rows}),
        "digital_pins": sorted(set(digital_pins)),
        "power_pins": sorted(set(power_pins)),
    }
    return {
        "direction_counts": dict(sorted(by_direction.items())),
        "pin_groups": pin_groups,
        "site_count": len({r["site"] for r in rows}),
    }


def _planning_agent(template_tests: list[str], datasheet_hints: dict[str, Any]) -> dict[str, Any]:
    wanted = set(datasheet_hints.get("recommended_tests", []))
    available = set(template_tests)

    selected = sorted(wanted & available)
    if not selected and available:
        preferred = [t for t in template_tests if t in wanted]
        selected = preferred or template_tests[:20]
    if not selected:
        selected = sorted(wanted) if wanted else ["Continuity_Pos", "Continuity_Neg", "MDIO_Read"]

    return {
        "selected_tests": selected,
        "template_test_count": len(template_tests),
        "selected_count": len(selected),
    }


def _render_flow_table_seed(selected_tests: list[str]) -> str:
    lines = [
        "DTFlowtableSheet,version=2.3:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\tFlow Table",
        "\t\t\t\tGate\t\t\t\tCommand\t\t\t\t\tComment",
        "\tLabel\tEnable\tJob\tPart\tEnv\tOpcode\tParameter\tTName\tComment",
    ]
    lines.append("\t\t\t\t\tset-error-bin\t\t\tAuto-generated A2A seed")
    for test_name in selected_tests:
        lines.append(f"\t\t\t\t\tTest\t{test_name}\t\t")
    return "\n".join(lines) + "\n"


def _render_test_instances_seed(selected_tests: list[str], protocol_hint: str) -> str:
    lines = [
        "DTTestInstancesSheet,version=2.4:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\tTest Instances",
        "\tTest Name\tType\tName\tTime Sets\tEdge Sets\tPin Levels\tComment",
    ]
    for test_name in selected_tests:
        lines.append(f"\t{test_name}\tVBT\t{test_name}\t{protocol_hint}\tTyp\tLVL_DEFAULT\tAuto-generated")
    return "\n".join(lines) + "\n"


def _render_vbt_module(project_name: str, selected_tests: list[str]) -> str:
    lines = [
        'Attribute VB_Name = "VBT_AutoTests"',
        "",
        "' Auto-generated by ultraflex-testprog A2A pipeline",
        f"' Project: {project_name}",
        "",
    ]
    for test_name in selected_tests:
        lines += [
            f"Public Function {test_name}()",
            "    On Error GoTo errHandler",
            "    Dim Result As New SiteLong",
            "",
            "    ' TODO: Add instrument setup and pattern execution logic.",
            "    TheExec.Flow.TestLimit Resultval:=0, unit:=unitNone, ScaleType:=scaleNone, ForceResults:=tlForceFlow",
            "    Exit Function",
            "",
            "errHandler:",
            "    If AbortTest Then Exit Function Else Resume Next",
            "End Function",
            "",
        ]
    return "\n".join(lines)


def _render_generation_summary(
    project_name: str,
    exports_info: dict[str, Any],
    datasheet_hints: dict[str, Any],
    channel_info: dict[str, Any],
    plan: dict[str, Any],
    warnings: list[str],
) -> str:
    out = [
        f"# {project_name} A2A Generation Summary",
        "",
        "## Inputs",
        "",
        f"- Exports directory: {exports_info.get('exports_dir', 'N/A')}",
        f"- Export files discovered: {exports_info.get('file_count', 0)}",
        f"- Sites detected: {channel_info.get('site_count', 0)}",
        "",
        "## Datasheet Hints",
        "",
        f"- Interfaces: {', '.join(datasheet_hints.get('interfaces', [])) or 'None detected'}",
        f"- Supplies parsed: {', '.join(datasheet_hints.get('supplies', [])) or 'None parsed'}",
        "",
        "## Planned Tests",
        "",
    ]
    for t in plan.get("selected_tests", []):
        out.append(f"- {t}")

    out += ["", "## Warnings", ""]
    if warnings:
        out.extend(f"- {w}" for w in warnings)
    else:
        out.append("- None")
    return "\n".join(out) + "\n"


def _write_a2a_outputs(
    output_dir: str,
    project_name: str,
    flow_table_txt: str,
    test_instances_txt: str,
    vbt_bas: str,
    summary_md: str,
) -> dict[str, str]:
    target = Path(output_dir).expanduser().resolve()
    target.mkdir(parents=True, exist_ok=True)

    slug = _slug(project_name)
    flow_path = target / f"{slug}_Flow_Table_Auto.txt"
    ti_path = target / f"{slug}_Test_Instances_Auto.txt"
    vbt_path = target / f"{slug}_VBT_AutoTests.bas"
    summary_path = target / f"{slug}_A2A_Generation_Summary.md"

    flow_path.write_text(flow_table_txt, encoding="utf-8")
    ti_path.write_text(test_instances_txt, encoding="utf-8")
    vbt_path.write_text(vbt_bas, encoding="utf-8")
    summary_path.write_text(summary_md, encoding="utf-8")

    return {
        "flow_table": str(flow_path),
        "test_instances": str(ti_path),
        "vbt_module": str(vbt_path),
        "summary": str(summary_path),
    }


def _slug(value: str) -> str:
    return re.sub(r"[^a-zA-Z0-9._-]+", "_", value.strip()) or "unnamed"


def _normalize_key(key: str) -> str:
    return _KEY_ALIASES.get(key.strip().lower(), key.strip().lower())


def _normalize_record(record: dict[str, Any], default_instrument: str) -> dict[str, str]:
    out: dict[str, str] = {k: "" for k in _DEFAULT_COLUMNS}
    for raw_key, raw_value in record.items():
        norm_key = _normalize_key(str(raw_key))
        if norm_key in out:
            out[norm_key] = "" if raw_value is None else str(raw_value).strip()

    out["site"] = out["site"] or "1"
    out["instrument"] = out["instrument"] or default_instrument
    out["direction"] = (out["direction"] or "IO").upper()
    return out


def _parse_json_assignments(raw_text: str, default_instrument: str) -> tuple[list[dict[str, str]], list[str]]:
    warnings: list[str] = []
    payload = json.loads(raw_text)

    rows: list[Any]
    if isinstance(payload, list):
        rows = payload
    elif isinstance(payload, dict):
        if isinstance(payload.get("assignments"), list):
            rows = payload["assignments"]
        elif isinstance(payload.get("channels"), list):
            rows = payload["channels"]
        else:
            rows = [payload]
    else:
        raise ValueError("Unsupported JSON structure for channel assignment.")

    normalized: list[dict[str, str]] = []
    for idx, row in enumerate(rows, start=1):
        if not isinstance(row, dict):
            warnings.append(f"Row {idx}: skipped non-object entry.")
            continue
        normalized.append(_normalize_record(row, default_instrument))
    return normalized, warnings


def _token_line_to_record(line: str, default_instrument: str) -> dict[str, str]:
    # Example: PIN=RESET_N CH=CH12 SITE=1 DIR=INSTR
    parts = re.split(r"\s+", line.strip())
    row: dict[str, Any] = {}
    for part in parts:
        if "=" not in part:
            continue
        k, v = part.split("=", 1)
        row[_normalize_key(k)] = v
    return _normalize_record(row, default_instrument)


def _parse_delimited_assignments(raw_text: str, default_instrument: str) -> tuple[list[dict[str, str]], list[str]]:
    warnings: list[str] = []
    lines = [ln for ln in raw_text.splitlines() if ln.strip() and not ln.strip().startswith("#")]
    if not lines:
        return [], ["No channel assignment lines were found."]

    first = lines[0]
    if "=" in first and "," not in first and "\t" not in first:
        rows = [_token_line_to_record(ln, default_instrument) for ln in lines]
        return rows, warnings

    delimiter = ","
    if "\t" in first:
        delimiter = "\t"
    elif "|" in first:
        delimiter = "|"

    has_header = any(tok in first.lower() for tok in ("pin", "channel", "signal", "net"))
    text_buffer = io.StringIO("\n".join(lines))

    rows: list[dict[str, str]] = []
    if has_header:
        reader = csv.DictReader(text_buffer, delimiter=delimiter)
        for idx, row in enumerate(reader, start=2):
            if not row:
                continue
            rows.append(_normalize_record(row, default_instrument))
            if not rows[-1]["pin_name"] or not rows[-1]["channel"]:
                warnings.append(f"Line {idx}: missing pin_name or channel.")
    else:
        reader2 = csv.reader(text_buffer, delimiter=delimiter)
        for idx, row in enumerate(reader2, start=1):
            if not row:
                continue
            pin_name = row[0].strip() if len(row) > 0 else ""
            channel = row[1].strip() if len(row) > 1 else ""
            site = row[2].strip() if len(row) > 2 else "1"
            direction = row[3].strip().upper() if len(row) > 3 else "IO"
            instrument = row[4].strip() if len(row) > 4 else default_instrument
            comment = row[5].strip() if len(row) > 5 else ""
            rows.append(
                {
                    "pin_name": pin_name,
                    "channel": channel,
                    "site": site or "1",
                    "direction": direction or "IO",
                    "instrument": instrument or default_instrument,
                    "comment": comment,
                }
            )
            if not pin_name or not channel:
                warnings.append(f"Line {idx}: missing pin_name or channel.")

    return rows, warnings


def _parse_assignments(raw_text: str, default_instrument: str) -> tuple[list[dict[str, str]], list[str]]:
    text = raw_text.strip()
    if not text:
        return [], ["Channel assignment input is empty."]

    if text.startswith("{") or text.startswith("["):
        try:
            return _parse_json_assignments(text, default_instrument)
        except Exception as exc:
            return [], [f"JSON parse failed: {exc}"]

    return _parse_delimited_assignments(text, default_instrument)


def _validate_assignments(rows: list[dict[str, str]]) -> list[str]:
    warnings: list[str] = []
    if not rows:
        warnings.append("No valid assignment rows parsed.")
        return warnings

    by_site_channel: dict[tuple[str, str], list[str]] = defaultdict(list)
    pin_counter: Counter[str] = Counter()

    for row in rows:
        pin = row["pin_name"]
        ch = row["channel"]
        site = row["site"]
        if not pin or not ch:
            continue
        by_site_channel[(site, ch)].append(pin)
        pin_counter[pin] += 1

    for (site, ch), pins in by_site_channel.items():
        if len(pins) > 1:
            warnings.append(f"Conflict: site {site} channel {ch} assigned to multiple pins: {', '.join(pins)}")

    for pin, count in pin_counter.items():
        if count > 1:
            warnings.append(f"Pin {pin} appears {count} times.")

    return warnings


def _to_csv(rows: list[dict[str, str]], columns: list[str]) -> str:
    buff = io.StringIO()
    writer = csv.DictWriter(buff, fieldnames=columns)
    writer.writeheader()
    for row in rows:
        writer.writerow({col: row.get(col, "") for col in columns})
    return buff.getvalue()


def _build_channel_map_rows(rows: list[dict[str, str]]) -> list[dict[str, str]]:
    out = []
    grouped: dict[tuple[str, str], list[dict[str, str]]] = defaultdict(list)
    for row in rows:
        grouped[(row["site"], row["instrument"])].append(row)

    for (site, instrument), items in sorted(grouped.items(), key=lambda x: (x[0][0], x[0][1])):
        for row in sorted(items, key=lambda r: r["channel"]):
            out.append(
                {
                    "site": site,
                    "instrument": instrument,
                    "channel": row["channel"],
                    "pin_name": row["pin_name"],
                    "direction": row["direction"],
                    "comment": row["comment"],
                }
            )
    return out


def _build_limits_rows(rows: list[dict[str, str]], default_test_name: str = "DC_SANITY") -> list[dict[str, str]]:
    out: list[dict[str, str]] = []
    seen: set[tuple[str, str]] = set()

    for row in sorted(rows, key=lambda r: (r["site"], r["pin_name"], r["channel"])):
        key = (row["site"], row["pin_name"])
        if key in seen:
            continue
        seen.add(key)
        out.append(
            {
                "site": row["site"],
                "pin_name": row["pin_name"],
                "test_name": default_test_name,
                "lo_limit": "",
                "hi_limit": "",
                "unit": "",
                "bin": "",
                "comment": "seed row - fill limits",
            }
        )
    return out


def _build_levels_rows(rows: list[dict[str, str]], default_levelset_name: str = "LVL_DEFAULT") -> list[dict[str, str]]:
    out: list[dict[str, str]] = []
    seen: set[tuple[str, str, str]] = set()

    for row in sorted(rows, key=lambda r: (r["site"], r["instrument"], r["pin_name"])):
        key = (row["site"], row["instrument"], row["pin_name"])
        if key in seen:
            continue
        seen.add(key)
        out.append(
            {
                "site": row["site"],
                "pin_name": row["pin_name"],
                "instrument": row["instrument"],
                "levelset_name": default_levelset_name,
                "vih": "",
                "vil": "",
                "voh": "",
                "vol": "",
                "vt": "",
                "comment": "seed row - fill level values",
            }
        )
    return out


def _build_timing_rows(
    rows: list[dict[str, str]],
    default_timingset_name: str = "TIM_DEFAULT",
    default_period_ns: str = "100",
) -> list[dict[str, str]]:
    out: list[dict[str, str]] = []
    seen: set[tuple[str, str]] = set()

    for row in sorted(rows, key=lambda r: (r["site"], r["pin_name"], r["channel"])):
        key = (row["site"], row["pin_name"])
        if key in seen:
            continue
        seen.add(key)
        out.append(
            {
                "site": row["site"],
                "pin_name": row["pin_name"],
                "timingset_name": default_timingset_name,
                "period_ns": default_period_ns,
                "drive_edge_ns": "",
                "compare_edge_ns": "",
                "format": "NRZ",
                "comment": "seed row - fill timing edges",
            }
        )
    return out


def _normalize_protocol_preset(protocol_preset: str) -> str:
    return (protocol_preset or "").strip().upper()


def _infer_pin_role(pin_name: str) -> str:
    p = pin_name.strip().lower()
    if any(tok in p for tok in ("sclk", "clk", "scl", "tck")):
        return "clock"
    if any(tok in p for tok in ("mosi", "tx", "tdi", "tms")):
        return "drive_data"
    if any(tok in p for tok in ("miso", "rx", "tdo")):
        return "capture_data"
    if any(tok in p for tok in ("sda",)):
        return "bidir_data"
    if any(tok in p for tok in ("cs", "ncs", "ss", "reset", "rst", "trst")):
        return "control"
    return "generic"


def _apply_protocol_preset(
    levels_rows: list[dict[str, str]],
    timing_rows: list[dict[str, str]],
    protocol_preset: str,
) -> list[str]:
    notes: list[str] = []
    preset_name = _normalize_protocol_preset(protocol_preset)
    if not preset_name or preset_name == "NONE":
        return notes

    preset = _PROTOCOL_PRESETS.get(preset_name)
    if not preset:
        notes.append(f"Unknown protocol_preset={protocol_preset}; supported: SPI, I2C, UART, JTAG.")
        return notes

    for row in levels_rows:
        row["VIH"] = preset["vih"]
        row["VIL"] = preset["vil"]
        row["VOH"] = preset["voh"]
        row["VOL"] = preset["vol"]
        row["VT"] = preset["vt"]
        existing = row.get("Comment", "")
        suffix = f"preset={preset_name}"
        row["Comment"] = f"{existing}; {suffix}".strip("; ") if existing else suffix

    for row in timing_rows:
        role = _infer_pin_role(row.get("PinName", ""))
        row["Format"] = preset["format"]
        row["Period_ns"] = preset["period_ns"]

        # Role-driven edge placement defaults.
        if role == "clock":
            row["DriveEdge_ns"] = preset["drive_edge_ns"]
            row["CompareEdge_ns"] = preset["compare_edge_ns"]
        elif role == "capture_data":
            row["DriveEdge_ns"] = ""
            row["CompareEdge_ns"] = preset["compare_edge_ns"]
        elif role in {"drive_data", "control"}:
            row["DriveEdge_ns"] = preset["drive_edge_ns"]
            row["CompareEdge_ns"] = ""
        else:
            row["DriveEdge_ns"] = preset["drive_edge_ns"]
            row["CompareEdge_ns"] = preset["compare_edge_ns"]

        existing = row.get("Comment", "")
        suffix = f"preset={preset_name}; role={role}"
        row["Comment"] = f"{existing}; {suffix}".strip("; ") if existing else suffix

    notes.append(f"Applied protocol preset {preset_name} to levels and timing seeds.")
    return notes


def _build_datasheet_markdown(project_name: str, rows: list[dict[str, str]], warnings: list[str]) -> str:
    by_instrument: Counter[str] = Counter(row["instrument"] for row in rows)
    by_direction: Counter[str] = Counter(row["direction"] for row in rows)
    sites = sorted({row["site"] for row in rows})

    lines = [
        f"# {project_name} Pin and Channel Datasheet",
        "",
        "## Summary",
        "",
        f"- Generated at: {datetime.utcnow().isoformat()}Z",
        f"- Total assignments: {len(rows)}",
        f"- Sites: {', '.join(sites) if sites else 'N/A'}",
        f"- Instruments: {', '.join(sorted(by_instrument.keys())) if by_instrument else 'N/A'}",
        "",
        "## Direction Breakdown",
        "",
        "| Direction | Count |",
        "|---|---:|",
    ]

    for direction, count in sorted(by_direction.items()):
        lines.append(f"| {direction} | {count} |")

    lines += [
        "",
        "## Instrument Breakdown",
        "",
        "| Instrument | Count |",
        "|---|---:|",
    ]
    for instrument, count in sorted(by_instrument.items()):
        lines.append(f"| {instrument} | {count} |")

    lines += [
        "",
        "## Assignment Table",
        "",
        "| Site | Instrument | Channel | Pin | Direction | Comment |",
        "|---|---|---|---|---|---|",
    ]

    for row in _build_channel_map_rows(rows):
        lines.append(
            f"| {row['site']} | {row['instrument']} | {row['channel']} | {row['pin_name']} | {row['direction']} | {row['comment']} |"
        )

    lines += ["", "## Validation Notes", ""]
    if warnings:
        for note in warnings:
            lines.append(f"- WARNING: {note}")
    else:
        lines.append("- No structural conflicts detected by the parser.")

    lines += [
        "",
        "## Next Steps",
        "",
        "- Add level-set definitions for each pin group.",
        "- Add timing-set associations per interface/bus.",
        "- Link each pin group to test methods and limits tables.",
    ]
    return "\n".join(lines)


def _write_outputs(output_dir: str, project_name: str, pinmap_csv: str, channelmap_csv: str, datasheet_md: str) -> dict[str, str]:
    target = Path(output_dir).expanduser().resolve()
    target.mkdir(parents=True, exist_ok=True)

    slug = _slug(project_name)
    pinmap_path = target / f"{slug}_pinmap.csv"
    channelmap_path = target / f"{slug}_channelmap.csv"
    datasheet_path = target / f"{slug}_datasheet.md"

    pinmap_path.write_text(pinmap_csv, encoding="utf-8")
    channelmap_path.write_text(channelmap_csv, encoding="utf-8")
    datasheet_path.write_text(datasheet_md, encoding="utf-8")

    return {
        "pinmap": str(pinmap_path),
        "channelmap": str(channelmap_path),
        "datasheet": str(datasheet_path),
    }


def _write_seed_outputs(
    output_dir: str,
    project_name: str,
    pinmap_csv: str,
    channelmap_csv: str,
    limits_csv: str,
    levels_csv: str,
    timing_csv: str,
) -> dict[str, str]:
    target = Path(output_dir).expanduser().resolve()
    target.mkdir(parents=True, exist_ok=True)

    slug = _slug(project_name)
    pinmap_path = target / f"{slug}_igxl_pinmap.csv"
    channelmap_path = target / f"{slug}_igxl_channelmap.csv"
    limits_path = target / f"{slug}_igxl_limits_seed.csv"
    levels_path = target / f"{slug}_igxl_levels_seed.csv"
    timing_path = target / f"{slug}_igxl_timing_seed.csv"

    pinmap_path.write_text(pinmap_csv, encoding="utf-8")
    channelmap_path.write_text(channelmap_csv, encoding="utf-8")
    limits_path.write_text(limits_csv, encoding="utf-8")
    levels_path.write_text(levels_csv, encoding="utf-8")
    timing_path.write_text(timing_csv, encoding="utf-8")

    return {
        "pinmap": str(pinmap_path),
        "channelmap": str(channelmap_path),
        "limits_seed": str(limits_path),
        "levels_seed": str(levels_path),
        "timing_seed": str(timing_path),
    }


@mcp.tool()
def build_pinmap_package(
    channel_assignment: str,
    project_name: str = "unnamed_dut",
    default_instrument: str = "UltraPin1600",
    write_files: bool = False,
    output_dir: str = "",
) -> str:
    """
    Build pin map, channel map, and datasheet from user-provided channel assignment.

    Inputs:
    - channel_assignment: JSON array/object or text/CSV lines.
    - project_name: Used in generated headers and filenames.
    - default_instrument: Fallback card name if not provided per-row.
    - write_files: If true, writes artifacts to output_dir.
    - output_dir: Target directory. Defaults to page_indexing_RAG/data/generated.
    """
    rows, parse_warnings = _parse_assignments(channel_assignment, default_instrument)
    rows = [r for r in rows if r.get("pin_name") and r.get("channel")]
    validate_warnings = _validate_assignments(rows)
    warnings = [*parse_warnings, *validate_warnings]

    pinmap_csv = _to_csv(rows, ["pin_name", "site", "channel", "instrument", "direction", "comment"])
    channel_rows = _build_channel_map_rows(rows)
    channelmap_csv = _to_csv(channel_rows, ["site", "instrument", "channel", "pin_name", "direction", "comment"])
    datasheet_md = _build_datasheet_markdown(project_name, rows, warnings)

    result: dict[str, Any] = {
        "project_name": project_name,
        "summary": {
            "rows": len(rows),
            "sites": sorted({r["site"] for r in rows}),
            "instruments": sorted({r["instrument"] for r in rows}),
            "warnings": len(warnings),
        },
        "warnings": warnings,
        "artifacts": {
            "pinmap_csv": pinmap_csv,
            "channelmap_csv": channelmap_csv,
            "datasheet_md": datasheet_md,
        },
        "assumptions": [
            "When site is missing, site=1 is assumed.",
            f"When instrument is missing, instrument={default_instrument} is assumed.",
            "Direction defaults to IO if missing.",
        ],
    }

    if write_files:
        target_dir = output_dir.strip() or str(PROJECT_ROOT / "data" / "generated" / _slug(project_name))
        result["written_files"] = _write_outputs(
            target_dir,
            project_name,
            pinmap_csv,
            channelmap_csv,
            datasheet_md,
        )

    return json.dumps(result, indent=2)


@mcp.tool()
def generate_igxl_seed_package(
    channel_assignment: str,
    project_name: str = "unnamed_dut",
    default_instrument: str = "UltraPin1600",
    default_test_name: str = "DC_SANITY",
    default_levelset_name: str = "LVL_DEFAULT",
    default_timingset_name: str = "TIM_DEFAULT",
    default_period_ns: str = "100",
    protocol_preset: str = "",
    write_files: bool = True,
    output_dir: str = "",
) -> str:
    """
    Generate IG-XL seed sheets (PinMap, ChannelMap, Limits, Levels, Timing) from channel assignment.

    Inputs:
    - channel_assignment: JSON array/object or text/CSV lines.
    - project_name: Used in generated headers and filenames.
    - default_instrument: Fallback card name when per-row instrument is missing.
    - default_test_name: Seed test name used in limits rows.
    - default_levelset_name: Seed level set name used in levels rows.
    - default_timingset_name: Seed timing set name used in timing rows.
    - default_period_ns: Seed period in ns used in timing rows.
    - protocol_preset: Optional protocol defaults: SPI|I2C|UART|JTAG.
    - write_files: If true, writes seed CSV files to output_dir.
    - output_dir: Target directory. Defaults to page_indexing_RAG/data/generated/<project>/igxl_seed.
    """
    rows, parse_warnings = _parse_assignments(channel_assignment, default_instrument)
    rows = [r for r in rows if r.get("pin_name") and r.get("channel")]
    validate_warnings = _validate_assignments(rows)
    warnings = [*parse_warnings, *validate_warnings]

    pinmap_rows = [
        {
            "PinName": r["pin_name"],
            "Site": r["site"],
            "Channel": r["channel"],
            "Instrument": r["instrument"],
            "Direction": r["direction"],
            "Comment": r["comment"],
        }
        for r in rows
    ]
    channel_rows = [
        {
            "Site": r["site"],
            "Instrument": r["instrument"],
            "Channel": r["channel"],
            "PinName": r["pin_name"],
            "Direction": r["direction"],
            "Comment": r["comment"],
        }
        for r in _build_channel_map_rows(rows)
    ]
    limits_rows = [
        {
            "Site": r["site"],
            "PinName": r["pin_name"],
            "TestName": r["test_name"],
            "LoLimit": r["lo_limit"],
            "HiLimit": r["hi_limit"],
            "Unit": r["unit"],
            "Bin": r["bin"],
            "Comment": r["comment"],
        }
        for r in _build_limits_rows(rows, default_test_name=default_test_name)
    ]
    levels_rows = [
        {
            "Site": r["site"],
            "PinName": r["pin_name"],
            "Instrument": r["instrument"],
            "LevelSetName": r["levelset_name"],
            "VIH": r["vih"],
            "VIL": r["vil"],
            "VOH": r["voh"],
            "VOL": r["vol"],
            "VT": r["vt"],
            "Comment": r["comment"],
        }
        for r in _build_levels_rows(rows, default_levelset_name=default_levelset_name)
    ]
    timing_rows = [
        {
            "Site": r["site"],
            "PinName": r["pin_name"],
            "TimingSetName": r["timingset_name"],
            "Period_ns": r["period_ns"],
            "DriveEdge_ns": r["drive_edge_ns"],
            "CompareEdge_ns": r["compare_edge_ns"],
            "Format": r["format"],
            "Comment": r["comment"],
        }
        for r in _build_timing_rows(
            rows,
            default_timingset_name=default_timingset_name,
            default_period_ns=default_period_ns,
        )
    ]

    preset_notes = _apply_protocol_preset(levels_rows, timing_rows, protocol_preset=protocol_preset)
    warnings.extend(preset_notes)

    pinmap_csv = _to_csv(pinmap_rows, ["PinName", "Site", "Channel", "Instrument", "Direction", "Comment"])
    channelmap_csv = _to_csv(channel_rows, ["Site", "Instrument", "Channel", "PinName", "Direction", "Comment"])
    limits_csv = _to_csv(limits_rows, ["Site", "PinName", "TestName", "LoLimit", "HiLimit", "Unit", "Bin", "Comment"])
    levels_csv = _to_csv(
        levels_rows,
        ["Site", "PinName", "Instrument", "LevelSetName", "VIH", "VIL", "VOH", "VOL", "VT", "Comment"],
    )
    timing_csv = _to_csv(
        timing_rows,
        ["Site", "PinName", "TimingSetName", "Period_ns", "DriveEdge_ns", "CompareEdge_ns", "Format", "Comment"],
    )

    result: dict[str, Any] = {
        "project_name": project_name,
        "summary": {
            "assignment_rows": len(rows),
            "pinmap_rows": len(pinmap_rows),
            "channelmap_rows": len(channel_rows),
            "limits_rows": len(limits_rows),
            "levels_rows": len(levels_rows),
            "timing_rows": len(timing_rows),
            "sites": sorted({r["site"] for r in rows}),
            "instruments": sorted({r["instrument"] for r in rows}),
            "warnings": len(warnings),
        },
        "warnings": warnings,
        "artifacts": {
            "pinmap_csv": pinmap_csv,
            "channelmap_csv": channelmap_csv,
            "limits_seed_csv": limits_csv,
            "levels_seed_csv": levels_csv,
            "timing_seed_csv": timing_csv,
        },
        "assumptions": [
            "Template headers are IG-XL-friendly CSV seeds and may require workbook-specific column tuning.",
            f"When instrument is missing, instrument={default_instrument} is assumed.",
            f"Limits seed uses TestName={default_test_name} for all rows.",
            f"Levels seed uses LevelSetName={default_levelset_name} for all rows.",
            f"Timing seed uses TimingSetName={default_timingset_name} and Period_ns={default_period_ns}.",
            "When protocol_preset is provided, timing and level numeric defaults are auto-filled.",
        ],
    }

    if write_files:
        target_dir = output_dir.strip() or str(PROJECT_ROOT / "data" / "generated" / _slug(project_name) / "igxl_seed")
        result["written_files"] = _write_seed_outputs(
            target_dir,
            project_name,
            pinmap_csv,
            channelmap_csv,
            limits_csv,
            levels_csv,
            timing_csv,
        )

    return json.dumps(result, indent=2)


@mcp.tool()
def ping() -> str:
    """Health check tool for MCP connectivity checks."""
    return "ultraflex-testprog MCP server is running"


@mcp.tool()
def analyze_exports_structure(exports_dir: str = "") -> str:
    """
    Inspect an IG-XL Exports directory and return inventory metadata.

    Inputs:
    - exports_dir: Folder with exported IG-XL sheets/modules.
      Defaults to ADIN1300_CHAR_1/Exports reference path.
    """
    target = exports_dir.strip() or str(_DEFAULT_EXPORTS_DIR)
    info = _parse_exports_structure(target)
    return json.dumps(info, indent=2)


@mcp.tool()
def generate_test_program_agent_package(
    channel_assignment: str,
    datasheet_text: str = "",
    datasheet_path: str = "",
    exports_dir: str = "",
    project_name: str = "unnamed_dut",
    default_instrument: str = "UltraPin1600",
    default_test_name: str = "DC_SANITY",
    default_levelset_name: str = "LVL_DEFAULT",
    default_timingset_name: str = "TIM_DEFAULT",
    default_period_ns: str = "100",
    protocol_preset: str = "",
    write_files: bool = True,
    output_dir: str = "",
) -> str:
    """
    Multi-agent test-program generation from channel assignment + datasheet.

    Agent stages:
    1) Datasheet Agent: parse interfaces and supply hints.
    2) Channel Agent: parse assignment and derive pin groups.
    3) Template Agent: inspect Exports structure and extract available tests.
    4) Planning Agent: choose a test subset aligned with templates.
    5) Codegen Agent: emit Flow Table, Test Instances, and VBT module seeds.

    Outputs include IG-XL seed sheets and auto-generated code/text artifacts.
    """
    warnings: list[str] = []

    ds_text, ds_warnings = _load_datasheet_text(datasheet_text, datasheet_path)
    warnings.extend(ds_warnings)

    rows, parse_warnings = _parse_assignments(channel_assignment, default_instrument)
    rows = [r for r in rows if r.get("pin_name") and r.get("channel")]
    validate_warnings = _validate_assignments(rows)
    warnings.extend(parse_warnings)
    warnings.extend(validate_warnings)

    target_exports = Path(exports_dir.strip() or str(_DEFAULT_EXPORTS_DIR)).expanduser()
    exports_info = _parse_exports_structure(str(target_exports))
    if "error" in exports_info:
        warnings.append(exports_info["error"])
        template_tests: list[str] = []
    else:
        template_tests = _extract_template_tests(target_exports)

    datasheet_hints = _datasheet_agent(ds_text)
    channel_info = _channel_agent(rows)
    plan = _planning_agent(template_tests, datasheet_hints)
    selected_tests = plan["selected_tests"]

    pinmap_rows = [
        {
            "PinName": r["pin_name"],
            "Site": r["site"],
            "Channel": r["channel"],
            "Instrument": r["instrument"],
            "Direction": r["direction"],
            "Comment": r["comment"],
        }
        for r in rows
    ]
    channel_rows = [
        {
            "Site": r["site"],
            "Instrument": r["instrument"],
            "Channel": r["channel"],
            "PinName": r["pin_name"],
            "Direction": r["direction"],
            "Comment": r["comment"],
        }
        for r in _build_channel_map_rows(rows)
    ]
    limits_rows = [
        {
            "Site": r["site"],
            "PinName": r["pin_name"],
            "TestName": r["test_name"],
            "LoLimit": r["lo_limit"],
            "HiLimit": r["hi_limit"],
            "Unit": r["unit"],
            "Bin": r["bin"],
            "Comment": r["comment"],
        }
        for r in _build_limits_rows(rows, default_test_name=default_test_name)
    ]
    levels_rows = [
        {
            "Site": r["site"],
            "PinName": r["pin_name"],
            "Instrument": r["instrument"],
            "LevelSetName": r["levelset_name"],
            "VIH": r["vih"],
            "VIL": r["vil"],
            "VOH": r["voh"],
            "VOL": r["vol"],
            "VT": r["vt"],
            "Comment": r["comment"],
        }
        for r in _build_levels_rows(rows, default_levelset_name=default_levelset_name)
    ]
    timing_rows = [
        {
            "Site": r["site"],
            "PinName": r["pin_name"],
            "TimingSetName": r["timingset_name"],
            "Period_ns": r["period_ns"],
            "DriveEdge_ns": r["drive_edge_ns"],
            "CompareEdge_ns": r["compare_edge_ns"],
            "Format": r["format"],
            "Comment": r["comment"],
        }
        for r in _build_timing_rows(
            rows,
            default_timingset_name=default_timingset_name,
            default_period_ns=default_period_ns,
        )
    ]
    warnings.extend(_apply_protocol_preset(levels_rows, timing_rows, protocol_preset=protocol_preset))

    pinmap_csv = _to_csv(pinmap_rows, ["PinName", "Site", "Channel", "Instrument", "Direction", "Comment"])
    channelmap_csv = _to_csv(channel_rows, ["Site", "Instrument", "Channel", "PinName", "Direction", "Comment"])
    limits_csv = _to_csv(limits_rows, ["Site", "PinName", "TestName", "LoLimit", "HiLimit", "Unit", "Bin", "Comment"])
    levels_csv = _to_csv(
        levels_rows,
        ["Site", "PinName", "Instrument", "LevelSetName", "VIH", "VIL", "VOH", "VOL", "VT", "Comment"],
    )
    timing_csv = _to_csv(
        timing_rows,
        ["Site", "PinName", "TimingSetName", "Period_ns", "DriveEdge_ns", "CompareEdge_ns", "Format", "Comment"],
    )

    protocol_hint = protocol_preset.strip().upper() or "Typ"
    flow_table_txt = _render_flow_table_seed(selected_tests)
    test_instances_txt = _render_test_instances_seed(selected_tests, protocol_hint=protocol_hint)
    vbt_bas = _render_vbt_module(project_name, selected_tests)
    summary_md = _render_generation_summary(
        project_name,
        exports_info if isinstance(exports_info, dict) else {},
        datasheet_hints,
        channel_info,
        plan,
        warnings,
    )

    result: dict[str, Any] = {
        "project_name": project_name,
        "agent_pipeline": {
            "datasheet_agent": datasheet_hints,
            "channel_agent": channel_info,
            "template_agent": {
                "exports_dir": str(target_exports),
                "template_tests_found": len(template_tests),
            },
            "planning_agent": plan,
        },
        "summary": {
            "assignment_rows": len(rows),
            "selected_tests": len(selected_tests),
            "warnings": len(warnings),
        },
        "warnings": warnings,
        "artifacts": {
            "pinmap_csv": pinmap_csv,
            "channelmap_csv": channelmap_csv,
            "limits_seed_csv": limits_csv,
            "levels_seed_csv": levels_csv,
            "timing_seed_csv": timing_csv,
            "flow_table_txt": flow_table_txt,
            "test_instances_txt": test_instances_txt,
            "vbt_tests_bas": vbt_bas,
            "summary_md": summary_md,
        },
        "assumptions": [
            "Generated files are seed scaffolds aligned to Exports naming style and may need project-specific tuning.",
            "Datasheet parsing is keyword-based unless richer structured datasheet input is provided.",
            "Channel assignment is treated as source-of-truth for pin/channel/site mappings.",
        ],
    }

    if write_files:
        target_dir = output_dir.strip() or str(PROJECT_ROOT / "data" / "generated" / _slug(project_name) / "a2a_seed")
        seed_files = _write_seed_outputs(
            target_dir,
            project_name,
            pinmap_csv,
            channelmap_csv,
            limits_csv,
            levels_csv,
            timing_csv,
        )
        a2a_files = _write_a2a_outputs(
            target_dir,
            project_name,
            flow_table_txt,
            test_instances_txt,
            vbt_bas,
            summary_md,
        )
        result["written_files"] = {**seed_files, **a2a_files}

    return json.dumps(result, indent=2)


if __name__ == "__main__":
    mcp.run()
