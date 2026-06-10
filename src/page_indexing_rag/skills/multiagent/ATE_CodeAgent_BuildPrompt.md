# ATE CODE AGENT — Build Prompt for Claude Agent SDK
## Feed this entire file to your AI coding agent to build the project

---

## WHAT YOU ARE BUILDING

You are building an **ATE Code Agent** — a Cursor AI equivalent for Teradyne UltraFLEX
IG-XL test programs. It works exactly like Claude Code:
- Engineer types a task in natural language
- Agent reads the codebase (Excel + VBA files)
- Agent plans, edits, runs, reads logs, and fixes autonomously
- Agent pauses for human approval before touching files or running hardware

The agent is built using the **Claude Agent SDK** (`claude_agent_sdk` Python package),
which is the same infrastructure that powers Claude Code — you get the agent loop,
tool execution, context management, and MCP integration out of the box.

---

## TECH STACK

- **Python 3.11+**
- **claude_agent_sdk** — Anthropic's Agent SDK (replaces LangGraph for orchestration)
- **openpyxl** — Read/write IG-XL Excel files
- **pywin32** — Read/write VBA modules via Windows COM
- **pystdf** — Parse STDF test result log files
- **GitPython** — Auto-commit backups before edits
- **fastapi + uvicorn** — REST API backend
- **python-dotenv** — Environment variable management

---

## PROJECT STRUCTURE TO CREATE

```
ate-code-agent/
├── main.py                  # Entry point — run the agent
├── agent.py                 # Core Claude Agent SDK setup and query loop
├── system_prompt.py         # ATE-specific system prompt for Claude
├── tools/
│   ├── __init__.py
│   ├── excel_tools.py       # read_excel_sheets, write_excel_cell
│   ├── vba_tools.py         # read_vba_module, write_vba_module
│   ├── run_tools.py         # run_igxl_program (trigger tester)
│   ├── log_tools.py         # parse_stdf_log
│   └── git_tools.py         # git_commit_backup
├── safety/
│   ├── __init__.py
│   └── rules.py             # SAFETY_RULES + validate_plan()
├── api/
│   ├── __init__.py
│   ├── main.py              # FastAPI app
│   └── models.py            # Pydantic models
├── tests/
│   ├── test_excel_tools.py
│   ├── test_vba_tools.py
│   └── test_log_tools.py
├── .env                     # ANTHROPIC_API_KEY
├── requirements.txt
└── README.md
```

---

## STEP 1 — Install Dependencies

Create `requirements.txt`:

```txt
# Claude Agent SDK
claude-agent-sdk>=0.2.0

# Excel + VBA
openpyxl>=3.1.0
pywin32>=306

# ATE log parsing
pystdf>=1.0.0

# Version control
GitPython>=3.1.0

# API
fastapi>=0.110.0
uvicorn>=0.28.0

# Utilities
python-dotenv>=1.0.0
anyio>=4.0.0
```

Install:
```bash
pip install -r requirements.txt
```

---

## STEP 2 — Create `tools/excel_tools.py`

Build these two functions exactly:

```python
"""
Excel tools for reading and writing IG-XL test program Excel workbooks.
Uses openpyxl — does NOT require Excel to be installed.
"""
import openpyxl
from pathlib import Path
from safety.rules import SAFETY_RULES


SHEETS_TO_READ = [
    'Test Instances', 'Test Suites', 'Limits',
    'Flow', 'Levels', 'Timing', 'Pin Map', 'Spec'
]


def read_excel_sheets(file_path: str) -> dict:
    """
    Read all IG-XL Excel sheets into a structured dict.

    Args:
        file_path: Absolute path to the IG-XL .xlsx program file

    Returns:
        Dict of {sheet_name: [list of row dicts keyed by column header]}
        Example: {'Limits': [{'Test Name': 'IDDQ', 'Hi Limit': 100, 'Lo Limit': 0}]}
    """
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f'IG-XL program file not found: {file_path}')

    wb = openpyxl.load_workbook(str(path), data_only=True)
    result = {}

    for sheet_name in SHEETS_TO_READ:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        headers = [cell.value for cell in ws[1]]
        rows = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if any(cell is not None for cell in row):
                rows.append(dict(zip(headers, row)))
        result[sheet_name] = rows

    wb.close()
    return result


def write_excel_cell(
    file_path: str,
    sheet_name: str,
    row: int,
    col: int,
    value: any,
    backup: bool = True
) -> dict:
    """
    Write a value to a specific cell in an IG-XL Excel sheet.

    Safety checks:
    - Raises ValueError if sheet_name is in SAFETY_RULES['readonly_sheets']
    - Raises ValueError if writing to Levels sheet with voltage > max allowed

    Args:
        file_path:   Path to the IG-XL .xlsx file
        sheet_name:  Name of the Excel sheet to edit
        row:         1-based row number
        col:         1-based column number
        value:       New value to write
        backup:      Whether to git-commit a backup before writing (default True)

    Returns:
        Dict: {sheet, row, col, old_value, new_value, status}
    """
    # SAFETY: Block read-only sheets
    if sheet_name in SAFETY_RULES['readonly_sheets']:
        raise ValueError(
            f"SAFETY BLOCK: '{sheet_name}' is read-only. "
            f"Editing this sheet risks hardware damage."
        )

    # SAFETY: Block unsafe voltage values
    if sheet_name == 'Levels' and isinstance(value, (int, float)):
        max_v = SAFETY_RULES['voltage_limits']['VCC_max']
        if value > max_v:
            raise ValueError(
                f"SAFETY BLOCK: Voltage {value}V exceeds maximum {max_v}V. "
                f"This could destroy DUT devices."
            )

    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f'IG-XL file not found: {file_path}')

    # Backup before editing
    if backup and SAFETY_RULES['git_backup_enabled']:
        from tools.git_tools import git_commit_backup
        git_commit_backup(
            file_path,
            f'Pre-edit backup: {sheet_name} row={row} col={col}'
        )

    wb = openpyxl.load_workbook(str(path))
    ws = wb[sheet_name]
    old_value = ws.cell(row=row, column=col).value
    ws.cell(row=row, column=col, value=value)
    wb.save(str(path))
    wb.close()

    return {
        'sheet':     sheet_name,
        'row':       row,
        'col':       col,
        'old_value': old_value,
        'new_value': value,
        'status':    'ok'
    }
```

---

## STEP 3 — Create `tools/vba_tools.py`

```python
"""
VBA tools for reading and writing IG-XL Visual Basic for Applications modules.
Uses pywin32 (win32com) — REQUIRES Windows with Excel installed.
"""
import win32com.client
from pathlib import Path
from tools.git_tools import git_commit_backup
from safety.rules import SAFETY_RULES


def get_vba_module_names(excel_path: str) -> list[str]:
    """
    List all VBA module names in the Excel workbook.

    Args:
        excel_path: Absolute path to the .xlsx file

    Returns:
        List of module name strings
    """
    excel = win32com.client.Dispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(str(Path(excel_path).resolve()))
        names = [comp.Name for comp in wb.VBProject.VBComponents]
        return names
    finally:
        try:
            wb.Close(SaveChanges=False)
        except Exception:
            pass
        excel.Quit()


def read_vba_module(excel_path: str, module_name: str) -> str:
    """
    Read the full source code of a VBA module.

    Args:
        excel_path:  Absolute path to the .xlsx file
        module_name: Name of the VBA module (e.g. 'TestMethods')

    Returns:
        String containing the full VBA source code
    """
    excel = win32com.client.Dispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(str(Path(excel_path).resolve()))
        for component in wb.VBProject.VBComponents:
            if component.Name == module_name:
                line_count = component.CodeModule.CountOfLines
                if line_count == 0:
                    return ''
                return component.CodeModule.Lines(1, line_count)
        raise ValueError(f"VBA module '{module_name}' not found in {excel_path}")
    finally:
        try:
            wb.Close(SaveChanges=False)
        except Exception:
            pass
        excel.Quit()


def write_vba_module(
    excel_path: str,
    module_name: str,
    new_code: str
) -> dict:
    """
    Replace the entire source code of a VBA module with new code.
    Always commits a git backup before modifying.

    Args:
        excel_path:  Absolute path to the .xlsx file
        module_name: Name of the VBA module to overwrite
        new_code:    New VBA source code to write

    Returns:
        Dict: {module, lines_written, status}
    """
    # Backup before editing
    if SAFETY_RULES['git_backup_enabled']:
        git_commit_backup(excel_path, f'VBA backup before edit: {module_name}')

    excel = win32com.client.Dispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(str(Path(excel_path).resolve()))
        component = wb.VBProject.VBComponents(module_name)
        cm = component.CodeModule
        # Clear existing code
        if cm.CountOfLines > 0:
            cm.DeleteLines(1, cm.CountOfLines)
        # Write new code
        if new_code.strip():
            cm.InsertLines(1, new_code)
        wb.Save()
        lines = len(new_code.splitlines())
        return {'module': module_name, 'lines_written': lines, 'status': 'ok'}
    finally:
        try:
            wb.Close(SaveChanges=False)
        except Exception:
            pass
        excel.Quit()
```

---

## STEP 4 — Create `tools/run_tools.py`

```python
"""
Tools for triggering IG-XL test program execution on the UltraFLEX tester.
Uses the IG-XL command line interface (igxl_run.exe).
"""
import subprocess
import os
from pathlib import Path
from safety.rules import SAFETY_RULES


def run_igxl_program(
    program_path: str,
    log_output_path: str,
    timeout_seconds: int = 600,
    mode: str = 'engineering'
) -> dict:
    """
    Execute an IG-XL test program and wait for completion.

    Args:
        program_path:    Path to the IG-XL .xlsx program file
        log_output_path: Path where the STDF log file will be written
        timeout_seconds: Max seconds to wait (default 600 = 10 minutes)
        mode:            Run mode — must be in SAFETY_RULES['allowed_run_modes']

    Returns:
        Dict: {return_code, stdout, log_path, status}
    """
    # SAFETY: Only allow engineering/characterization modes
    if mode not in SAFETY_RULES['allowed_run_modes']:
        raise ValueError(
            f"SAFETY BLOCK: Run mode '{mode}' not allowed. "
            f"Allowed modes: {SAFETY_RULES['allowed_run_modes']}"
        )

    cmd = [
        'igxl_run.exe',
        '--program', str(Path(program_path).resolve()),
        '--output',  str(Path(log_output_path).resolve()),
        '--format',  'stdf',
        '--mode',    mode,
    ]

    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            timeout=timeout_seconds,
            cwd=str(Path(program_path).parent)
        )
        return {
            'return_code': result.returncode,
            'stdout':      result.stdout.decode(errors='replace'),
            'stderr':      result.stderr.decode(errors='replace'),
            'log_path':    log_output_path,
            'status':      'complete' if result.returncode == 0 else 'error'
        }
    except subprocess.TimeoutExpired:
        return {
            'return_code': -1,
            'stdout':      '',
            'stderr':      f'Timed out after {timeout_seconds}s',
            'log_path':    log_output_path,
            'status':      'timeout'
        }
    except FileNotFoundError:
        return {
            'return_code': -1,
            'stdout':      '',
            'stderr':      'igxl_run.exe not found — is IG-XL installed?',
            'log_path':    log_output_path,
            'status':      'error'
        }
```

---

## STEP 5 — Create `tools/log_tools.py`

```python
"""
Tools for parsing IG-XL and STDF test result log files.
"""
from pathlib import Path


def parse_stdf_log(log_path: str) -> dict:
    """
    Parse a STDF binary log file into structured test results.

    Args:
        log_path: Path to the .stdf log file

    Returns:
        Dict: {test_name: {result, measured, limit_lo, limit_hi, units, test_num}}
        result is 'PASS' or 'FAIL'
    """
    path = Path(log_path)
    if not path.exists():
        raise FileNotFoundError(f'STDF log not found: {log_path}')

    results = {}
    try:
        from pystdf.Importer import STDF2DataFrame
        df = STDF2DataFrame(str(path))

        # Parametric Test Records (measured value tests)
        for record in df.get('PTR', []):
            name = record.get('TEST_TXT') or f"test_{record.get('TEST_NUM', 'unknown')}"
            flag = record.get('TEST_FLG', 1)
            results[name] = {
                'result':   'PASS' if flag == 0 else 'FAIL',
                'measured': record.get('RESULT'),
                'limit_lo': record.get('LO_LIMIT'),
                'limit_hi': record.get('HI_LIMIT'),
                'units':    record.get('UNITS', ''),
                'test_num': record.get('TEST_NUM'),
            }

        # Functional Test Records (pass/fail digital tests)
        for record in df.get('FTR', []):
            name = record.get('TEST_TXT') or f"ftest_{record.get('TEST_NUM', 'unknown')}"
            flag = record.get('TEST_FLG', 1)
            results[name] = {
                'result':   'PASS' if flag == 0 else 'FAIL',
                'measured': None,
                'limit_lo': None,
                'limit_hi': None,
                'units':    '',
                'test_num': record.get('TEST_NUM'),
            }

    except Exception as e:
        # Fallback: try parsing as plain text log
        results = _parse_text_log(str(path), str(e))

    return results


def _parse_text_log(log_path: str, stdf_error: str) -> dict:
    """Fallback text log parser if STDF binary parsing fails."""
    results = {
        '__parse_error__': {
            'result':   'FAIL',
            'measured': None,
            'limit_lo': None,
            'limit_hi': None,
            'units':    '',
            'test_num': 0,
            'note':     f'STDF parse failed: {stdf_error}. Raw log at {log_path}',
        }
    }
    try:
        with open(log_path, 'r', errors='replace') as f:
            for line in f:
                line = line.strip()
                if 'PASS' in line or 'FAIL' in line:
                    parts = line.split()
                    if len(parts) >= 2:
                        name   = parts[0]
                        result = 'PASS' if 'PASS' in line else 'FAIL'
                        results[name] = {
                            'result':   result,
                            'measured': None,
                            'limit_lo': None,
                            'limit_hi': None,
                            'units':    '',
                            'test_num': 0,
                        }
    except Exception:
        pass
    return results


def summarize_results(results: dict) -> dict:
    """
    Summarize test results into pass/fail counts.

    Returns:
        Dict: {total, passed, failed, pass_rate, failed_tests, passed_tests}
    """
    passed = [t for t, r in results.items() if r['result'] == 'PASS']
    failed = [t for t, r in results.items() if r['result'] == 'FAIL']
    total  = len(results)
    return {
        'total':        total,
        'passed':       len(passed),
        'failed':       len(failed),
        'pass_rate':    f'{(len(passed)/total*100):.1f}%' if total else '0%',
        'failed_tests': failed,
        'passed_tests': passed,
    }
```

---

## STEP 6 — Create `tools/git_tools.py`

```python
"""
Git backup tools — auto-commit before any file edit.
"""
import git
from pathlib import Path
from datetime import datetime


def git_commit_backup(file_path: str, message: str) -> str:
    """
    Commit the current state of a file to git before editing.
    Creates the repo if it doesn't exist yet.

    Args:
        file_path: Path to the file to back up
        message:   Description of why this backup was made

    Returns:
        Git commit hash string
    """
    path = Path(file_path).resolve()

    try:
        repo = git.Repo(str(path.parent), search_parent_directories=True)
    except git.InvalidGitRepositoryError:
        repo = git.Repo.init(str(path.parent))
        # Create initial .gitignore
        gitignore = path.parent / '.gitignore'
        if not gitignore.exists():
            gitignore.write_text('*.pyc\n__pycache__/\n.env\n')
        repo.index.add(['.gitignore'])

    try:
        repo.index.add([str(path)])
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        commit = repo.index.commit(
            f'[ATE Agent Backup] {timestamp} — {message}'
        )
        return commit.hexsha
    except Exception as e:
        # Non-fatal — log but don't block the edit
        print(f'Warning: Git backup failed: {e}')
        return 'backup-failed'
```

---

## STEP 7 — Create `safety/rules.py`

```python
"""
Safety rules for the ATE Code Agent.
These rules are enforced in tool implementations and CANNOT be overridden
by Claude or the engineer without modifying this file directly.
"""

SAFETY_RULES = {
    # ── Iteration control ────────────────────────────────────────
    'max_iterations': 5,           # Max debug loops before escalating
    'require_human_approval': True, # Always pause before editing

    # ── Read-only sheets — NEVER write to these ──────────────────
    # Pin Map is hardware — wrong values destroy loadboards
    # Spec is reference data — should never be agent-modified
    'readonly_sheets': ['Pin Map', 'Spec'],

    # ── Voltage safety limits (Volts) ────────────────────────────
    # Exceeding these limits can burn DUT devices
    'voltage_limits': {
        'VCC_max':   5.0,
        'VDD_max':   3.6,
        'VDDIO_max': 3.6,
        'VDDA_max':  5.0,
    },

    # ── Run mode whitelist ───────────────────────────────────────
    # Production mode is NEVER allowed — would run wafer-level quantities
    'allowed_run_modes': ['engineering', 'characterization'],

    # ── Git backup ───────────────────────────────────────────────
    'git_backup_enabled': True,
}
```

---

## STEP 8 — Create `system_prompt.py`

This is the ATE-domain system prompt that turns Claude into an expert IG-XL engineer.

```python
"""
ATE-specific system prompt for the Claude Agent SDK.
This tells Claude what kind of agent it is and what rules to follow.
"""

ATE_SYSTEM_PROMPT = """
You are an expert Teradyne UltraFLEX ATE (Automatic Test Equipment) engineer.
You are acting as an AI coding agent — like Cursor AI — but specialized for
IG-XL test programs running on the UltraFLEX tester.

## What You Work With

IG-XL test programs consist of two parts:
1. **Excel workbooks** — define test structure:
   - Test Instances sheet: test names, method classes, parameters
   - Limits sheet: hi/lo pass-fail limits per test (MOST COMMON EDIT TARGET)
   - Flow sheet: test execution sequence and branching logic
   - Levels sheet: supply voltages (VCC, VDD, VDDIO) — SAFETY CRITICAL
   - Timing sheet: clock edge definitions
   - Test Suites sheet: grouping of test instances
   - Pin Map sheet: DUT pin to tester channel mapping — READ ONLY, NEVER EDIT
   - Spec sheet: device specification reference — READ ONLY

2. **VBA modules** — Visual Basic for Applications code containing test method logic.
   These are called by test instances defined in the Excel sheets.

## Your Workflow

When given a task:
1. READ: Use read_excel_sheets() and read_vba_module() to load the full program
2. PLAN: Analyze what needs to change and explain it clearly to the engineer
3. WAIT: Always describe your plan BEFORE making any changes. Wait for approval.
4. EDIT: Use write_excel_cell() or write_vba_module() to apply approved changes
5. RUN: Use run_igxl_program() to execute the test on the tester
6. ANALYZE: Use parse_stdf_log() to read results — list PASS and FAIL tests
7. DEBUG: If failures, diagnose root cause and propose fixes. Loop back to PLAN.

## Safety Rules — NON-NEGOTIABLE

- NEVER edit the Pin Map sheet — hardware definition, wrong values damage loadboards
- NEVER write voltage values above safety maximums (VCC > 5V, VDD > 3.6V)
- ALWAYS describe your full plan before editing ANY file
- ALWAYS run in 'engineering' mode, never 'production'
- STOP after 5 debug iterations and escalate to engineer if unfixed
- Every change you make is git-committed as a backup automatically

## Code Style for VBA

When writing VBA code:
- Use descriptive Sub names matching IG-XL conventions
- Include error handling with On Error GoTo
- Comment each measurement block
- Use meaningful variable names (not i, j, k)
- Follow existing module style when modifying

## Communication Style

- Be specific: say WHICH sheet, WHICH row, WHICH column, WHAT old value, WHAT new value
- Always explain WHY you are making a change
- When tests fail, show the measured value vs the limit side by side
- If confidence is low, say so clearly and ask the engineer for guidance
- Keep responses concise — engineers are busy
"""
```

---

## STEP 9 — Create `agent.py`

This is the CORE of the project — the Claude Agent SDK setup with all custom ATE tools registered.

```python
"""
Core ATE Code Agent using Claude Agent SDK.
This is the equivalent of Cursor AI but for IG-XL test programs.
"""
import anyio
import json
from claude_agent_sdk import (
    tool,
    create_sdk_mcp_server,
    ClaudeAgentOptions,
    ClaudeSDKClient,
    AssistantMessage,
    TextBlock,
    ResultMessage
)
from system_prompt import ATE_SYSTEM_PROMPT
from tools.excel_tools import read_excel_sheets, write_excel_cell
from tools.vba_tools import read_vba_module, write_vba_module, get_vba_module_names
from tools.run_tools import run_igxl_program
from tools.log_tools import parse_stdf_log, summarize_results
from tools.git_tools import git_commit_backup


# ── Define ATE Custom Tools ───────────────────────────────────────────────────
# Each @tool decorator registers a Python function as a tool Claude can call.

@tool(
    name="read_igxl_excel",
    description="Read all IG-XL Excel sheets from the test program workbook. "
                "Returns Test Instances, Limits, Flow, Levels, Timing, Test Suites. "
                "Use this first to understand the current state of the program.",
    input_schema={"file_path": str}
)
async def tool_read_excel(args: dict) -> dict:
    try:
        data = read_excel_sheets(args['file_path'])
        return {"content": [{"type": "text", "text": json.dumps(data, indent=2, default=str)}]}
    except Exception as e:
        return {"content": [{"type": "text", "text": f"ERROR: {str(e)}"}], "isError": True}


@tool(
    name="write_igxl_cell",
    description="Write a value to a specific cell in an IG-XL Excel sheet. "
                "Safety checks: Pin Map is blocked, voltage limits enforced. "
                "Always backs up via git before writing.",
    input_schema={
        "file_path":  str,
        "sheet_name": str,
        "row":        int,
        "col":        int,
        "value":      str,  # Accept as string, convert in handler
    }
)
async def tool_write_excel(args: dict) -> dict:
    try:
        # Try to convert value to number if possible
        value = args['value']
        try:
            value = float(value) if '.' in str(value) else int(value)
        except (ValueError, TypeError):
            pass  # Keep as string

        result = write_excel_cell(
            args['file_path'],
            args['sheet_name'],
            args['row'],
            args['col'],
            value
        )
        return {"content": [{"type": "text", "text": json.dumps(result)}]}
    except Exception as e:
        return {"content": [{"type": "text", "text": f"ERROR: {str(e)}"}], "isError": True}


@tool(
    name="list_vba_modules",
    description="List all VBA module names in the IG-XL Excel workbook.",
    input_schema={"file_path": str}
)
async def tool_list_vba_modules(args: dict) -> dict:
    try:
        names = get_vba_module_names(args['file_path'])
        return {"content": [{"type": "text", "text": json.dumps(names)}]}
    except Exception as e:
        return {"content": [{"type": "text", "text": f"ERROR: {str(e)}"}], "isError": True}


@tool(
    name="read_vba_module",
    description="Read the full VBA source code of a named module. "
                "Use this to understand the current test method logic.",
    input_schema={"file_path": str, "module_name": str}
)
async def tool_read_vba(args: dict) -> dict:
    try:
        code = read_vba_module(args['file_path'], args['module_name'])
        return {"content": [{"type": "text", "text": code}]}
    except Exception as e:
        return {"content": [{"type": "text", "text": f"ERROR: {str(e)}"}], "isError": True}


@tool(
    name="write_vba_module",
    description="Overwrite a VBA module with new source code. "
                "Git-commits a backup first. "
                "WARNING: Replaces the ENTIRE module — provide the full code.",
    input_schema={
        "file_path":   str,
        "module_name": str,
        "new_code":    str
    }
)
async def tool_write_vba(args: dict) -> dict:
    try:
        result = write_vba_module(
            args['file_path'],
            args['module_name'],
            args['new_code']
        )
        return {"content": [{"type": "text", "text": json.dumps(result)}]}
    except Exception as e:
        return {"content": [{"type": "text", "text": f"ERROR: {str(e)}"}], "isError": True}


@tool(
    name="run_igxl_program",
    description="Execute the IG-XL test program on the UltraFLEX tester. "
                "Runs in engineering mode only. "
                "Returns the path to the STDF log file when complete.",
    input_schema={
        "program_path":    str,
        "log_output_path": str,
        "timeout_seconds": int
    }
)
async def tool_run_program(args: dict) -> dict:
    try:
        result = run_igxl_program(
            args['program_path'],
            args['log_output_path'],
            timeout_seconds=args.get('timeout_seconds', 600)
        )
        return {"content": [{"type": "text", "text": json.dumps(result)}]}
    except Exception as e:
        return {"content": [{"type": "text", "text": f"ERROR: {str(e)}"}], "isError": True}


@tool(
    name="parse_test_log",
    description="Parse the STDF test result log file after a run. "
                "Returns pass/fail status, measured values, and limits for every test. "
                "Always call this after run_igxl_program completes.",
    input_schema={"log_path": str}
)
async def tool_parse_log(args: dict) -> dict:
    try:
        results = parse_stdf_log(args['log_path'])
        summary = summarize_results(results)
        output = {
            'summary': summary,
            'results': results
        }
        return {"content": [{"type": "text", "text": json.dumps(output, indent=2, default=str)}]}
    except Exception as e:
        return {"content": [{"type": "text", "text": f"ERROR: {str(e)}"}], "isError": True}


@tool(
    name="git_backup",
    description="Create a git backup commit of the current state of a file. "
                "Useful for manual checkpoints. Edits auto-backup already.",
    input_schema={"file_path": str, "message": str}
)
async def tool_git_backup(args: dict) -> dict:
    try:
        commit_hash = git_commit_backup(args['file_path'], args['message'])
        return {"content": [{"type": "text", "text": f"Backup committed: {commit_hash}"}]}
    except Exception as e:
        return {"content": [{"type": "text", "text": f"ERROR: {str(e)}"}], "isError": True}


# ── Build MCP Server with all ATE tools ──────────────────────────────────────

ate_tool_server = create_sdk_mcp_server(
    name="ate-tools",
    version="1.0.0",
    tools=[
        tool_read_excel,
        tool_write_excel,
        tool_list_vba_modules,
        tool_read_vba,
        tool_write_vba,
        tool_run_program,
        tool_parse_log,
        tool_git_backup,
    ]
)


# ── Agent Options ─────────────────────────────────────────────────────────────

def get_agent_options(program_path: str) -> ClaudeAgentOptions:
    """
    Build ClaudeAgentOptions for an ATE agent session.
    Combines built-in Claude Code tools with our custom ATE tools.
    """
    return ClaudeAgentOptions(
        # ATE domain system prompt
        system_prompt=ATE_SYSTEM_PROMPT,

        # Working directory — set to program folder so relative paths work
        cwd=str(program_path),

        # Our custom ATE MCP tools
        mcp_servers={"ate-tools": ate_tool_server},

        # Auto-approve our ATE tools (they have internal safety checks)
        allowed_tools=[
            "mcp__ate-tools__read_igxl_excel",
            "mcp__ate-tools__list_vba_modules",
            "mcp__ate-tools__read_vba_module",
            "mcp__ate-tools__parse_test_log",
            "mcp__ate-tools__git_backup",
            # Write/run tools are NOT auto-approved — Claude must ask
            # "mcp__ate-tools__write_igxl_cell",
            # "mcp__ate-tools__write_vba_module",
            # "mcp__ate-tools__run_igxl_program",
        ],

        # Built-in Claude tools available
        tools=["Read", "Glob", "Grep"],  # File reading only — no built-in Bash/Write

        # Max turns before stopping
        max_turns=50,

        # Use Claude Sonnet — best balance of speed and code quality
        model="claude-sonnet-4-6",
    )


# ── Interactive Agent Loop ────────────────────────────────────────────────────

async def run_interactive_agent(program_path: str):
    """
    Run an interactive ATE code agent session.
    Engineer types tasks, agent reads/edits/runs the test program.

    This is the equivalent of opening Cursor AI on your IG-XL project.
    """
    print("\n" + "="*60)
    print("  ATE CODE AGENT — UltraFLEX / IG-XL")
    print("  Powered by Claude Agent SDK")
    print("="*60)
    print(f"  Program: {program_path}")
    print("  Type your task. Type 'quit' to exit.")
    print("="*60 + "\n")

    options = get_agent_options(program_path)

    async with ClaudeSDKClient(options=options) as client:
        while True:
            # Get task from engineer
            try:
                task = input("You: ").strip()
            except (EOFError, KeyboardInterrupt):
                print("\nExiting ATE Agent.")
                break

            if task.lower() in ('quit', 'exit', 'q'):
                print("Exiting ATE Agent.")
                break

            if not task:
                continue

            print("\nAgent: ", end="", flush=True)

            # Send task to Claude Agent SDK — it handles the full agentic loop
            async for message in client.query(task):
                if isinstance(message, AssistantMessage):
                    for block in message.content:
                        if isinstance(block, TextBlock):
                            print(block.text, end="", flush=True)
                elif isinstance(message, ResultMessage):
                    if message.is_error:
                        print(f"\n[Agent stopped with error: {message.result}]")

            print("\n")  # New line after response


# ── Single-shot Query (for API use) ──────────────────────────────────────────

async def run_single_query(program_path: str, task: str) -> str:
    """
    Run a single task and return the result as a string.
    Used by the FastAPI backend.
    """
    options = get_agent_options(program_path)
    full_response = []

    async for message in client.query(task, options=options):
        if isinstance(message, AssistantMessage):
            for block in message.content:
                if isinstance(block, TextBlock):
                    full_response.append(block.text)

    return "\n".join(full_response)
```

---

## STEP 10 — Create `main.py`

```python
"""
Entry point for the ATE Code Agent.
Run this file to start an interactive agent session.
"""
import anyio
import sys
from pathlib import Path
from dotenv import load_dotenv

# Load ANTHROPIC_API_KEY from .env
load_dotenv()

from agent import run_interactive_agent


def main():
    if len(sys.argv) < 2:
        print("Usage: python main.py <path-to-igxl-program.xlsx>")
        print("Example: python main.py /programs/MyDevice/MyDevice_TP.xlsx")
        sys.exit(1)

    program_path = Path(sys.argv[1])
    if not program_path.exists():
        print(f"Error: Program file not found: {program_path}")
        sys.exit(1)

    anyio.run(run_interactive_agent, str(program_path.parent))


if __name__ == "__main__":
    main()
```

---

## STEP 11 — Create `.env`

```env
ANTHROPIC_API_KEY=your_api_key_here
```

---

## STEP 12 — Create `README.md`

```markdown
# ATE Code Agent

Cursor AI for Teradyne UltraFLEX / IG-XL test programs.
Built using the Claude Agent SDK.

## What it does
- Reads your IG-XL Excel program (Test Instances, Limits, Flow, etc.)
- Reads and writes VBA test method code
- Plans changes, waits for your approval, then edits
- Runs the test program on the UltraFLEX tester
- Reads STDF log files, analyzes failures, proposes fixes
- Loops until all tests pass (max 5 iterations)

## Quick Start

1. Install dependencies:
   pip install -r requirements.txt

2. Add your API key to .env:
   ANTHROPIC_API_KEY=your_key_here

3. Run the agent on your IG-XL program:
   python main.py /path/to/your/program.xlsx

4. Talk to the agent:
   You: The IDDQ test is failing. The limit should be 150uA not 100uA. Fix it and verify.
   Agent: [reads program, plans change, waits for approval, edits, runs, reports back]

## Safety

- Pin Map sheet is ALWAYS read-only
- Voltage limits enforced (VCC max 5V, VDD max 3.6V)
- Git backup before every edit
- Max 5 debug iterations before escalating to you
- Engineering mode only — never production
```

---

## STEP 13 — Write Unit Tests

### `tests/test_excel_tools.py`

```python
"""Unit tests for Excel tools."""
import pytest
import openpyxl
import tempfile
import os
from pathlib import Path

# Create a minimal test .xlsx file
def create_test_workbook(path: str):
    wb = openpyxl.Workbook()
    # Test Instances sheet
    ws1 = wb.active
    ws1.title = 'Test Instances'
    ws1.append(['Test Name', 'Method', 'Parameter'])
    ws1.append(['IDDQ_TEST', 'IDDQ', 'default'])
    # Limits sheet
    ws2 = wb.create_sheet('Limits')
    ws2.append(['Test Name', 'Hi Limit', 'Lo Limit', 'Units'])
    ws2.append(['IDDQ_TEST', 100, 0, 'uA'])
    # Pin Map sheet (read-only)
    ws3 = wb.create_sheet('Pin Map')
    ws3.append(['Pin', 'Channel'])
    ws3.append(['VCC', 'CH1'])
    # Levels sheet
    ws4 = wb.create_sheet('Levels')
    ws4.append(['Level Name', 'VCC', 'VDD'])
    ws4.append(['nom', 3.3, 1.8])
    wb.save(path)


def test_read_excel_sheets():
    from tools.excel_tools import read_excel_sheets
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        path = f.name
    try:
        create_test_workbook(path)
        result = read_excel_sheets(path)
        assert 'Test Instances' in result
        assert 'Limits' in result
        assert result['Limits'][0]['Test Name'] == 'IDDQ_TEST'
        assert result['Limits'][0]['Hi Limit'] == 100
    finally:
        os.unlink(path)


def test_write_excel_cell_roundtrip():
    from tools.excel_tools import write_excel_cell, read_excel_sheets
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        path = f.name
    try:
        create_test_workbook(path)
        # Write new limit value (row 2, col 2 = Hi Limit for IDDQ_TEST)
        result = write_excel_cell(path, 'Limits', 2, 2, 150, backup=False)
        assert result['status'] == 'ok'
        assert result['old_value'] == 100
        assert result['new_value'] == 150
        # Verify it was actually written
        data = read_excel_sheets(path)
        assert data['Limits'][0]['Hi Limit'] == 150
    finally:
        os.unlink(path)


def test_write_pin_map_blocked():
    from tools.excel_tools import write_excel_cell
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        path = f.name
    try:
        create_test_workbook(path)
        with pytest.raises(ValueError, match='read-only'):
            write_excel_cell(path, 'Pin Map', 2, 1, 'CH99', backup=False)
    finally:
        os.unlink(path)


def test_write_voltage_limit_blocked():
    from tools.excel_tools import write_excel_cell
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        path = f.name
    try:
        create_test_workbook(path)
        with pytest.raises(ValueError, match='Voltage'):
            write_excel_cell(path, 'Levels', 2, 2, 99.9, backup=False)
    finally:
        os.unlink(path)
```

---

## STEP 14 — How to Run and Test

```bash
# 1. Create virtual environment
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run unit tests (no tester needed)
pytest tests/ -v

# 4. Run the agent on your IG-XL program
python main.py /path/to/YourDevice.xlsx

# 5. Try these example tasks:
# "Read the test program and give me a summary of all tests and limits"
# "The IDDQ_TEST Hi Limit is 100uA. Change it to 150uA and explain why that is safe"
# "List all VBA modules in this program"
# "Read the TestMethods VBA module and explain what IDDQ_TEST does"
# "Run the test program and tell me which tests are failing"
# "The continuity test is failing — diagnose and fix it"
```

---

## IMPORTANT NOTES FOR YOUR CODING AGENT

1. **Phase 1 only**: Build Steps 1–9 (tools + agent.py). Do NOT build FastAPI yet.

2. **Windows required**: `vba_tools.py` uses `win32com` which needs Windows + Excel.
   On other OS, implement read_vba_module using `oletools` as fallback.

3. **No real tester needed yet**: `run_tools.py` calls `igxl_run.exe`.
   For testing without a tester, add a `mock_run_igxl_program()` that writes
   a fake STDF log with some PASS and FAIL results.

4. **GitHub MCP**: Connect the GitHub MCP server to let the agent read your repo:
   ```python
   options = ClaudeAgentOptions(
       mcp_servers={
           "ate-tools": ate_tool_server,
           "github": {
               "command": "npx",
               "args": ["-y", "@modelcontextprotocol/server-github"],
               "env": {"GITHUB_TOKEN": os.environ["GITHUB_TOKEN"]}
           }
       }
   )
   ```

5. **Your existing RAG**: Connect your LangChain RAG as a custom tool:
   ```python
   @tool("query_ultraflex_docs", "Search UltraFLEX documentation", {"question": str})
   async def tool_rag_query(args):
       answer = igxl_rag_chain.invoke({"question": args["question"]})
       return {"content": [{"type": "text", "text": answer['answer']}]}
   ```
   Add `tool_rag_query` to the `ate_tool_server` tools list.

6. **Build order**:
   - Step 1: safety/rules.py + tools/git_tools.py (no dependencies)
   - Step 2: tools/excel_tools.py + tests/test_excel_tools.py (verify it works)
   - Step 3: tools/vba_tools.py
   - Step 4: tools/run_tools.py + tools/log_tools.py
   - Step 5: system_prompt.py
   - Step 6: agent.py (wires everything together)
   - Step 7: main.py (entry point)
   - Step 8: Run it!
```
