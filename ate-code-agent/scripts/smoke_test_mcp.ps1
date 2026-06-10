param(
    [string]$PythonExe = "c:/AI Projects/.venv/Scripts/python.exe"
)

$ErrorActionPreference = "Stop"

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Resolve-Path (Join-Path $scriptRoot "..")
$serverScript = Join-Path $projectRoot "rag_mcp_server.py"

if (-not (Test-Path $PythonExe)) {
    throw "Python executable not found: $PythonExe"
}

if (-not (Test-Path $serverScript)) {
    throw "MCP server script not found: $serverScript"
}

Write-Host "[1/3] Syntax check..."
& $PythonExe -m py_compile $serverScript
if ($LASTEXITCODE -ne 0) {
    throw "py_compile failed for $serverScript"
}

Write-Host "[2/3] Import and ping tool check..."
$pingCode = @"
import os, sys
from pathlib import Path
root = Path(r'''$projectRoot''').resolve()
os.chdir(root)
import rag_mcp_server as srv
print(srv.ping())
"@
& $PythonExe -c $pingCode
if ($LASTEXITCODE -ne 0) {
    throw "Import/ping check failed"
}

Write-Host "[3/3] Process startup check (stdio server)..."
$tmpScript = Join-Path $env:TEMP "mcp_stdio_startup_check.py"
$startupTemplate = @'
import subprocess
import time
from pathlib import Path

python_exe = r"__PYTHON_EXE__"
server_script = r"__SERVER_SCRIPT__"
project_root = Path(r"__PROJECT_ROOT__").resolve()

p = subprocess.Popen(
    [python_exe, server_script],
    cwd=str(project_root),
    stdin=subprocess.PIPE,
    stdout=subprocess.PIPE,
    stderr=subprocess.PIPE,
    text=True,
)

time.sleep(1.5)
still_running = p.poll() is None
if still_running:
    p.terminate()
    try:
        p.wait(timeout=3)
    except Exception:
        p.kill()
        p.wait(timeout=3)

if not still_running:
    stderr_text = ""
    try:
        stderr_text = (p.stderr.read() or "").strip()
    except Exception:
        stderr_text = ""
    raise SystemExit("Server exited early. stderr: " + stderr_text)

print("stdio startup OK")
'@

$startupCode = $startupTemplate
$startupCode = $startupCode.Replace("__PYTHON_EXE__", $PythonExe.Replace("\", "\\"))
$startupCode = $startupCode.Replace("__SERVER_SCRIPT__", [string]$serverScript).Replace("\", "\\")
$startupCode = $startupCode.Replace("__PROJECT_ROOT__", [string]$projectRoot).Replace("\", "\\")

Set-Content -Path $tmpScript -Value $startupCode -Encoding UTF8
& $PythonExe $tmpScript
Remove-Item -Path $tmpScript -ErrorAction SilentlyContinue
if ($LASTEXITCODE -ne 0) {
    throw "Server process startup check failed"
}

Write-Host "Smoke test passed: MCP server imports, responds to ping(), and stays up as a stdio process."
