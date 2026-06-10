param(
    [ValidateSet("setup", "run", "test", "check")]
    [string]$Task = "run"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$workspaceRoot = Split-Path -Parent $projectRoot

$pythonCandidates = @(
    "$projectRoot\.venv\Scripts\python.exe",
    "$workspaceRoot\.venv\Scripts\python.exe",
    "python"
)

$python = $null
foreach ($candidate in $pythonCandidates) {
    if ($candidate -eq "python" -or (Test-Path $candidate)) {
        $python = $candidate
        break
    }
}

if (-not $python) {
    throw "Python executable not found."
}

function Invoke-Checked {
    param([Parameter(Mandatory = $true)][scriptblock]$Command)
    & $Command
    if ($LASTEXITCODE -ne 0) {
        throw "Command failed with exit code $LASTEXITCODE"
    }
}

switch ($Task) {
    "setup" {
        Invoke-Checked { & $python -m pip install --upgrade pip }
        Invoke-Checked { & $python -m pip install -r "$projectRoot\requirements.txt" }
        Invoke-Checked { & $python -m pip install pytest }
    }
    "run" {
        Invoke-Checked { & $python -m streamlit run "$projectRoot\app\streamlit_app.py" }
    }
    "test" {
        Invoke-Checked { & $python -m pytest "$projectRoot\tests" -q -p no:cacheprovider }
    }
    "check" {
        Invoke-Checked { & $python -c "import ast, pathlib; root=pathlib.Path(r'$projectRoot'); files=[root/'app'/'streamlit_app.py']+list((root/'src'/'page_indexing_rag').glob('*.py')); [ast.parse(p.read_text(encoding='utf-8', errors='replace')) for p in files]; print('AST_OK')" }
    }
}
