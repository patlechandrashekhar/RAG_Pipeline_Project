#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
VENV_PYTHON="${ROOT_DIR}/../.venv/bin/python"

if [[ ! -x "${VENV_PYTHON}" ]]; then
  echo "Python venv not found at ${VENV_PYTHON}."
  echo "Create it first: python3 -m venv ../.venv"
  exit 1
fi

export PYTHONHTTPSVERIFY="0"
export CURL_CA_BUNDLE=""
export REQUESTS_CA_BUNDLE=""
export ULTRAFLEX_MCP_VECTOR_ONLY="${ULTRAFLEX_MCP_VECTOR_ONLY:-true}"

exec "${VENV_PYTHON}" -m uvicorn rag_http_api:app --host 0.0.0.0 --port 8081 --app-dir "${ROOT_DIR}"
