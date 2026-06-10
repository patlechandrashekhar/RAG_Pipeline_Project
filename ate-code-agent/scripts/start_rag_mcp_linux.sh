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
export ULTRAFLEX_MCP_CACHE_TTL="${ULTRAFLEX_MCP_CACHE_TTL:-300}"
export ULTRAFLEX_MCP_FAST_N_PER_QUERY="${ULTRAFLEX_MCP_FAST_N_PER_QUERY:-2}"

exec "${VENV_PYTHON}" "${ROOT_DIR}/rag_mcp_server.py"
