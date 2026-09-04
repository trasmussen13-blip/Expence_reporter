#!/usr/bin/env bash
set -euo pipefail

export PORT="${PORT:-5000}"
export OLLAMA_HOST="${OLLAMA_HOST:-http://127.0.0.1:11434}"

PYTHON_BIN="${PYTHON_BIN:-python3}"
if [[ -x ".venv/bin/python" ]]; then
  PYTHON_BIN=".venv/bin/python"
fi

echo "Starter Kvitteringer · Månedsrapport på port ${PORT}"
echo "Bruger lokal Ollama på ${OLLAMA_HOST}"
exec "${PYTHON_BIN}" app.py