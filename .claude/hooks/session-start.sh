#!/bin/bash
# SessionStart hook: install Python dependencies so the reconciliation engine's
# tests and CLI work in a fresh Claude Code (web) container.
# Idempotent and non-interactive; safe to re-run.
set -euo pipefail

# Only needed in the remote (web) environment; local sessions manage their own.
if [ "${CLAUDE_CODE_REMOTE:-}" != "true" ]; then
  exit 0
fi

cd "${CLAUDE_PROJECT_DIR:-.}"

python3 -m pip install --quiet --upgrade pip >/dev/null 2>&1 || true
python3 -m pip install --quiet -r requirements.txt

# Let `python3 recon_engine.py ...` and `import recon_engine` resolve from root.
echo 'export PYTHONPATH="."' >> "${CLAUDE_ENV_FILE:-/dev/null}"
