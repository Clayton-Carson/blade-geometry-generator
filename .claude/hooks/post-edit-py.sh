#!/usr/bin/env bash
# Non-blocking: run ruff in the background on Python file edits.
# Prefer a per-project venv's ruff over whatever's on PATH, because
# Claude Code's shell does not inherit venv activation.
set -u
payload="$(cat)"
path="$(printf '%s' "$payload" | tr -d '\n' \
  | sed -n 's/.*"file_path"[[:space:]]*:[[:space:]]*"\(\(\\.\|[^"\\]\)*\)".*/\1/p')"
[ -z "$path" ] && exit 0

case "$path" in
  *.py) ;;
  *) exit 0 ;;
esac

cd "$CLAUDE_PROJECT_DIR" 2>/dev/null || exit 0

# Find ruff — prefer project venv, fall back to PATH.
if   [ -x "./venv/Scripts/ruff.exe" ];  then ruff="./venv/Scripts/ruff.exe"
elif [ -x "./venv/bin/ruff" ];          then ruff="./venv/bin/ruff"
elif [ -x "./.venv/Scripts/ruff.exe" ]; then ruff="./.venv/Scripts/ruff.exe"
elif [ -x "./.venv/bin/ruff" ];         then ruff="./.venv/bin/ruff"
elif command -v ruff >/dev/null 2>&1;   then ruff="ruff"
else exit 0
fi

# Fire and forget — do not block the agent on lint output.
( "$ruff" check "$path" >/dev/null 2>&1 ) &
exit 0
