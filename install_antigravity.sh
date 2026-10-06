#!/usr/bin/env bash
set -euo pipefail

PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

LOCAL_PRETENDARD_DIR="${HOME}/.local/share/fonts/pretendard"
if [[ ! -f "${LOCAL_PRETENDARD_DIR}/Pretendard-Regular.otf" ]] && command -v curl >/dev/null 2>&1; then
  echo "[1/3] Downloading Pretendard fonts for 1:1 HTML rendering compatibility..."
  mkdir -p "${LOCAL_PRETENDARD_DIR}"
  for w in Regular Medium SemiBold Bold ExtraBold Black; do
    curl -fsSL "https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/packages/pretendard/dist/public/static/Pretendard-${w}.otf" \
      -o "${LOCAL_PRETENDARD_DIR}/Pretendard-${w}.otf" 2>/dev/null || true
  done
  if command -v fc-cache >/dev/null 2>&1; then
    fc-cache -f "${LOCAL_PRETENDARD_DIR}" >/dev/null 2>&1 || true
  fi
else
  echo "[1/3] Pretendard fonts verified."
fi

if [[ ! -x "${PROJECT_DIR}/.venv/bin/python" ]]; then
  echo "[2/3] Creating Python virtual environment in ${PROJECT_DIR}/.venv ..."
  if command -v uv >/dev/null 2>&1; then
    (cd "${PROJECT_DIR}" && uv venv .venv && uv pip install --index-url https://pypi.org/simple -e .)
  else
    python3 -m venv "${PROJECT_DIR}/.venv"
    "${PROJECT_DIR}/.venv/bin/pip" install --index-url https://pypi.org/simple --upgrade pip
    "${PROJECT_DIR}/.venv/bin/pip" install --index-url https://pypi.org/simple -e "${PROJECT_DIR}"
  fi
else
  echo "[2/3] Updating local package in ${PROJECT_DIR}/.venv ..."
  if command -v uv >/dev/null 2>&1; then
    (cd "${PROJECT_DIR}" && uv pip install --index-url https://pypi.org/simple -e .)
  else
    "${PROJECT_DIR}/.venv/bin/pip" install --index-url https://pypi.org/simple -e "${PROJECT_DIR}"
  fi
fi

CONFIG_DIR="${HOME}/.gemini/config"
mkdir -p "${CONFIG_DIR}/skills"
ln -sfn "${PROJECT_DIR}" "${CONFIG_DIR}/skills/html-to-pptx"
rm -f "${CONFIG_DIR}/plugins/html-to-pptx"

GLOBAL_MCP_CONFIG="${CONFIG_DIR}/mcp_config.json"
if [[ -f "${GLOBAL_MCP_CONFIG}" ]]; then
  PYTHON_FOR_JSON="$(command -v python3 || echo "${PROJECT_DIR}/.venv/bin/python")"
  "${PYTHON_FOR_JSON}" - <<PY
import json
import os

global_path = "${GLOBAL_MCP_CONFIG}"
try:
    with open(global_path, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    if isinstance(cfg.get("mcpServers"), dict) and "html-to-pptx" in cfg["mcpServers"]:
        del cfg["mcpServers"]["html-to-pptx"]
        with open(global_path, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2, ensure_ascii=False)
except Exception:
    pass
PY
fi

echo "[3/3] Linked Skill to ${CONFIG_DIR}/skills/html-to-pptx"
echo "✅ Done! 'html-to-pptx' is ready as a pure local Agent Skill (no Cloud Run, MCP server, or GCP Project required)."
