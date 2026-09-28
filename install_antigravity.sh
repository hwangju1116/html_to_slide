#!/usr/bin/env bash
set -euo pipefail

PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REMOTE_URL=""
EXPLICIT_GCP_PROJECT=""
GCP_LOCATION="global"

while [[ $# -gt 0 ]]; do
  case "$1" in
    --remote-url)
      REMOTE_URL="$2"
      shift 2
      ;;
    --project)
      EXPLICIT_GCP_PROJECT="$2"
      shift 2
      ;;
    --location)
      GCP_LOCATION="$2"
      shift 2
      ;;
    *)
      shift
      ;;
  esac
done

CONFIG_DIR="${HOME}/.gemini/config"
mkdir -p "${CONFIG_DIR}/plugins" "${CONFIG_DIR}/skills"

# 1. Link Plugin & Skill into Antigravity (~/.gemini/config)
ln -sfn "${PROJECT_DIR}" "${CONFIG_DIR}/plugins/html-to-pptx"
ln -sfn "${PROJECT_DIR}/skills/html-to-pptx" "${CONFIG_DIR}/skills/html-to-pptx"
echo "[1/3] Linked Skill & Plugin into ${CONFIG_DIR}/skills/html-to-pptx"

# 2. Build mcp_config.json (either Remote Cloud Run SSE mode or Local Stdio mode)
if [[ -n "${REMOTE_URL}" ]]; then
  echo "[2/3] Configuring Remote Cloud Run SSE MCP endpoint: ${REMOTE_URL}"
  cat > "${PROJECT_DIR}/mcp_config.json" <<EOF
{
  "mcpServers": {
    "html-to-pptx": {
      "serverUrl": "${REMOTE_URL}"
    }
  }
}
EOF
else
  # Resolve active GCP Project dynamically (explicit flag -> env var -> gcloud config)
  DETECTED_PROJECT="${EXPLICIT_GCP_PROJECT:-${GOOGLE_CLOUD_PROJECT:-}}"
  if [[ -z "${DETECTED_PROJECT}" ]] && command -v gcloud >/dev/null 2>&1; then
    DETECTED_PROJECT="$(gcloud config get-value project 2>/dev/null | tr -d '[:space:]' || true)"
    if [[ "${DETECTED_PROJECT}" == "(unset)" ]]; then
      DETECTED_PROJECT=""
    fi
  fi

  # Ensure local virtual environment exists
  if [[ ! -x "${PROJECT_DIR}/.venv/bin/python" ]]; then
    echo "[Setup] Creating Python virtual environment in ${PROJECT_DIR}/.venv ..."
    if command -v uv >/dev/null 2>&1; then
      (cd "${PROJECT_DIR}" && uv venv .venv && uv pip install --index-url https://pypi.org/simple -e .)
    else
      python3 -m venv "${PROJECT_DIR}/.venv"
      "${PROJECT_DIR}/.venv/bin/pip" install --index-url https://pypi.org/simple --upgrade pip
      "${PROJECT_DIR}/.venv/bin/pip" install --index-url https://pypi.org/simple -e "${PROJECT_DIR}"
    fi
  fi

  PYTHON_BIN="${PROJECT_DIR}/.venv/bin/python"

  echo "[2/3] Updating ${PROJECT_DIR}/mcp_config.json (GCP Project: ${DETECTED_PROJECT}, AI Model Region: ${GCP_LOCATION})"
  cat > "${PROJECT_DIR}/mcp_config.json" <<EOF
{
  "mcpServers": {
    "html-to-pptx": {
      "command": "${PYTHON_BIN}",
      "args": [
        "${PROJECT_DIR}/app/mcp_server.py"
      ],
      "env": {
        "GOOGLE_CLOUD_PROJECT": "${DETECTED_PROJECT}",
        "GOOGLE_CLOUD_LOCATION": "${GCP_LOCATION}",
        "VERTEX_AI_LOCATION": "${GCP_LOCATION}",
        "GOOGLE_GENAI_USE_VERTEXAI": "True",
        "PYTHONPATH": "${PROJECT_DIR}"
      }
    }
  }
}
EOF
fi

# 3. Merge into ~/.gemini/config/mcp_config.json safely
GLOBAL_MCP_CONFIG="${CONFIG_DIR}/mcp_config.json"
PYTHON_FOR_JSON="$(command -v python3 || echo "${PROJECT_DIR}/.venv/bin/python")"
"${PYTHON_FOR_JSON}" - <<PY
import json
import os

global_path = "${GLOBAL_MCP_CONFIG}"
local_path = "${PROJECT_DIR}/mcp_config.json"

with open(local_path, "r", encoding="utf-8") as f:
    local_cfg = json.load(f)

if os.path.exists(global_path):
    try:
        with open(global_path, "r", encoding="utf-8") as f:
            global_cfg = json.load(f)
    except Exception:
        global_cfg = {}
else:
    global_cfg = {}

global_cfg.setdefault("mcpServers", {})
global_cfg["mcpServers"]["html-to-pptx"] = local_cfg["mcpServers"]["html-to-pptx"]

with open(global_path, "w", encoding="utf-8") as f:
    json.dump(global_cfg, f, indent=2, ensure_ascii=False)

print(f"[3/3] Registered 'html-to-pptx' MCP server in {global_path}")
PY

echo "✅ Done! Restart Antigravity or reload window to use the 'html-to-pptx' Skill & MCP tools globally."

