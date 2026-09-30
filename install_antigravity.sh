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

# 0. Verify GCP Account Authentication (Shared Skill Access Gate)
if [[ -z "${REMOTE_URL}" ]]; then
  if ! command -v gcloud >/dev/null 2>&1; then
    echo "❌ Error: 'gcloud' CLI가 설치되어 있지 않습니다. 이 공유 스킬은 GCP 계정이 있는 사용자만 사용할 수 있습니다."
    exit 1
  fi

  ACTIVE_ACCOUNT="$(gcloud auth list --filter=status:ACTIVE --format="value(account)" 2>/dev/null | head -n 1 || true)"
  ADC_FILE="${CLOUDSDK_CONFIG:-${HOME}/.config/gcloud}/application_default_credentials.json"
  if [[ -z "${ACTIVE_ACCOUNT}" ]] && [[ ! -f "${ADC_FILE}" ]]; then
    echo "❌ Error: 로그인된 GCP 계정이 없습니다. 먼저 아래 명령어로 GCP 계정에 로그인해 주세요:"
    echo "   gcloud auth login"
    echo "   gcloud auth application-default login"
    exit 1
  fi

  echo "🔐 Verified GCP Account: ${ACTIVE_ACCOUNT:-ADC}"
fi

# Ensure Pretendard fonts are cached locally in ~/.local/share/fonts/pretendard for 1:1 HTML font compatibility
LOCAL_PRETENDARD_DIR="${HOME}/.local/share/fonts/pretendard"
if [[ ! -f "${LOCAL_PRETENDARD_DIR}/Pretendard-Regular.otf" ]] && command -v curl >/dev/null 2>&1; then
  echo "[Fonts] Downloading official Pretendard fonts for 1:1 HTML rendering compatibility..."
  mkdir -p "${LOCAL_PRETENDARD_DIR}"
  for w in Regular Medium SemiBold Bold ExtraBold Black; do
    curl -fsSL "https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/packages/pretendard/dist/public/static/Pretendard-${w}.otf" \
      -o "${LOCAL_PRETENDARD_DIR}/Pretendard-${w}.otf" 2>/dev/null || true
  done
  if command -v fc-cache >/dev/null 2>&1; then
    fc-cache -f "${LOCAL_PRETENDARD_DIR}" >/dev/null 2>&1 || true
  fi
fi

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

  if [[ -z "${DETECTED_PROJECT}" ]]; then
    echo "❌ Error: GCP Project ID를 찾을 수 없습니다. 본인 GCP 프로젝트 ID를 지정해 주세요:"
    echo "   ./install_antigravity.sh --project YOUR_GCP_PROJECT_ID"
    exit 1
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
