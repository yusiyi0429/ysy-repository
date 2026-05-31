#!/bin/bash
# 示例：docker run 部署（必须挂载 workspace，否则重启丢流水线）
set -euo pipefail

IMAGE_TAG="${1:-tacit-knowledge-externalization:arm64-v2026.05.25-latest}"
HOST_PORT="${2:-5001}"
DATA_ROOT="${3:-/data/tacit-knowledge}"

mkdir -p "${DATA_ROOT}/workspace" "${DATA_ROOT}/logs"
if [ ! -f "${DATA_ROOT}/llm-config.yaml" ]; then
  cp -n config/llm-config.yaml "${DATA_ROOT}/llm-config.yaml" 2>/dev/null || true
fi

docker rm -f tacit-knowledge-externalization 2>/dev/null || true

docker run -d --name tacit-knowledge-externalization \
  --platform linux/arm64 \
  -p "${HOST_PORT}:5000" \
  -v "${DATA_ROOT}/workspace:/app/workspace" \
  -v "${DATA_ROOT}/logs:/app/logs" \
  -v "${DATA_ROOT}/llm-config.yaml:/app/config/llm-config.yaml:ro" \
  -e WORKSPACE_DIR=/app/workspace \
  -e TZ=Asia/Shanghai \
  --restart unless-stopped \
  "${IMAGE_TAG}"

echo "Health: curl http://127.0.0.1:${HOST_PORT}/api/health"
