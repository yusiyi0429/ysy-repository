# =============================================================================
# 隐性知识显性化流水线 - Docker 镜像（linux/arm64 / linux/amd64）
#
# 构建示例（ARM64 离线包）:
#   .\scripts\build-docker-arm64.ps1
# 或:
#   docker buildx build --platform linux/arm64 \
#     --build-arg APP_VERSION=$(head -1 VERSION) \
#     --build-arg IMAGE_VERSION=2.0.0-arm64-202605291800 \
#     -t tacit-knowledge-externalization:2.0.0-arm64-202605291800 \
#     -f docker/Dockerfile --load .
# =============================================================================

FROM python:3.11-slim

ARG APP_VERSION=0.0.0
ARG IMAGE_VERSION=dev
ARG BUILD_DATE=
ARG VCS_REF=
ARG APP_PLATFORM=linux/arm64

LABEL maintainer="tacit-knowledge-team"
LABEL org.opencontainers.image.title="tacit-knowledge-externalization"
LABEL org.opencontainers.image.description="隐性知识显性化四步法流水线 Web 应用"
LABEL org.opencontainers.image.version="${APP_VERSION}"
LABEL org.opencontainers.image.revision="${VCS_REF}"
LABEL org.opencontainers.image.created="${BUILD_DATE}"
LABEL com.tacit-knowledge.app.version="${APP_VERSION}"
LABEL com.tacit-knowledge.image.version="${IMAGE_VERSION}"
LABEL com.tacit-knowledge.platform="${APP_PLATFORM}"

WORKDIR /app

# 不安装 curl/apt 依赖，避免 ARM64 交叉构建时访问 deb.debian.org 失败（内网/受限网络）
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY VERSION ./
COPY backend/ ./backend/
COPY config/ ./config/
COPY data/ ./data/
COPY frontend/ ./frontend/
COPY scripts/copy-frontend-vendor.js ./scripts/copy-frontend-vendor.js

# 若仓库已含 frontend/vendor 则跳过 npm；否则在镜像内安装并裁剪 Luckysheet
RUN set -eux; \
    if [ -f frontend/vendor/luckysheet/luckysheet.umd.js ] \
       && [ -f frontend/vendor/jquery/jquery.min.js ]; then \
      echo "Using prebuilt frontend/vendor"; \
    else \
      apt-get update && apt-get install -y --no-install-recommends nodejs npm; \
      cd frontend && npm install luckysheet@2.1.13 jquery@3.6.4 --no-save; \
      node ../scripts/copy-frontend-vendor.js; \
      rm -rf node_modules; \
      apt-get purge -y nodejs npm && apt-get autoremove -y && rm -rf /var/lib/apt/lists/*; \
    fi

ENV PYTHONUNBUFFERED=1 \
    TZ=Asia/Shanghai \
    WORKSPACE_DIR=/app/workspace \
    APP_VERSION=${APP_VERSION} \
    IMAGE_VERSION=${IMAGE_VERSION} \
    APP_BUILD=${IMAGE_VERSION} \
    BUILD_DATE=${BUILD_DATE} \
    VCS_REF=${VCS_REF} \
    APP_PLATFORM=${APP_PLATFORM}

RUN mkdir -p /app/logs /app/workspace

EXPOSE 5000

HEALTHCHECK --interval=30s --timeout=5s --start-period=15s --retries=3 \
    CMD python -c "import urllib.request; urllib.request.urlopen('http://127.0.0.1:5000/api/health', timeout=3)" || exit 1

CMD ["python", "backend/app_server.py", "--host", "0.0.0.0", "--port", "5000"]
