#!/usr/bin/env bash
# 构建 linux/arm64 镜像并导出离线 tar（Linux/macOS）
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"

IMAGE_NAME="tacit-knowledge-externalization"
PLATFORM="linux/arm64"
PLATFORM_SLUG="arm64"
DATE_TAG="$(date +%Y%m%d-%H%M)"
UTC_NOW="$(date -u +%Y-%m-%dT%H:%M:%SZ)"

APP_VERSION="${APP_VERSION:-}"
if [[ -z "$APP_VERSION" && -f VERSION ]]; then
  APP_VERSION="$(grep -v '^#' VERSION | head -1 | tr -d '[:space:]')"
fi
[[ -n "$APP_VERSION" ]] || { echo "VERSION 为空"; exit 1; }

VERSION_TAG="${APP_VERSION}-${PLATFORM_SLUG}-${DATE_TAG}"
TAR_FILE="${ROOT}/${IMAGE_NAME}-${VERSION_TAG}.tar"
TAR_LATEST="${ROOT}/tacit-knowledge-externalization-arm64.tar"
MANIFEST="${ROOT}/${IMAGE_NAME}-${VERSION_TAG}.manifest.json"

echo "=== ARM64 镜像构建 ==="
echo "APP_VERSION=$APP_VERSION"
echo "IMAGE_TAG=$VERSION_TAG"

docker buildx build --platform "$PLATFORM" \
  --build-arg "APP_VERSION=$APP_VERSION" \
  --build-arg "IMAGE_VERSION=$VERSION_TAG" \
  --build-arg "BUILD_DATE=$UTC_NOW" \
  --build-arg "APP_PLATFORM=$PLATFORM" \
  -t "${IMAGE_NAME}:${VERSION_TAG}" \
  -t "${IMAGE_NAME}:${APP_VERSION}-arm64" \
  -f docker/Dockerfile \
  --load .

docker save "${IMAGE_NAME}:${VERSION_TAG}" -o "$TAR_FILE"
cp -f "$TAR_FILE" "$TAR_LATEST"

cat > "$MANIFEST" <<EOF
{
  "app_name": "$IMAGE_NAME",
  "app_version": "$APP_VERSION",
  "image": "${IMAGE_NAME}:${VERSION_TAG}",
  "image_version": "$VERSION_TAG",
  "platform": "$PLATFORM",
  "built_at_utc": "$UTC_NOW",
  "tar_file": "$(basename "$TAR_FILE")",
  "compose_image_line": "image: ${IMAGE_NAME}:${VERSION_TAG}"
}
EOF

echo "完成: $TAR_FILE"
echo "清单: $MANIFEST"
