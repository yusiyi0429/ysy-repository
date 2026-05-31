"""应用与镜像版本信息（单一来源：项目根目录 VERSION + 构建时环境变量）。"""

from __future__ import annotations

import os
from datetime import datetime, timezone
from pathlib import Path

# Step2 Excel 产出格式兼容标识（与历史 health.build 字段一致）
STEP2_EXCEL_BUILD = "20260525-optional-sub-scenario"

PROJECT_DIR = Path(__file__).resolve().parent.parent
_VERSION_FILE = PROJECT_DIR / "VERSION"


def _read_app_version() -> str:
    if not _VERSION_FILE.exists():
        return "0.0.0"
    for line in _VERSION_FILE.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line and not line.startswith("#"):
            return line
    return "0.0.0"


def get_release_info() -> dict:
    app_version = os.environ.get("APP_VERSION", "").strip() or _read_app_version()
    image_version = (
        os.environ.get("IMAGE_VERSION", "").strip()
        or os.environ.get("APP_BUILD", "").strip()
        or "dev"
    )
    build_date = os.environ.get("BUILD_DATE", "").strip()
    if not build_date:
        build_date = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    platform = os.environ.get("APP_PLATFORM", "").strip() or "unknown"
    vcs_ref = os.environ.get("VCS_REF", "").strip()
    return {
        "app_name": "tacit-knowledge-externalization",
        "app_version": app_version,
        "image_version": image_version,
        "build_id": image_version,
        "build_date": build_date,
        "platform": platform,
        "vcs_ref": vcs_ref,
        "step2_excel_build": STEP2_EXCEL_BUILD,
        # 兼容旧前端/脚本仅读取 build 字段
        "build": STEP2_EXCEL_BUILD,
    }


def format_image_tag(app_version: str, platform_slug: str, timestamp: str) -> str:
    """标准镜像 tag：{app_version}-{platform}-{yyyyMMdd-HHmm}"""
    return f"{app_version}-{platform_slug}-{timestamp}"
