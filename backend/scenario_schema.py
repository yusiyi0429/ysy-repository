#!/usr/bin/env python3
"""加载 scenario-schema.yaml，解析知识列定义，供 Step1 从结构方案生成 Excel 骨架。"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import yaml

ANCHOR_COLUMNS = ("场景", "场景说明", "子场景", "子场景说明")

# schema 未配置 fields 时使用的默认知识列（与 Step2 映射、修订流程兼容）
DEFAULT_KNOWLEDGE_COLUMNS = (
    "环节",
    "访谈方向",
    "具体方法",
    "知识类型",
    "知识引用",
    "知识描述",
    "适用条件",
    "判断逻辑",
    "反模式/踩坑提示",
    "来源文档",
    "置信度",
    "备注",
)

# Markdown 萃取默认补齐的富语义列（与 DEFAULT 一致，可单独调整顺序）
RICH_MARKDOWN_EXTRACTION_COLUMNS = DEFAULT_KNOWLEDGE_COLUMNS

_ABSTRACT_COLUMN_RE = None


def _abstract_column_re():
    import re
    global _ABSTRACT_COLUMN_RE
    if _ABSTRACT_COLUMN_RE is None:
        _ABSTRACT_COLUMN_RE = re.compile(
            r"^(列[a-zA-Z0-9]{1,3}|(column|field|字段)\s*\d+|[a-zA-Z]\d?)$",
            re.IGNORECASE,
        )
    return _ABSTRACT_COLUMN_RE


def is_abstract_column_name(name: str) -> bool:
    """无业务语义的占位列名（如 列A、字段1）。"""
    s = str(name or "").strip()
    if not s:
        return True
    if s in ANCHOR_COLUMNS:
        return True
    return bool(_abstract_column_re().match(s))


def enrich_knowledge_columns_for_markdown(
    user_columns: list | None,
    schema: dict[str, Any] | None = None,
) -> tuple[list[str], bool, list[str]]:
    """
    Markdown 路径：在用户列基础上补齐富语义列，提升 Step2 萃取深度。
    返回 (有效列, 是否发生补齐, 用户原始有效列)。
    """
    user_norm = normalize_knowledge_columns(user_columns)
    substantive_user = [c for c in user_norm if not is_abstract_column_name(c)]
    rich_pool = list(resolve_knowledge_columns(schema))
    seen: set[str] = set()
    effective: list[str] = []
    for c in substantive_user:
        if c not in seen:
            seen.add(c)
            effective.append(c)
    before_len = len(effective)
    for c in rich_pool:
        if c not in seen:
            seen.add(c)
            effective.append(c)
    enriched = len(effective) > before_len or (before_len == 0 and len(user_norm) > 0)
    if not effective:
        effective = list(rich_pool)
    return effective, enriched, substantive_user


def load_scenario_schema(schema_path: Path | str) -> dict[str, Any]:
    path = Path(schema_path)
    if not path.exists():
        return {}
    with open(path, encoding="utf-8") as f:
        data = yaml.safe_load(f)
    return data if isinstance(data, dict) else {}


def normalize_knowledge_columns(names: list | None) -> list[str]:
    """用户自定义知识列：去重、去空、排除锚定列名。"""
    anchor_set = set(ANCHOR_COLUMNS)
    out: list[str] = []
    seen: set[str] = set()
    for raw in names or []:
        name = str(raw or "").strip()
        if not name or name in seen or name in anchor_set:
            continue
        seen.add(name)
        out.append(name)
    return out[:40]


def resolve_knowledge_columns(schema: dict[str, Any] | None) -> list[str]:
    """从 schema.fields 解析知识列；无有效字段时回退默认列集。"""
    schema = schema or {}
    names: list[str] = []
    seen: set[str] = set()
    for raw in schema.get("fields") or []:
        if not isinstance(raw, dict):
            continue
        name = str(raw.get("name", "")).strip()
        if not name or name in seen:
            continue
        seen.add(name)
        names.append(name)
    if names:
        return names
    return list(DEFAULT_KNOWLEDGE_COLUMNS)


def resolve_knowledge_columns_for_request(
    schema: dict[str, Any] | None,
    user_columns: list | None,
) -> list[str]:
    """优先使用用户列；否则 schema / 默认列。"""
    normalized = normalize_knowledge_columns(user_columns)
    if normalized:
        return normalized
    return resolve_knowledge_columns(schema)


def schema_summary(schema: dict[str, Any] | None, *, schema_path: Path | str | None = None) -> dict[str, Any]:
    schema = schema or {}
    knowledge_columns = resolve_knowledge_columns(schema)
    rich_md = list(RICH_MARKDOWN_EXTRACTION_COLUMNS)
    return {
        "schema_path": str(schema_path) if schema_path else "",
        "scenario_name": str(schema.get("scenario_name", "")).strip(),
        "display_name": str(schema.get("display_name", "")).strip() or "默认知识结构",
        "domain": str(schema.get("domain", "")).strip(),
        "version": str(schema.get("version", "v1.0")).strip() or "v1.0",
        "categories": list(schema.get("categories") or []),
        "anchor_columns": list(ANCHOR_COLUMNS),
        "knowledge_columns": knowledge_columns,
        "rich_markdown_columns": rich_md,
        "column_count": len(ANCHOR_COLUMNS) + len(knowledge_columns),
    }
