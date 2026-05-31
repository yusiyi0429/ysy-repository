#!/usr/bin/env python3
"""根据场景锚定信息与自定义知识列生成 Markdown 骨架。"""

from __future__ import annotations

from pathlib import Path

from scenario_schema import ANCHOR_COLUMNS
from step1_template import normalize_sub_scenarios


def _md_cell(value: str) -> str:
    return str(value or "").replace("|", "\\|").replace("\n", " ")


def generate_markdown_skeleton(
    output_path: Path | str,
    scenario_name: str,
    scenario_content: str,
    sub_scenarios: list,
    knowledge_columns: list[str],
) -> dict:
    """生成 Markdown 场景骨架（每个子场景一张表）。"""
    anchor = list(ANCHOR_COLUMNS)
    headers = anchor + list(knowledge_columns)
    subs = normalize_sub_scenarios(sub_scenarios)
    if not subs:
        subs = [{"name": "", "content": ""}]

    lines = [
        f"# 场景锚定骨架 · {scenario_name}",
        "",
        "## 场景说明",
        "",
        scenario_content or "（待补充）",
        "",
        "## 知识列",
        "",
        "、".join(knowledge_columns) if knowledge_columns else "（无）",
        "",
    ]

    for i, sub in enumerate(subs, start=1):
        sub_name = sub.get("name", "") or f"子场景{i}"
        sub_content = sub.get("content", "")
        lines.append(f"## 子场景：{sub_name}")
        if sub_content:
            lines.append("")
            lines.append(sub_content)
        lines.append("")
        lines.append("| " + " | ".join(headers) + " |")
        lines.append("| " + " | ".join(["---"] * len(headers)) + " |")
        row = [
            scenario_name,
            scenario_content,
            sub.get("name", ""),
            sub.get("content", ""),
        ] + [""] * len(knowledge_columns)
        lines.append("| " + " | ".join(_md_cell(c) for c in row) + " |")
        lines.append("")
        lines.append("> 可在下方继续追加知识条目行，或进入 Step2 使用同结构 Excel 萃取。")
        lines.append("")

    text = "\n".join(lines).strip() + "\n"
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(text, encoding="utf-8")

    return {
        "markdown_path": str(out),
        "headers": headers,
        "sub_scenario_count": len(subs),
        "knowledge_column_count": len(knowledge_columns),
    }
