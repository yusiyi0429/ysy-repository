#!/usr/bin/env python3
"""校验 Step1 从 scenario-schema 生成骨架（无网络）。"""

import sys
import tempfile
from pathlib import Path

BACKEND = Path(__file__).resolve().parent.parent
ROOT = BACKEND.parent
sys.path.insert(0, str(BACKEND))

from scenario_schema import enrich_knowledge_columns_for_markdown, is_abstract_column_name  # noqa: E402
from step1_schema_builder import generate_skeleton_from_schema  # noqa: E402


def main():
    schema_path = ROOT / "config" / "scenario-schema.yaml"
    if not schema_path.exists():
        print("FAIL: missing", schema_path)
        return 1

    if not is_abstract_column_name("列A"):
        print("FAIL: 列A should be abstract")
        return 1
    effective, enriched, _ = enrich_knowledge_columns_for_markdown(["列A", "列B"], None)
    if not enriched or len(effective) < 8:
        print("FAIL: markdown enrich", len(effective), enriched)
        return 1
    if "判断逻辑" not in effective:
        print("FAIL: missing 判断逻辑 in enriched columns")
        return 1

    tmp = Path(tempfile.mkdtemp(prefix="step1_schema_"))
    try:
        out = tmp / "template_test.xlsx"
        custom_cols = ["具体方法", "判断逻辑", "自定义列A"]
        result = generate_skeleton_from_schema(
            schema_path,
            out,
            "测试场景",
            "测试场景说明内容",
            [{"name": "子场景A", "content": "子场景说明"}],
            knowledge_columns=custom_cols,
        )
        if not out.exists():
            print("FAIL: output xlsx not created")
            return 1
        fields = result.get("fields_info") or []
        if not fields:
            print("FAIL: fields_info empty")
            return 1
        headers = fields[0].get("headers") or []
        if "场景" not in headers:
            print("FAIL: missing anchor column 场景", headers)
            return 1
        if "自定义列A" not in headers:
            print("FAIL: missing custom column", headers)
            return 1
        md_out = tmp / "template_test.md"
        from step1_markdown_builder import generate_markdown_skeleton  # noqa: E402
        generate_markdown_skeleton(md_out, "测试场景", "说明", [{"name": "子场景A", "content": ""}], custom_cols)
        if not md_out.exists() or "自定义列A" not in md_out.read_text(encoding="utf-8"):
            print("FAIL: markdown skeleton")
            return 1
        print("OK: schema skeleton", out.name, "cols=", len(headers), "subs=", result.get("sub_scenario_count"))
        return 0
    finally:
        import shutil
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    raise SystemExit(main())
