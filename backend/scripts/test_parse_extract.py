#!/usr/bin/env python3
"""Unit tests for Step2 LLM JSON parsing (no server)."""

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from app_server import _parse_extracted_items, _normalize_extracted_items  # noqa: E402


def test_preamble_array():
    raw = '说明如下：\n[{"环节-具体方法": "测试规则", "环节-知识引用": "doc1"}]'
    items, mode = _parse_extracted_items(raw)
    assert len(items) == 1, (items, mode)
    assert mode != "parse_failed"


def test_wrapper_items():
    raw = '{"items": [{"content": "hello"}]}'
    items, mode = _parse_extracted_items(raw)
    assert len(items) == 1


def test_trailing_comma():
    raw = '[{"content": "a",}, {"content": "b",}]'
    items, mode = _parse_extracted_items(raw)
    assert len(items) >= 1


def test_align_keys():
    cols = ["环节-具体方法", "环节-知识引用"]
    items = _normalize_extracted_items([{"具体方法": "x", "知识引用": "y"}], cols)
    assert items[0].get("环节-具体方法") == "x"


if __name__ == "__main__":
    test_preamble_array()
    test_wrapper_items()
    test_trailing_comma()
    test_align_keys()
    print("All parse tests passed.")
