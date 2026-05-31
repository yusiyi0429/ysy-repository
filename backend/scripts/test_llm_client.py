#!/usr/bin/env python3
"""Unit tests for llm_client adapters (no network)."""

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from llm_client import (
    API_TYPE_CCB,
    API_TYPE_OPENAI,
    _ccb_validate_gateway,
    extract_assistant_content,
    extract_delta_from_chunk,
    normalize_llm_url,
    parse_sse_data_line,
)


def test_normalize_openai_url():
    assert normalize_llm_url("http://h/v1", API_TYPE_OPENAI).endswith("/chat/completions")
    assert normalize_llm_url("http://h/ai-service/ainlplm/chat", API_TYPE_CCB) == "http://h/ai-service/ainlplm/chat"


def test_extract_openai_content():
    result = {"choices": [{"message": {"role": "assistant", "content": "hello"}}]}
    assert extract_assistant_content(result) == "hello"


def test_extract_ccb_content():
    result = {"choices": [{"messages": {"role": "assistant", "content": "1+1等于2"}}]}
    assert extract_assistant_content(result) == "1+1等于2"


def test_ccb_gateway_parse():
    inner = {"choices": [{"messages": {"content": "ok", "role": "assistant"}}]}
    outer = {
        "C-API-Status": "00",
        "C-Response-Code": "000000000000",
        "C-Response-Desc": "成功",
        "C-Response-Body": {"codeid": "20000", "Data_Enqr_Rslt": json.dumps(inner, ensure_ascii=False)},
    }
    parsed = _ccb_validate_gateway(outer)
    assert extract_assistant_content(parsed) == "ok"


def test_sse_line():
    chunk = parse_sse_data_line('data: {"choices":[{"delta":{"content":"x"}}]}')
    assert extract_delta_from_chunk(chunk) == "x"


def main():
    test_normalize_openai_url()
    test_extract_openai_content()
    test_extract_ccb_content()
    test_ccb_gateway_parse()
    test_sse_line()
    print("All llm_client tests passed.")


if __name__ == "__main__":
    main()
