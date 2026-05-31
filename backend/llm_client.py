"""LLM client adapters: OpenAI-compatible and CCB ainlplm gateway."""

import json
import time
import uuid
from typing import Any, Dict, Generator, Iterator, List, Optional, Union

import requests as http_requests


API_TYPE_OPENAI = "openai"
API_TYPE_CCB = "ccb_ainlplm"

# Retry configuration for transient LLM API failures
DEFAULT_RETRY_MAX = 2
DEFAULT_RETRY_BACKOFF_BASE = 2.0
RETRYABLE_HTTP_STATUS = {429, 500, 502, 503, 504}
RETRYABLE_EXCEPTIONS = (
    http_requests.Timeout,
    http_requests.ConnectionError,
)


class LlmApiError(Exception):
    """Raised when LLM gateway or business layer returns an error."""

    def __init__(self, message: str, *, api_status=None, response_code=None, codeid=None, desc=None):
        super().__init__(message)
        self.api_status = api_status
        self.response_code = response_code
        self.codeid = codeid
        self.desc = desc


def normalize_llm_url(url: str, api_type: str = API_TYPE_OPENAI) -> str:
    """Ensure OpenAI-compatible chat completions endpoint URL."""
    if api_type == API_TYPE_CCB:
        return (url or "").strip().rstrip("/")
    url = (url or "").strip().rstrip("/")
    if not url:
        return url
    if url.endswith("/chat/completions"):
        return url
    if url.endswith("/v1"):
        return url + "/chat/completions"
    # Bare hostname or custom base path → append standard OpenAI-compatible suffix
    return url + "/v1/chat/completions"


def extract_assistant_content(result: Any) -> str:
    """Extract assistant text from OpenAI or CCB inner response dict."""
    if not isinstance(result, dict):
        return ""
    choices = result.get("choices") or []
    if not choices:
        return ""
    first = choices[0] if isinstance(choices, list) else {}
    if not isinstance(first, dict):
        return ""
    # CCB: choices[0].messages.content; OpenAI: choices[0].message.content
    msg = first.get("messages") or first.get("message") or {}
    if isinstance(msg, dict):
        content = msg.get("content")
        if content is not None:
            return str(content)
    # Streaming delta variants
    delta = first.get("delta")
    if isinstance(delta, dict) and delta.get("content") is not None:
        return str(delta["content"])
    return ""


def extract_delta_from_chunk(chunk: Any) -> str:
    """Extract incremental text from a single stream chunk dict."""
    if not isinstance(chunk, dict):
        return ""
    choices = chunk.get("choices") or []
    if not choices or not isinstance(choices, list):
        return ""
    first = choices[0]
    if not isinstance(first, dict):
        return ""
    delta = first.get("delta")
    if isinstance(delta, dict) and delta.get("content"):
        return str(delta["content"])
    msg = first.get("messages") or first.get("message") or {}
    if isinstance(msg, dict) and msg.get("content"):
        return str(msg["content"])
    return ""


def parse_sse_data_line(line: str) -> Optional[dict]:
    """Parse one SSE data line into JSON object, or None."""
    line = (line or "").strip()
    if not line or line.startswith(":"):
        return None
    if line.startswith("data:"):
        line = line[5:].strip()
    if not line or line == "[DONE]":
        return None
    try:
        return json.loads(line)
    except json.JSONDecodeError:
        return None


def _ccb_validate_gateway(body: dict) -> dict:
    """Validate CCB outer response and return inner enquiry result dict."""
    # Some environments may return inner model response directly
    # (without C-* gateway envelope). In that case, pass through.
    if "C-API-Status" not in body and "C-Response-Body" not in body:
        return body

    api_status = body.get("C-API-Status", "")
    response_code = body.get("C-Response-Code", "")
    response_desc = body.get("C-Response-Desc", "")
    if api_status != "00":
        raise LlmApiError(
            f"网关调用失败: C-API-Status={api_status}, {response_desc}",
            api_status=api_status,
            response_code=response_code,
            desc=response_desc,
        )
    if response_code != "000000000000":
        raise LlmApiError(
            f"应答异常: C-Response-Code={response_code}, {response_desc}",
            api_status=api_status,
            response_code=response_code,
            desc=response_desc,
        )
    resp_body = body.get("C-Response-Body") or {}
    if not isinstance(resp_body, dict):
        raise LlmApiError("C-Response-Body 格式无效", api_status=api_status, response_code=response_code)
    codeid = str(resp_body.get("codeid", ""))
    if codeid != "20000":
        raise LlmApiError(
            f"业务处理失败: codeid={codeid}",
            api_status=api_status,
            response_code=response_code,
            codeid=codeid,
            desc=response_desc,
        )
    raw = resp_body.get("Data_Enqr_Rslt")
    if raw is None:
        raise LlmApiError("缺少 Data_Enqr_Rslt", codeid=codeid)
    if isinstance(raw, str):
        try:
            inner = json.loads(raw)
        except json.JSONDecodeError as e:
            raise LlmApiError(f"Data_Enqr_Rslt 解析失败: {e}", codeid=codeid) from e
    elif isinstance(raw, dict):
        inner = raw
    else:
        raise LlmApiError("Data_Enqr_Rslt 类型无效", codeid=codeid)
    return inner


def _ccb_headers(model_cfg: dict, trace_id: str, tx_serial_no: str) -> dict:
    return {
        "Content-Type": "application/json",
        "Access_Key_Id": model_cfg["api_key"],
        "Tx-Code": model_cfg["tx_code"],
        "Sec-Node-No": model_cfg["sec_node_no"],
        "Trace-Id": trace_id,
        "Tx-Serial-No": tx_serial_no,
    }


def _ccb_payload(model_cfg: dict, messages: list, stream: bool) -> dict:
    data_cntnt = {"messages": messages}
    payload = {
        "Data_cntnt": data_cntnt,
        "stream": stream,
        "model_config": {"model": model_cfg["model"]},
    }
    if model_cfg.get("data_cntnt_stringify"):
        payload["Data_cntnt"] = json.dumps(data_cntnt, ensure_ascii=False)
    fst = model_cfg.get("fst_attr_rmrk") or model_cfg.get("api_key")
    if fst:
        payload["Fst_Attr_Rmrk"] = fst
    return payload


def call_openai(
    model_cfg: dict,
    messages: list,
    stream: bool = False,
    temperature=None,
    max_tokens=None,
):
    """Call OpenAI-compatible endpoint. Returns dict or streaming Response."""
    url = normalize_llm_url(model_cfg["url"], API_TYPE_OPENAI)
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {model_cfg['api_key']}",
    }
    payload = {
        "model": model_cfg["model"],
        "messages": messages,
        "stream": stream,
        "temperature": temperature if temperature is not None else model_cfg.get("temperature", 0.7),
        "max_tokens": max_tokens if max_tokens is not None else model_cfg.get("max_tokens", 4096),
    }
    timeout = model_cfg.get("timeout", 300)
    if stream:
        resp = http_requests.post(url, json=payload, headers=headers, stream=True, timeout=timeout)
        resp.raise_for_status()
        return resp
    resp = http_requests.post(url, json=payload, headers=headers, timeout=timeout)
    resp.raise_for_status()
    return resp.json()


def call_ccb_ainlplm(
    model_cfg: dict,
    messages: list,
    stream: bool = False,
    temperature=None,
    max_tokens=None,
):
    """Call CCB ainlplm gateway. Non-stream returns inner dict; stream returns Response."""
    url = normalize_llm_url(model_cfg["url"], API_TYPE_CCB)
    trace_id = str(uuid.uuid4())
    tx_serial_no = str(uuid.uuid4())
    headers = _ccb_headers(model_cfg, trace_id, tx_serial_no)
    payload = _ccb_payload(model_cfg, messages, stream)
    timeout = model_cfg.get("timeout", 300)

    if stream:
        resp = http_requests.post(url, json=payload, headers=headers, stream=True, timeout=timeout)
        resp.raise_for_status()
        return resp

    resp = http_requests.post(url, json=payload, headers=headers, timeout=timeout)
    resp.raise_for_status()
    body = resp.json()
    return _ccb_validate_gateway(body)


def call_llm(
    model_cfg: dict,
    messages: list,
    stream: bool = False,
    temperature=None,
    max_tokens=None,
):
    """Dispatch to OpenAI or CCB adapter based on api_type."""
    api_type = (model_cfg.get("api_type") or API_TYPE_OPENAI).strip().lower()
    if api_type == API_TYPE_CCB:
        return call_ccb_ainlplm(model_cfg, messages, stream=stream, temperature=temperature, max_tokens=max_tokens)
    return call_openai(model_cfg, messages, stream=stream, temperature=temperature, max_tokens=max_tokens)


def call_llm_with_retry(
    model_cfg: dict,
    messages: list,
    stream: bool = False,
    temperature=None,
    max_tokens=None,
    *,
    max_retries: int = DEFAULT_RETRY_MAX,
    backoff_base: float = DEFAULT_RETRY_BACKOFF_BASE,
):
    """Call LLM with exponential backoff retry on transient failures.

    Retries on: HTTP 429/5xx, timeouts, and connection errors.
    Non-retryable exceptions (LlmApiError with 4xx status) propagate immediately.
    """
    last_error = None
    for attempt in range(max_retries + 1):
        try:
            return call_llm(
                model_cfg, messages, stream=stream,
                temperature=temperature, max_tokens=max_tokens,
            )
        except http_requests.HTTPError as e:
            status = e.response.status_code if hasattr(e, "response") and e.response is not None else 0
            if status in RETRYABLE_HTTP_STATUS and attempt < max_retries:
                last_error = e
                delay = backoff_base ** attempt
                time.sleep(delay)
                continue
            raise
        except RETRYABLE_EXCEPTIONS as e:
            if attempt < max_retries:
                last_error = e
                delay = backoff_base ** attempt
                time.sleep(delay)
                continue
            raise
        except LlmApiError:
            # Business-layer errors (wrong key, auth failure) are not retryable
            raise
    # Should not reach here, but safety net
    raise last_error  # type: ignore[misc]


def iter_openai_stream(resp) -> Generator[str, None, None]:
    """Yield text deltas from OpenAI-style SSE stream."""
    for line in resp.iter_lines(decode_unicode=True):
        if not line:
            continue
        chunk = parse_sse_data_line(line)
        if chunk is None:
            continue
        delta = extract_delta_from_chunk(chunk)
        if delta:
            yield delta


def iter_ccb_stream(resp) -> Generator[str, None, None]:
    """Yield text deltas from CCB gateway SSE stream."""
    for line in resp.iter_lines(decode_unicode=True):
        if not line:
            continue
        chunk = parse_sse_data_line(line)
        if chunk is None:
            continue
        # Outer gateway wrapper on each chunk
        if "C-API-Status" in chunk or "C-Response-Body" in chunk:
            try:
                inner = _ccb_validate_gateway(chunk)
                delta = extract_assistant_content(inner) or extract_delta_from_chunk(inner)
                if delta:
                    yield delta
            except LlmApiError:
                # Partial chunks may not be full gateway responses; try inner body only
                body = chunk.get("C-Response-Body") or {}
                if isinstance(body, dict):
                    raw = body.get("Data_Enqr_Rslt")
                    if isinstance(raw, str):
                        try:
                            inner = json.loads(raw)
                            delta = extract_delta_from_chunk(inner) or extract_assistant_content(inner)
                            if delta:
                                yield delta
                        except json.JSONDecodeError:
                            pass
            continue
        delta = extract_delta_from_chunk(chunk) or extract_assistant_content(chunk)
        if delta:
            yield delta


def iter_llm_stream(model_cfg: dict, resp) -> Generator[str, None, None]:
    """Yield text deltas from streaming response for configured api_type."""
    api_type = (model_cfg.get("api_type") or API_TYPE_OPENAI).strip().lower()
    if api_type == API_TYPE_CCB:
        yield from iter_ccb_stream(resp)
    else:
        yield from iter_openai_stream(resp)


def collect_stream_text(model_cfg: dict, resp) -> str:
    """Aggregate full assistant text from a streaming HTTP response."""
    parts = []
    for delta in iter_llm_stream(model_cfg, resp):
        parts.append(delta)
    return "".join(parts)
