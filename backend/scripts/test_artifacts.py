#!/usr/bin/env python3
"""离线校验 pipeline 产出物命名与 API 防护（不依赖 LLM）。"""

import os
import sys
import uuid
from pathlib import Path

import requests

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from pipeline_artifacts import (  # noqa: E402
    is_download_allowed,
    is_step1_filename,
    is_step2_preextract_filename,
    is_step3_revision_filename,
    is_step4_final_filename,
    validate_step_data_patch,
)

BASE = os.environ.get("E2E_BASE", "http://127.0.0.1:5000")


def ok(msg):
    print(f"[OK] {msg}")


def fail(msg):
    print(f"[FAIL] {msg}")
    return False


def test_offline_invariants():
    assert is_step1_filename("template_abc.xlsx")
    assert not is_step2_preextract_filename("template_abc.xlsx")
    assert is_step2_preextract_filename("preextract_abc.xlsx")
    assert is_step2_preextract_filename("edited_step2_abc.xlsx")
    assert is_step3_revision_filename("revision_abc.xlsx")
    assert is_step4_final_filename("final_abc.xlsx")
    assert not is_download_allowed("pipelines.json")
    assert is_download_allowed("preextract_abcd.xlsx")
    err = validate_step_data_patch({"step2_output_file": "template_bad.xlsx"})
    assert err and "step2" in err
    ok("offline naming invariants")


def test_api_guards():
    errors = []
    try:
        r = requests.get(f"{BASE}/api/health", timeout=5)
        r.raise_for_status()
    except Exception as e:
        fail(f"server not up: {e}")
        return 1

    r = requests.get(f"{BASE}/downloads/pipelines.json", timeout=10)
    if r.status_code == 403:
        ok("downloads blocks pipelines.json")
    else:
        fail(f"downloads should 403 pipelines.json, got {r.status_code}")
        errors.append("download_guard")

    r = requests.post(
        f"{BASE}/api/pipelines",
        json={"name": f"ArtifactTest-{uuid.uuid4().hex[:6]}", "scenario": "测试", "domain": "测试"},
        timeout=30,
    )
    data = r.json()
    if data.get("status") != "ok":
        fail(f"create pipeline: {data}")
        return 1
    pid = data["pipeline"]["id"]
    ok(f"pipeline {pid}")

    r = requests.put(
        f"{BASE}/api/pipelines/{pid}",
        json={"step_data": {"step2_output_file": "template_evil.xlsx"}},
        timeout=30,
    )
    put = r.json()
    if put.get("status") == "error" and "step2" in (put.get("error") or ""):
        ok("PUT rejects template as step2_output_file")
    else:
        fail(f"PUT should reject invalid step2: {put}")
        errors.append("put_validate")

    r = requests.post(f"{BASE}/api/pipelines/{pid}/rollback/1", timeout=30)
    if r.json().get("status") == "ok":
        ok("rollback/1")
    else:
        errors.append("rollback")

    r = requests.get(f"{BASE}/api/step1/templates", timeout=10)
    tpl = r.json()
    if tpl.get("status") == "ok" and tpl.get("templates"):
        ok(f"step1 templates: {len(tpl['templates'])}")
    else:
        fail(f"step1 templates: {tpl}")
        errors.append("templates")

    if errors:
        print("Failed:", ", ".join(errors))
        return 1
    print("All artifact API guards passed.")
    return 0


def main():
    test_offline_invariants()
    return test_api_guards()


if __name__ == "__main__":
    sys.exit(main())
