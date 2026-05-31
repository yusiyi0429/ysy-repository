#!/usr/bin/env python3
"""无 LLM 的产出物/下载/校验不变量测试。"""

import os
import sys
import uuid
from pathlib import Path

import requests

BASE = os.environ.get("E2E_BASE", "http://127.0.0.1:5000")
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from pipeline_artifacts import (  # noqa: E402
    is_download_allowed,
    is_step1_filename,
    is_step2_preextract_filename,
    is_step3_revision_filename,
    is_step4_final_filename,
    resolve_knowledge_workbook_path,
    validate_step_data_patch,
)


def ok(msg):
    print(f"[OK] {msg}")


def fail(msg):
    print(f"[FAIL] {msg}")
    return False


def test_module_invariants():
    assert is_step1_filename("template_abc.xlsx")
    assert not is_step2_preextract_filename("template_abc.xlsx")
    assert is_step2_preextract_filename("preextract_abc.xlsx")
    assert is_step2_preextract_filename("edited_step2_abc.xlsx")
    assert not is_step3_revision_filename("template_abc.xlsx")
    assert is_step3_revision_filename("revision_abc.xlsx")
    assert is_step4_final_filename("final_abc.xlsx")
    assert not is_download_allowed("pipelines.json")
    assert is_download_allowed("preextract_x.xlsx")
    err = validate_step_data_patch({"step2_output_file": "template_bad.xlsx"})
    assert err and "step2" in err
    import tempfile

    with tempfile.TemporaryDirectory() as td:
        ws = Path(td)
        pre = ws / "preextract_test_resolve.xlsx"
        fin = ws / "final_test_resolve.xlsx"
        pre.write_bytes(b"PK\x03\x04")
        fin.write_bytes(b"PK\x03\x04")
        sd = {"step2_output_file": pre.name, "step4_final_file": fin.name}
        p, key = resolve_knowledge_workbook_path(ws, sd, purpose="align")
        assert p and key == "step4_final_file"
        p2, key2 = resolve_knowledge_workbook_path(ws, {"step2_output_file": pre.name}, purpose="align")
        assert p2 and key2 == "step2_output_file"
    ok("module invariants")


def test_http_invariants():
    errors = []
    try:
        r = requests.get(f"{BASE}/api/health", timeout=5)
        r.raise_for_status()
    except Exception as e:
        fail(f"health: {e}")
        return 1

    r = requests.get(f"{BASE}/downloads/pipelines.json", timeout=10)
    if r.status_code == 403:
        ok("downloads blocks pipelines.json")
    else:
        fail(f"downloads pipelines.json expected 403 got {r.status_code}")
        errors.append("download_protect")

    r = requests.post(
        f"{BASE}/api/pipelines",
        json={"name": f"INV-{uuid.uuid4().hex[:6]}", "scenario": "测试", "domain": "测试"},
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
    if r.status_code == 200 and r.json().get("status") == "ok":
        fail("PUT accepted illegal step2_output_file")
        errors.append("put_step2")
    else:
        ok("PUT rejects template as step2_output_file")

    r = requests.put(
        f"{BASE}/api/pipelines/{pid}",
        json={"step_data": {"step3_revision_file": "template_evil.xlsx"}},
        timeout=30,
    )
    if r.status_code == 200 and r.json().get("status") == "ok":
        fail("PUT accepted illegal step3_revision_file")
        errors.append("put_step3")
    else:
        ok("PUT rejects template as step3_revision_file")

    r = requests.get(f"{BASE}/api/step3/prev_output", params={"pipeline_id": pid}, timeout=30)
    s3 = r.json()
    if s3.get("has_output"):
        fail("step3 prev_output should be false without step2 preextract")
        errors.append("step3_prev_empty")
    else:
        ok("step3 prev_output empty without step2")

    if errors:
        print("Failed:", ", ".join(errors))
        return 1
    print("All artifact invariant checks passed.")
    return 0


def main():
    test_module_invariants()
    return test_http_invariants()


if __name__ == "__main__":
    sys.exit(main())
