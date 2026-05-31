#!/usr/bin/env python3
"""端到端验证：Step1 -> Step2缓存上传 -> 刷新后Step3/Step4，不依赖浏览器。"""

import json
import os
import sys
import uuid
from pathlib import Path

import requests

BASE = os.environ.get("E2E_BASE", "http://127.0.0.1:5000")
ROOT = Path(__file__).resolve().parent.parent
SAMPLES = ROOT.parent / "data" / "samples"
WORKSPACE = Path(os.environ.get("TEMP", os.path.expanduser("~"))) / "tacit_knowledge_app"


def ok(msg):
    print(f"[OK] {msg}")


def fail(msg):
    print(f"[FAIL] {msg}")
    return False


def main():
    errors = []

    # health + download guard
    try:
        r = requests.get(f"{BASE}/api/health", timeout=5)
        r.raise_for_status()
        ok("health")
        r = requests.get(f"{BASE}/downloads/pipelines.json", timeout=10)
        if r.status_code == 403:
            ok("downloads blocks pipelines.json")
        else:
            fail(f"downloads should 403 pipelines.json, got {r.status_code}")
            errors.append("download_guard")
    except Exception as e:
        fail(f"health: {e}")
        return 1

    # 1) create pipeline
    r = requests.post(
        f"{BASE}/api/pipelines",
        json={"name": f"E2E链路测试-{uuid.uuid4().hex[:6]}", "scenario": "科技金融", "domain": "科技金融"},
        timeout=30,
    )
    data = r.json()
    if data.get("status") != "ok":
        fail(f"create pipeline: {data}")
        return 1
    pid = data["pipeline"]["id"]
    ok(f"pipeline id={pid}")

    # 2) step1 generate
    step1_name = (SAMPLES / "Step1_场景名称.txt").read_text(encoding="utf-8").strip()
    step1_content = (SAMPLES / "Step1_场景内容.txt").read_text(encoding="utf-8").strip()
    subs = []
    csv_path = SAMPLES / "Step1_子场景列表.csv"
    if csv_path.exists():
        import csv
        with open(csv_path, encoding="utf-8") as f:
            for row in csv.DictReader(f):
                subs.append({"name": row.get("子场景名称", ""), "content": row.get("子场景内容", "")})
    if not subs:
        subs = [{"name": "科技企业普惠贷款", "content": "贷前尽调与贷后管理"}]

    r = requests.post(
        f"{BASE}/api/step1/generate",
        data={
            "scenario_name": step1_name,
            "scenario_content": step1_content,
            "sub_scenarios": json.dumps(subs, ensure_ascii=False),
            "pipeline_id": pid,
        },
        timeout=120,
    )
    data = r.json()
    if data.get("status") != "success":
        fail(f"step1: {data.get('error', data)}")
        errors.append("step1")
    else:
        ok(f"step1 -> {data.get('file_name')}")

    # 3) step2 cache upload (simulate user picking file)
    doc_path = SAMPLES / "Step1_场景内容.txt"
    with open(doc_path, "rb") as f:
        r = requests.post(
            f"{BASE}/api/files/cache_upload",
            files={"file": ("source.txt", f.read(), "text/plain")},
            data={"pipeline_id": pid, "step": "2"},
            timeout=60,
        )
    cache_data = r.json()
    if cache_data.get("status") != "ok":
        fail(f"cache_upload step2: {cache_data}")
        errors.append("cache_step2")
    else:
        cached2 = cache_data["file_name"]
        ok(f"step2 cached -> {cached2}")

    # 4) simulate refresh: GET pipeline only
    r = requests.get(f"{BASE}/api/pipelines/{pid}", timeout=30)
    pdata = r.json()
    sd = pdata.get("pipeline", {}).get("step_data", {})
    if sd.get("step2_cached_file") != cached2:
        fail(f"refresh lost step2_cached: {sd.get('step2_cached_file')} != {cached2}")
        errors.append("refresh_step2")
    else:
        ok("refresh后 step2_cached_file 仍在")

    # 5) step2 prev_output
    r = requests.get(f"{BASE}/api/step2/prev_output", params={"pipeline_id": pid}, timeout=30)
    s2prev = r.json()
    if not s2prev.get("has_output"):
        fail(f"step2 prev_output: {s2prev}")
        errors.append("step2_prev")
    else:
        ok("step2 prev_output has_output")

    # 6) step2 execute WITHOUT re-upload (cached_file only) - needs LLM
    r = requests.post(
        f"{BASE}/api/skills/execute",
        data={
            "skill_id": "knowledge-extraction",
            "pipeline_id": pid,
            "style": "标准萃取",
            "cached_file": cached2,
        },
        timeout=300,
    )
    s2exec = r.json()
    if s2exec.get("status") not in ("ok", "success"):
        fail(f"step2 execute (cached only): {s2exec.get('error', s2exec)}")
        errors.append("step2_llm")
    else:
        dl = s2exec.get("download_name") or ""
        ok(f"step2 execute -> {dl}")
        if not (dl.startswith("preextract_") and dl.endswith(".xlsx")):
            fail(f"step2 download_name must be preextract_*.xlsx, got {dl}")
            errors.append("step2_name_invariant")
        if dl.startswith("template_"):
            fail("step2 must not return template_ skeleton as pre-extract")
            errors.append("step2_template_leak")
        if dl:
            p = WORKSPACE / dl
            if p.exists() and p.suffix == ".xlsx":
                ok("step2 xlsx exists on disk")
            else:
                fail(f"step2 file missing: {dl}")
                errors.append("step2_file")

    # refresh again - step2_output_file
    r = requests.get(f"{BASE}/api/pipelines/{pid}", timeout=30)
    sd = r.json()["pipeline"]["step_data"]
    if not sd.get("step2_output_file"):
        fail("step2_output_file not persisted")
        errors.append("step2_persist")
    else:
        ok(f"step2_output_file persisted: {sd.get('step2_output_file')}")

    # 7) step3 cache expert doc
    expert_path = SAMPLES / "Step3_专家会议纪要.txt"
    with open(expert_path, "rb") as f:
        r = requests.post(
            f"{BASE}/api/files/cache_upload",
            files={"file": ("expert.txt", f.read(), "text/plain")},
            data={"pipeline_id": pid, "step": "3"},
            timeout=60,
        )
    c3 = r.json()
    if c3.get("status") != "ok":
        fail(f"cache_upload step3: {c3}")
        errors.append("cache_step3")
    else:
        cached3 = c3["file_name"]
        ok(f"step3 cached -> {cached3}")

    r = requests.get(f"{BASE}/api/pipelines/{pid}", timeout=30)
    sd = r.json()["pipeline"]["step_data"]
    if sd.get("step3_cached_file") != cached3:
        fail("refresh lost step3_cached")
        errors.append("refresh_step3")
    else:
        ok("refresh后 step3_cached_file 仍在")

    # 8) step3 execute cached only
    r = requests.post(
        f"{BASE}/api/skills/execute",
        data={
            "skill_id": "knowledge-revision",
            "pipeline_id": pid,
            "style": "标准修订",
            "expert_text": "",
            "expert_cached_file": cached3,
        },
        timeout=300,
    )
    s3exec = r.json()
    if s3exec.get("status") not in ("ok", "success"):
        fail(f"step3 execute: {s3exec.get('error', s3exec)}")
        errors.append("step3_llm")
    else:
        rev = s3exec.get("download_name") or s3exec.get("output_file") or ""
        ok(f"step3 execute -> {rev}")
        if rev and not (rev.startswith("revision_") or rev.startswith("edited_step3_")):
            fail(f"step3 output must be revision_*.xlsx, got {rev}")
            errors.append("step3_name_invariant")

    # 9) step4 finalize (uses step3 revision as base) - cached expert only
    step4_text = (SAMPLES / "Step4_终审确认表.txt").read_text(encoding="utf-8")[:1500]
    with open(expert_path, "rb") as f:
        r = requests.post(
            f"{BASE}/api/files/cache_upload",
            files={"file": ("step4_expert.txt", f.read(), "text/plain")},
            data={"pipeline_id": pid, "step": "4"},
            timeout=60,
        )
    c4 = r.json()
    cached4 = c4.get("file_name", "") if c4.get("status") == "ok" else ""

    r = requests.post(
        f"{BASE}/api/step4/finalize",
        data={
            "pipeline_id": pid,
            "style": "标准修订",
            "expert_text": step4_text,
            "expert_cached_file": cached4,
        },
        timeout=300,
    )
    s4 = r.json()
    if s4.get("status") != "success":
        err = s4.get("error", "")
        if "MergedCell" in err:
            fail(f"step4 MergedCell STILL FAILS: {err}")
            errors.append("step4_mergedcell")
        elif "LLM" in err or "连接" in err:
            fail(f"step4 (LLM/network): {err[:200]}")
            errors.append("step4_llm")
        else:
            fail(f"step4: {err}")
            errors.append("step4_other")
    else:
        final = s4.get("download_name") or s4.get("output_file") or ""
        ok(f"step4 finalize -> {final} (no MergedCell error)")
        if final and not (final.startswith("final_") or final.startswith("edited_step4_")):
            fail(f"step4 output must be final_*.xlsx, got {final}")
            errors.append("step4_name_invariant")
        p = WORKSPACE / final
        if p.exists():
            ok("step4 xlsx on disk")

    # 11) step5 compile SKILL.md
    r = requests.post(
        f"{BASE}/api/step5/compile",
        data={"pipeline_id": pid},
        timeout=120,
    )
    s5 = r.json()
    if s5.get("status") != "success":
        fail(f"step5 compile: {s5.get('error', s5.get('message', s5))}")
        errors.append("step5_compile")
    else:
        km = s5.get("knowledge_count", 0)
        ok(f"step5 compile -> {km} knowledge items, download={s5.get('download_url', '')}")
        if km <= 0:
            fail("step5 compile: knowledge_count is 0")
            errors.append("step5_empty")

    # refresh after step3/4
    r = requests.get(f"{BASE}/api/pipelines/{pid}", timeout=30)
    sd = r.json()["pipeline"]["step_data"]
    if sd.get("step3_revision_file"):
        ok(f"step3_revision_file persisted: {sd.get('step3_revision_file')}")
    if sd.get("step4_final_file") or sd.get("step4_download_url"):
        ok("step4 persisted in pipeline")

    # 10) MergedCell unit test offline
    try:
        sys.path.insert(0, str(ROOT))
        from revision_processor import process_workbook

        step3_file = sd.get("step3_revision_file") or sd.get("step2_output_file")
        if step3_file and (WORKSPACE / step3_file).exists():
            out = WORKSPACE / f"e2e_final_{uuid.uuid4().hex[:8]}.xlsx"
            notes = [
                {
                    "sheet": "Sheet1",
                    "row": 3,
                    "col": 2,
                    "action": "modify",
                    "old_value": "test",
                    "new_value": "modified",
                    "note": "e2e",
                }
            ]
            process_workbook(str(WORKSPACE / step3_file), notes, str(out))
            ok(f"offline revision_processor -> {out.name} (MergedCell safe)")
        else:
            fail("no step3 file for offline revision test")
    except Exception as e:
        if "MergedCell" in str(e):
            fail(f"offline revision_processor MergedCell: {e}")
            errors.append("offline_merged")
        else:
            fail(f"offline revision_processor: {e}")

    print("\n=== SUMMARY ===")
    if errors:
        print("Failed checks:", ", ".join(errors))
        return 1
    print("All critical checks passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
