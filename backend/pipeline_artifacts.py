"""Pipeline workspace artifact naming and path safety invariants."""

from __future__ import annotations

import os
from pathlib import Path

PROTECTED_WORKSPACE_FILES = frozenset({
    "pipelines.json",
    "custom_models.json",
    "preset_overrides.json",
})

DOWNLOAD_ALLOWED_PREFIXES = (
    "template_",
    "preextract_",
    "revision_",
    "final_",
    "edited_step",
    "upload_",
    "edit_read_",
    "upload_tpl_",
    "SKILL_",
    "COT_",
    "QA_",
    "openclaw_",
    "quality_report_",
    "report_",
    "cache_",
    # 新 Skill 产出报告
    "pattern_mining_",
    "gap_analysis_",
    "freshness_audit_",
)

STEP_OUTPUT_KEYS_BY_STEP = {
    1: (
        "step1_output_file", "step1_download_url",
        "step1_md_file", "step1_md_download_url",
        "step1_knowledge_columns", "step1_output_format",
        "step1_template_source", "step1_template_name",
    ),
    2: (
        "step2_output_file", "step2_download_url",
        "step2_md_file", "step2_md_download_url",
        "step2_extracted_count", "skill_extract_result", "skill_extract_style",
    ),
    # UI 第 3 步「知识对齐」产出 final_*.xlsx；step3_revision_* 为旧五步流兼容
    3: (
        "step4_final_file", "step4_download_url", "step4_md_file", "step4_md_download_url",
        "step4_final_notes", "step4_final_style", "step4_final_count",
        "step3_revision_file", "step3_download_url", "step3_md_file", "step3_md_download_url",
        "step3_revision_notes", "step3_revision_style", "step3_revision_count", "step3_excel_path",
    ),
    4: (
        "step5_skill_file", "step5_download_url",
        "step5_cot_file", "step5_cot_download_url",
        "step5_qa_file", "step5_qa_download_url", "step5_qa_md_file", "step5_qa_md_download_url",
        "step5_openclaw_manifest_file", "step5_openclaw_manifest_url",
        "step5_quality_file", "step5_quality_url",
    ),
}


def basename_only(name: str) -> str:
    return os.path.basename((name or "").strip().replace("\\", "/"))


def is_step1_filename(name: str) -> bool:
    n = basename_only(name).lower()
    return n.endswith(".xlsx") and n.startswith("template_")


def is_step2_preextract_filename(name: str) -> bool:
    n = basename_only(name).lower()
    return n.endswith(".xlsx") and (n.startswith("preextract_") or n.startswith("edited_step2_"))


def is_step3_revision_filename(name: str) -> bool:
    n = basename_only(name).lower()
    return n.endswith(".xlsx") and (n.startswith("revision_") or n.startswith("edited_step3_"))


def is_step4_final_filename(name: str) -> bool:
    n = basename_only(name).lower()
    return n.endswith(".xlsx") and (n.startswith("final_") or n.startswith("edited_step4_"))


def resolve_knowledge_workbook_path(
    workspace: Path,
    step_data: dict,
    *,
    purpose: str = "compile",
) -> tuple[Path | None, str]:
    """Resolve the best on-disk knowledge Excel for align/compile.

    purpose=align: prefer latest alignment draft, then legacy revision, then Step2 preextract.
    purpose=compile: prefer final_*.xlsx, then legacy revision_*.xlsx, then preextract (smoke only).
    """
    if not isinstance(step_data, dict):
        return None, ""

    candidates = (
        ("step4_final_file", is_step4_final_filename),
        ("step3_revision_file", is_step3_revision_filename),
        ("step2_output_file", is_step2_preextract_filename),
    )

    for key, validator in candidates:
        raw = step_data.get(key, "")
        if not raw or not validator(str(raw)):
            continue
        resolved = safe_workspace_path(workspace, str(raw), must_exist=True)
        if resolved:
            return resolved, key
    return None, ""


def is_download_allowed(name: str) -> bool:
    base = basename_only(name)
    if not base or base in PROTECTED_WORKSPACE_FILES:
        return False
    return any(base.startswith(p) for p in DOWNLOAD_ALLOWED_PREFIXES)


def is_cache_filename(name: str) -> bool:
    """Uploaded source files cached for Step2/3/4 (cache_s2_*, cache_s3_*, etc.)."""
    base = basename_only(name)
    return bool(base) and base.startswith("cache_")


def resolve_cache_file_path(workspace: Path, name: str) -> Path | None:
    """Resolve a cached upload path; rejects non-cache and protected names."""
    if not is_cache_filename(name):
        return None
    return safe_workspace_path(workspace, name, must_exist=True)


def safe_workspace_path(workspace: Path, name: str, *, must_exist: bool = True) -> Path | None:
    base = basename_only(name)
    if not base or base in PROTECTED_WORKSPACE_FILES:
        return None
    try:
        root = workspace.resolve()
        path = (workspace / base).resolve()
        path.relative_to(root)
    except ValueError:
        return None
    if must_exist and not path.is_file():
        return None
    return path


def resolve_client_excel_path(workspace: Path, file_path: str, file_name: str = "") -> Path | None:
    """Resolve excel editor path: prefer basename file_name under workspace."""
    if file_name:
        return safe_workspace_path(workspace, file_name, must_exist=True)
    raw = (file_path or "").strip()
    if not raw:
        return None
    base = basename_only(raw)
    candidate = safe_workspace_path(workspace, base, must_exist=True)
    if candidate:
        return candidate
    try:
        p = Path(raw).resolve()
        p.relative_to(workspace.resolve())
        if p.is_file():
            return p
    except (ValueError, OSError):
        pass
    return None


def validate_step_data_patch(patch: dict) -> str | None:
    """Return error message if step_data output fields violate invariants."""
    if not isinstance(patch, dict):
        return None
    checks = (
        ("step1_output_file", is_step1_filename),
        ("step2_output_file", is_step2_preextract_filename),
        ("step3_revision_file", is_step3_revision_filename),
        ("step4_final_file", is_step4_final_filename),
    )
    for key, fn in checks:
        val = patch.get(key)
        if val and not fn(str(val)):
            return f"非法 {key}: {val}"
    for key in ("step1_download_url", "step2_download_url", "step3_download_url", "step4_download_url", "step5_download_url"):
        url = patch.get(key)
        if not url:
            continue
        part = str(url).split("/")[-1]
        if part and not is_download_allowed(part):
            return f"非法下载路径 {key}: {url}"
    return None


def downstream_output_keys(from_step: int) -> list[str]:
    keys = []
    for step, names in STEP_OUTPUT_KEYS_BY_STEP.items():
        if step > from_step:
            keys.extend(names)
    return keys


def auxiliary_step_data_keys(from_step: int) -> list[str]:
    """Non-output step_data keys to clear on rollback / upstream regenerate."""
    keys = []
    for step in range(from_step, 6):
        keys.extend((
            f"step{step}_cached_file",
            f"step{step}_excel_path",
            f"step{step}_preview_name",
            f"step{step}_preview_url",
        ))
    if from_step <= 4:
        keys.extend(("step5_quality_file", "step5_quality_url"))
    return keys


def keys_to_clear_from_step(from_step: int) -> list[str]:
    return list(dict.fromkeys(downstream_output_keys(from_step) + auxiliary_step_data_keys(from_step)))
