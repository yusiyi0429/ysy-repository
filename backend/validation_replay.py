"""
显性化校验闭环（方案四）
上传历史案例（含已知专家结论），用生成的 SKILL.md/QA 让 LLM 判断，
与专家结论比对，输出命中率/分歧清单。
"""

from __future__ import annotations

import json
from datetime import datetime


def build_validation_prompt(knowledge_text: str, cases: list[dict]) -> tuple[str, str]:
    """构建校验 system prompt 和 user prompt。"""
    system_prompt = (
        "你是一位银行信贷审批专家。请基于下面提供的知识库内容，对每个案例给出你的判断。\n\n"
        "规则：\n"
        "1. 仔细阅读知识库内容，理解其中每条规则的适用条件和判断逻辑\n"
        "2. 对每个案例，先列出你参考了知识库中的哪些规则，再给出结论\n"
        "3. 输出为 JSON 数组，每个案例一条\n\n"
        '格式：[{"case_id": "案例ID", '
        '"prediction": "通过|拒绝|条件通过", '
        '"reasoning": "你的推理过程（引用知识库规则）", '
        '"referenced_rules": ["KN-xxx", ...], '
        '"confidence": "高|中|低"}]\n\n'
        "不要输出任何前后说明文字，只输出 JSON 数组。"
    )

    cases_text = ""
    for c in cases:
        cases_text += f"\n---\n案例ID：{c.get('case_id', '?')}\n"
        cases_text += f"场景描述：{c.get('description', c.get('场景', ''))}\n"
        for k, v in c.items():
            if k not in ("case_id", "description", "场景", "conclusion", "结论", "expert_conclusion"):
                cases_text += f"{k}：{v}\n"

    user_prompt = f"知识库内容：\n{knowledge_text[:8000]}\n\n待判断案例：{cases_text}"

    return system_prompt, user_prompt


def compare_predictions(predictions: list[dict], cases: list[dict]) -> dict:
    """对比 LLM 预测与专家结论。"""
    total = max(len(predictions), 1)
    hits = 0
    mismatches = []
    for pred in predictions:
        case_id = pred.get("case_id", "")
        pred_label = (pred.get("prediction", "") or "").strip()
        # 找对应案例的专家结论
        expert_label = ""
        for c in cases:
            if c.get("case_id") == case_id:
                expert_label = (c.get("结论", c.get("conclusion", c.get("expert_conclusion", ""))) or "").strip()
                break
        pred_norm = _normalize_label(pred_label)
        expert_norm = _normalize_label(expert_label)
        if pred_norm and expert_norm and pred_norm == expert_norm:
            hits += 1
        else:
            mismatches.append({
                "case_id": case_id,
                "prediction": pred_label,
                "expert_conclusion": expert_label,
                "match": pred_norm == expert_norm,
                "reasoning": pred.get("reasoning", ""),
                "referenced_rules": pred.get("referenced_rules", []),
            })
    hit_rate = round(hits / total, 3)
    return {
        "total_cases": total,
        "hits": hits,
        "hit_rate": hit_rate,
        "mismatch_count": len(mismatches),
        "mismatches": mismatches,
    }


def _normalize_label(label: str) -> str:
    s = label.strip().lower()
    if not s:
        return ""
    if any(k in s for k in ("通过", "approve", "yes", "同意")):
        return "通过"
    if any(k in s for k in ("拒绝", "reject", "no", "否决")):
        return "拒绝"
    if any(k in s for k in ("条件", "conditional", "有条件", "附条件")):
        return "条件通过"
    return s


def generate_replay_report(result: dict, cases: list[dict], predictions: list[dict]) -> str:
    lines = [
        "# 显性化校验报告 · 决策回放",
        "",
        f"- 生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M')}",
        f"- 案例数：{result['total_cases']}",
        f"- **命中率：{result['hit_rate'] * 100:.1f}%**（{result['hits']}/{result['total_cases']}）",
        f"- 分歧数：{result['mismatch_count']}",
        "",
        "## 不一致案例分析",
        "",
    ]
    if not result["mismatches"]:
        lines.append("> 所有案例判断与专家结论一致。")
        lines.append("")
    else:
        for i, m in enumerate(result["mismatches"]):
            lines.append(f"### 案例 {m['case_id']}")
            lines.append(f"- **LLM 预测**：{m['prediction']}")
            lines.append(f"- **专家结论**：{m['expert_conclusion']}")
            lines.append(f"- 推理过程：{m.get('reasoning', '')[:200]}")
            rules = m.get("referenced_rules", [])
            if rules:
                lines.append(f"- 引用的规则：{', '.join(rules)}")
            lines.append("")
            lines.append("> 建议：请检查上述规则是否需要更新或补充例外情形。")
            lines.append("")
    return "\n".join(lines)
