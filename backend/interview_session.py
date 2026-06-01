"""
结构化访谈模块（方案三）
将 schema 中的访谈方法论落地为产品：案例反推 / 对比追问 / 极限假设。
"""

from __future__ import annotations

import json
import logging

_logger = logging.getLogger("tacit_knowledge")

INTERVIEW_METHODS = {
    "case_reverse": {
        "name": "案例反推法",
        "prompt": (
            "你是一位资深知识工程访谈者。请对下面这条知识进行「案例反推」追问，"
            "目标是从专家那里挖出隐性判断经验。\n\n"
            "追问规则：\n"
            "1. 假设有一个真实案例正好「突破」了这条知识——请构造一个逼真的反例场景\n"
            "2. 问专家：在这个反例中，你会怎么判断？和标准流程有什么不同？\n"
            "3. 追问输出应为 JSON 数组，每条问一个具体问题\n\n"
            '格式：[{{"question": "追问内容", "category": "经验判断|适用边界|例外情形", "hint": "为什么问这个"}}]\n'
            "要求：生成 2-3 个追问，优先触及专家没说清楚的灰色地带。"
        ),
    },
    "contrast_probe": {
        "name": "对比追问法",
        "prompt": (
            "你是一位资深知识工程访谈者。请对下面这条知识进行「对比追问」，"
            "通过构造相似但不同的场景，找到这条知识的真正边界。\n\n"
            "追问规则：\n"
            "1. 构造一个和当前场景高度相似但有细微差异的情况\n"
            "2. 问专家：这两个场景的差异点在哪？为什么会导致判断不同？\n"
            "3. 追问输出应为 JSON 数组\n\n"
            '格式：[{{"question": "追问内容", "category": "经验判断|适用边界|例外情形", "hint": "为什么问这个"}}]\n'
            "要求：生成 2-3 个追问，差异点要足够细。"
        ),
    },
    "limit_hypothesis": {
        "name": "极限假设法",
        "prompt": (
            "你是一位资深知识工程访谈者。请对下面这条知识进行「极限假设」追问，"
            "通过推演极端情况来找出规则的失效边界。\n\n"
            "追问规则：\n"
            "1. 假设某个关键条件推到极限值（如时间极长/极短、金额极大/极小等）\n"
            "2. 问专家：在极限情况下，这条知识还适用吗？如果不适用，为什么？\n"
            "3. 追问输出应为 JSON 数组\n\n"
            '格式：[{{"question": "追问内容", "category": "经验判断|适用边界|例外情形", "hint": "为什么问这个"}}]\n'
            "要求：生成 2-3 个追问。"
        ),
    },
}


def build_interview_prompt(method: str, knowledge_item: dict) -> tuple[str, str]:
    """构建访谈 system prompt 和 user prompt。"""
    method_cfg = INTERVIEW_METHODS.get(method, INTERVIEW_METHODS["case_reverse"])
    desc = knowledge_item.get("知识描述", knowledge_item.get("content", "未命名知识"))
    condition = knowledge_item.get("适用条件", "")
    logic = knowledge_item.get("判断逻辑", "")
    category = knowledge_item.get("知识分类", "")

    item_text = f"知识：{desc}\n"
    if category:
        item_text += f"分类：{category}\n"
    if condition:
        item_text += f"适用条件：{condition}\n"
    if logic:
        item_text += f"判断逻辑：{logic}\n"

    system_prompt = method_cfg["prompt"]
    user_prompt = f"请对以下知识条目进行{method_cfg['name']}追问：\n\n{item_text}"

    return system_prompt, user_prompt


def parse_interview_result(raw_text: str) -> list[dict]:
    """解析 LLM 返回的追问列表。"""
    text = (raw_text or "").strip()
    # 尝试提取 JSON 数组
    try:
        arr = json.loads(text)
        if isinstance(arr, list):
            return _normalize_probes(arr)
    except json.JSONDecodeError:
        pass
    # 尝试提取 markdown 代码块中的 JSON
    import re
    m = re.search(r"```(?:json)?\s*(\[.*?\])\s*```", text, re.DOTALL)
    if m:
        try:
            arr = json.loads(m.group(1))
            if isinstance(arr, list):
                return _normalize_probes(arr)
        except json.JSONDecodeError:
            pass
    # 兜底：返回一条全文追问
    return [{"question": text[:200], "category": "经验判断", "hint": "LLM 输出非标准 JSON，已截取前 200 字符"}]


def _normalize_probes(probes: list) -> list[dict]:
    result = []
    for p in probes:
        if isinstance(p, dict):
            result.append({
                "question": str(p.get("question", "")).strip(),
                "category": str(p.get("category", "经验判断")).strip(),
                "hint": str(p.get("hint", "")).strip(),
            })
    return [p for p in result if p["question"]]
