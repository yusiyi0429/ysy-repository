"""
字段别名表（quality_report.py 与 excel_to_skill.py 共用，消除两份别名表漂移）

L1 显性结论层 / L2 判断上下文层 / L3 证据层
"""

FIELD_ALIASES = {
    # L1 显性结论层
    "知识编号": ["知识编号", "编号", "KN编号", "知识ID", "knowledge_id"],
    "知识分类": ["知识分类", "分类", "类型", "知识要点", "知识"],
    "知识描述": [
        "知识描述", "描述", "知识内容", "内容", "数据规则", "具体方案",
        "具体方法", "场景说明", "子场景说明", "知识说明",
    ],
    "适用条件": ["适用条件", "触发条件", "条件"],
    "判断逻辑": ["判断逻辑", "判断规则", "逻辑", "规则引用"],
    "反模式/踩坑提示": ["反模式/踩坑提示", "反模式", "踩坑提示", "注意事项"],
    # L2 判断上下文层（隐性注释写入的目标列）
    "经验判断": ["经验判断", "专家经验判断"],
    "适用边界": ["适用边界", "适用边界/例外"],
    "例外情形": ["例外情形", "例外场景", "破例场景"],
    # L3 证据层
    "来源文档": ["来源文档", "来源"],
    "来源位置": ["来源位置", "位置", "页码"],
    "原文摘录": ["原文摘录", "摘录"],
    "置信度": ["置信度", "可信度"],
    "贡献专家": ["贡献专家", "贡献人"],
    "确认专家": ["确认专家", "确认人"],
    "证据数": ["证据数", "案例支撑数", "evidence_count"],
    "突破数": ["突破数", "被突破数", "break_count"],
    "备注": ["备注", "说明", "修订说明"],
}

REVISION_COLUMN_MARKERS = ("修订状态", "原始内容", "修订内容", "修订说明", "修订时间")


def resolve_header(raw_header: str) -> str:
    """将 Excel 表头文本解析为规范字段名。"""
    if not raw_header:
        return ""
    raw = str(raw_header).strip()
    for canonical, aliases in FIELD_ALIASES.items():
        if raw in aliases:
            return canonical
    return raw


def match_alias(header: str, canonical_key: str) -> bool:
    """检查 header 是否匹配某个 canonical 字段的任一别名。"""
    if not header:
        return False
    h = str(header).strip()
    return h in FIELD_ALIASES.get(canonical_key, [])


def iter_knowledge_columns(header_map: dict) -> dict:
    """从 {表头: col} 中筛选出知识字段列。"""
    result = {}
    for h, col in header_map.items():
        canonical = resolve_header(h)
        if canonical and canonical not in result:
            result[canonical] = (h, col)
    return result
