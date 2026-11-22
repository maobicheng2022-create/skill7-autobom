# -*- coding: utf-8 -*-
"""型材材质归类（2026-11-23 钦定：只区分 201 与 Q235镀锌管）

数据来源：零件自定义属性「材料」的**解析值**（SW 材质名，如 Q235 / SUS201 / 201不锈钢）。
调用方拿到原始材质名后，用 classify() 归一成两个采购/生产口径的桶。

规则（顺序敏感）：
    Q235 / 镀锌            -> Q235镀锌管
    201 / SUS201 / 不锈钢   -> 201
    304 / SUS304           -> 304（保留，但不参与两桶合并）
    其他                    -> 原样返回（便于发现新材质）
    空                      -> 未注明
"""

BUCKET_201 = "201"
BUCKET_Q235 = "Q235镀锌管"
BUCKET_UNKNOWN = "未注明"


def classify(raw):
    """把 SW 材质名归一成材质桶。"""
    s = str(raw or "").strip()
    if not s:
        return BUCKET_UNKNOWN
    low = s.lower()
    if "q235" in low or "镀锌" in s:
        return BUCKET_Q235
    if "201" in low or "sus201" in low or "不锈钢" in s:
        return BUCKET_201
    if "304" in low:
        return "304"
    return s


def short(raw):
    """打印/目录名用的短标签（与桶名一致）。"""
    return classify(raw)