# -*- coding: utf-8 -*-
"""
sw_categorize.py �?共享模块: 零件"采购类别"属�?�自动判�?+ �?��写入

设�?原则 (2026-08-05 皇上钦定):
- 已有值优�? 绝不覆盖用户已填�?采购类别"
- 已有值为空时, 按文件命名�?则自动判�?+ 写入"配置特定"�?��义属�?- 命名规则�?AutoBOM setting.json 模板对齐:
    GB/T ××××_...  �?标准�?    ST_...          �?外购�?    SZ.xxxx.xx.xx_... �?�?���?    *.sldasm        �?装配�?    *.slddrw        �?工程�?    其余             �?其他

用法:
    from sw_categorize import judge_by_name, ensure_purchasing_category, ensure_for_parts
    categ = judge_by_name(filename)
    ensure_purchasing_category(sw, part_path)   # 单件, 返回判定�?已有或新�?
    ensure_for_parts(sw, [path1, path2, ...])   # 批量, 返回 dict[path] -> categ
"""
import os
import re

PROP_NAME = "采购类别"

# 规则顺序很重�? 先匹配更具体�?(GB/T �?ST / SZ 都可�? SZ �?GB 更具�?
_RULES = [
    (re.compile(r"^GB[/／╱][ ]?T[ ]?\d", re.IGNORECASE), "标准件"),
    (re.compile(r"^ST_", re.IGNORECASE), "外购件"),
    (re.compile(r"^SZ\.\d", re.IGNORECASE), "自制件"),
]

_EXT_CATEG = {
    ".sldasm": "装配体",
    ".slddrw": "工程图",
}


def judge_by_name(filename):
    """按文件名规则判定采购类别 (不�?零件, �?��符串)�?    返回: str (标准�?外购�?�?���?装配�?工程�?其他)"""
    base = os.path.splitext(filename)[0]
    ext = os.path.splitext(filename)[1].lower()

    # 扩展名优�?(装配�?工程�?
    if ext in _EXT_CATEG:
        return _EXT_CATEG[ext]

    # 命名规则
    for pat, c in _RULES:
        if pat.match(base):
            return c
    return "其他"


def _get_active_config_name(model):
    try:
        ac = model.ConfigurationManager.ActiveConfiguration
        n = ac.Name
        return n() if callable(n) else n
    except Exception:
        return ""


def _read_categ(model, cfg_name):
    try:
        cpm = model.Extension.CustomPropertyManager(cfg_name)
        if not cpm:
            return ""
        res = cpm.Get3(PROP_NAME, False, "", "")
        # Get3 返回多�?形式:
        #   (True, "value", "resolved") 标准 tuple
        #   (True, "") 删除后空�?        #   True / False (�?�� COM 绑定)
        if isinstance(res, tuple):
            if len(res) >= 2:
                v = res[1]
                if isinstance(v, bool):
                    return ""
                return (v or "") if v is not None else ""
            return ""
        if isinstance(res, bool):
            return ""
        return str(res) if res else ""
    except Exception:
        return ""


def _write_categ(model, cfg_name, value):
    """�?Add3 4�?force overwrite) 写入, 返回�?��成功(�?Get4 复查为准)"""
    try:
        cpm = model.Extension.CustomPropertyManager(cfg_name)
        if not cpm:
            return False
        cpm.Add3(PROP_NAME, 30, value, True)  # 30 = swCustomInfoText
        # Save3
        try:
            model.Save3(0, 0, 0)
        except Exception:
            pass
        # 复查
        actual = _read_categ(model, cfg_name).strip()
        return actual == value
    except Exception:
        return False


def ensure_purchasing_category(sw, part_path, force_judge=False, log=lambda *a: None):
    """
    单件: �?采购类别" �?已有且非空直接返�?
          �?�?force_judge=True) �?按�?则判定并写入"配置特定"属�?? 返回判定值�??    写入失败则返回判定�??不报�? 调用方自己决定�?�么处理)�?
    参数:
        sw: SldWorks.Application COM 对象
        part_path: 零件绝�?�?��
        force_judge: True 时即使已有属性也按�?则�?�?(默�? False, 已有优先)
        log: 日志回调, 默�?静默

    返回: 判定/读取到的采购类别字�?�?(永不为空, 至少返回 "其他")
    """
    import pythoncom
    import win32com.client
    filename = os.path.basename(part_path)
    judged = judge_by_name(filename)

    # 先打�?
    try:
        r = sw.OpenDoc6(part_path, 1, 1, "", 0, 0)
    except Exception as e:
        log(f"  [CATEG] OpenDoc6 失败 {filename}: {e}")
        return judged
    if isinstance(r, tuple):
        r = r[0]
    if not r:
        log(f"  [CATEG] OpenDoc6 返回�?{filename}")
        return judged

    try:
        cfg = _get_active_config_name(r)

        # 读已有�??        existing = ""
        if not force_judge:
            existing = _read_categ(r, cfg).strip()

        if existing:
            log(f"  [CATEG] {filename} 已有 '{existing}' (跳过)")
            return existing

        # 写入
        if _write_categ(r, cfg, judged):
            log(f"  [CATEG] {filename} 写入 '{judged}'")
            # 2026-08-15 修复: 写入必须立即 Save 落盘, 否则随后 CloseDoc 静默丢弃改动
            try:
                r.Save()
            except Exception as e:
                log(f"  [CATEG] {filename} 保存失败: {repr(e)[:50]}")
        else:
            log(f"  [CATEG] {filename} 写入失败, 仍按规则判定 '{judged}'")
        return judged
    finally:
        # 关文�?用无扩展�?title)
        try:
            base = os.path.splitext(filename)[0]
            sw.CloseDoc(base)
        except Exception:
            try:
                t = r.GetTitle()
                sw.CloseDoc(t() if callable(t) else t)
            except Exception:
                pass


def ensure_for_parts(sw, paths, force_judge=False, log=print):
    """批量: 对每�?��件确�?采购类别"已�?�?    返回 dict[abs_path] -> categ_string"""
    out = {}
    written = 0
    kept = 0
    for p in paths:
        if not p or not os.path.exists(p):
            continue
        ext = os.path.splitext(p)[1].lower()
        if ext not in (".sldprt", ".sldasm", ".slddrw"):
            continue
        before = _quick_read(sw, p)
        c = ensure_purchasing_category(sw, p, force_judge=force_judge, log=log)
        out[p] = c
        if before:
            kept += 1
        else:
            written += 1
    log(f"  [CATEG] 批量完成: 已有 {kept}, 新写�?{written}, �?{len(out)}")
    return out


def _quick_read(sw, path):
    """快速读一次"采购类别", 不写入。用于统计"""
    import pythoncom
    import win32com.client
    try:
        r = sw.OpenDoc6(path, 1, 1, "", 0, 0)
        if isinstance(r, tuple):
            r = r[0]
        if not r:
            return ""
        cfg = _get_active_config_name(r)
        return _read_categ(r, cfg).strip()
    except Exception:
        return ""
    finally:
        try:
            base = os.path.splitext(os.path.basename(path))[0]
            sw.CloseDoc(base)
        except Exception:
            pass