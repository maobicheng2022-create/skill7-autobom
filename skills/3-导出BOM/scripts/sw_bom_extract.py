# -*- coding: utf-8 -*-
"""
sw_bom_extract.py — SW 装配体组件枚举 + 属性读取内核 (路线3)
适配 pywin32 新版 + SW 2022。用 IAssemblyDoc.GetComponents(False) 枚举所有组件。
"""
import os
import sys
import io
import collections

import win32com.client
import pythoncom


def _find_sw_tlb():
    """定位 sldworks.tlb (注册表安装目录优先, 其次常见盘符路径)。"""
    cands = []
    try:
        import winreg
        base = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\SolidWorks")
        i = 0
        while True:
            try:
                sub = winreg.EnumKey(base, i)
            except OSError:
                break
            i += 1
            if not sub.lower().startswith("solidworks"):
                continue
            try:
                k = winreg.OpenKey(base, sub + r"\Setup")
                folder, _ = winreg.QueryValueEx(k, "SolidWorks Folder")
                cands.append(os.path.join(folder, "sldworks.tlb"))
            except OSError:
                continue
    except OSError:
        pass
    for drv in ("E:", "D:", "C:", "F:"):
        cands.append(os.path.join(drv + os.sep, "Program Files",
                                  "SOLIDWORKS Corp", "SOLIDWORKS", "sldworks.tlb"))
    for p in cands:
        if p and os.path.exists(p):
            return p
    return None


def ensure_gen():
    import win32com.client.gencache as gencache
    base = gencache.GetGeneratePath()
    try:
        os.makedirs(base, exist_ok=True)
    except OSError:
        pass

    def _hits():
        try:
            return [os.path.join(base, n) for n in os.listdir(base)
                    if n.startswith("83A33D31-27C5-11CE-BFD4-00400513BB57x0x")
                    and n.endswith(".py")]
        except OSError:
            return []

    hits = _hits()
    if not hits:
        # 2026-10-17 修复: gen_py 在 %TEMP% 下, 会被系统/清理工具删除; 缓存丢失后
        # win32com 退化为后期绑定, OpenDoc6 报 '类型不匹配'(0x80020005) 导致 skill7 步骤A 崩。
        # 原实现直接 raise 要求人工跑 makepy; 改为自动重建缓存(幂等, 找不到 tlb 才报错)。
        tlb = _find_sw_tlb()
        if tlb:
            try:
                from win32com.client import makepy
                makepy.GenerateFromTypeLibSpec(tlb)
            except Exception:
                pass
            hits = _hits()
    if not hits:
        raise RuntimeError("未找到 SW 类型库缓存且自动重建失败,请手动运行: "
                           "python -m win32com.client.makepy <sldworks.tlb 完整路径>")
    modname = "win32com.gen_py." + os.path.splitext(os.path.basename(hits[0]))[0]
    return __import__(modname, fromlist=['IModelDoc2', 'IAssemblyDoc', 'IComponent2'])


_GEN = None


def gen():
    global _GEN
    if _GEN is None:
        _GEN = ensure_gen()
    return _GEN


def wrap_model(model):
    if model is None:
        return None
    try:
        return gen().IModelDoc2(model._oleobj_)
    except Exception:
        return model


def mtype(model):
    t = model.GetType
    return t() if callable(t) else t


def mtitle(model):
    t = model.GetTitle
    return t() if callable(t) else t


def mpath(model):
    t = model.GetPathName
    return t() if callable(t) else t


def open_doc(sw, path, doc_type=1):
    # 2026-10-08 修复: SW 刚启动(4s内) COM 未完全初始化时 OpenDoc6 偶发
    # com_error 类型不匹配(0x80020005)。加重试: 失败等 3s 重试, 最多 4 次。
    import time as _t
    last = None
    for _i in range(4):
        try:
            r = sw.OpenDoc6(path, doc_type, 1, "", 0, 0)
            break
        except Exception as _e:
            last = _e
            _t.sleep(3)
    else:
        raise last
    if isinstance(r, tuple):
        r = r[0]
    return wrap_model(r) if r else None

def config_name(model):
    try:
        cm = model.ConfigurationManager
        ac = cm.ActiveConfiguration
        n = ac.Name
        return n() if callable(n) else n
    except Exception:
        return ""


def read_props(model, names, config=""):
    out = {}
    try:
        ext = model.Extension
        cpm = ext.CustomPropertyManager(config)
        if not cpm:
            return {n: "" for n in names}
        for n in names:
            try:
                r = cpm.Get3(n, False, "", "")
                if isinstance(r, (tuple, list)) and len(r) >= 3:
                    v = r[2] if r[2] else r[1]
                    out[n] = v or ""
                elif isinstance(r, (tuple, list)) and len(r) == 2:
                    out[n] = r[1] or ""
                else:
                    out[n] = r or ""
            except Exception:
                out[n] = ""
    except Exception:
        return {n: "" for n in names}
    return out


def _clean_name(doc, prop_names, path, title):
    """2026-09-19: 名称优先用 SW 属性「名称」(干净名如 称重传感器), 缺省才回退文件名去扩展名。
    修复: 之前用 os.path.basename(path) 带扩展名(ST_称重传感器.SLDPRT) 一路污染到 BOM 表。"""
    try:
        pn = read_props(doc, prop_names, config_name(doc)).get("名称", "") or ""
        if str(pn).strip():
            return str(pn).strip()
    except Exception:
        pass
    if path:
        return os.path.splitext(os.path.basename(path))[0]
    return title


def _bom_excluded(c2, doc):
    """2026-08-13 皇上钦定: 是否"不包括在材料明细表中"(Exclude from BOM)。
    双通道判定:
      1. 组件级 — 装配体里右键组件->零部件属性->不包括在材料明细表中 (IComponent2.ExcludeFromBOM)
      2. 文档级 — 配置特定属性 "Exclude from bill of materials" = TRUE/1/是
    任一命中返回 True → 不出现在任何 BOM 中, 且其整棵子树也排除。"""
    try:
        v = c2.ExcludeFromBOM
        if callable(v):
            v = v()
        if v:
            return True
    except Exception:
        pass
    try:
        cn = config_name(doc)
        cpm = doc.Extension.CustomPropertyManager(cn)
        if cpm:
            for pn in ("Exclude from bill of materials", "不包括在材料明细表中"):
                try:
                    r = cpm.Get(pn)
                    if isinstance(r, (tuple, list)):
                        v = r[1] if len(r) > 1 else ""
                    else:
                        v = r
                    if str(v or "").strip().upper() in ("TRUE", "1", "是", "YES"):
                        return True
                except Exception:
                    continue
    except Exception:
        pass
    return False


def extract_components(sw, prop_names):
    model = wrap_model(sw.ActiveDoc)
    if model is None:
        raise RuntimeError("SolidWorks 没有打开文档")
    root = {"path": mpath(model), "title": mtitle(model),
            "type": mtype(model), "config": config_name(model)}
    rows = []
    if mtype(model) != 2:
        rows.append({"path": mpath(model), "code": os.path.splitext(os.path.basename(mpath(model)))[0] if mpath(model) else mtitle(model),
                     "name": _clean_name(model, prop_names, mpath(model), mtitle(model)),
                     "config": config_name(model), "is_asm": False,
                     "props": read_props(model, prop_names)})
        return root, rows
    try:
        asm = gen().IAssemblyDoc(model._oleobj_)
        try:
            asm.ResolveAllLightWeightComponents(False)
        except Exception:
            pass
        IComponent2 = gen().IComponent2

        def _walk(adoc, excluded=False):
            """2026-08-13: 递归 GetComponents(True) 逐层遍历(本环境 True=顶层, False=全树平铺)。
            父组件被排除 → 整棵子树跳过(排除标记在父实例上, 平铺会漏子树零件)。"""
            try:
                comps = gen().IAssemblyDoc(adoc._oleobj_).GetComponents(True)
            except Exception:
                comps = []
            for comp in (comps or []):
                try:
                    c2 = IComponent2(comp._oleobj_)
                    doc = c2.GetModelDoc2()
                except Exception:
                    continue
                if doc is None:
                    continue
                try:
                    doc = gen().IModelDoc2(doc._oleobj_)
                except Exception:
                    pass
                if excluded or _bom_excluded(c2, doc):
                    continue  # 2026-08-13: 不包括在材料明细表中 → 整棵子树跳过
                p = mpath(doc)
                is_asm = mtype(doc) == 2
                rows.append({"path": p,
                             "code": os.path.splitext(os.path.basename(p))[0] if p else mtitle(doc),
                             "name": _clean_name(doc, prop_names, p, mtitle(doc)),
                             "config": config_name(doc),
                             "is_asm": is_asm,
                             "props": read_props(doc, prop_names, config_name(doc))})
                if is_asm:
                    _walk(doc, False)

        _walk(model)
    except Exception:
        pass
    return root, rows


if __name__ == "__main__":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    try:
        sw = win32com.client.GetActiveObject("SldWorks.Application")
    except Exception:
        sw = win32com.client.Dispatch("SldWorks.Application")  # exec 场景兼容
    props = ["采购类别", "型材类型", "厚度", "代号", "名称", "$总数量"]
    root, rows = extract_components(sw, props)
    print("装配体:", root["title"], "| 组件实例总数:", len(rows))
    pc_cnt = collections.Counter()
    xingcai = collections.OrderedDict()
    asm_cnt = 0
    for r in rows:
        if r["is_asm"]:
            asm_cnt += 1
            continue
        pc = r["props"].get("采购类别", "") or "(空)"
        pc_cnt[pc] += 1
        if r["props"].get("采购类别", "") == "型材":
            k = r["path"] or r["name"]
            if k not in xingcai:
                xingcai[k] = [0, r]
            xingcai[k][0] += 1
    print(f"零件实例(非装配): {len(rows)-asm_cnt}, 子装配实例: {asm_cnt}")
    print("零件采购类别分布:", dict(pc_cnt))
    print(f"\n=== 采购类别=型材 ({len(xingcai)} 种) ===")
    for k, (cnt, r) in xingcai.items():
        p = r["props"]
        print(f"  {r['name']} | 汇总数量={cnt} | 型材类型={p.get('型材类型')} | 厚度={p.get('厚度')} | 代号={p.get('代号')}")
