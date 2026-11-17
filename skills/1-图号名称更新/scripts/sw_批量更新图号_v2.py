# -*- coding: utf-8 -*-
"""
批量更新图号 v2 - 纯 COM 直写, 不调用宏
遍历当前装配体所有文件(含子装配全层级) -> 读真实文件名 -> 按 "_" 分离代号/名称
-> 直接写入每个文件配置特定的「代号」「名称」属性
★铁律: 只写「代号」「名称」两个属性, 其他属性禁止新增/删除/修改
"""
import io, sys, os, glob, time, collections
import pythoncom
import win32com.client
import win32com.client.gencache as gencache

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
pythoncom.CoInitialize()

swCustomInfoText = 30

def val(x):
    try: return x() if callable(x) else x
    except Exception: return x

def ensure_gen():
    base = gencache.GetGeneratePath()
    hits = glob.glob(base + r"\\83A33D31-27C5-11CE-BFD4-00400513BB57x0x*.py")
    return __import__("win32com.gen_py."+os.path.splitext(os.path.basename(hits[0]))[0],
                      fromlist=['IAssemblyDoc','IComponent2','IModelDoc2'])

def split_name(fname):
    base = fname
    for ext in (".SLDPRT",".SLDASM",".sldprt",".sldasm"):
        if base.upper().endswith(ext.upper()):
            base = base[:-len(ext)]; break
    if "_" not in base:
        return None
    parts = base.split("_", 1)
    code, name = parts[0].strip(), parts[1].strip()
    if not code or not name:
        return None
    return (code, name)

def write_prop(doc, name, value):
    """★白名单硬保护: 只允许代号/名称"""
    if name not in ("代号", "名称"):
        raise ValueError(f"禁止写非代号/名称属性: {name!r}")
    try:
        cfg = doc.ConfigurationManager.ActiveConfiguration
        cpm = cfg.CustomPropertyManager
        try: cpm.Delete(name)
        except Exception: pass
        cpm.Add2(name, swCustomInfoText, value)
        return True
    except Exception:
        return False

def main():
    import argparse
    _ap = argparse.ArgumentParser(description="批量更新图号(代号/名称)")
    _ap.add_argument("--asm", default="", help="总装配体 .sldasm 路径(可选; 缺省用 SW 当前激活装配体)")
    _args = _ap.parse_args()
    try:
        sw = win32com.client.GetActiveObject("SldWorks.Application")
    except Exception:
        sw = win32com.client.Dispatch("SldWorks.Application")  # exec 场景兼容
    print("SW:", sw.RevisionNumber(), flush=True)
    md = None
    if _args.asm:
        if not os.path.isfile(_args.asm):
            print(f"错误: 装配体不存在: {_args.asm}"); return
        _o = sw.OpenDoc6(_args.asm, 2, 0, "", 0, 0)
        md = _o[0] if isinstance(_o,(tuple,list)) else _o
        if md is None:
            print(f"错误: 打不开装配体: {_args.asm}"); return
        try:
            md_nm = os.path.splitext(os.path.basename(_args.asm))[0]
            sw.ActivateDoc(md_nm)
        except Exception:
            pass
    else:
        md = sw.ActiveDoc
    if md is None:
        print("错误: SW 没有活动文档!"); return
    print("当前文档:", val(md.GetTitle), "type=", val(md.GetType), flush=True)
    if val(md.GetType) != 2:
        print("错误: 需要装配体!"); return

    G = ensure_gen()
    asm = G.IAssemblyDoc(md._oleobj_)
    try: asm.ResolveAllLightWeightComponents(False)
    except Exception: pass
    comps = asm.GetComponents(False)
    print("组件实例总数:", len(comps), flush=True)

    paths = set()
    for c in comps:
        try:
            c2 = G.IComponent2(c._oleobj_)
            d = c2.GetModelDoc2()
            if d is None: continue
            p = val(d.GetPathName)
            if p and os.path.exists(p):
                paths.add(p)
        except Exception:
            pass
    print("唯一文件数:", len(paths), flush=True)

    dist = collections.Counter()
    for p in paths:
        ext = os.path.splitext(p)[1].lower()
        dist["装配" if ext==".sldasm" else "零件"] += 1
    for k,v in dist.items():
        print(f"   {k}: {v}", flush=True)

    ok=0; skip=0; fail=[]
    top_ok = 0          # 2026-10-19: 顶层装配体自身是否更新成功
    top_note = ""
    filt = os.environ.get("SW_FILTER", "")
    for p in sorted(paths):
        fname = os.path.basename(p)
        if filt and filt not in fname:
            continue
        spl = split_name(fname)
        if spl is None:
            skip += 1
            print(f"  [跳过-格式不规范] {fname}", flush=True)
            continue
        code, name = spl
        ext = os.path.splitext(p)[1].lower()
        dt = 1 if ext==".sldprt" else 2
        _o = sw.OpenDoc6(p, dt, 1, "", 0, 0)
        doc = _o[0] if isinstance(_o,(tuple,list)) else _o
        if not doc:
            fail.append(p); print(f"  [打开失败] {fname}", flush=True); continue
        w1 = write_prop(doc, "代号", code)
        w2 = write_prop(doc, "名称", name)
        if w1 and w2:
            ok += 1
            print(f"  [OK] {fname}\n       代号={code!r}  名称={name!r}", flush=True)
        else:
            fail.append(p)
            print(f"  [写失败] {fname}", flush=True)
        # ★写完立即保存到磁盘(否则改动只存内存, CloseDoc 后丢失), 用标题关闭
        try:
            doc.Save()
        except Exception as e:
            print(f"  [保存警告] {fname}: {repr(e)[:50]}", flush=True)
        try:
            # 兼容强类型(GetTitle 为属性)与动态(为方法)两种绑定
            _t = doc.GetTitle
            title = _t() if callable(_t) else _t
            sw.CloseDoc(title)
        except Exception as e:
            print(f"  [关文档警告] {fname}: {repr(e)[:50]}", flush=True)

    # ===== 2026-10-19 修复: 顶层装配体自身也要更新 =====
    # 旧版只遍历 GetComponents(组件), 顶层装配体不在其中 → 它的「代号」一直是旧值。
    # 实测: 炉具03 顶层 SZLJ.03.0_445×527炉具.SLDASM 的代号仍为旧系列 'SZLJ.01.0'(文件名已是 SZLJ.03.0)。
    # 注: 工程图(SLDDRW)不在装配体层级里, 且本套图纸不写「代号/名称」属性(标题栏引用模型属性), 故不处理。
    if not filt:
        _top_p = os.path.basename(val(md.GetPathName) or "")
        _tspl = split_name(_top_p) if _top_p else None
        if _tspl:
            _tc, _tn = _tspl
            if write_prop(md, "代号", _tc) and write_prop(md, "名称", _tn):
                try:
                    md.Save()
                    top_ok = 1
                    print(f"  [OK-顶层装配体] {_top_p}\n       代号={_tc!r}  名称={_tn!r}", flush=True)
                except Exception as e:
                    top_note = f"顶层保存失败: {repr(e)[:40]}"
                    print(f"  [保存警告] {_top_p}: {repr(e)[:50]}", flush=True)
            else:
                top_note = "顶层写失败"
                print(f"  [顶层写失败] {_top_p}", flush=True)
        else:
            top_note = "顶层文件名无下划线, 跳过"
            print(f"  [跳过-顶层文件名无下划线] {_top_p}", flush=True)

    print("\n=== 汇总(纯COM直写) ===", flush=True)
    print(f"成功: {ok + top_ok}   跳过(格式不规范): {skip}   失败: {len(fail)}" + (f"   ({top_note})" if top_note else ""), flush=True)
    for f in fail[:20]:
        print("  ", f, flush=True)

if __name__ == "__main__":
    main()