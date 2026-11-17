# -*- coding: utf-8 -*-
"""
sw_tuhao_categorize.py — 图号更新 + 零件分类 一次遍历合并版 (2026-08-13 皇上钦定)
打开每个文件一次, 同时完成:
  ① 图号: 按文件名 "_" 拆分写「代号」「名称」(白名单硬保护, 只写这两个)
  ② 类别: 读「采购类别」; 空则按文件名规则判定并写入(已有值优先, 绝不覆盖用户已填)
  ③ 保存(有改动才存) → 关闭 → 按类别复制到分类文件夹
不分固定批次, 内存<15% 自动重启 SW 继续(不设固定批, 皇上钦定)。
用法:
  python sw_tuhao_categorize.py --dir <图纸文件夹> --asm <总装.sldasm>
"""
import argparse, io, os, shutil, subprocess, sys, time, glob
import pythoncom
import win32com.client
import win32com.client.gencache as gencache

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
pythoncom.CoInitialize()

swCustomInfoText = 30
VALID = {"钣金件", "型材", "外购件", "标准件", "机加工", "机加件"}
SLDWORKS_EXE = r"E:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\sldworks.exe"
MEM_TRIGGER = 15.0

# 复用 sw_categorize 的判定规则(空属性时按文件名判定)
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "..", "4-导出BOM带截图", "scripts"))
import sw_categorize as sc


def log(msg):
    print(msg, flush=True)


def val(x):
    try:
        return x() if callable(x) else x
    except Exception:
        return x


def ensure_gen():
    base = gencache.GetGeneratePath()
    hits = [os.path.join(base, n) for n in os.listdir(base) if n.startswith("83A33D31-27C5-11CE-BFD4-00400513BB57x0x") and n.endswith(".py")]
    return __import__("win32com.gen_py." + os.path.splitext(os.path.basename(hits[0]))[0],
                      fromlist=["IAssemblyDoc", "IComponent2", "IModelDoc2"])


def split_name(fname):
    base = fname
    for ext in (".SLDPRT", ".SLDASM", ".sldprt", ".sldasm"):
        if base.upper().endswith(ext.upper()):
            base = base[:-len(ext)]
            break
    if "_" not in base:
        return None
    parts = base.split("_", 1)
    code, name = parts[0].strip(), parts[1].strip()
    if not code or not name:
        return None
    return (code, name)


def write_prop(doc, name, value):
    """白名单硬保护: 只允许代号/名称"""
    if name not in ("代号", "名称"):
        raise ValueError(f"禁止写非代号/名称属性: {name!r}")
    try:
        cfg = doc.ConfigurationManager.ActiveConfiguration
        cpm = cfg.CustomPropertyManager
        try:
            cpm.Delete(name)
        except Exception:
            pass
        cpm.Add2(name, swCustomInfoText, value)
        return True
    except Exception:
        return False


def connect_sw():
    for _ in range(30):
        try:
            return win32com.client.GetActiveObject("SldWorks.Application")
        except Exception:
            time.sleep(3)
    raise SystemExit("SW 连接失败")


def restart_sw(sw):
    log("[MEM] 剩余内存低, 重启 SW 释放内存...")
    try:
        sw.ExitApp()
    except Exception:
        pass
    time.sleep(4)
    subprocess.run(["powershell", "-NoProfile", "-Command",
                    "Stop-Process -Name sldworks -Force -ErrorAction SilentlyContinue"],
                   capture_output=True, text=True)
    time.sleep(2)
    subprocess.Popen([SLDWORKS_EXE])
    time.sleep(6)
    sw = connect_sw()
    log("[MEM] SW 已重启")
    return sw


def mem_free_pct():
    import ctypes
    class MEMORYSTATUSEX(ctypes.Structure):
        _fields_ = [("dwLength", ctypes.c_ulong), ("dwMemoryLoad", ctypes.c_ulong),
                    ("ullTotalPhys", ctypes.c_ulonglong), ("ullAvailPhys", ctypes.c_ulonglong),
                    ("ullTotalPageFile", ctypes.c_ulonglong), ("ullAvailPageFile", ctypes.c_ulonglong),
                    ("ullTotalVirtual", ctypes.c_ulonglong), ("ullAvailVirtual", ctypes.c_ulonglong),
                    ("ullAvailExtendedVirtual", ctypes.c_ulonglong)]
    try:
        ms = MEMORYSTATUSEX()
        ms.dwLength = ctypes.sizeof(MEMORYSTATUSEX)
        ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(ms))
        return ms.ullAvailPhys / ms.ullTotalPhys * 100.0
    except Exception:
        return 100.0


def _localize_pdm_path(p, folder):
    """PDM 库副本 -> 桌面夹本地可写副本重定向 (2026-08-30 根治)。
    总装引用解析可能把零件指到 D:\feedingHero PDM 库只读副本,
    对它 Save() 会弹 PDM 另存为框阻塞流程(SZ.Z3.0/Z4.0 绞龙轴组件02 实证)。
    返回 (本地路径, 是否可写处理):
      - 非 PDM 路径: (p, True) 原样处理
      - PDM 库副本 + 桌面夹有同名可写副本: (桌面副本, True) 重定向, 属性可正常刷全
      - PDM 库副本 + 桌面夹无同名副本(真库件): (p, False) 标记只读跳过(库件属性原本已对, 不动)
    """
    low = p.lower().replace("/", "\\")
    if "d:\\feedinghero" not in low:
        return p, True
    fname = os.path.basename(p)
    local = os.path.join(folder, fname)
    if os.path.exists(local) and os.access(local, os.W_OK):
        log(f"  [PDM重定向] {fname}: 库副本 -> 桌面可写副本")
        return local, True
    log(f"  [PDM跳过] {fname}: 库引用且桌面无可写副本, 跳过(库件属性已对)")
    return p, False


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dir", required=True)
    ap.add_argument("--asm", default="")
    args = ap.parse_args()
    folder = args.dir

    sw = connect_sw()
    G = ensure_gen()

    # 1. 打开总装收集全树唯一文件(含子装配全层级)
    paths = set()
    if args.asm:
        _o = sw.OpenDoc6(args.asm, 2, 1, "", 0, 0)
        md = _o[0] if isinstance(_o, (tuple, list)) else _o
        if md is None:
            raise SystemExit(f"打不开总装: {args.asm}")
        try:
            sw.ActivateDoc(os.path.splitext(os.path.basename(args.asm))[0])
        except Exception:
            pass
        asm = G.IAssemblyDoc(md._oleobj_)
        try:
            asm.ResolveAllLightWeightComponents(False)
        except Exception:
            pass
        comps = asm.GetComponents(False)
        for c in comps:
            try:
                c2 = G.IComponent2(c._oleobj_)
                d = c2.GetModelDoc2()
                if d is None:
                    continue
                p = val(d.GetPathName)
                if p and os.path.exists(p):
                    paths.add(p)
            except Exception:
                continue
        try:
            sw.CloseDoc(os.path.splitext(os.path.basename(args.asm))[0])
        except Exception:
            pass
        # ★2026-08-30 根治: PDM 库引用重定向到本地可写副本, 杜绝 Save() 弹 PDM 另存为
        # 有桌面可写副本 -> 重定向处理; 真库件(桌面无同名副本) -> 剔除不处理(库件属性已对)
        localized = set()
        for p in paths:
            lp, w = _localize_pdm_path(p, folder)
            if w:
                localized.add(lp)
            else:
                log(f"  [PDM剔除] {os.path.basename(p)}: 库件且桌面无可写副本, 不处理")
        paths = localized
        log(f"总装收集唯一文件: {len(paths)}")
    else:
        # 无总装: 根目录直接 .sldprt/.sldasm
        for f in os.listdir(folder):
            if f.lower().endswith((".sldprt", ".sldasm")) and not f.startswith("~"):
                paths.add(os.path.join(folder, f))
        log(f"根目录文件: {len(paths)}")

    tuhao_ok = tuhao_skip = 0
    cat = {}
    fail = []
    processed = 0
    # 2026-10-17: 记录本轮应存在的分类副本, 循环后清理上一轮残留(零件改名/删除/换分类时
    # 旧副本会永久留在 钣金件\ 等目录 → 步骤8 会为它多展开一个 DXF → 下游合并出现重复件。
    # 实测 0901: 大门板改图号后, 钣金件\ 里仍留着旧图号副本, 其 DXF 代号与"小门板"撞号被混入别组)
    expected_copies = {}
    for p in sorted(paths):
        fname = os.path.basename(p)
        ext = os.path.splitext(p)[1].lower()
        dt = 1 if ext == ".sldprt" else 2
        doc = None
        for attempt in range(3):
            try:
                r = sw.OpenDoc6(p, dt, 1, "", 0, 0)
                doc = r[0] if isinstance(r, (tuple, list)) else r
            except Exception:
                doc = None
            if doc is not None:
                break
            time.sleep(2)
        if doc is None:
            fail.append(fname)
            log(f"  [打开失败] {fname}")
            continue
        changed = False
        # ① 图号
        spl = split_name(fname)
        if spl is None:
            tuhao_skip += 1
        else:
            code, name = spl
            if write_prop(doc, "代号", code) and write_prop(doc, "名称", name):
                tuhao_ok += 1
                changed = True
            else:
                fail.append(fname)
        # ③ 保存(有改动才存) — ★必须在类别步骤之前(2026-08-15 修复):
        #   ensure_purchasing_category 结束会 CloseDoc, 之后 doc 已断连 Save 必失败
        #   (实测 51/51 保存警告 RPC_E_DISCONNECTED → 代号/名称从未落盘)
        if changed:
            try:
                doc.Save()
            except Exception as e:
                log(f"  [保存警告] {fname}: {repr(e)[:50]}")
        # ② 类别(已有值优先, 空则按文件名判定写入; 内部写完会 Save 落盘)
        try:
            cat_val = sc.ensure_purchasing_category(sw, p)
            cat_val = (cat_val or "").strip()
        except Exception:
            cat_val = ""
        if cat_val not in VALID:
            cat_val = "未分类"
        # 关闭(可能已被类别步骤关闭, 静默容错)
        try:
            _t = doc.GetTitle
            sw.CloseDoc(_t() if callable(_t) else _t)
        except Exception:
            pass
        # ④ 复制到分类文件夹(仅零件)
        if ext == ".sldprt" and cat_val in VALID:
            td = os.path.join(folder, cat_val)
            try:
                os.makedirs(td, exist_ok=True)
                shutil.copy2(p, os.path.join(td, fname))
                expected_copies.setdefault(cat_val, set()).add(fname.lower())
            except Exception:
                pass
        cat[fname] = cat_val
        processed += 1
        if processed % 20 == 0:
            log(f"  进度 {processed}/{len(paths)}, 剩余内存 {mem_free_pct():.0f}%")
        # 内存触发重启(不固定分批)
        if processed < len(paths) and mem_free_pct() < MEM_TRIGGER:
            sw = restart_sw(sw)

    log("" + chr(10) + "=== 汇总(图号+分类一次遍历) ===")
    log(f"图号: 成功 {tuhao_ok}, 跳过(格式不规范) {tuhao_skip}, 打开/写失败 {len(fail)}")
    dist = {}
    for v in cat.values():
        dist[v] = dist.get(v, 0) + 1
    log("分类分布: " + ", ".join(f"{k}={n}" for k, n in sorted(dist.items())))

    # 2026-10-17: 清理分类目录里上一轮残留的 .SLDPRT(只清分类目录根层, 不动 8.8 建的焊接组子目录,
    # 也不动 csv/dxf/合并图纸等其它步骤的产物)
    stale_total = 0
    for cval in sorted(VALID):
        cd = os.path.join(folder, cval)
        if not os.path.isdir(cd):
            continue
        keep = expected_copies.get(cval, set())
        stale = [f for f in os.listdir(cd)
                 if f.lower().endswith(".sldprt") and os.path.isfile(os.path.join(cd, f))
                 and f.lower() not in keep]
        for f in stale:
            try:
                os.remove(os.path.join(cd, f))
                log(f"  [清理] {cval}\\{f} (上轮残留, 本装配已无此件)")
                stale_total += 1
            except Exception as e:
                log(f"  [清理失败] {cval}\\{f}: {repr(e)[:50]}")
    if stale_total:
        log(f"分类目录残留副本已清理: {stale_total} 个")


if __name__ == "__main__":
    main()
