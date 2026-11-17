# -*- coding: utf-8 -*-
"""
sw_preshot.py — 预生成全量等轴测截图 (2026-08-12 钦定: 一步截图, 后续全流程零截图)
流程: 打开总装 -> GetComponents(False) 全树 -> 收集唯一(路径,配置) -> 循环截图库:
      命中跳过, 未命中截图入库(15张/内存<12% 重启 SW)。纯截图, 无 BOM 逻辑。
用法:
  python sw_preshot.py --dir <图纸文件夹> [--asm <总装.sldasm>] [--cache-dir <库>]
"""
import argparse
import io
import os
import subprocess
import sys
import time

import win32com.client
import pythoncom

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import shot_cache

# ---- skill7 路径中枢（开源部署版）：所有绝对路径集中到 shared/skill7_paths.py ----
import os as _os, sys as _sys
_SK_ROOT = _os.path.abspath(_os.path.join(_os.path.dirname(_os.path.abspath(__file__)), "..", "..", ".."))
if _os.path.join(_SK_ROOT, "shared") not in _sys.path:
    _sys.path.insert(0, _os.path.join(_SK_ROOT, "shared"))
from skill7_paths import (  # noqa: E402
    AUTOBOM_DIR, MICDATAS_DIR, AUTOBOM_CFG, SHOT_CACHE, ZHANKAI_EXE,
    MACRO_DIR, MACRO_CLASSIFY, MACRO_PROFILE, SOLIDWORKS_EXE, SLDWORKS_EXE,
    SLDWORKS_CANDIDATES, POPUP_GUARD, SKILLS_DIR, ACAD_EXE, PDM_MARKERS,
)
GLOBAL_SKILLS = SKILLS_DIR  # 兼容旧命名：依赖 skill 脚本统一走本仓库 skills/

SHOT_BATCH = 50  # 2026-08-12: 实测每张~90MB, 50张才4.5GB; 15张过于保守(15张时内存还有26%)
# 2026-08-30: 12% 太晚, 大装配体(1000+实例)打开瞬间可吃满内存到 4% 触发系统不稳; 提前到 20% 重启
SHOT_MEM_TRIGGER = 20.0
# 2026-08-30: 危险线 — 打开装配体后低于此值立即止损(中断该组内部截图 + 重启 SW)
SHOT_MEM_DANGER = 6.0


def _gen():
    import glob as _glob
    import win32com.client.gencache as gencache
    base = gencache.GetGeneratePath()
    hits = [os.path.join(base, n) for n in os.listdir(base) if n.startswith("83A33D31-27C5-11CE-BFD4-00400513BB57x0x") and n.endswith(".py")]
    modname = "win32com.gen_py." + os.path.splitext(os.path.basename(hits[0]))[0]
    return __import__(modname, fromlist=["IModelDoc2", "IAssemblyDoc", "IComponent2", "IPartDoc"])


gen = _gen()


def _mtype(m):
    t = m.GetType
    return t() if callable(t) else t


def _mtitle(m):
    t = m.GetTitle
    return t() if callable(t) else t


def _mpath(m):
    t = m.GetPathName
    return t() if callable(t) else t


def _config_name(m):
    try:
        n = m.ConfigurationManager.ActiveConfiguration.Name
        return n() if callable(n) else n
    except Exception:
        return ""


def _bom_excluded(c2, doc):
    """2026-08-13 皇上钦定: 是否"不包括在材料明细表中"(Exclude from BOM)。
    组件级 IComponent2.ExcludeFromBOM 或 文档配置属性 "Exclude from bill of materials"=TRUE → True。"""
    try:
        v = c2.ExcludeFromBOM
        if callable(v):
            v = v()
        if v:
            return True
    except Exception:
        pass
    try:
        cn = _config_name(doc)
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


def _wrap(m):
    if m is None:
        return None
    try:
        return gen.IModelDoc2(m._oleobj_)
    except Exception:
        return m


def open_doc(sw, path, dtype=1):
    r = sw.OpenDoc6(path, dtype, 1, "", 0, 0)
    if isinstance(r, tuple):
        r = r[0]
    return _wrap(r) if r else None

def _mem_free_pct():
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


def _restart_sw(sw):
    try:
        sw.ExitApp()
    except Exception:
        pass
    for _i in range(20):
        time.sleep(1)
        r = subprocess.run(["powershell", "-NoProfile", "-Command",
                            "if (Get-Process sldworks -ErrorAction SilentlyContinue) { 'ALIVE' } else { 'GONE' }"],
                           capture_output=True, text=True)
        if "GONE" in r.stdout:
            break
    try:
        subprocess.run(["powershell", "-NoProfile", "-Command",
                        f"Stop-Process -Name sldworks -Force -ErrorAction SilentlyContinue"],
                       capture_output=True, text=True, timeout=20)
    except Exception:
        pass
    time.sleep(2)
    subprocess.Popen([SLDWORKS_EXE])
    time.sleep(5)
    for i in range(120):
        time.sleep(1)
        try:
            sw = win32com.client.GetActiveObject("SldWorks.Application")
            print(f"[MEM] SW 重启完成(等了 {i+1}s)", flush=True)
            time.sleep(3)
            return sw
        except Exception:
            continue
    raise RuntimeError("SW 重启后 120s 未连上")


def shot_iso(sw, path, img_dir, config_name=None):
    """打开零件 -> 等轴测 -> 缩放适应 -> SaveAs2 存 jpg。返回图片路径。"""
    os.makedirs(img_dir, exist_ok=True)
    if os.path.exists(path) and path.lower().endswith((".sldprt", ".sldasm")):
        dtype = 2 if path.lower().endswith(".sldasm") else 1
        try:
            doc = open_doc(sw, path, dtype)
        except Exception:
            doc = None
    else:
        doc = None
    if doc is None:
        return None
    cfg = config_name or _config_name(doc)
    out = shot_cache.shot_path_of(img_dir, path, cfg)
    name = os.path.splitext(os.path.basename(path))[0]
    try:
        doc.Visible = True
        try:
            sw.ActivateDoc(name)
        except Exception:
            pass
        try:
            doc.ShowConfiguration2(cfg)
        except Exception:
            pass
        doc.ShowNamedView2("*等轴测", 7)
        doc.ViewZoomtofit2()
        try:
            doc.ActiveView.DisplayMode = 5
        except Exception:
            pass
        try:
            sw.Refresh(0)
        except Exception:
            pass
        doc.SaveAs2(out, 0, True, True)
        if os.path.exists(out):
            try:
                from PIL import Image as _PIL
                _PIL.MAX_IMAGE_PIXELS = None
                _im = _PIL.open(out)
                if max(_im.size) > 1000:
                    _im.thumbnail((1000, 1000), _PIL.LANCZOS)
                    _im.convert("RGB").save(out, "JPEG", quality=88)
                _im.close()
            except Exception as _e:
                print("    [shot] 缩图失败(保留原图):", repr(_e)[:80])
            return out
    except Exception as e:
        print("    [shot] 失败:", path, "|", repr(e)[:80])
    finally:
        try:
            if name:
                sw.CloseDoc(name)
        except Exception:
            pass
    return None


def _shot_opened(sw, doc, path, cfg, img_dir):
    """对已打开的文档(子装配内部零件已加载)切配置+等轴测截图。返回 img 路径或 None。"""
    out = shot_cache.shot_path_of(img_dir, path, cfg)
    name = os.path.splitext(os.path.basename(path))[0]
    try:
        doc.Visible = True
        try:
            sw.ActivateDoc(name)
        except Exception:
            pass
        try:
            doc.ShowConfiguration2(cfg)
        except Exception:
            pass
        doc.ShowNamedView2("*等轴测", 7)
        doc.ViewZoomtofit2()
        try:
            doc.ActiveView.DisplayMode = 5
        except Exception:
            pass
        try:
            sw.Refresh(0)
        except Exception:
            pass
        doc.SaveAs2(out, 0, True, True)
        if os.path.exists(out):
            try:
                from PIL import Image as _PIL
                _PIL.MAX_IMAGE_PIXELS = None
                _im = _PIL.open(out)
                if max(_im.size) > 1000:
                    _im.thumbnail((1000, 1000), _PIL.LANCZOS)
                    _im.convert("RGB").save(out, "JPEG", quality=88)
                _im.close()
            except Exception:
                pass
            return out
    except Exception as e:
        print("    [shot] 失败:", name, cfg, "|", repr(e)[:80])
    return None


def _collect_inner(sw, doc, gen):
    """枚举已打开装配体全树组件的 {路径: 文档对象} 映射(内部零件/子装配都算, 已加载直接截)。
    2026-08-13: 排除 ExcludeFromBOM 组件(不出现在任何 BOM, 不预生成截图)。"""
    out = {}
    IComponent2 = gen.IComponent2

    def _walk(adoc, excluded=False):
        try:
            comps = gen.IAssemblyDoc(adoc._oleobj_).GetComponents(True)
        except Exception:
            comps = []
        for comp in (comps or []):
            try:
                c2 = IComponent2(comp._oleobj_)
                cd = c2.GetModelDoc2()
                if cd is None:
                    continue
                try:
                    cd = gen.IModelDoc2(cd._oleobj_)
                except Exception:
                    pass
                if excluded or _bom_excluded(c2, cd):
                    continue  # 2026-08-13: 不包括在材料明细表中 → 整棵子树跳过
                p = _mpath(cd) or _mtitle(cd)
                if p:
                    out[p] = cd
                try:
                    t = _mtype(cd)
                except Exception:
                    t = 1
                if t == 2:
                    _walk(cd, False)
            except Exception:
                continue

    try:
        _walk(doc)
    except Exception:
        pass
    return out


def find_top_assembly(folder):
    """找总装 .sldasm(排除 ~$/ST_/Pipe_)。2026-08-30 改 listdir 避免中文目录 glob 失效。"""
    best = None
    try:
        names = sorted(os.listdir(folder))
    except OSError:
        names = []
    for f in names:
        if not f.lower().endswith(".sldasm"):
            continue
        b = f
        if b.startswith("~") or b.startswith("ST_") or b.startswith("Pipe_"):
            continue
        if best is None or len(os.path.splitext(b)[0].split("_")[0]) < len(
                os.path.splitext(os.path.basename(best))[0].split("_")[0]):
            best = os.path.join(folder, f)
    return best


def main():
    ap = argparse.ArgumentParser(description="预生成全量等轴测截图到缓存库(一次建库, 后续流程零截图)")
    ap.add_argument("--dir", required=True, help="图纸文件夹(自动找总装)")
    ap.add_argument("--asm", default="", help="显式总装路径(优先)")
    ap.add_argument("--cache-dir", default=shot_cache.DEFAULT_CACHE_DIR)
    ap.add_argument("--limit", type=int, default=0, help="仅处理前N个(测试用)")
    ap.add_argument("--refresh", action="store_true", help="仅刷新截图库记录的文件大小/时间(属性写入后防缓存失效), 不截图")
    args = ap.parse_args()

    cache_dir = args.cache_dir
    if args.refresh:
        n = shot_cache.refresh_stats(cache_dir)
        print(f"[PRESHOT] 刷新截图库 {n} 条记录的 size/mtime (属性写入后防缓存失效)")
        return
    os.makedirs(cache_dir, exist_ok=True)
    asm = args.asm or find_top_assembly(args.dir)
    if not asm:
        print("错误: 未找到总装配体"); sys.exit(1)
    print(f"[PRESHOT] 总装: {os.path.basename(asm)}")
    print(f"[PRESHOT] 截图库: {cache_dir}")

    try:
        sw = win32com.client.GetActiveObject("SldWorks.Application")
    except Exception:
        sw = win32com.client.Dispatch("SldWorks.Application")
        time.sleep(8)
    a = open_doc(sw, asm, 2)
    if a is None:
        print("错误: 打不开总装"); sys.exit(1)
    try:
        gen.IAssemblyDoc(a._oleobj_).ResolveAllLightWeightComponents(False)
    except Exception:
        pass
    comps = gen.IAssemblyDoc(a._oleobj_).GetComponents(False)
    print(f"[PRESHOT] 组件实例: {len(comps)}")

    targets = []
    seen = set()
    IComponent2 = gen.IComponent2

    def _walk(adoc, excluded=False):
        """2026-08-13: 递归 GetComponents(True) 逐层; 父排除 → 整棵子树跳过。"""
        try:
            comps = gen.IAssemblyDoc(adoc._oleobj_).GetComponents(True)
        except Exception:
            comps = []
        for comp in (comps or []):
            try:
                c2 = IComponent2(comp._oleobj_)
                doc = c2.GetModelDoc2()
                if doc is None:
                    continue
                try:
                    doc = gen.IModelDoc2(doc._oleobj_)
                except Exception:
                    pass
                if excluded or _bom_excluded(c2, doc):
                    continue  # 2026-08-13: 不包括在材料明细表中 → 整棵子树跳过
                p = _mpath(doc) or _mtitle(doc)
                if p and os.path.exists(p):
                    # 含装配体组件(焊接组件等本体也要截图, 6.6/6.7 阶段1 会命中)
                    cn = _config_name(doc)
                    key = (p, cn)
                    if key not in seen:
                        seen.add(key)
                        targets.append((p, cn))
                try:
                    t = _mtype(doc)
                except Exception:
                    t = 1
                if t == 2:
                    _walk(doc, False)
            except Exception:
                continue

    _walk(a)
    try:
        sw.CloseDoc(_mtitle(a))
    except Exception:
        pass
    print(f"[PRESHOT] 唯一零件(路径,配置): {len(targets)}")
    # ★方案B(2026-08-13 钦定): 关闭总装后立即重启 SW, 释放总装常驻内存(6.9GB级),
    #   后续截图阶段内存 = 单个装配体 + 单张截图, 不再受总装拖累
    print("[MEM] 关闭总装, 重启 SW 释放内存...", flush=True)
    sw = _restart_sw(sw)

    # 截图清单按装配体分组: 先打开每个装配体截它自己+内部清单项, 再兜底独立零件
    target_map = {}
    asm_list = []
    for p, cfg in targets:
        dt = 2 if p.lower().endswith(".sldasm") else 1
        target_map[p] = (cfg, dt)
        if dt == 2:
            asm_list.append(p)
    # 浅路径优先(外层装配体先开, 覆盖内部)
    asm_list.sort(key=len)

    hit_n = 0
    new_n = 0
    fail_n = 0
    done = set()

    def _record(p, cfg, img):
        nonlocal new_n, fail_n
        if img:
            new_n += 1
            st = None
            try:
                st = os.stat(p)
            except Exception:
                st = None
            if st:
                shot_cache.put_shot(cache_dir, p, cfg,
                                    os.path.splitext(os.path.basename(p))[0], cfg,
                                    img, st.st_size, st.st_mtime)
            print(f"    [图] {os.path.basename(p)} ({cfg})", flush=True)
        else:
            fail_n += 1
            print(f"    [失败] {os.path.basename(p)}", flush=True)

    def _need_shot(p, cfg):
        nonlocal hit_n
        st = None
        try:
            st = os.stat(p)
        except Exception:
            st = None
        if st:
            img = shot_cache.get_shot(cache_dir, p, cfg, st.st_size, st.st_mtime)
            if img:
                hit_n += 1
                return False
        return True

    groups_done = 0
    _last_restart = [0.0]

    def _maybe_restart(why):
        """2026-08-30: 内存不足则重启 SW(带20s冷却防频繁重启); 重启后仍 < 危险线则报错让 skill7 干净停止。"""
        free = _mem_free_pct()
        if free >= SHOT_MEM_TRIGGER:
            return
        now = time.time()
        if now - _last_restart[0] < 20:
            return
        print(f"[MEM] 剩余内存 {free:.0f}% < {SHOT_MEM_TRIGGER:.0f}%, 重启 SW ({why})...", flush=True)
        sw2 = _restart_sw(sw)
        _last_restart[0] = time.time()
        free2 = _mem_free_pct()
        if free2 < SHOT_MEM_DANGER:
            raise RuntimeError(
                f"重启 SW 后剩余内存仍仅 {free2:.0f}% (< {SHOT_MEM_DANGER:.0f}%), 系统资源不足, 请关闭占内存程序后重跑")
        return sw2

    # ---- 第一轮: 装配体分组 ----
    for p in asm_list:
        if args.limit and len(done) >= args.limit:
            break
        if p in done:
            continue
        cfg, dt = target_map[p]
        if not _need_shot(p, cfg):
            done.add(p)
            continue
        # 2026-08-30: 每次打开装配体前都查内存(不等 5 组), 不够先重启, 防止大装配体把内存瞬间吃满
        r = _maybe_restart("打开装配体前")
        if r is not None:
            sw = r
        doc = open_doc(sw, p, 2)
        if doc is None:
            fail_n += 1
            print(f"    [失败] 装配体打不开: {os.path.basename(p)}", flush=True)
            done.add(p)
            continue
        groups_done += 1
        # 截装配体自己
        img = _shot_opened(sw, doc, p, cfg, cache_dir)
        _record(p, cfg, img)
        done.add(p)
        # 截内部所有清单项(已加载, ActivateDoc 直接截)
        inner = _collect_inner(sw, doc, gen)
        for ip, idoc in inner.items():
            if args.limit and len(done) >= args.limit:
                break
            if ip not in target_map or ip in done:
                continue
            icfg, idt = target_map[ip]
            if not _need_shot(ip, icfg):
                done.add(ip)
                continue
            # 2026-08-30: 大装配体打开后内存可能已被吃满, 低于危险线及时止损
            #   (未截的内部件落到第二轮独立打开, 单件内存占用小, 更安全)
            if _mem_free_pct() < SHOT_MEM_DANGER:
                print(f"    [MEM] 内存危险 {_mem_free_pct():.0f}% < {SHOT_MEM_DANGER:.0f}%, 中断本组内部截图", flush=True)
                break
            try:
                img2 = _shot_opened(sw, idoc, ip, icfg, cache_dir)
                _record(ip, icfg, img2)
            except Exception:
                img2 = shot_iso(sw, ip, cache_dir, config_name=icfg)
                _record(ip, icfg, img2)
            done.add(ip)
        try:
            sw.CloseDoc(os.path.splitext(os.path.basename(p))[0])
        except Exception:
            pass
        # 2026-08-30: 处理完一组(尤其大组)后 SW 内存不还给 OS, 立即评估重启
        r = _maybe_restart("组后释放")
        if r is not None:
            sw = r

    # ---- 第二轮: 独立零件兜底(未被装配体覆盖的) ----
    for p, cfg in targets:
        if args.limit and len(done) >= args.limit:
            break
        if p in done:
            continue
        if not _need_shot(p, cfg):
            done.add(p)
            continue
        r = _maybe_restart("独立零件前")
        if r is not None:
            sw = r
        img = shot_iso(sw, p, cache_dir, config_name=cfg)
        _record(p, cfg, img)
        done.add(p)
        if new_n % 10 == 0 and new_n > 0:
            n, mb = shot_cache.stats(cache_dir)
            print(f"[PRESHOT] 进度: 新截 {new_n}, 命中 {hit_n}, 失败 {fail_n}; 库 {n} 条 / {mb:.1f}MB", flush=True)

    n, mb = shot_cache.stats(cache_dir)
    print(f"\n[PRESHOT] 完成: 命中 {hit_n}, 新截 {new_n}, 失败 {fail_n}; 截图库共 {n} 条 / {mb:.1f} MB", flush=True)


if __name__ == "__main__":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    main()
