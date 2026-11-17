# -*- coding: utf-8 -*-
"""sw_weld_export.py — 从装配体按 AutoBOM 模板以「焊接或装配标记」属性分类导出 BOM(含三维等轴测截图 WPS 嵌入 Excel)
继承 skill4(sw_buyout_export.py v4) 截图嵌入机制, 仅分类属性不同: FILTER_PROP=焊接或装配标记, 不做采购类别自动判定。

内嵌 SW 正常方式启动 + 弹窗守卫 + 完成后释放内存(2026-08-07 从 skill7 融合)。

用法:
  python sw_weld_export.py --dir <图纸文件夹> -o <输出目录> --template <模板名> --filter-value <焊接或装配标记值> [--qty 套数]
"""
import argparse
import json
import os
import subprocess
import sys
import io
import time
import collections
import uuid as _uuid
import tempfile

import win32com.client
import pythoncom

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import shot_cache
import prop_cache

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

FILTER_PROP = "焊接或装配标记"

# SW 安装位置候选(2026-08-06 实测装于 E:\Program Files; C 盘亦常见)

PY = sys.executable or "python"

# ---- 弹窗守卫 ----
_guard_proc = None


def start_popup_guard():
    """后台常驻 SW 弹窗守卫, 自动点掉 PDM未检出/另存为/完成提示, 避免无人点卡死。
    2026-08-12: 幂等 — 已有存活守卫不重复启动(重复 start 会覆盖 _guard_proc 引用, 旧守卫变孤儿残留)。"""
    global _guard_proc
    if _guard_proc is not None and _guard_proc.poll() is None:
        print("[GUARD] 弹窗守卫已在运行, 跳过重复启动")
        return
    if not os.path.exists(POPUP_GUARD):
        print("[WARN] 未找到 sw_popup_guard.py, 跳过弹窗守卫(PDM库图纸可能有弹窗风险)")
        return
    print(f"[GUARD] 启动弹窗守卫: {POPUP_GUARD}")
    # 2026-10-03: 传父进程 PID 给守卫 → 本进程退出(含被强杀/崩溃)后守卫自动退出, 防残留干扰人工用 SW
    _guard_proc = subprocess.Popen(
        [PY, "-X", "utf8", POPUP_GUARD, "1.0", str(os.getpid())],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

def stop_popup_guard():
    """停止所有 sw_popup_guard.py 弹窗守卫进程(含历史孤儿)。
    2026-08-12: 抄 skill7 全量清理版 — 之前只 terminate 本次 _guard_proc,
    中途异常退出或多次 start 覆盖引用产生的孤儿守卫会残留成死循环进程, 一直自动点 SW 弹窗, 干扰皇上正常用 SW。"""
    global _guard_proc
    # 1) 先杀本次启动的(优雅) + 兜底强杀
    if _guard_proc is not None:
        try:
            _guard_proc.terminate()
            _guard_proc.wait(timeout=3)
        except Exception:
            try:
                _guard_proc.kill()
            except Exception:
                pass
        _guard_proc = None
    # 2) 全量清理所有 sw_popup_guard.py 残留(含历史孤儿)
    killed = 0
    try:
        def _cmdline_pids(keyword):
            pids = []
            try:
                ps = (
                    "Get-CimInstance Win32_Process | "
                    "Where-Object { $_.CommandLine -like '*" + keyword + "*' -and $_.ProcessId -ne $PID } | "
                    "Select-Object -ExpandProperty ProcessId"
                )
                out = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
                                     capture_output=True, timeout=30)
                raw = (out.stdout or b"")
                for ln in raw.decode("ascii", errors="ignore").splitlines():
                    ln = ln.strip()
                    if ln.isdigit():
                        pids.append(ln)
            except Exception:
                pass
            return pids

        for pid in _cmdline_pids("sw_popup_guard"):
            try:
                subprocess.run(["taskkill", "/F", "/PID", str(pid)], capture_output=True, text=True, timeout=10)
                killed += 1
            except Exception:
                pass
    except Exception as e:
        print(f"[WARN] 清理残留守卫失败: {repr(e)[:60]}")
    if killed:
        print(f"[GUARD] 已额外清理 {killed} 个残留弹窗守卫进程")
    print("[GUARD] 弹窗守卫全部停止")

def _sw_exe_from_registry():
    r"""从注册表读 SW 安装目录 (HKLM\SOFTWARE\SolidWorks\SolidWorks XXXX\Setup 的 SolidWorks Folder 值)。"""
    import winreg
    try:
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
                exe = os.path.join(folder, "sldworks.exe")
                if os.path.exists(exe):
                    return exe
            except OSError:
                continue
    except OSError:
        pass
    return None


def ensure_sw_normal_start():
    """确保 SW 以正常方式启动且可被 GetActiveObject 连上。
    根因(2026-08-06): exec 后台进程 COM 上下文读不到桌面交互式 SW 的 ROT,
    SW 须以正常方式启动(非 -Embedding)才注册进桌面 ROT。若连不上则用 Start-Process 拉起并等注册。"""
    try:
        win32com.client.GetActiveObject("SldWorks.Application")
        print("[SW] 已连上运行中的 SW (GetActiveObject 成功)")
        return
    except Exception as e:
        print(f"[SW] GetActiveObject 失败({repr(e)[:50]}), 尝试正常方式启动 SW...")
    exe = _sw_exe_from_registry()
    cands = [exe] if exe else SLDWORKS_CANDIDATES
    for cand in cands:
        if os.path.exists(cand):
            print(f"[SW] Start-Process 正常方式启动: {cand}")
            subprocess.Popen([cand])
            break
    # 轮询等待 SW 注册进 ROT(最长120s)
    for i in range(120):
        time.sleep(1)
        try:
            win32com.client.GetActiveObject("SldWorks.Application")
            print(f"[SW] SW 已正常启动并可连接(等了 {i+1}s)")
            time.sleep(3)  # 多等几秒让 COM 稳定
            return
        except Exception:
            continue
    raise RuntimeError("SW 未能在120s内以正常方式启动并连上, 请手动打开 SolidWorks 后重试")


def _free_ram_pct():
    """返回当前剩余物理内存百分比(0-100)。"""
    try:
        import ctypes
        class MEMORYSTATUSEX(ctypes.Structure):
            _fields_ = [("dwLength", ctypes.c_ulong),
                        ("dwMemoryLoad", ctypes.c_ulong),
                        ("ullTotalPhys", ctypes.c_ulonglong),
                        ("ullAvailPhys", ctypes.c_ulonglong),
                        ("ullTotalPageFile", ctypes.c_ulonglong),
                        ("ullAvailPageFile", ctypes.c_ulonglong),
                        ("ullTotalVirtual", ctypes.c_ulonglong),
                        ("ullAvailVirtual", ctypes.c_ulonglong),
                        ("ullAvailExtendedVirtual", ctypes.c_ulonglong)]
        ms = MEMORYSTATUSEX()
        ms.dwLength = ctypes.sizeof(MEMORYSTATUSEX)
        ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(ms))
        total = float(ms.ullTotalPhys)
        avail = float(ms.ullAvailPhys)
        if total <= 0:
            return 100.0
        return avail / total * 100.0
    except Exception:
        return 100.0


def _maybe_restart_sw(sw, idx, every=12):
    """2026-08-11: 截图内存保护: 每 N 张或剩余内存<12% 重启 SW 释放内存, 返回(新)sw。
    ExitApp 失败时强制 Stop-Process 兜底(实测 0.8GB 空闲时系统会 SIGKILL 进程)。"""
    if idx > 0 and (idx % every == 0 or _free_ram_pct() < 12.0):
        print(f"[MEM] 截图 {idx}, 剩余内存 {_free_ram_pct():.0f}%, 重启 SW 释放内存...")
        try:
            sw.ExitApp()
        except Exception:
            pass
        import subprocess as _sp
        for _i in range(20):
            time.sleep(1)
            r = _sp.run(["powershell", "-NoProfile", "-Command",
                         "if (Get-Process sldworks -ErrorAction SilentlyContinue) { 'ALIVE' } else { 'GONE' }"],
                        capture_output=True, text=True)
            if "GONE" in (r.stdout or ""):
                break
        else:
            print("[MEM] ExitApp 未退出 SW, 强制 Stop-Process sldworks")
            _sp.run(["powershell", "-NoProfile", "-Command",
                     "Stop-Process -Name sldworks -Force -ErrorAction SilentlyContinue"], capture_output=True)
            time.sleep(3)
        exe = _sw_exe_from_registry()
        if exe and os.path.exists(exe):
            subprocess.Popen([exe])
        for _i in range(120):
            time.sleep(1)
            try:
                sw2 = win32com.client.GetActiveObject("SldWorks.Application")
                time.sleep(3)
                print(f"[MEM] SW 已重启并可连接(等了 {_i+1}s)")
                return sw2
            except Exception:
                continue
        raise RuntimeError("SW 重启后 120s 未连上")
    return sw


def release_sw_memory(force_exit=True):
    """释放 SW 内存: 关闭全部文档 + 退出 SW 进程。"""
    try:
        import win32com.client
        sw = win32com.client.GetActiveObject("SldWorks.Application")
    except Exception as e:
        print(f"[MEM] 无可连接的 SW(无需释放): {repr(e)[:40]}")
        return
    try:
        sw.ExitApp()
        print("[MEM] 已请求退出 SW (释放内存)")
        import subprocess as _sp
        _sp.run(["powershell", "-NoProfile", "-Command",
                 "Start-Sleep -Seconds 3; "
                 "if (Get-Process sldworks -ErrorAction SilentlyContinue) { "
                 "Stop-Process -Name sldworks -Force -ErrorAction SilentlyContinue }"],
                capture_output=True)
    except Exception as e:
        print(f"[WARN] 退出SW失败({repr(e)[:40]})")


def _gen():
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
        # 2026-10-17 修复(同 skill3): gen_py 在 %TEMP% 下会被清理工具删除, 缓存丢失后
        # win32com 退化为后期绑定, OpenDoc6 报 '类型不匹配'(0x80020005)。改为自动重建缓存。
        tlb = None
        try:
            exe = _sw_exe_from_registry()
            if exe:
                cand = os.path.join(os.path.dirname(exe), "sldworks.tlb")
                if os.path.exists(cand):
                    tlb = cand
        except Exception:
            tlb = None
        if not tlb:
            for drv in ("E:", "D:", "C:", "F:"):
                p = os.path.join(drv + os.sep, "Program Files",
                                 "SOLIDWORKS Corp", "SOLIDWORKS", "sldworks.tlb")
                if os.path.exists(p):
                    tlb = p
                    break
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
        _GEN = _gen()
    return _GEN


def _get_bounding_dimensions(doc):
    """实时读取文档长/宽/高包络尺寸(mm), 返回 {"$包络尺寸":"长x宽x高", "$长":"", "$宽":"", "$高":""}。
    零件(1)用 GetPartBox(False) 取米->mm; 装配体(2)用 GetBox(2) 取毫米。
    """
    out = {"$包络尺寸": "", "$长": "", "$宽": "", "$高": ""}
    try:
        mt = _mtype(doc)
        if mt == 1:  # 零件: GetPartBox(False) 返回 mm
            pd = gen().IPartDoc(doc._oleobj_)
            box = pd.GetPartBox(False)
            if box and len(box) >= 6:
                dx = abs(box[3] - box[0])
                dy = abs(box[4] - box[1])
                dz = abs(box[5] - box[2])
            else:
                return out
        elif mt == 2:  # 装配体: GetBox(2) 返回 米 -> 乘 1000 转 mm
            ad = gen().IAssemblyDoc(doc._oleobj_)
            box = ad.GetBox(2)
            if box and len(box) >= 6:
                dx = abs(box[3] - box[0]) * 1000
                dy = abs(box[4] - box[1]) * 1000
                dz = abs(box[5] - box[2]) * 1000
            else:
                return out
        else:
            return out
        l, w, h = sorted([dx, dy, dz], reverse=True)
        out["$长"] = f"{l:.2f}"
        out["$宽"] = f"{w:.2f}"
        out["$高"] = f"{h:.2f}"
        out["$包络尺寸"] = f"{l:.0f}x{w:.0f}x{h:.0f}"
    except Exception:
        pass
    return out


def _wrap(m):
    if m is None:
        return None
    try:
        return gen().IModelDoc2(m._oleobj_)
    except Exception:
        return m


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


def open_doc(sw, path, dtype=1):
    # 2026-08-12: 强绑定(gen_py)传 0,0; 动态绑定(Dispatch)需显式 VARIANT — try 强绑定, 失败自动降级
    if sw is None:
        return None
    try:
        r = sw.OpenDoc6(path, dtype, 1, "", 0, 0)
    except Exception:
        errs = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        warns = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        r = sw.OpenDoc6(path, dtype, 1, "", errs, warns)
    if isinstance(r, tuple):
        r = r[0]
    return _wrap(r) if r else None

def read_props(model, names, config):
    out = {}
    try:
        cpm = model.Extension.CustomPropertyManager(config)
        if not cpm:
            return {n: "" for n in names}
        for n in names:
            try:
                r = cpm.Get3(n, False, "", "")
                if isinstance(r, (tuple, list)) and len(r) >= 3:
                    out[n] = (r[2] if r[2] else r[1]) or ""
                elif isinstance(r, (tuple, list)) and len(r) == 2:
                    out[n] = r[1] or ""
                else:
                    out[n] = r or ""
            except Exception:
                out[n] = ""
    except Exception:
        return {n: "" for n in names}
    return out


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


def load_template(name):
    with open(AUTOBOM_CFG, encoding="utf-8") as f:
        cfg = json.load(f)
    for t in cfg.get("Templates", []):
        if t.get("Name") == name:
            return [p["Name"] for p in t.get("Props", [])]
    raise ValueError("未找到 AutoBOM 模板: " + name + " (可选: " +
                     ", ".join(t.get("Name", "?") for t in cfg.get("Templates", [])) + ")")


def find_top_assembly(directory):
    cands = []
    for fn in os.listdir(directory):
        if not fn.lower().endswith(".sldasm"):
            continue
        if fn.startswith("~") or fn.startswith("ST_") or fn.lower().startswith("st_"):
            continue  # 排除 ~$ 锁文件 / ST_ 标准件
        code = fn.split("_")[0]
        cands.append((len(code), code, os.path.join(directory, fn)))
    if not cands:
        raise ValueError("目录下未找到装配体(.sldasm, 排除 ST_): " + directory)
    cands.sort(key=lambda x: (x[0], x[1]))
    return cands[0][2]


def collect_parts(sw, asm_path, cache_dir=None, asm_cache_hit=None):
    """打开总装, 遍历所有组件, 按文件路径聚合 (去重+累计数量)。
    cache_dir 给定时扫描结果写属性缓存; asm_cache_hit 为预取缓存时直接返回(零 SW)。"""
    if asm_cache_hit:
        agg = collections.OrderedDict()
        for k, e in asm_cache_hit.items():
            if k == "_v" or not isinstance(e, dict):
                continue  # 2026-08-12: 跳过版本号字段
            if "\x00" in k:
                p, cfg = k.split("\x00", 1)
            else:
                p, cfg = k, e.get("config", "")
            if p not in agg:
                agg[p] = {"count": e.get("count", 1), "props": e.get("props", {}),
                          "name": e.get("name", os.path.splitext(os.path.basename(p))[0]),
                          "config": cfg}
            else:
                agg[p]["count"] += e.get("count", 1)
        return agg
    asm = open_doc(sw, asm_path, 2)
    if asm is None or _mtype(asm) != 2:
        raise RuntimeError("无法打开总装配体: " + asm_path)
    try:
        gen().IAssemblyDoc(asm._oleobj_).ResolveAllLightWeightComponents(False)
    except Exception:
        pass
    agg = collections.OrderedDict()
    meta = {}
    IComponent2 = gen().IComponent2

    def _walk(adoc, excluded=False):
        """2026-08-13: 递归 GetComponents(True) 逐层; 父排除 → 整棵子树跳过。"""
        try:
            comps = gen().IAssemblyDoc(adoc._oleobj_).GetComponents(True)
        except Exception:
            comps = []
        for comp in (comps or []):
            try:
                c2 = IComponent2(comp._oleobj_)
                doc = c2.GetModelDoc2()
                if doc is None:
                    continue
                try:
                    doc = gen().IModelDoc2(doc._oleobj_)
                except Exception:
                    pass
                if excluded or _bom_excluded(c2, doc):
                    continue  # 2026-08-13: 不包括在材料明细表中 → 整棵子树跳过
                t = _mtype(doc)
                # 不跳过子装配(type=2): 焊接/装配标记常标在焊接组件(子装配)上, 须一并检查
                p = _mpath(doc) or _mtitle(doc)
                if not p:
                    continue
                cn = _config_name(doc)
                # 2026-08-12: 加 单位/网址/采购链接 — 命中 skill3 缓存时字段要全; 未命中直扫也顺手读全
                props = read_props(doc, [FILTER_PROP, "名称", "代号", "单位", "网址", "采购链接"], cn)
                if os.path.exists(p):
                    _st = os.stat(p)
                    meta[p] = {"size": _st.st_size, "mtime": _st.st_mtime}
            except Exception:
                continue
            if p not in agg:
                agg[p] = {"count": 0, "props": props,
                          "name": os.path.splitext(os.path.basename(p))[0] if p else _mtitle(doc),
                          "config": cn}
            agg[p]["count"] += 1
            if not agg[p].get("config"):
                agg[p]["config"] = cn
            for k in ("名称", "代号"):
                if not agg[p]["props"].get(k) and props.get(k):
                    agg[p]["props"][k] = props[k]
            if t == 2:
                _walk(doc, False)  # 子装配递归展开内部

    _walk(asm)
    # 2026-08-13: 总装自身也进聚合 — 焊接/装配标记可能打在总装根文档上(如 150KG 整机标焊接)
    try:
        _p0 = asm_path
        _cn0 = _config_name(asm)
        _props0 = read_props(asm, [FILTER_PROP, "名称", "代号", "单位", "网址", "采购链接"], _cn0)
        if _p0 not in agg:
            agg[_p0] = {"count": 1, "props": _props0,
                        "name": os.path.splitext(os.path.basename(_p0))[0],
                        "config": _cn0}
    except Exception:
        pass
    # 2026-08-12: 不写属性缓存 — 缓存统一由 skill3 步骤2 维护(含 is_asm/全字段); skill8 直扫时只读不写, 避免覆盖
    return agg


def shot_iso(sw, path, img_dir, config_name=None):
    """打开零件 -> ActivateDoc切前台 -> 等轴测 -> 缩放适应 -> SaveAs2 存 jpg。返回图片路径。"""
    os.makedirs(img_dir, exist_ok=True)
    if _path_exists(path) and path.lower().endswith((".sldprt", ".sldasm")):
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
            try:
                if os.path.exists(path):
                    sw.CloseDoc(os.path.basename(path))
            except Exception:
                pass
    return None


def _path_exists(p):
    return bool(p) and os.path.exists(p)


def _col_start_emu(colw_chars):
    x = 0; out = {}
    for i, w in enumerate(colw_chars, start=1):
        out[i] = x
        x += int(w * 7.0 * 9525)
    return out


def export_xlsx(cols, rows_with_img, out_path, img_col_key="$三维截图", header_h_pt=24, row_h_pt=70):
    """生成真正嵌入单元格图片的 xlsx(WPS/Excel365 识别)。 rows_with_img 每项含 {各列值..., '_img': 图片路径}"""
    from PIL import Image as PILImage
    PILImage.MAX_IMAGE_PIXELS = None

    colw = [20 if c == img_col_key else 12 for c in cols]
    col_start = _col_start_emu(colw)
    img_col_idx = None
    for i, c in enumerate(cols, start=1):
        if c == img_col_key:
            img_col_idx = i
            break
    if img_col_idx is None:
        raise ValueError("模板列缺少图片列: " + img_col_key)

    tmpdir = tempfile.mkdtemp(prefix="wps_emb_")  # 临时打包目录用系统临时区, 不在输出目录建子文件夹
    gids = []
    media = []
    cellimg_pics = []
    cell_rels = []
    for rid, item in enumerate(rows_with_img, start=1):
        src = item.get("_img")
        if not src or not os.path.exists(src):
            gids.append(None)
            continue
        gid = "ID_" + _uuid.uuid4().hex.upper()
        gids.append(gid)
        im = PILImage.open(src)
        im.thumbnail((120, 90), PILImage.LANCZOS)
        wpx, hpx = im.size
        pp = os.path.join(tmpdir, f"e{rid}.png")
        im.save(pp, "PNG")
        cy = int((header_h_pt + (rid - 1) * row_h_pt) * 12700)
        cx = col_start[img_col_idx]
        ecx, ecy = int(wpx * 9525), int(hpx * 9525)
        media.append((pp, f"xl/media/image{rid}.png"))
        cellimg_pics.append(
            f'<etc:cellImage><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{rid}" name="{gid}"/><xdr:cNvPicPr/></xdr:nvPicPr>'
            f'<xdr:blipFill><a:blip r:embed="rId{rid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>'
            f'<xdr:spPr><a:xfrm><a:off x="{cx}" y="{cy}"/><a:ext cx="{ecx}" cy="{ecy}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic></etc:cellImage>')
        cell_rels.append(
            f'<Relationship Id="rId{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image{rid}.png"/>')

    def xml_esc(s):
        return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                .replace('"', "&quot;"))

    rows = ['<row r="1">']
    for c, col in enumerate(cols, start=1):
        L = chr(64 + c)
        rows.append(f'<c r="{L}1" s="0" t="inlineStr"><is><t>{xml_esc(col)}</t></is></c>')
    rows.append('</row>')
    for r_off, item in enumerate(rows_with_img):
        r = r_off + 2
        gid = gids[r_off]
        row = [f'<row r="{r}" ht="{row_h_pt}" customHeight="1">']
        for c, col in enumerate(cols, start=1):
            L = chr(64 + c)
            if col == img_col_key and gid:
                row.append(f'<c r="{L}{r}" s="1" t="str"><f>_xlfn.DISPIMG(&quot;{gid}&quot;,1)</f>'
                           f'<v>=DISPIMG(&quot;{gid}&quot;,1)</v></c>')
            else:
                v = item.get(col, "")
                if isinstance(v, (int, float)):
                    row.append(f'<c r="{L}{r}" s="1"><v>{v}</v></c>')
                else:
                    row.append(f'<c r="{L}{r}" s="1" t="inlineStr"><is><t>{xml_esc(v)}</t></is></c>')
        row.append('</row>')
        rows.append("".join(row))
    sheet1 = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<dimension ref="A1:' + chr(64 + len(cols)) + str(len(rows_with_img) + 1) + '"/>'
        '<sheetData>' + "".join(rows) + '</sheetData></worksheet>')

    styles_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="2"><font><sz val="11"/><name val="宋体"/></font>'
        '<font><b/><sz val="11"/><name val="宋体"/></font></fonts>'
        '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="2">'
        '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
        '</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '</styleSheet>')

    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="png" ContentType="image/png"/>'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/cellimages.xml" ContentType="application/vnd.wps-officedocument.cellimage+xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '</Types>')
    workbook = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="BOM" sheetId="1" r:id="rId1"/></sheets></workbook>')
    wb_rels = ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '<Relationship Id="rId3" Type="http://www.wps.cn/officeDocument/2020/cellImage" Target="cellimages.xml"/>'
        '</Relationships>')
    sheet_rels = ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>')
    cellimages_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<etc:cellImages xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:etc="http://www.wps.cn/officeDocument/2017/etCustomData">'
        + "".join(cellimg_pics) + '</etc:cellImages>')
    cellimages_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(cell_rels) + '</Relationships>')
    root_rels = ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>')

    import zipfile
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", workbook)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/worksheets/sheet1.xml", sheet1)
        z.writestr("xl/styles.xml", styles_xml)
        z.writestr("xl/worksheets/_rels/sheet1.xml.rels", sheet_rels)
        z.writestr("xl/cellimages.xml", cellimages_xml)
        z.writestr("xl/_rels/cellimages.xml.rels", cellimages_rels)
        for sp, dp in media:
            z.write(sp, dp)
    try:
        for fn in os.listdir(tmpdir):
            os.remove(os.path.join(tmpdir, fn))
        os.rmdir(tmpdir)
    except Exception:
        pass
    return out_path



def export_multi_sheet_xlsx(cols, sheets, out_path, img_col_key="$三维截图", header_h_pt=24, row_h_pt=70):
    """多 sheet 的 xlsx, sheets = { sheet_name: [row_dicts] }。每个 sheet 内图片从全局rid递增。"""
    from PIL import Image as PILImage
    PILImage.MAX_IMAGE_PIXELS = None

    colw = [20 if c == img_col_key else (14 if c == "采购类别" else 12) for c in cols]
    col_start = _col_start_emu(colw)
    img_col_idx = None
    for i, c in enumerate(cols, start=1):
        if c == img_col_key:
            img_col_idx = i
            break
    if img_col_idx is None:
        raise ValueError("模板列缺少图片列: " + img_col_key)

    tmpdir = tempfile.mkdtemp(prefix="wps_multi_")
    global_rid = 0
    all_media = []
    all_cellimg_pics = []
    all_cellimg_rels = []
    sheet_xmls = {}
    sheet_rels_xmls = {}

    def xml_esc(s):
        return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                .replace('"', "&quot;"))

    COL_CNT = len(cols)
    MAX_COL_LETTER = chr(64 + COL_CNT)

    for sheet_id, (sheet_name, rows_with_img) in enumerate(sheets.items(), start=1):
        gids = []
        sheet_media = []
        sheet_cellimg_pics = []
        sheet_cell_rels = []

        for r_off, item in enumerate(rows_with_img):
            global_rid += 1
            src = item.get("_img")
            if not src or not os.path.exists(src):
                gids.append(None)
                continue
            gid = "ID_" + _uuid.uuid4().hex.upper()
            gids.append(gid)
            im = PILImage.open(src)
            im.thumbnail((120, 90), PILImage.LANCZOS)
            wpx, hpx = im.size
            pp = os.path.join(tmpdir, f"s{sheet_id}_r{r_off+1}.png")
            im.save(pp, "PNG")
            cy = int((header_h_pt + r_off * row_h_pt) * 12700)
            cx = col_start[img_col_idx]
            ecx, ecy = int(wpx * 9525), int(hpx * 9525)
            sheet_media.append((pp, f"xl/media/image{global_rid}.png"))
            sheet_cellimg_pics.append(
                f'<etc:cellImage><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{global_rid}" name="{gid}"/><xdr:cNvPicPr/></xdr:nvPicPr>'
                f'<xdr:blipFill><a:blip r:embed="rId{global_rid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>'
                f'<xdr:spPr><a:xfrm><a:off x="{cx}" y="{cy}"/><a:ext cx="{ecx}" cy="{ecy}"/></a:xfrm>'
                f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic></etc:cellImage>')
            sheet_cell_rels.append(
                f'<Relationship Id="rId{global_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image{global_rid}.png"/>')

        all_media.extend(sheet_media)
        all_cellimg_pics.extend(sheet_cellimg_pics)
        all_cellimg_rels.extend(sheet_cell_rels)

        # 构建此 sheet 的 xml
        rows_xml = ['<row r="1">']
        for c, col in enumerate(cols, start=1):
            L = chr(64 + c)
            rows_xml.append(f'<c r="{L}1" s="0" t="inlineStr"><is><t>{xml_esc(col)}</t></is></c>')
        rows_xml.append('</row>')

        for r_off, item in enumerate(rows_with_img):
            r = r_off + 2
            gid = gids[r_off]
            row = [f'<row r="{r}" ht="{row_h_pt}" customHeight="1">']
            for c, col in enumerate(cols, start=1):
                L = chr(64 + c)
                if col == img_col_key and gid:
                    row.append(f'<c r="{L}{r}" s="1" t="str"><f>_xlfn.DISPIMG(&quot;{gid}&quot;,1)</f>'
                               f'<v>=DISPIMG(&quot;{gid}&quot;,1)</v></c>')
                else:
                    v = item.get(col, "")
                    if isinstance(v, (int, float)):
                        row.append(f'<c r="{L}{r}" s="1"><v>{v}</v></c>')
                    else:
                        row.append(f'<c r="{L}{r}" s="1" t="inlineStr"><is><t>{xml_esc(v)}</t></is></c>')
            row.append('</row>')
            rows_xml.append("".join(row))

        sheet_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f'<dimension ref="A1:{MAX_COL_LETTER}{len(rows_with_img)+1}"/>'
            '<sheetData>' + "".join(rows_xml) + '</sheetData></worksheet>')

        sheet_rel = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>')

        sheet_xmls[sheet_id] = sheet_xml
        sheet_rels_xmls[sheet_id] = sheet_rel

    # workbook.xml (多sheet)
    sheets_wb = []
    wb_rels_items = []
    for sheet_id, (sheet_name, _) in enumerate(sheets.items(), start=1):
        sheets_wb.append(f'<sheet name="{xml_esc(sheet_name)}" sheetId="{sheet_id}" r:id="rId{sheet_id}"/>')
        wb_rels_items.append(
            f'<Relationship Id="rId{sheet_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{sheet_id}.xml"/>')
    wb_rels_items.append(
        f'<Relationship Id="rId{len(sheets)+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>')
    wb_rels_items.append(
        f'<Relationship Id="rId{len(sheets)+2}" Type="http://www.wps.cn/officeDocument/2020/cellImage" Target="cellimages.xml"/>')

    workbook = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets>' + "".join(sheets_wb) + '</sheets></workbook>')
    wb_rels = ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(wb_rels_items) + '</Relationships>')

    styles_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="2"><font><sz val="11"/><name val="\u5b8b\u4f53"/></font>'
        '<font><b/><sz val="11"/><name val="\u5b8b\u4f53"/></font></fonts>'
        '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="2">'
        '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
        '</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '</styleSheet>')

    # Content_Types (多sheet)
    ct_overrides = []
    for sheet_id in range(1, len(sheets)+1):
        ct_overrides.append(
            f'<Override PartName="/xl/worksheets/sheet{sheet_id}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="png" ContentType="image/png"/>'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/cellimages.xml" ContentType="application/vnd.wps-officedocument.cellimage+xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        + "".join(ct_overrides) + '</Types>')

    root_rels = ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>')

    cellimages_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<etc:cellImages xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:etc="http://www.wps.cn/officeDocument/2017/etCustomData">'
        + "".join(all_cellimg_pics) + '</etc:cellImages>')
    cellimages_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(all_cellimg_rels) + '</Relationships>')

    import zipfile
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", workbook)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/styles.xml", styles_xml)
        z.writestr("xl/cellimages.xml", cellimages_xml)
        z.writestr("xl/_rels/cellimages.xml.rels", cellimages_rels)
        for sheet_id in range(1, len(sheets)+1):
            z.writestr(f"xl/worksheets/sheet{sheet_id}.xml", sheet_xmls[sheet_id])
            z.writestr(f"xl/worksheets/_rels/sheet{sheet_id}.xml.rels", sheet_rels_xmls[sheet_id])
        for sp, dp in all_media:
            z.write(sp, dp)
    try:
        for fn in os.listdir(tmpdir):
            os.remove(os.path.join(tmpdir, fn))
        os.rmdir(tmpdir)
    except Exception:
        pass
    return out_path


def route_detail_rows(sw, asm_path, qty, max_depth=6):
    """2026-08-11: 若部件(焊接/装配)含步路装配体, 展开其 零部件+线路零件 明细行。
    线路零件规格优先取属性 规格/直径/尺寸, 缺则用配置名; 长度取 SW 自带"长度"属性(mm)。"""
    import re as _re
    try:
        asm = open_doc(sw, asm_path, 2)
    except Exception:
        return []
    if asm is None:
        return []
    out = collections.OrderedDict()  # (title,spec,len) -> {count,...}

    def _add(cdoc):
        try:
            cdoc = gen().IModelDoc2(cdoc._oleobj_)
        except Exception:
            pass
        p = _mpath(cdoc) or _mtitle(cdoc)
        nm = os.path.splitext(os.path.basename(p))[0] if p else _mtitle(cdoc)
        cn = _config_name(cdoc)
        pr = read_props(cdoc, ["名称", "规格", "直径", "尺寸", "长度", "Length", "SWPipeLength",
                               "Pipe Identifier", "SWbompartno"], cn)
        # 规格: Pipe Identifier → SWbompartno → 规格/直径/尺寸 → 配置名
        spec = (pr.get("Pipe Identifier") or pr.get("SWbompartno")
                or pr.get("规格") or pr.get("直径") or pr.get("尺寸") or cn)
        # 长度: SW Routing 自带 Length(带mm后缀) → SWPipeLength → 长度
        ln = str(pr.get("Length") or pr.get("SWPipeLength") or pr.get("长度") or "").replace(",", "")
        m = _re.search(r"-?[\d.]+(?:[eE][+-]?\d+)?", ln)
        ln = m.group(0) if m else ""
        k = (nm, spec, ln)
        if k not in out:
            out[k] = {"title": nm, "spec": spec, "length": ln, "count": 0}
        out[k]["count"] += 1

    def _walk(adoc, depth, in_route):
        if depth > max_depth:
            return
        try:
            comps = gen().IAssemblyDoc(adoc._oleobj_).GetComponents(False)
        except Exception:
            return
        if not comps:
            return
        IComponent2 = gen().IComponent2
        for comp in comps:
            try:
                c2 = IComponent2(comp._oleobj_)
                cdoc = c2.GetModelDoc2()
            except Exception:
                continue
            if cdoc is None:
                continue
            try:
                cdoc = gen().IModelDoc2(cdoc._oleobj_)
            except Exception:
                pass
            try:
                t = _mtype(cdoc)
            except Exception:
                continue
            if _bom_excluded(c2, cdoc):
                continue  # 2026-08-13: 不包括在材料明细表中 → 整棵子树跳过
            p = _mpath(cdoc) or _mtitle(cdoc)
            nm = os.path.splitext(os.path.basename(p))[0] if p else _mtitle(cdoc)
            is_route = ("线路" in nm) or ("步路" in nm) or nm.upper().startswith("PIPE_")
            if t == 2:
                if in_route or is_route:
                    # 步路装配体内的子装配(如 Pipe_2-...装配)也递归展开, 直到线路零件/零部件
                    _walk(cdoc, depth + 1, True)
                else:
                    _walk(cdoc, depth + 1, False)  # 继续往里找步路装配体
            else:
                if in_route or is_route:
                    _add(cdoc)  # 线路零件/步路内零件/零部件

    self_name = os.path.splitext(os.path.basename(asm_path))[0]
    try:
        _walk(asm, 0, ("线路" in self_name) or ("步路" in self_name) or self_name.upper().startswith("PIPE_"))
    finally:
        try:
            sw.CloseDoc(_mtitle(asm))
        except Exception:
            pass
    rows = []
    for (title, spec, length), d in out.items():
        rows.append({"名称": "　└ " + d["title"], "代号": "", "$总数量": d["count"] * qty,
                     "单位": "", "规格": d["spec"], "长度(mm)": d["length"],
                     "$三维截图": "", "_img": None})
    return rows



def _get_mark(cd):
    """2026-08-12: 读"焊接或装配标记"值 — 遍历所有配置用 CPM.Get 简单取值(Get3 的 ByRef 参数有类型坑)。"""
    try:
        cfgs = None
        try:
            cfgs = list(cd.GetConfigurationNames())
        except TypeError:
            cfgs = list(cd.GetConfigurationNames or [])
        for cn in (cfgs or [""]):
            try:
                cpm = cd.Extension.CustomPropertyManager(cn)
                v = cpm.Get(FILTER_PROP)
                if v and str(v).strip():
                    return str(v).strip()
            except Exception:
                continue
    except Exception:
        pass
    return ""


def _agg_component(cd, cat_agg, export_cols, PROCURE_PROP):
    """聚合单个组件(零件或焊接标记装配体)到 cat_agg。2026-08-12 钦定: 焊接件显示自己一行。"""
    try:
        cn = _config_name(cd)
        extra_props = [c for c in export_cols if c not in ("名称", "代号", "$总数量", "$三维截图")]
        props = read_props(cd, [PROCURE_PROP, "名称", "代号"] + extra_props, cn)
        if any(c in extra_props for c in ("$包络尺寸", "$长", "$宽", "$高")):
            props.update(_get_bounding_dimensions(cd))
        cat_val = str(props.get(PROCURE_PROP, "")).strip()
        comp_path = _mpath(cd) or _mtitle(cd)
        if not comp_path:
            return
    except Exception:
        return
    key = (cat_val, comp_path)
    if key not in cat_agg:
        dname = os.path.splitext(os.path.basename(comp_path))[0]
        if not os.path.exists(comp_path):
            dname = props.get("名称") or _mtitle(cd) or dname
        cat_agg[key] = {
            "count": 0,
            "name": dname,
            "code": props.get("代号", ""),
            "unit": props.get("单位", ""),
            "extra": {c: props.get(c, "") for c in extra_props},
            "category": cat_val,
            "path": comp_path,
            "config": cn,
            "doc": cd,
        }
    cat_agg[key]["count"] += 1


def _collect_tree(adoc, cat_agg, export_cols, g, IComp, IMDoc, PROCURE_PROP):
    """2026-08-12 钦定: 递归收集组件(GetComponents(True)=当前装配体顶层; False 返回全树会重复)。
    纪律: 只认"焊接或装配标记"属性 —
      · 标记=焊接的装配体/零件 → 显示自己一行(不展开其内部零件)
      · 无焊接标记的装配体 → 不显示自己, 递归展开内部零件
      · Pipe_ 管段(虚拟线路零件) → 不聚合(已由 PVC管 明细行代表)"""
    try:
        comps = g.IAssemblyDoc(adoc._oleobj_).GetComponents(True)
    except Exception:
        comps = []
    if not comps:
        return
    for comp in comps:
        try:
            c2 = IComp(comp._oleobj_)
            cd = c2.GetModelDoc2()
            if cd is None:
                continue
            cd = IMDoc(cd._oleobj_)
            try:
                t = cd.GetType
                dtype = t() if callable(t) else t
            except Exception:
                dtype = 1
            if _bom_excluded(c2, cd):
                continue  # 2026-08-13: 不包括在材料明细表中 → 整棵子树跳过
            if _get_mark(cd) == "焊接":
                # 焊接标记(装配体或零件): 显示自己一行, 不展开内部
                _agg_component(cd, cat_agg, export_cols, PROCURE_PROP)
                continue
            if dtype == 2:
                # 无焊接标记的装配体: 不显示自己, 递归展开内部零件
                _collect_tree(cd, cat_agg, export_cols, g, IComp, IMDoc, PROCURE_PROP)
                continue
            # 零件: Pipe_ 虚拟线路零件(管段)不聚合, PVC管 明细行已代表
            _p = _mpath(cd) or _mtitle(cd)
            _bn = os.path.basename(str(_p))
            if str(_bn).startswith("Pipe_"):
                continue
            _agg_component(cd, cat_agg, export_cols, PROCURE_PROP)
        except Exception:
            continue


def _shot_opened(doc, sw, title, cfg, out):
    """对已打开的文档(虚拟件)切配置+等轴测截图。返回 img 路径或 None。"""
    try:
        doc.Visible = True
        try:
            sw.ActivateDoc(title)
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
        print("    [步路截图] 失败:", title, cfg, "|", repr(e)[:80])
    return None


def _seg_display_id(cd, rc):
    """管段代号 = 显示状态名里的数字(如 显示状态-694 → 694)。读引用配置的显示状态列表。"""
    import re as _re2
    try:
        dyn = win32com.client.Dispatch(cd._oleobj_)
        cfg = dyn.GetConfigurationByName(rc)
        gds = cfg.GetDisplayStates
        dss = gds() if callable(gds) else gds
        for ds in (dss or []):
            m = _re2.search(r"显示状态[-_]?(\d+)", str(ds))
            if m:
                return m.group(1)
    except Exception:
        pass
    return ""


def pipe_route_rows(sw, asm_doc, comps, qty, img_dir, no_cache=False, max_open=4):
    """2026-08-12 钦定: 装配部件 sheet 展开 Pipe_ 步路装配体内的线路零件(管段)明细。
    每条管段一行 PVC 管:
      名称=PVC管 | 代号=显示状态后缀(694/697/698) | 规格=引用配置 Pipe Identifier | 长度=引用配置 Length
      数量=同代号管段实例数 = Pipe 装配体实例数 × 该装配体内同代号管段组件数
      截图=管段文档切引用配置的等轴测图 | 采购类别=外购件
    管段判定: 引用配置有 Length/SWPipeLength 属性(管件/弯头等外购件无此属性, 不进明细)。"""
    import re as _re
    rows = []
    try:
        from win32com.client import gencache
        g = gencache.EnsureModule('{83A33D31-27C5-11CE-BFD4-00400513BB57}', 0, 30, 0)
    except Exception:
        g = None
    if g is None or sw is None:
        return rows

    # 1. 统计 Pipe_ 装配体在装配树中的实例数 (GetComponents(True) 递归: 全树对虚拟装配体实例会去重, 递归才数得全)
    pipe_counts = collections.Counter()
    pipe_titles = {}

    def _count_pipes(adoc):
        try:
            comps2 = g.IAssemblyDoc(adoc._oleobj_).GetComponents(True)
        except Exception:
            comps2 = []
        for comp in (comps2 or []):
            try:
                c2 = g.IComponent2(comp._oleobj_)
                cd = c2.GetModelDoc2()
                if cd is None:
                    continue
                try:
                    cd = g.IModelDoc2(cd._oleobj_)
                except Exception:
                    pass
                if _bom_excluded(c2, cd):
                    continue  # 2026-08-13: 不包括在材料明细表中 → 整棵子树跳过
                try:
                    t = cd.GetType
                    dtype = t() if callable(t) else t
                except Exception:
                    dtype = 1
                if dtype != 2:
                    continue
                p = _mpath(cd) or _mtitle(cd)
                base = os.path.basename(str(p))
                if "Pipe_" in base or "线路" in base or "步路" in base:
                    pipe_counts[str(p)] += 1
                    pipe_titles[str(p)] = os.path.splitext(base)[0]
                _count_pipes(cd)
            except Exception:
                continue

    try:
        _count_pipes(asm_doc)
    except Exception:
        pass
    if not pipe_counts:
        return rows

    opened = 0
    for p, asm_count in pipe_counts.items():
        if opened >= max_open:
            break
        adoc = open_doc(sw, p, 2)
        if adoc is None:
            continue
        opened += 1
        try:
            try:
                pcomps = g.IAssemblyDoc(adoc._oleobj_).GetComponents(True)
            except Exception:
                pcomps = []
            # 内部管段按 (代号, 规格, 长度) 聚合内部实例数
            seg_agg = collections.OrderedDict()
            for pc in (pcomps or []):
                try:
                    c2 = g.IComponent2(pc._oleobj_)
                    pcd = c2.GetModelDoc2()
                    if pcd is None:
                        continue
                    try:
                        pcd = g.IModelDoc2(pcd._oleobj_)
                    except Exception:
                        pass
                    if _bom_excluded(c2, pcd):
                        continue  # 2026-08-13: 不包括在材料明细表中 → 跳过该管段
                    try:
                        t = pcd.GetType
                        dtype = t() if callable(t) else t
                    except Exception:
                        dtype = 1
                    if dtype != 1:
                        continue
                    rc = str(c2.ReferencedConfiguration or "")
                    if not rc:
                        continue
                    # 管段判定: 引用配置有 Length/SWPipeLength (管件如弯头/球阀无)
                    try:
                        cpm = pcd.Extension.CustomPropertyManager(rc)
                        ln_raw = str(cpm.Get("Length") or cpm.Get("SWPipeLength") or "").strip()
                    except Exception:
                        ln_raw = ""
                    if not ln_raw:
                        continue
                    m = _re.search(r"-?[\d.]+", ln_raw)
                    ln = m.group(0) if m else ""
                    try:
                        cpm2 = pcd.Extension.CustomPropertyManager(rc)
                        spec = str(cpm2.Get("Pipe Identifier") or cpm2.Get("SWbompartno") or "DN 15 ,TK 2.3").strip()
                    except Exception:
                        spec = "DN 15 ,TK 2.3"
                    did = _seg_display_id(pcd, rc) or ""
                    k = (did, spec, ln)
                    if k not in seg_agg:
                        seg_agg[k] = {"code": did, "spec": spec, "length": ln, "count": 0, "doc": pcd, "rc": rc}
                    seg_agg[k]["count"] += 1
                except Exception:
                    continue
            # 输出行: 数量 = 内部实例数 × Pipe 装配体实例数 × qty
            for (did, spec, ln), d in seg_agg.items():
                total = d["count"] * asm_count * qty
                img = ""
                if img_dir and d["doc"] is not None and not no_cache:
                    import time as _time
                    safe = (did or "seg") + "_" + _re.sub(r"[^0-9A-Za-z]", "_", spec) + "_" + (ln or "0")
                    out = os.path.join(img_dir, "pipe_" + safe + ".jpg")
                    if os.path.exists(out) and os.path.getsize(out) > 0:
                        img = out
                    else:
                        img = _shot_opened(d["doc"], sw, pipe_titles.get(p, ""), d["rc"], out)
                rows.append({"名称": "PVC管", "代号": did, "$总数量": total, "单位": "",
                             "规格": spec, "长度(mm)": ln, "$三维截图": img, "_img": img,
                             "采购类别": "外购件"})
        finally:
            try:
                sw.CloseDoc(pipe_titles.get(p, os.path.splitext(os.path.basename(p))[0]))
            except Exception:
                pass
    return rows


def main():
    ap = argparse.ArgumentParser(description="从装配体按 AutoBOM 模板以「焊接或装配标记」属性分类导出 BOM(含等轴测截图 WPS 嵌入)")
    ap.add_argument("--dir", required=True, help="图纸文件夹(自动找总装)")
    ap.add_argument("-o", "--output", required=True, help="输出目录(不放下级)")
    ap.add_argument("--template", required=True, help="AutoBOM模板名(必填), 如 焊接部件清点清单")
    ap.add_argument("--filter-value", required=True, help="焊接或装配标记过滤值(必填), 如 焊接/装配")
    ap.add_argument("--asm", default="", help="可选, 显式指定总装配体.sldasm完整路径(优先于自动识别)")
    ap.add_argument("--qty", type=int, default=1, help="套数/数量, 导出表格的$总数量会乘以此值(默认1, 建议显式指定)")
    ap.add_argument("--img-dir", default="", help="截图存放目录(默认 %s)" % SHOT_CACHE)
    ap.add_argument("--cache-dir", default=shot_cache.DEFAULT_CACHE_DIR, help="截图缓存库目录(默认 %s)" % SHOT_CACHE)
    ap.add_argument("--no-cache", action="store_true", help="禁用截图缓存(强制全部重新截图)")
    ap.add_argument("--no-prop-cache", action="store_true", help="禁用属性缓存(强制重新打开总装扫属性)")
    ap.add_argument("--limit", type=int, default=0, help="仅处理前N个零件(测试用)")
    args = ap.parse_args()
    if args.qty < 1:
        print("错误: --qty 必须 >= 1"); sys.exit(1)
    if not args.dir or not os.path.isdir(args.dir):
        print("错误: --dir 图纸文件夹必填且需存在"); sys.exit(1)
    FILTER_VALUE = args.filter_value

    cols = load_template(args.template)
    if FILTER_PROP not in cols:
        cols = cols + [FILTER_PROP]
    # 2026-08-12: AutoBOM 模板已去掉 $三维截图(防 AutoBOM 自动截图), 本脚本自动补列出图
    if "$三维截图" not in cols:
        cols = cols + ["$三维截图"]
    # 2026-08-11: 焊接部件清单加 规格/长度(mm) 列(步路/线路零件明细行用)
    # 2026-08-13: 加 $长/$宽/$高 包络尺寸列(用户要求清单带长宽高) — 强制加, 不依赖模板
    for _add_col in ("规格", "长度(mm)", "$长", "$宽", "$高"):
        if _add_col not in cols:
            idx = cols.index("$三维截图") if "$三维截图" in cols else len(cols)
            cols.insert(idx, _add_col)

    os.makedirs(args.output, exist_ok=True)
    cache_dir = args.cache_dir or shot_cache.DEFAULT_CACHE_DIR
    os.makedirs(cache_dir, exist_ok=True)
    from datetime import datetime as _dt
    out_file = os.path.join(args.output, f"{args.template}_{_dt.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    print(f"[BOM] 带时间戳输出: {os.path.basename(out_file)}")
    print(f"[FILTER] {FILTER_PROP}={FILTER_VALUE}")
    print(f"[CACHE] 截图库: {cache_dir}")

    # ★ 内存检查(2026-08-07): 可用内存过低直接拒绝
    pct = _free_ram_pct()
    print(f"[MEM] 启动前剩余内存: {pct:.0f}%")
    if pct < 8:
        raise SystemExit(f"错误: 系统可用内存过低 ({pct:.0f}% < 8%). 请先关闭占内存程序后重试.")

    asm = args.asm if args.asm else find_top_assembly(args.dir)
    print("[SCAN] 总装配体:", os.path.basename(asm))
    sw = None
    agg = None
    # ★ 属性缓存优先: 命中则连 SW 都不需要 (2026-08-12 钦定)
    if not args.no_prop_cache:
        _hit, _cdata = prop_cache.get_asm_cache(cache_dir, asm)
        if _hit:
            agg = collect_parts(None, asm, cache_dir, asm_cache_hit=_cdata)
            print(f"[PROPCACHE] 命中属性缓存: {len(agg)} 项 (零 SW, 未打开总装)")
    if agg is None:
        # ★ SW 正常方式启动（自动完成，无需手动打开 SW）
        ensure_sw_normal_start()
        # ★ 启动弹窗守卫
        start_popup_guard()
        try:
            sw = win32com.client.GetActiveObject("SldWorks.Application")
        except Exception:
            sw = win32com.client.Dispatch("SldWorks.Application")  # exec 场景兼容
        agg = collect_parts(sw, asm, cache_dir)
        print(f"[PROPCACHE] 未命中, 打开总装重扫并更新缓存: {len(agg)} 项")
    print(f"[BOM] 聚合零件种类: {len(agg)}")

    hit = collections.OrderedDict()
    for p, e in agg.items():
        if str(e["props"].get(FILTER_PROP, "")).strip() == FILTER_VALUE:
            hit[p] = e
    print(f"[BOM] 过滤 {FILTER_PROP}={FILTER_VALUE}: {len(hit)} 行")

    export_cols = [c for c in cols if c != FILTER_PROP]

    if not hit:
        # 2026-08-12: 0 行也必须输出空表+告警, 不许静默不生成
        # 2026-08-13: 文件名按 filter 区分 — 焊接时=焊接部件清点清单_xxx, 装配时=装配部件清点清单_xxx
        _empty_name = args.template.replace("焊接", FILTER_VALUE) if "焊接" in args.template else f"{args.template}_{FILTER_VALUE}"
        _empty_out = os.path.join(args.output, f"{_empty_name}_{_dt.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        print(f"[WARN] 过滤 {FILTER_PROP}={FILTER_VALUE} 为 0 行! 输出空表: {os.path.basename(_empty_out)}")
        try:
            export_multi_sheet_xlsx(export_cols, {"空清单": []}, _empty_out)
        except Exception as _e:
            print(f"[WARN] 空表写出失败: {repr(_e)[:100]}")

    export_cols = [c for c in cols if c != FILTER_PROP]

    rows = []
    cache_hit = 0
    cache_new = 0
    for i, (p, e) in enumerate(hit.items()):
        if args.limit and i >= args.limit:
            break
        name = os.path.splitext(os.path.basename(p))[0]
        cfg = e.get("config", "")
        st = None
        try:
            st = os.stat(p)
        except Exception:
            st = None
        img = None
        # 缓存优先: 命中则完全不开 SW
        if st and not args.no_cache:
            img = shot_cache.get_shot(cache_dir, p, cfg, st.st_size, st.st_mtime)
            if img:
                cache_hit += 1
                print(f"    [缓存] {name} ({cfg})")
        if img is None:
            # 2026-08-12: 必须强类型(gen_py) — Dispatch 动态绑定下 OpenDoc6(0,0) 类型不匹配, 且毒化后续 sw
            if sw is None:
                ensure_sw_normal_start()
                start_popup_guard()
                for _i in range(30):
                    try:
                        sw = win32com.client.GetActiveObject("SldWorks.Application")
                        break
                    except Exception:
                        time.sleep(3)
                print("[SW] 阶段1 懒连接 SW (强类型)")
            sw = _maybe_restart_sw(sw, i)
            img = shot_iso(sw, p, cache_dir, config_name=cfg)
            if img:
                cache_new += 1
                if st:
                    shot_cache.put_shot(cache_dir, p, cfg, name, cfg, img, st.st_size, st.st_mtime)
                print(f"    [图] {os.path.basename(img)}")
        row = {"_img": img}
        # 实时算焊接组件(装配体)的长宽高包络尺寸: 打开 -> GetBox -> 关闭
        box_dims = {}
        if any(c in export_cols for c in ("$包络尺寸", "$长", "$宽", "$高")):
            try:
                # 2026-08-13: 按扩展名定文档类型 — 焊接标记零件(.sldprt)用 type2 打开会失败
                _dtype = 2 if p.lower().endswith(".sldasm") else 1
                bd = open_doc(sw, p, _dtype)
                if bd is not None:
                    box_dims = _get_bounding_dimensions(bd)
                    try:
                        sw.CloseDoc(_mtitle(bd))
                    except Exception:
                        pass
            except Exception:
                pass
        for c in export_cols:
            if c == "名称":
                row[c] = name
            elif c == "代号":
                row[c] = e["props"].get("代号", "")
            elif c == "$总数量":
                row[c] = e["count"] * args.qty
            elif c in ("$长", "$宽", "$高", "$包络尺寸"):
                row[c] = box_dims.get(c, e["props"].get(c, ""))
            elif c == "单位":
                row[c] = e["props"].get(c, "")
            elif c == "$三维截图":
                row[c] = img or ""
            else:
                row[c] = e["props"].get(c, "")
        rows.append(row)
        # 2026-08-12: 汇总 sheet 不再展开步路明细 — 外购件/虚拟装配体不进清单汇总, 明细只在各组件 sheet

    phase1_rows = rows
    phase1_cols = export_cols
    # 暂存阶段1数据, 待阶段2一起合并到同一个xlsx
    print(f"\n[OK] 阶段1: 共 {len(hit)} 个焊接组件, 截图库: {cache_dir}")
    print(f"[OK] 共 {len(hit)} 个焊接组件, 截图库: {cache_dir}")

    
    # ===== 第二阶段: 对每个焊接子装配体导出分类 BOM (所有采购类别合并到同一sheet) =====
    if hit:
        print(f"\n===== 阶段2: 焊接子装配体分类BOM (全部采购类别合并) =====")
        # 2026-08-12: 属性缓存命中时 sw 为 None, 阶段2 要开子装配必须延迟连 SW
        # 必须 GetActiveObject 拿强类型(gen_py): Dispatch 动态绑定下 OpenDoc6(0,0) 会类型不匹配
        if sw is None:
            ensure_sw_normal_start()
            start_popup_guard()
            sw = None
            for _i in range(30):
                try:
                    sw = win32com.client.GetActiveObject("SldWorks.Application")
                    break
                except Exception:
                    time.sleep(3)
            if sw is None:
                sw = win32com.client.Dispatch("SldWorks.Application")
            print("[SW] 阶段2 延迟连接 SW (强类型)")
        CATEGORIES = ["外购件", "标准件", "钣金件", "型材", "机加件"]
        PROCURE_PROP = "采购类别"
        CATEGORY_COL = "采购类别"
        IComp = gen().IComponent2
        IMDoc = gen().IModelDoc2
        weld_date = _dt.now().strftime("%Y%m%d")

        # 阶段2列 = 模板列(剔除FILTER_PROP) + 采购类别
        phase2_cols = [c for c in export_cols if c != CATEGORY_COL] + [CATEGORY_COL]

        # 关闭总装释放锁
        try:
            top_title = os.path.splitext(os.path.basename(asm))[0]
            sw.CloseDoc(top_title)
            print(f"[CLOSE] 已关闭总装: {top_title}")
        except Exception as e:
            print(f"[WARN] 关闭总装失败: {repr(e)[:50]}")
        time.sleep(1)

        # 收集所有 sheet: { sheet_name: [row_dicts] }
        all_sheets = collections.OrderedDict()

        for weld_idx, (weld_path, weld_info) in enumerate(hit.items()):
            weld_name = os.path.splitext(os.path.basename(weld_path))[0]
            if not weld_path.lower().endswith(".sldasm"):
                print(f"  [{weld_idx+1}/{len(hit)}] {weld_name} 不是装配体, 跳过")
                continue
            print(f"\n  [{weld_idx+1}/{len(hit)}] 焊接子装配体: {weld_name}")

            # 打开焊接子装配体
            weld_asm = open_doc(sw, weld_path, 2)
            if weld_asm is None:
                print(f"    [FAIL] 无法打开: {weld_name}")
                continue
            try:
                gen().IAssemblyDoc(weld_asm._oleobj_).ResolveAllLightWeightComponents(False)
            except Exception:
                pass

            try:
                weld_comps = gen().IAssemblyDoc(weld_asm._oleobj_).GetComponents(False)
            except Exception:
                weld_comps = []
            print(f"    [SCAN] 组件数: {len(weld_comps) if weld_comps else 0}")

            if not weld_comps:
                sw.CloseDoc(weld_name)
                continue

            # 按 (采购类别, 文件路径) 聚合 (2026-08-12: 递归收集 GetComponents(True)=顶层, 焊接标记子树整棵剔除)
            cat_agg = collections.OrderedDict()
            _collect_tree(weld_asm, cat_agg, export_cols, gen(), IComp, IMDoc, PROCURE_PROP)

            # 按采购类别排序
            cat_order = {c: i for i, c in enumerate(CATEGORIES)}
            sorted_keys = sorted(cat_agg.keys(), key=lambda k: (cat_order.get(k[0], 99), k[1]))

            # 汇总到此焊接组件的 sheet
            sheet_rows = []
            for ki, key in enumerate(sorted_keys):
                info = cat_agg[key]
                comp_path = info["path"]
                cat_name = info["category"]
                cfg = info.get("config", "")
                st = None
                try:
                    st = os.stat(comp_path)
                except Exception:
                    st = None
                img = None
                # 缓存优先: 命中则完全不开 SW
                if st and not args.no_cache:
                    img = shot_cache.get_shot(cache_dir, comp_path, cfg, st.st_size, st.st_mtime)
                    if img:
                        cache_hit += 1
                        print(f"      [{cat_name}] [缓存] {info['name']}")
                if img is None:
                    if sw is None:
                        try:
                            sw = win32com.client.GetActiveObject("SldWorks.Application")
                        except Exception:
                            sw = win32com.client.Dispatch("SldWorks.Application")
                    sw = _maybe_restart_sw(sw, ki)
                    # 2026-08-13 方案B: 子装配还开着时内部零件文档已加载 → ActivateDoc 直接截(不反复开关文档)
                    doc = info.get("doc")
                    img = None
                    if doc is not None:
                        img = _shot_opened(doc, sw, os.path.splitext(os.path.basename(comp_path))[0], cfg,
                                           shot_cache.shot_path_of(cache_dir, comp_path, cfg))
                    if img is None:
                        # SW 重启后旧 doc 对象 RPC 失效 → 回退重新打开截图
                        img = shot_iso(sw, comp_path, cache_dir, config_name=cfg)
                    if img:
                        cache_new += 1
                        if st:
                            shot_cache.put_shot(cache_dir, comp_path, cfg, info["name"], cfg, img,
                                                st.st_size, st.st_mtime)
                        print(f"      [{cat_name}] [图] {os.path.basename(img)}")
                item = {"_img": img, CATEGORY_COL: cat_name}
                item["名称"] = info["name"]
                item["代号"] = info["code"]
                item["$总数量"] = info["count"] * args.qty
                # 其余模板列从 info 读取
                for c in export_cols:
                    if c not in ("名称", "代号", "$总数量", "$三维截图"):
                        if c == "单位":
                            item[c] = info.get("unit", "")
                        elif c in ("$包络尺寸", "$长", "$宽", "$高"):
                            item[c] = info.get("extra", {}).get(c, "")
                        elif c not in item:
                            item[c] = ""
                item["$三维截图"] = img or ""
                sheet_rows.append(item)

            # 2026-08-12 钦定: 装配部件 sheet 展开 Pipe 虚拟装配体内线路零件(管)明细 名称/规格/长度
            pipe_rows = pipe_route_rows(sw, weld_asm, weld_comps, args.qty, args.cache_dir, args.no_cache)
            if pipe_rows:
                print(f"    [步路] {weld_name}: 线路零件明细 {len(pipe_rows)} 行")
                sheet_rows.extend(pipe_rows)

            if sheet_rows:
                sheet_name = weld_name
                if len(sheet_name) > 31:
                    sheet_name = sheet_name[:28] + "..."
                base_name = sheet_name
                counter = 1
                while sheet_name in all_sheets:
                    suffix = f"_{counter}"
                    sheet_name = base_name[:31-len(suffix)] + suffix
                    counter += 1
                all_sheets[sheet_name] = sheet_rows
                total_in_sheet = sum(r["$总数量"] for r in sheet_rows)
                print(f"    [汇总] {len(sheet_rows)}行, 总数量={total_in_sheet}(x{args.qty}套)")

            try:
                sw.CloseDoc(weld_name)
            except Exception:
                pass
            time.sleep(1)

        # 合并阶段1+阶段2到一个xlsx
        _kind = "装配" if FILTER_VALUE == "装配" else "焊接"
        merged_out = os.path.join(
            args.output,
            f"{_kind}部件分类BOM_{weld_date}_{_dt.now().strftime('%H%M%S')}.xlsx"
        )

        # 合并sheets: 第1个=阶段1焊接清单, 后续=各焊接组件分类
        merged_sheets = collections.OrderedDict()
        if phase1_rows:
            merged_sheets[f"{_kind}部件清单"] = phase1_rows
        merged_sheets.update(all_sheets)

        # 统一列: 名称 代号 $总数量 单位 采购类别 $三维截图
        unified_cols = phase2_cols
        for sn, srows in merged_sheets.items():
            for row in srows:
                if "采购类别" not in row:
                    row["采购类别"] = ""

        export_multi_sheet_xlsx(unified_cols, merged_sheets, merged_out)
        print(f"\n[OK] 合并BOM(阶段1+阶段2): {merged_out}")
        print(f"    Sheet [焊接部件清单]: {len(phase1_rows)} 个焊接组件")
        for sn, sr in all_sheets.items():
            print(f"    Sheet [{sn}]: {len(sr)} 行")
        n, mb = shot_cache.stats(cache_dir)
        print(f"[CACHE] 本次: 命中 {cache_hit} 张, 新截 {cache_new} 张; 截图库共 {n} 条记录 / {mb:.1f} MB")


    # ★ 清理: 关闭弹窗守卫 + 退出 SW 释放内存
    stop_popup_guard()
    release_sw_memory(force_exit=True)
    print("\n[DONE] skill8 焊接分类BOM导出完成 (含阶段1焊接清单 + 阶段2分类BOM), SW 已退出释放内存.")


if __name__ == "__main__":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    try:
        main()
    finally:
        # 2026-08-12: 异常路径也必须清守卫+退SW, 否则残留守卫干扰皇上用 SW
        try:
            stop_popup_guard()
        except Exception:
            pass
        try:
            release_sw_memory(force_exit=True)
        except Exception:
            pass
