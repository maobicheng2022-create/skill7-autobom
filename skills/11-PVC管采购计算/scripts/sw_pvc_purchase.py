# -*- coding: utf-8 -*-
"""skill11-PVC管采购计算
从 SW 总装配体出发, 读取所有步路/线路零件(PVC管段, 含装配体内的虚拟线路零件):
- 规格: 优先取属性 规格/直径/尺寸, 缺则用配置名 (如 DN15)
- 长度: SW 自带"长度"属性(mm) (SW Routing 自动写入每个线路零件), 缺则用零件包围盒最长边
×套数 qty → 按规格独立做 6米(6000mm)管 1D 下料(FFD): 一根6m管切完剩余不够再切就浪费,
算每规格最少需要几根6m管 + 总余料, 输出 PVC管采购表.xlsx (汇总 + 切割明细)。

用法:
  python -X utf8 sw_pvc_purchase.py --dir <图纸文件夹> [--asm <总装.sldasm>] [--qty 套数] [-o 输出目录]
"""
import argparse
import collections
import io
import math
import os
import re
import subprocess
import sys
import time

import win32com.client
import pythoncom

AUTOBOM_CFG = r"C:\micDatas\pluginSettings\BOM\setting.json"
POPUP_GUARD = r"C:\Users\mao\.openclaw\workspace\turing-workspace\scripts\sw_popup_guard.py"
SLDWORKS_CANDIDATES = [
    r"E:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\sldworks.exe",
    r"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\sldworks.exe",
    r"D:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\sldworks.exe",
    "sldworks.exe",
]
STOCK_MM = 6000.0  # 一根 PVC 管标准 6 米

_guard_proc = None


def start_popup_guard():
    global _guard_proc
    if _guard_proc is not None and _guard_proc.poll() is None:
        print("[GUARD] 弹窗守卫已在运行, 跳过重复启动")
        return
    if not os.path.exists(POPUP_GUARD):
        print("[WARN] 未找到 sw_popup_guard.py, 跳过弹窗守卫")
        return
    _guard_proc = subprocess.Popen([sys.executable, "-X", "utf8", POPUP_GUARD, "1.0"],
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
    try:
        win32com.client.GetActiveObject("SldWorks.Application")
        print("[SW] 已连上运行中的 SW")
        return
    except Exception as e:
        print(f"[SW] GetActiveObject 失败({repr(e)[:50]}), 尝试正常方式启动 SW...")
    exe = _sw_exe_from_registry()
    cands = [exe] if exe else SLDWORKS_CANDIDATES
    for cand in cands:
        if os.path.exists(cand):
            subprocess.Popen([cand])
            break
    for i in range(120):
        time.sleep(1)
        try:
            win32com.client.GetActiveObject("SldWorks.Application")
            print(f"[SW] SW 已正常启动并可连接(等了 {i+1}s)")
            time.sleep(3)
            return
        except Exception:
            continue
    raise RuntimeError("SW 未能在120s内启动并连上")


def _gen():
    import glob as _glob
    import win32com.client.gencache as gencache
    base = gencache.GetGeneratePath()
    hits = [os.path.join(base, n) for n in os.listdir(base) if n.startswith("83A33D31-27C5-11CE-BFD4-00400513BB57x0x") and n.endswith(".py")]
    if not hits:
        raise RuntimeError("未找到 SW 类型库缓存,请先运行: python -m win32com.client.makepy sldworks.tlb")
    modname = "win32com.gen_py." + os.path.splitext(os.path.basename(hits[0]))[0]
    return __import__(modname, fromlist=['IModelDoc2', 'IAssemblyDoc', 'IComponent2'])


_GEN = None
def gen():
    global _GEN
    if _GEN is None:
        _GEN = _gen()
    return _GEN


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
    r = sw.OpenDoc6(path, dtype, 1, "", 0, 0)
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


def find_top_assembly(directory):
    cands = []
    for fn in os.listdir(directory):
        if not fn.lower().endswith(".sldasm"):
            continue
        if fn.startswith("~") or fn.startswith("ST_") or fn.lower().startswith("st_") \
                or fn.upper().startswith("PIPE_"):
            continue
        code = fn.split("_")[0]
        cands.append((len(code), code, os.path.join(directory, fn)))
    if not cands:
        raise ValueError("目录下未找到总装配体(.sldasm, 排除 ~/ST_/Pipe_): " + directory)
    cands.sort(key=lambda x: (x[0], x[1]))
    return cands[0][2]


def _to_float(s):
    if s is None:
        return 0.0
    m = re.search(r"-?[\d.]+(?:[eE][+-]?\d+)?", str(s).replace(",", "").replace("mm", "").strip())
    return float(m.group(0)) if m else 0.0


def collect_pipes(sw, asm_path):
    """打开总装, 收集所有 线路零件/Pipe_ 管段: (名称, 规格, 长度mm)。含虚拟件。
    判据: 名称含 线路/Pipe_ 前缀, 或有 SW Routing 自动写入的 Length/SWPipeLength 属性。"""
    asm = open_doc(sw, asm_path, 2)
    if asm is None or _mtype(asm) != 2:
        raise RuntimeError("无法打开总装配体: " + asm_path)
    try:
        gen().IAssemblyDoc(asm._oleobj_).ResolveAllLightWeightComponents(False)
    except Exception:
        pass
    comps = gen().IAssemblyDoc(asm._oleobj_).GetComponents(False)
    pipes = []
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
                doc = _wrap(doc)
                if excluded or _bom_excluded(c2, doc):
                    continue  # 2026-08-13: 不包括在材料明细表中 → 整棵子树跳过
                if _mtype(doc) == 2:
                    _walk(doc, False)  # 装配体不是管段, 但递归内部
                    continue
                p = _mpath(doc) or _mtitle(doc)
                nm = os.path.splitext(os.path.basename(p))[0] if p else _mtitle(doc)
                cn = _config_name(doc)
                props = read_props(doc, ["名称", "规格", "直径", "尺寸", "长度", "Length", "SWPipeLength",
                                         "Pipe Identifier", "SWbompartno"], cn)
                has_len = bool(props.get("Length") or props.get("SWPipeLength") or props.get("长度"))
                is_pipe = (("线路" in nm) or nm.upper().startswith("PIPE_")) or has_len
                if not is_pipe:
                    continue
                # 规格: Pipe Identifier → SWbompartno → 规格/直径/尺寸 → 配置名
                spec = (props.get("Pipe Identifier") or props.get("SWbompartno")
                        or props.get("规格") or props.get("直径") or props.get("尺寸") or cn)
                spec = str(spec).strip() or "(未知名)"
                # 长度: SW Routing 自带 Length(带mm后缀) → SWPipeLength → 长度 → 包围盒最长边
                ln = _to_float(props.get("Length") or props.get("SWPipeLength") or props.get("长度"))
                if not ln:
                    try:
                        pd = gen().IPartDoc(doc._oleobj_)
                        box = pd.GetPartBox(False)
                        if box and len(box) >= 6:
                            ln = max(abs(box[3] - box[0]), abs(box[4] - box[1]), abs(box[5] - box[2]))
                    except Exception:
                        ln = 0.0
                pipes.append({"name": nm, "spec": spec, "length_mm": round(ln, 1)})
            except Exception:
                continue

    _walk(asm)
    try:
        sw.CloseDoc(_mtitle(asm))
    except Exception:
        pass
    return pipes


def pack_6m(segments):
    """segments: [(长度mm, 数量)] → (根数, 余料mm, 单段超6m数量)。FFD 一维下料, 6m 独立算。"""
    items = []
    for L, c in segments:
        items.extend([float(L)] * int(round(c)))
    items.sort(reverse=True)
    bins = []  # 每根剩余可切长度
    over = 0
    for L in items:
        if L > STOCK_MM:
            bins.append(0.0)  # 单段超6m: 独占一根(余料0, 需另购非标)
            over += 1
            continue
        placed = False
        for i in range(len(bins)):
            if bins[i] >= L:
                bins[i] -= L
                placed = True
                break
        if not placed:
            bins.append(STOCK_MM - L)
    return len(bins), round(sum(bins), 1), over


def main():
    ap = argparse.ArgumentParser(description="PVC管采购计算: 总装 → 线路零件规格/长度 ×套数 → 6m管切割根数")
    ap.add_argument("--dir", required=True, help="图纸文件夹(自动找总装)")
    ap.add_argument("--asm", default="", help="可选, 显式指定总装配体.sldasm完整路径")
    ap.add_argument("--qty", type=int, default=1, help="套数/数量(默认1)")
    ap.add_argument("-o", "--output", default="", help="输出目录(默认图纸文件夹)")
    args = ap.parse_args()
    if args.qty < 1:
        raise SystemExit("错误: --qty 必须 >= 1")
    if not os.path.isdir(args.dir):
        raise SystemExit(f"错误: 文件夹不存在: {args.dir}")
    out_dir = args.output or args.dir
    os.makedirs(out_dir, exist_ok=True)

    start_popup_guard()
    try:
        ensure_sw_normal_start()
        asm = args.asm if args.asm else find_top_assembly(args.dir)
        print("[SCAN] 总装配体:", os.path.basename(asm))
        sw = win32com.client.GetActiveObject("SldWorks.Application")
        pipes = collect_pipes(sw, asm)
        print(f"[PIPE] 线路/Pipe_ 管段实例: {len(pipes)}")
        if not pipes:
            raise SystemExit("未找到任何线路零件/Pipe_ 管段, 请检查总装与步路零件")

        # 按规格聚合: spec -> {length_mm: count}
        by_spec = collections.OrderedDict()
        for p in pipes:
            if p["spec"] not in by_spec:
                by_spec[p["spec"]] = collections.OrderedDict()
            by_spec[p["spec"]][p["length_mm"]] = by_spec[p["spec"]].get(p["length_mm"], 0) + 1

        from datetime import datetime as _dt
        ts = _dt.now().strftime("%Y%m%d_%H%M%S")
        out_file = os.path.join(out_dir, f"PVC管采购表_{ts}.xlsx")

        # ===== 汇总打印 + 生成 xlsx =====
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Font
        from openpyxl.utils import get_column_letter
        wb = Workbook()
        ws1 = wb.active
        ws1.title = "采购汇总"
        ws1.append(["规格", "段数(单套)", f"段数(x{args.qty}套)", "总长度(m)",
                    f"需6m管(根)", "余料(m)", "备注"])
        ws2 = wb.create_sheet("切割明细")
        ws2.append(["规格", "每段长度(mm)", "数量(单套)", f"数量(x{args.qty}套)", "小计(m)", "来源"])
        grand_bins = 0
        for spec, segs in by_spec.items():
            cnt1 = sum(segs.values())
            segs_qty = {L: c * args.qty for L, c in segs.items()}
            total_mm = sum(L * c for L, c in segs_qty.items())
            bins, left, over = pack_6m(list(segs_qty.items()))
            grand_bins += bins
            detail = " + ".join(f"{int(L)}mm×{c}" for L, c in sorted(segs_qty.items(), reverse=True))
            note = ""
            if over:
                note = f"含{over}段单段超6m(需另购)"
            ws1.append([spec, cnt1, sum(segs_qty.values()), round(total_mm / 1000, 2),
                        bins, round(left / 1000, 2), note])
            for L, c in sorted(segs_qty.items(), reverse=True):
                ws2.append([spec, L, segs.get(L, 0), c, round(L * c / 1000, 2), "步路/线路零件"])
            print(f"  {spec:<22} x{qty}套 {sum(segs_qty.values()):>3}段 总长{total_mm/1000:>7.2f}m "
                  f"→ 需6m管 {bins} 根 (余料 {left/1000:.2f}m){note}")
        print(f"\n合计: {len(by_spec)} 种规格, 共需 {grand_bins} 根 6m PVC管")
        print(f"[OK] 已输出: {out_file}")

        # 样式
        wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for ws in (ws1, ws2):
            for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
                for cell in row:
                    cell.alignment = wrap
            for col in range(1, ws.max_column + 1):
                ws.column_dimensions[get_column_letter(col)].width = 16
            ws.column_dimensions["A"].width = 10
        for cell in ws1[1]:
            cell.font = Font(bold=True)
        for cell in ws2[1]:
            cell.font = Font(bold=True)
        wb.save(out_file)
    finally:
        stop_popup_guard()


if __name__ == "__main__":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    main()
