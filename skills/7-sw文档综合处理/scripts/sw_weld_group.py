# -*- coding: utf-8 -*-
"""sw_weld_group.py — skill7 焊接装配体分组输出 (2026-08-29 新增)
命令:
  collect          : SW 开总装递归收集「焊接或装配标记=焊接」装配体分组 -> folder\\焊接分组.json
  merge-ban        : 钣金件\\ 按组分 DXF + 各焊接组独立合并图纸 (纯文件, 调 dwg_merge)
  prep-profile     : 型材\\ 建 _weld_tmp 各组工作目录 (复制组型材件 + 写过滤 型材1.CSV)
  collect-profile  : 汇总各组带数量零件 -> 型材\\带数量零件\\<组名>\\ , 删 _weld_tmp
归属规则(2026-08-29 皇上钦定):
  1) 标记=焊接的装配体 -> 一个组, 组内全部零件归它
  2) 无焊接标记 -> 非焊接组(ungrouped), 走原总合并/总带数量零件(焊接组已剔除)
  3) 嵌套焊接(A 内套 B): B 零件归 B 组, A 组不含 B 的零件(最近焊接祖先)
用法:
  python sw_weld_group.py collect --dir <folder> [--asm <asm>]
  python sw_weld_group.py merge-ban --dir <folder>
  python sw_weld_group.py prep-profile --dir <folder>
  python sw_weld_group.py collect-profile --dir <folder>
"""
import argparse
import json
import os
import shutil
import subprocess
import sys
import time

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

PY = sys.executable or "python"
S6 = os.path.normpath(os.path.join(GLOBAL_SKILLS, "6-合并DXF", "scripts"))
DWG_MERGE = os.path.join(S6, "dwg_merge.py")
FILTER_PROP = "焊接或装配标记"

_guard_proc = None


def log(msg):
    print(msg, flush=True)


# ---------------- 弹窗守卫(全量清理版, 抄 skill8) ----------------
def start_popup_guard():
    global _guard_proc
    if _guard_proc is not None and _guard_proc.poll() is None:
        log("[GUARD] 弹窗守卫已在运行, 跳过")
        return
    if not os.path.exists(POPUP_GUARD):
        log(f"[WARN] 未找到弹窗守卫 {POPUP_GUARD}")
        return
    log(f"[GUARD] 启动弹窗守卫: {POPUP_GUARD}")
    _guard_proc = subprocess.Popen([PY, "-X", "utf8", POPUP_GUARD, "1.0"],
                                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def stop_popup_guard():
    global _guard_proc
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
    killed = 0
    try:
        ps = ("Get-CimInstance Win32_Process | "
              "Where-Object { $_.CommandLine -like '*sw_popup_guard*' -and $_.ProcessId -ne $PID } | "
              "Select-Object -ExpandProperty ProcessId")
        out = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
                             capture_output=True, timeout=30)
        raw = (out.stdout or b"")
        for ln in raw.decode("ascii", errors="ignore").splitlines():
            ln = ln.strip()
            if ln.isdigit():
                try:
                    subprocess.run(["taskkill", "/F", "/PID", str(ln)], capture_output=True, timeout=10)
                    killed += 1
                except Exception:
                    pass
    except Exception:
        pass
    if killed:
        log(f"[GUARD] 已清理 {killed} 个残留守卫")
    log("[GUARD] 弹窗守卫全部停止")


# ---------------- SW 正常方式启动(抄 skill8) ----------------
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
    import win32com.client
    try:
        win32com.client.GetActiveObject("SldWorks.Application")
        log("[SW] 已连上运行中的 SW")
        return
    except Exception as e:
        log(f"[SW] GetActiveObject 失败({repr(e)[:50]}), 尝试正常方式启动 SW...")
    exe = _sw_exe_from_registry()
    cands = [exe] if exe else SLDWORKS_CANDIDATES
    for cand in cands:
        if os.path.exists(cand):
            log(f"[SW] Start-Process 正常方式启动: {cand}")
            subprocess.Popen([cand])
            break
    for i in range(120):
        time.sleep(1)
        try:
            win32com.client.GetActiveObject("SldWorks.Application")
            log(f"[SW] SW 已正常启动并可连接(等了 {i+1}s)")
            time.sleep(3)
            return
        except Exception:
            continue
    raise RuntimeError("SW 未能在120s内以正常方式启动并连上")


# ---------------- SW 对象辅助(抄 skill8) ----------------
def _gen():
    import glob as _glob
    import win32com.client.gencache as gencache
    base = gencache.GetGeneratePath()
    hits = [os.path.join(base, n) for n in os.listdir(base) if n.startswith("83A33D31-27C5-11CE-BFD4-00400513BB57x0x") and n.endswith(".py")]
    if not hits:
        raise RuntimeError("未找到 SW 类型库缓存, 请先运行: python -m win32com.client.makepy sldworks.tlb")
    modname = "win32com.gen_py." + os.path.splitext(os.path.basename(hits[0]))[0]
    return __import__(modname, fromlist=['IModelDoc2', 'IAssemblyDoc', 'IComponent2'])


_GEN = None


def gen():
    global _GEN
    if _GEN is None:
        _GEN = _gen()
    return _GEN


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


def _get_mark(cd):
    """遍历所有配置读「焊接或装配标记」值(CPM.Get 简单取值, Get3 的 ByRef 有类型坑)。"""
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


def _bom_excluded(c2, doc):
    """组件级 ExcludeFromBOM 或 文档配置属性 排除 -> True。2026-08-13 钦定全链路过滤。"""
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


def _sanitize(name):
    for ch in '\\/:*?"<>|':
        name = name.replace(ch, "_")
    return name.strip() or "焊接组"


def _code_of(path):
    return os.path.splitext(os.path.basename(path))[0].split("_")[0]


# ---------------- collect: SW 递归收集分组 ----------------
def find_top_assembly(directory):
    cands = []
    for fn in os.listdir(directory):
        if not fn.lower().endswith(".sldasm"):
            continue
        if fn.startswith("~") or fn.startswith("ST_") or fn.upper().startswith("PIPE_"):
            continue
        code = fn.split("_")[0]
        cands.append((len(code), code, os.path.join(directory, fn)))
    if not cands:
        raise RuntimeError("目录下未找到装配体(.sldasm, 排除 ~$/ST_/Pipe_): " + directory)
    cands.sort(key=lambda x: (x[0], x[1]))
    return cands[0][2]


def collect_groups(sw, asm_path):
    """开总装递归遍历 -> (groups: {组名: [零件路径]}, ungrouped: [零件路径])"""
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    G = gen()
    IModelDoc2 = G.IModelDoc2
    IAssemblyDoc = G.IAssemblyDoc
    IComponent2 = G.IComponent2

    r = sw.OpenDoc6(asm_path, 2, 0, "", 0, 0)
    md = r[0] if isinstance(r, (tuple, list)) else r
    if md is None:
        raise RuntimeError("打不开总装: " + asm_path)
    try:
        sw.ActivateDoc(os.path.splitext(os.path.basename(asm_path))[0])
    except Exception:
        pass
    try:
        adoc = IAssemblyDoc(md._oleobj_)
    except Exception:
        raise RuntimeError("总装不是装配体文档: " + asm_path)
    # 2026-08-29: 子装配体默认轻化, GetModelDoc2/GetComponents 拿不到内部组件 -> 必须解析全部
    try:
        adoc.ResolveAllLightWeightComponents(False)
    except Exception:
        pass

    groups = {}
    ungrouped = []

    def walk(doc, group):
        # 2026-08-29: doc 可能是 IModelDoc2, GetComponents 必须转 IAssemblyDoc(否则抛异常被吞成空)
        try:
            doc_asm = IAssemblyDoc(doc._oleobj_)
            comps = doc_asm.GetComponents(True)
        except Exception:
            comps = []
        if not comps:
            return
        for comp in comps:
            try:
                c2 = IComponent2(comp._oleobj_)
                raw = c2.GetModelDoc2()
                if raw is None:
                    continue
                cd = IModelDoc2(raw._oleobj_)
                dtype = _mtype(cd)
                if _bom_excluded(c2, cd):
                    continue
                mark = _get_mark(cd)
                if dtype == 2:
                    if mark == "焊接":
                        gname = _sanitize(_mtitle(cd) or _code_of(_mpath(cd)))
                        if gname not in groups:
                            groups[gname] = []
                        walk(cd, gname)  # 嵌套焊接: 内部零件归新组
                    else:
                        walk(cd, group)
                    continue
                # 零件
                p = _mpath(cd)
                if not p or not os.path.exists(p):
                    continue
                if group:
                    groups[group].append(p)
                else:
                    ungrouped.append(p)
            except Exception:
                continue

    walk(adoc, None)

    try:
        sw.CloseDoc(os.path.splitext(os.path.basename(asm_path))[0])
    except Exception:
        pass
    # 去重保序
    for k in groups:
        groups[k] = list(dict.fromkeys(groups[k]))
    ungrouped = list(dict.fromkeys(ungrouped))
    return groups, ungrouped


def cmd_collect(args):
    import win32com.client
    start_popup_guard()
    try:
        ensure_sw_normal_start()
        sw = win32com.client.GetActiveObject("SldWorks.Application")
        folder = os.path.abspath(args.dir)
        asm = args.asm if args.asm else find_top_assembly(folder)
        log(f"[COLLECT] 总装: {asm}")
        groups, ungrouped = collect_groups(sw, asm)
        out = os.path.join(folder, "焊接分组.json")
        with open(out, "w", encoding="utf-8") as f:
            json.dump({"groups": groups, "ungrouped": ungrouped}, f, ensure_ascii=False, indent=1)
        log(f"[COLLECT] 焊接组 {len(groups)} 个, 非焊接零件 {len(ungrouped)} 个")
        for k, v in groups.items():
            log(f"  [组] {k}: {len(v)} 零件")
        log(f"[COLLECT] 已写出: {out}")
    finally:
        stop_popup_guard()


# ---------------- 纯文件操作 ----------------
def load_groups(folder):
    p = os.path.join(folder, "焊接分组.json")
    if not os.path.exists(p):
        raise RuntimeError("缺少 焊接分组.json, 请先跑 collect: " + p)
    with open(p, encoding="utf-8") as f:
        return json.load(f)


def read_csv_map(path):
    """读 csv(utf-8-sig): 返回 (header, rows); rows 为 dict 列表。"""
    import csv as _csv
    with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
        rdr = list(_csv.reader(f))
    if not rdr:
        return [], []
    header = rdr[0]
    rows = []
    for r in rdr[1:]:
        if r and r[0].strip():
            rows.append(dict(zip(header, r)))
    return header, rows


def run_dwg_merge(gdir):
    log(f"[MERGE] 合并: {gdir}")
    subprocess.run([PY, "-X", "utf8", DWG_MERGE, gdir], check=False)
    time.sleep(1)
    outs = [f for f in os.listdir(gdir) if f.startswith("合并图纸_") and f.endswith(".dxf")]
    if not outs:
        log(f"[WARN] 组目录未产出合并图纸: {gdir}")
    else:
        log(f"[MERGE] 产出: {outs}")


def cmd_merge_ban(args):
    data = load_groups(args.dir)
    folder = os.path.abspath(args.dir)
    ban_dir = os.path.join(folder, "钣金件")
    if not os.path.isdir(ban_dir):
        raise RuntimeError("缺少钣金件目录: " + ban_dir)
    csv_path = os.path.join(ban_dir, "钣金件.csv")
    if not os.path.exists(csv_path):
        raise RuntimeError("缺少配料表: " + csv_path)
    header, rows = read_csv_map(csv_path)
    # 2026-10-17 修复: 本步骤末尾会把根 钣金件.csv 覆写成"非焊接过滤版"(供步骤9总合并),
    # 因此本步骤自身不可重跑 —— 再跑一次时读到的是过滤后的表(只剩非焊接件), 会认为所有焊接组
    # 都没有钣金件, 于是静默写出 0 行组配料表、并把 DXF 全留在根目录(实测 0901 --from 8 重跑踩到)。
    # 对策: ①每次把全量配料表另存为 钣金件_全量.csv; ②取行数更多的那份作为来源;
    #       ③步骤2 重导 BOM 后会删除该备份, 保证新 BOM 永远优先。
    csv_full = os.path.join(ban_dir, "钣金件_全量.csv")
    if os.path.exists(csv_full):
        h2, r2 = read_csv_map(csv_full)
        if len(r2) > len(rows):
            log("[MERGE-BAN] 根配料表仅 %d 行 < 全量备份 %d 行 → 改用全量备份(根表已被上轮过滤版覆写)"
                % (len(rows), len(r2)))
            header, rows = h2, r2
    if not rows:
        raise RuntimeError("配料表为空, 拒绝执行(疑似已被过滤版覆写): " + csv_path)
    try:
        with open(csv_full, "w", encoding="utf-8-sig", newline="") as f:
            import csv as _csv
            w = _csv.writer(f)
            w.writerow(header)
            for r in rows:
                w.writerow([r.get(c, "") for c in header])
    except Exception as e:
        log("[WARN] 写全量备份失败: %s" % repr(e)[:60])
    all_codes = {r["代号"].strip() for r in rows if r.get("代号")}
    # 根目录 dxf 列表(排除合并产物/_dxf_tmp)
    dxf_list = [f for f in os.listdir(ban_dir)
                if f.lower().endswith(".dxf") and not f.startswith("合并图纸") and not f.startswith("_")]
    by_code = {}
    for f in dxf_list:
        base = os.path.splitext(f)[0]
        code = base.split("-")[0].split("_")[0]
        by_code.setdefault(code, []).append(f)
    total_moved = 0
    for gname, paths in sorted(data["groups"].items()):
        gdir = os.path.join(ban_dir, gname)
        os.makedirs(gdir, exist_ok=True)
        # 2026-10-17 修复(重要): 先清空本组目录里上一轮遗留的 DXF / 合并图纸 / _* 临时目录。
        # 旧版只"移入"不清理: 一旦分组发生变化(零件从A组改到B组、图号改名、件被删),
        # 旧文件就永久残留在组目录里 —— 组内会混入别组的件和旧图号件。
        # 组自身的合并图纸按配料csv过滤, 所以组合并图仍是对的(掩盖了问题);
        # 但下游 skill16 是按 *.dxf 全量归拢的, 会把同一件合并两次(实测 0901 大门板重复)。
        purged = 0
        for fn in os.listdir(gdir):
            p = os.path.join(gdir, fn)
            try:
                if os.path.isdir(p):
                    if fn.startswith("_"):
                        shutil.rmtree(p, ignore_errors=True)
                        purged += 1
                elif fn.lower().endswith(".dxf"):
                    os.remove(p)
                    purged += 1
            except Exception as e:
                log(f"[WARN] 清理失败 {fn}: {repr(e)[:60]}")
        if purged:
            log(f"[组] {gname}: 已清理上轮遗留 {purged} 个文件(旧DXF/旧合并图纸/_tmp)")
        gcodes = {_code_of(p) for p in paths}
        ban_codes = gcodes & all_codes
        moved = []
        for code in sorted(ban_codes):
            for f in by_code.get(code, []):
                try:
                    shutil.move(os.path.join(ban_dir, f), os.path.join(gdir, f))
                    moved.append(f)
                except Exception as e:
                    log(f"[WARN] 移动失败 {f}: {repr(e)[:60]}")
        # 组配料 csv
        g_rows = [r for r in rows if r.get("代号") and r["代号"].strip() in ban_codes]
        g_csv = os.path.join(gdir, os.path.basename(csv_path))
        with open(g_csv, "w", encoding="utf-8-sig", newline="") as f:
            import csv as _csv
            w = _csv.writer(f)
            w.writerow(header)
            for r in g_rows:
                w.writerow([r.get(c, "") for c in header])
        # 校验: 组内钣金件应有 DXF(缺件警告)
        missing = sorted(ban_codes - set(by_code.keys()))
        log(f"[组] {gname}: 钣金件 {len(ban_codes)} 个, DXF 移入 {len(moved)} 个, csv 行 {len(g_rows)} 行"
            + (f", 缺 DXF: {missing}" if missing else ""))
        if moved:
            run_dwg_merge(gdir)
        total_moved += len(moved)
    # 根目录 csv 写回非焊接过滤版, 供步骤9 总合并(避免“CSV有配料但无dxf”警告噪音)
    un_codes = {_code_of(p) for p in data["ungrouped"]} & all_codes
    un_rows = [r for r in rows if r.get("代号") and r["代号"].strip() in un_codes]
    with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
        import csv as _csv
        w = _csv.writer(f)
        w.writerow(header)
        for r in un_rows:
            w.writerow([r.get(c, "") for c in header])
    log(f"[MERGE-BAN] 根目录钣金件.csv 已写回非焊接过滤版({len(un_rows)} 行)")
    log(f"[MERGE-BAN] 完成: 焊接组 {len(data['groups'])} 个, 共移入 {total_moved} 个 DXF; 根目录剩余 dxf 供总合并")


def cmd_prep_profile(args):
    data = load_groups(args.dir)
    folder = os.path.abspath(args.dir)
    pro_dir = os.path.join(folder, "型材")
    csv_path = os.path.join(pro_dir, "型材.csv")
    if not os.path.isdir(pro_dir):
        raise RuntimeError("缺少型材目录: " + pro_dir)
    if not os.path.exists(csv_path):
        raise RuntimeError("缺少型材.csv: " + csv_path)
    header, rows = read_csv_map(csv_path)
    all_codes = {r["代号"].strip() for r in rows if r.get("代号")}
    tmp_root = os.path.join(pro_dir, "_weld_tmp")
    os.makedirs(tmp_root, exist_ok=True)
    entries = [(gname, paths) for gname, paths in sorted(data["groups"].items())] + [("__ungrouped__", data["ungrouped"])]
    made = 0
    for gname, paths in entries:
        gcodes = {_code_of(p) for p in paths}
        prof_codes = sorted(gcodes & all_codes)
        if not prof_codes:
            log(f"[组] {gname}: 无型材件, 跳过")
            continue
        gdir = os.path.join(tmp_root, gname)
        os.makedirs(gdir, exist_ok=True)
        copied = 0
        for p in paths:
            if _code_of(p) in prof_codes and os.path.exists(p):
                try:
                    shutil.copy2(p, os.path.join(gdir, os.path.basename(p)))
                    copied += 1
                except Exception:
                    pass
        # 过滤 csv: 保留全列(宏只读前6列: 代号,数量,宽,高,厚,类型)
        g_rows = [r for r in rows if r.get("代号") and r["代号"].strip() in prof_codes]
        g_csv = os.path.join(gdir, "型材1.CSV")
        with open(g_csv, "w", encoding="utf-8-sig", newline="") as f:
            import csv as _csv
            w = _csv.writer(f)
            w.writerow(header)
            for r in g_rows:
                w.writerow([r.get(c, "") for c in header])
        missing = sorted(set(prof_codes) - {_code_of(p) for p in paths})
        log(f"[组] {gname}: 型材件 {len(prof_codes)} 个, 复制 {copied}, csv 行 {len(g_rows)}"
            + (f", 源文件缺失: {missing}" if missing else ""))
        made += 1
    log(f"[PREP-PROFILE] 完成: 建 {made} 个组工作目录 @ {tmp_root}")


def cmd_collect_profile(args):
    data = load_groups(args.dir)
    folder = os.path.abspath(args.dir)
    pro_dir = os.path.join(folder, "型材")
    tmp_root = os.path.join(pro_dir, "_weld_tmp")
    dst_root = os.path.join(pro_dir, "带数量零件")
    if not os.path.isdir(tmp_root):
        log("[COLLECT-PROFILE] 无 _weld_tmp, 无事可做")
        return
    total = 0
    for gname in list(data["groups"].keys()) + ["__ungrouped__"]:
        src = os.path.join(tmp_root, gname, "带数量零件")
        if not os.path.isdir(src):
            log(f"[组] {gname}: 无带数量零件产物, 跳过")
            continue
        if gname == "__ungrouped__":
            os.makedirs(dst_root, exist_ok=True)
            for item in os.listdir(src):
                shutil.move(os.path.join(src, item), os.path.join(dst_root, item))
            total += 1
            log(f"[组] 非焊接: 带数量零件 -> {dst_root}")
        else:
            dst = os.path.join(dst_root, gname)
            shutil.move(src, dst)
            total += 1
            log(f"[组] {gname}: 带数量零件 -> {dst}")
    shutil.rmtree(tmp_root, ignore_errors=True)
    log(f"[COLLECT-PROFILE] 完成: 汇总 {total} 组, 已删 {tmp_root}")


def main():
    ap = argparse.ArgumentParser(description="skill7 焊接装配体分组输出")
    sub = ap.add_subparsers(dest="cmd", required=True)
    p1 = sub.add_parser("collect", help="SW 收集焊接分组 -> 焊接分组.json")
    p1.add_argument("--dir", required=True)
    p1.add_argument("--asm", default="")
    p1.set_defaults(fn=cmd_collect)
    p2 = sub.add_parser("merge-ban", help="钣金分组 + 各焊接组独立合并")
    p2.add_argument("--dir", required=True)
    p2.set_defaults(fn=cmd_merge_ban)
    p3 = sub.add_parser("prep-profile", help="型材分组工作目录准备")
    p3.add_argument("--dir", required=True)
    p3.set_defaults(fn=cmd_prep_profile)
    p4 = sub.add_parser("collect-profile", help="带数量零件汇总 + 清理临时目录")
    p4.add_argument("--dir", required=True)
    p4.set_defaults(fn=cmd_collect_profile)
    args = ap.parse_args()
    args.fn(args)


if __name__ == "__main__":
    main()