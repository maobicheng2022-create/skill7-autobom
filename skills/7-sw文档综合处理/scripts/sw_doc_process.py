# -*- coding: utf-8 -*-
"""sw文档综合处理 orchestrator (v4).
必填: 文件夹路径 folder + 套数 qty。
串行: 图号名称更新 -> 零件分类宏 -> 钣金/型材/型材采购 BOM(×qty) -> 总外购件/标准件带截图BOM(×qty)
      -> 钣金展开 DXF -> 合并 DXF -> 型材合成宏。
数量 qty 只乘各 BOM 表格的 $总数量; 合并DXF/型材宏复用已乘过的 csv。缺 folder/qty 即拒绝运行。
各步骤均计时统计, 末尾输出耗时表(便于提效优化)。
"""
import argparse
import os
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


BASE = os.path.dirname(os.path.abspath(__file__))

# 依赖 skill 脚本根 — 统一指向本仓库 skills/（由 skill7_paths 解析，无需维护多份副本）
S1 = os.path.normpath(os.path.join(GLOBAL_SKILLS, "1-图号名称更新", "scripts"))
S3 = os.path.normpath(os.path.join(GLOBAL_SKILLS, "3-导出BOM", "scripts"))
S4 = os.path.normpath(os.path.join(GLOBAL_SKILLS, "4-导出BOM带截图", "scripts"))
S5 = os.path.normpath(os.path.join(GLOBAL_SKILLS, "5-运行SW宏", "scripts"))
S6 = os.path.normpath(os.path.join(GLOBAL_SKILLS, "6-合并DXF", "scripts"))
S8 = os.path.normpath(os.path.join(GLOBAL_SKILLS, "8-焊接装配标志分类输出BOM带截图", "scripts"))
S9 = os.path.normpath(os.path.join(GLOBAL_SKILLS, "9-整理型材", "scripts"))
S10 = os.path.normpath(os.path.join(GLOBAL_SKILLS, "10-虚拟件外部化与Pipe标记", "scripts"))


PY = sys.executable or "python"

# ---- 内存保护(2026-08-07): skill7 打开大装配体加载全部引用件后 SW 内存不还给 OS,
#      曾导致系统内存 96% 爆掉。启动前检查 + 跑完自动退 SW 释放内存 ----
MIN_FREE_RAM_PCT = float(os.environ.get("SKILL7_MIN_FREE_RAM_PCT", "15") or 15)   # 启动门槛: 可用内存低于此百分比则拒绝启动(环境变量/--min-mem 可调)
FREE_SW_ON_EXIT = True    # 全流程跑完自动退出 SW 释放内存(可 --keep-sw 关闭)


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


def check_mem_before_start(required=MIN_FREE_RAM_PCT):
    """启动前内存闸: 可用内存过低则拒绝启动, 避免跑到一半系统卡死/失败。"""
    pct = _free_ram_pct()
    log(f"[MEM] 启动前剩余内存: {pct:.0f}% (门槛 {required:.0f}%)")
    if pct < required:
        raise SystemExit(
            f"错误: 系统可用内存过低 ({pct:.0f}% < {required:.0f}%). "
            f"Skill7 需打开大装配体(可能加载上百引用件占数GB内存). "
            f"请先关闭占内存程序(尤其 SolidWorks/浏览器)后重试.")


def release_sw_memory(force_exit=False):
    """释放 SW 内存: 关闭全部文档 + 退出 SW 进程(默认), 因为 skill7 跑完后 SW 已用不到。
    force_exit=False 时只关文档不退出; True 时 ExitApp 释放 SW 全部分配内存。"""
    try:
        import win32com.client
        sw = win32com.client.GetActiveObject("SldWorks.Application")
    except Exception as e:
        log(f"[MEM] 无可连接的 SW(无需释放): {repr(e)[:40]}")
        return
    # 关闭全部打开的文档(逐层, 避免残留锁定)
    try:
        for _ in range(50):
            n = int(sw.GetDocumentCount())
            if n <= 0:
                break
            try:
                d = sw.ActiveDoc
                if d is not None:
                    title = str(d.GetTitle())
                    sw.CloseDoc(title)
                else:
                    break
            except Exception:
                break
    except Exception as e:
        log(f"[WARN] 关闭SW文档失败({repr(e)[:40]})")
    if force_exit:
        try:
            sw.ExitApp()
            log("[MEM] 已请求退出 SW (释放大装配体占用的内存)")
            # 等待 SW 进程退出释放内存
            import subprocess as _sp
            _sp.run(["powershell", "-NoProfile", "-Command",
                     "Start-Sleep -Seconds 5; "
                     "if (Get-Process sldworks -ErrorAction SilentlyContinue) { "
                     "Stop-Process -Name sldworks -Force -ErrorAction SilentlyContinue }"],
                    capture_output=True)
        except Exception as e:
            log(f"[WARN] 退出SW失败({repr(e)[:40]})")

# 型材宏实际定位(2026-08-05 实测): 模块是 AutoBOM_型材批量处理, 方法 main
PROFILE_MODULE = "AutoBOM_型材批量处理"
PROFILE_METHOD = "main"

# ---- 步骤耗时统计
_timings = []

# ---- 步骤执行控制 (2026-08-30 断点续跑: --from/--to/--skip) ----
# 中断后不必整条重跑: 如中断在8.8, 用 --from 9 从步骤9续跑(前面产物已全)。
_STEP_ORDER = {
    "A": 1, "0": 2, "1.5": 3, "1.6": 4, "2": 5, "2.1": 6, "2.2": 6.5, "2.5": 7,
    "5": 8, "5.5": 9, "6": 10, "6.6": 11, "6.7": 12, "6.9": 13, "7": 14,
    "7.5": 15, "8.0": 16, "8": 17, "8.8": 18, "9": 19, "9.5": 20,
    "10.0": 21, "10": 22, "10.9": 23,
}
_run_from = None
_run_to = None
_run_skip = set()


def _step_order(n):
    return _STEP_ORDER.get(str(n), 999)


def _step_enabled(n):
    """按 --from/--to/--skip 判断步骤是否执行。"""
    key = str(n)
    if _run_from is not None and _step_order(n) < _step_order(_run_from):
        return False
    if _run_to is not None and _step_order(n) > _step_order(_run_to):
        return False
    if key in _run_skip:
        return False
    return True


def log(msg):
    print(msg, flush=True)


# ---- 前置: 弹窗守卫 + SW 正常方式启动(2026-08-06 三大根因之一/二) ----

_guard_proc = None


def start_popup_guard():
    """后台常驻 SW 弹窗守卫, 自动点掉 PDM未检出/另存为/完成提示, 避免 exec 无人点卡死。"""
    global _guard_proc
    if not os.path.exists(POPUP_GUARD):
        log("[WARN] 未找到 sw_popup_guard.py, 跳过弹窗守卫(PDM库图纸可能有弹窗风险)")
        return
    log(f"[GUARD] 启动弹窗守卫: {POPUP_GUARD}")
    # 2026-10-03: 传父进程 PID 给守卫 → 本进程退出(含被强杀/崩溃)后守卫自动退出, 防残留干扰人工用 SW
    _guard_proc = subprocess.Popen(
        [PY, "-X", "utf8", POPUP_GUARD, "1.0", str(os.getpid())],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def stop_popup_guard():
    """停止所有 sw_popup_guard.py 弹窗守卫进程(含历史孤儿)。
    2026-08-08修复: 之前只 terminate 本次 _guard_proc; 若 skill7 中途异常退出或
    之前几轮遗留孤儿, 守卫会残留成死循环进程, 一直自动点 SW 弹窗/关窗, 干扰皇上正常用 SW。
    故这里改成全量清理: 按命令行匹配所有 sw_popup_guard.py 并强杀(taskkill /F), 不留任何残留。"""
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
    #    ——按进程命令行精确定位 sw_popup_guard.py, 逐个强杀；绝不碰其它 python 进程。
    killed = 0
    try:
        def _cmdline_pids(keyword):
            """返回命令行含 keyword 的进程 PID 列表(仅 python.exe)。"""
            pids = []
            try:
                ps = (
                    "Get-CimInstance Win32_Process | "
                    "Where-Object { $_.CommandLine -like '*" + keyword + "*' -and $_.ProcessId -ne $PID } | "
                    "Select-Object -ExpandProperty ProcessId"
                )
                out = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
                                     capture_output=True, timeout=30)  # 字节模式, 避免中文系统GBK解码崩
                raw = (out.stdout or b"")
                # 输出只有纯数字PID, 安全按 ascii 容忍解码
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
        log(f"[WARN] 清理残留守卫失败: {repr(e)[:60]}")
    if killed:
        log(f"[GUARD] 已额外清理 {killed} 个残留弹窗守卫进程")
    log("[GUARD] 弹窗守卫全部停止")



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
    """确保 SW 以正常方式启动且可被 exec 后台进程 GetActiveObject 连上。
    根因(2026-08-06): exec 后台进程 COM 上下文读不到桌面交互式 SW 的 ROT,
    SW 须以正常方式启动(非 -Embedding)才注册进桌面 ROT。若连不上则用 Start-Process 拉起并等注册。"""
    import win32com.client
    try:
        win32com.client.GetActiveObject("SldWorks.Application")
        log("[SW] 已连上运行中的 SW (GetActiveObject 成功)")
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
    # 轮询等待 SW 注册进 ROT(最长120s)
    for i in range(120):
        time.sleep(1)
        try:
            win32com.client.GetActiveObject("SldWorks.Application")
            log(f"[SW] SW 已正常启动并可连接(等了 {i+1}s)")
            time.sleep(3)  # 多等几秒让 COM 稳定
            return
        except Exception:
            continue
    raise RuntimeError("SW 未能在120s内以正常方式启动并连上, 请手动打开 SolidWorks 后重试")


def run(cmd, cwd=None, check=True):
    log("[EXEC] " + " ".join(cmd))
    r = subprocess.run(cmd, cwd=cwd, text=True, encoding="utf-8", errors="replace")
    if check and r.returncode != 0:
        raise RuntimeError(f"步骤失败 (rc={r.returncode}): {' '.join(cmd)}")
    return r


def py_script(script, args):
    return run([PY, "-X", "utf8", script] + args)


def ensure_dir(d):
    os.makedirs(d, exist_ok=True)


def close_asm(asm_path):
    """关闭总装配体(用标题), 释放它锁定的全部关联零件, 否则分类副本打不开(65536)。"""
    if not asm_path:
        return
    title = os.path.splitext(os.path.basename(asm_path))[0]
    try:
        import win32com.client
        try:
            sw = win32com.client.GetActiveObject("SldWorks.Application")
        except Exception:
            sw = win32com.client.Dispatch("SldWorks.Application")  # exec 场景兼容
        sw.CloseDoc(title)
        log(f"[CLOSE-ASM] 已关闭总装: {title}")
    except Exception as e:
        log(f"[WARN] 关闭总装失败({repr(e)[:60]}), 继续")
    finally:
        time.sleep(1)


def restart_sw():
    """重启 SW 释放内存(2026-08-11 内存优化)。
    背景: SW CloseDoc 后内存不还给 OS, 连续处理几百文件后 SW 内存 6-10GB,
    16GB 机器+用户程序 → 可用<10% → SW Not Responding 假死(20260811 实战 3 次)。
    关键步骤间调用, 从干净 SW 继续。"""
    import subprocess as _sp
    import win32com.client
    try:
        sw = win32com.client.GetActiveObject("SldWorks.Application")
        try:
            sw.ExitApp()
            log("[RESTART] SW 已退出")
        except Exception:
            pass
    except Exception:
        pass
    time.sleep(5)
    try:
        _sp.run(["powershell", "-NoProfile", "-Command",
                 f"Start-Process '{SLDWORKS_EXE}'"], capture_output=True, text=True, timeout=30)
    except Exception as e:
        log(f"[WARN] 重启 SW 失败({repr(e)[:60]}), 继续")
        return
    log("[RESTART] SW 重启中...")
    time.sleep(8)
    try:
        ensure_sw_normal_start()
    except Exception as e:
        log(f"[WARN] 等待新 SW 失败({repr(e)[:60]})")
    log("[RESTART] SW 就绪(内存已释放)")


def step(n, title, fn):
    if not _step_enabled(n):
        log(f"--- 步骤{n} {title} 已跳过(--from/--to/--skip) ---")
        return
    t0 = time.time()
    log(f"\n===== 步骤{n} {title} ====")
    fn()
    dt = time.time() - t0
    _timings.append((n, title, dt))
    log(f"---- 步骤{n} {title} 完成, 耗时 {dt:.1f}s ----")


def main():
    ap = argparse.ArgumentParser(description="sw文档综合处理 orchestrator (强制: folder + qty)")
    ap.add_argument("folder", help="SW 图纸所在文件夹(必填)")
    ap.add_argument("--qty", type=int, default=1, help="这套图纸的数量/套数(必填语义, 默认1); 各BOM $总数量会×此值")
    ap.add_argument("--asm", default="", help="总装配体 .sldasm 路径(缺省自动找/用 SW 当前激活装配体)")
    ap.add_argument("--no-asm-close", action="store_true", help="跳过关闭装配体步骤(调试用)")
    ap.add_argument("--keep-sw", action="store_true", help="跑完不自动退出 SW(默认跑完退出以释放内存)")
    ap.add_argument("--from", dest="run_from", default="", help="断点续跑: 从某步骤号开始执行(如 --from 9 从步骤9续跑, 跳过前面已完成步骤)")
    ap.add_argument("--to", dest="run_to", default="", help="只执行到某步骤号为止(如 --from 8 --to 10 只跑8/8.8/9/10)")
    ap.add_argument("--skip", default="", help="跳过步骤号列表(逗号分隔, 如 --skip 1.6,2.1)")
    ap.add_argument("--min-mem", type=float, default=None,
                    help="启动前最低可用内存百分比(默认 15, 也可用环境变量 SKILL7_MIN_FREE_RAM_PCT); 小装配内存紧张时可下调")
    args = ap.parse_args()

    global _run_from, _run_to, _run_skip
    if args.run_from:
        _run_from = str(args.run_from)
    if args.run_to:
        _run_to = str(args.run_to)
    if args.skip:
        _run_skip = set(x.strip() for x in args.skip.split(",") if x.strip())
        log(f"[STEP-CTRL] 跳过步骤: {sorted(_run_skip)}")
    if _run_from or _run_to:
        log(f"[STEP-CTRL] 断点续跑: from={_run_from} to={_run_to} (已完成的步骤产物需存在, 否则对应脚本会自行校验)")

    t_total0 = time.time()

    # ===== 强制前置校验: 文件夹 + 数量 =====
    if not args.folder:
        raise SystemExit("错误: 必须提供文件夹路径 folder (SW 图纸位置)")
    folder = os.path.abspath(args.folder)
    if not os.path.isdir(folder):
        raise SystemExit(f"错误: 文件夹不存在: {folder}")
    if args.qty < 1:
        raise SystemExit("错误: --qty 数量必须 >= 1")
    if not os.path.exists(POPUP_GUARD):
        raise SystemExit(f"错误: 缺依赖 弹窗守卫 {POPUP_GUARD} (turing-workspace/scripts/sw_popup_guard.py)")

    # ★内存闸(2026-08-07): 可用内存过低直接拒绝, 避免跑一半系统卡死成上次 96%
    check_mem_before_start(args.min_mem if args.min_mem is not None else MIN_FREE_RAM_PCT)

    qty_args = ["--qty", str(args.qty)]

    # 找总装配体(省略则用 SW 当前激活装配体)
    # 2026-08-11: 排除 Pipe_ 前缀(管道外购装配, 会被误当总装) + ST_ + ~$ 残留
    # 2026-11-23: 改为「文件夹名优先」——先去掉末尾 _做N套，再找同名/同尾的 .sldasm；
    #             原先按"首个下划线前的字符"算代号，遇到 100-DX.xxx / 100-ST_xxx 这类
    #             统一前缀会全部平局，误把 ST_ 螺钉装配当总装。ST_/Pipe_ 改为任意位置包含即排除。
    asm = args.asm
    if not asm:
        names = [f for f in os.listdir(folder)
                 if f.lower().endswith(".sldasm") and not f.startswith("~")]
        base = os.path.basename(os.path.normpath(folder))
        if "_做" in base:
            base = base.split("_做")[0]
        stems = {n: os.path.splitext(n)[0] for n in names}
        exact = [n for n in names if stems[n] == base]
        tail = [n for n in names if stems[n].endswith(base)]
        if exact:
            asm = os.path.join(folder, exact[0])
        elif len(tail) == 1:
            asm = os.path.join(folder, tail[0])
        else:
            cands = [os.path.join(folder, n) for n in names
                     if "st_" not in n.lower() and "pipe_" not in n.lower()]
            if cands:
                asm = sorted(cands, key=lambda p: len(os.path.splitext(os.path.basename(p))[0].split("_")[0]))[0]
        if asm:
            log(f"[ASM] 自动识别总装: {os.path.basename(asm)}")
    asm_args = ["--asm", asm] if asm else []

    # 分类宏实际建的子目录名是 钣金件 / 型材 (2026-08-05 实测)
    ban_folder = os.path.join(folder, "钣金件")
    xing_folder = os.path.join(folder, "型材")

    # ===== 前置(2026-08-06 三大根因): 弹窗守卫 + SW 正常方式启动 =====
    start_popup_guard()
    try:
        ensure_sw_normal_start()
    except Exception as e:
        stop_popup_guard()
        raise

    # ===== 步骤A(2026-08-11 钦定首位): 虚拟零部件外部化 + Pipe_/ST_ 外购件标记 =====
    # 虚拟件会导致: BOM聚合失败 / $宽$高读不到 / 截图打不开; Pipe_ 装配会干扰找总装。
    # 先外部化(生成磁盘文件)再标记, 再进图号更新。
    step("A", "虚拟件外部化+Pipe标记", lambda: py_script(
        os.path.join(S10, "sw_virtual_extract_and_mark.py"),
        ["--dir", folder] + (["--asm", asm] if asm else []) + ["--no-guard"]))

    # ===== 步骤0 图号+分类 一次遍历合并版(2026-08-13 皇上钦定) =====
    # 每个文件只打开一次: ①按文件名写「代号/名称」 ②读/判「采购类别」(空则按规则写)
    # ③有改动才保存 ④按类别复制到分类文件夹。不分固定批次, 内存<15% 自动重启 SW 继续。
    step(0, "图号+分类(一次遍历)", lambda: py_script(
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "sw_tuhao_categorize.py"),
        ["--dir", folder] + asm_args))
    ensure_dir(ban_folder)
    ensure_dir(xing_folder)

    # ===== 步骤1.5 (2026-08-12 钦定): 刷新截图库记录的文件统计 =====
    # 图号/分类步骤只改属性不改变几何, 截图内容不变; 刷新 size/mtime 防截图缓存批量失效重截
    step("1.5", "刷新截图库(防缓存失效)", lambda: py_script(
        os.path.join(S4, "sw_preshot.py"), ["--dir", folder, "--refresh"]))

    # ===== 步骤1.6 (2026-08-12 钦定): 预生成全量等轴测截图 =====
    # 从总装全树收集唯一(路径,配置), 命中缓存跳过, 未命中截图入库(15张/内存<12%自动重启SW)。
    # 首次建库耗时较长, 之后所有步骤(5/6/6.6/6.7)截图全命中, 零 SW 截图。
    step("1.6", "预生成全量截图", lambda: py_script(
        os.path.join(S4, "sw_preshot.py"), ["--dir", folder] + asm_args))

    # ===== 步骤2 多模板BOM一次扫装配体(钣金件.csv + 型材.csv + 型材采购.xlsx, ×qty) =====
    step(2, "BOM三表一次导出", lambda: py_script(
        os.path.join(S3, "autobom_export.py"),
        asm_args + qty_args + ["--templates-multi",
            ";".join([
                "钣金:钣金件" + ":" + os.path.join(ban_folder, "钣金件.csv"),
                "型材:型材" + ":" + os.path.join(xing_folder, "型材.csv"),
                "型材采购:型材" + ":" + os.path.join(folder, "型材采购.xlsx"),
            ])]))

    # 2026-10-17: 步骤8.8 用 钣金件_全量.csv 作为"根配料表被过滤版覆写"时的回退来源。
    # 本步骤刚按最新 BOM 重导了 钣金件.csv, 必须让新表优先 → 删除上一轮的全量备份。
    def _invalidate_ban_full():
        p = os.path.join(ban_folder, "钣金件_全量.csv")
        if os.path.exists(p):
            try:
                os.remove(p)
                log("[步骤2] 已删除上轮 钣金件_全量.csv 备份(本轮 BOM 为准)")
            except Exception as e:
                log("[步骤2] 删除全量备份失败: %s" % repr(e)[:60])
    if _step_enabled(2):
        _invalidate_ban_full()

    # ===== 步骤2.1 (2026-08-11 新增): 型材整理 = 读型材采购.xlsx 按规格归并,
    # 折6m管条数 + 每根6m最多开几整套/每根余料(按 --qty 套数算每套用量) =====
    step("2.1", "型材整理(6m折算+每根可开套数)", lambda: py_script(
        os.path.join(S9, "xingcai_summary.py"),
        [os.path.join(folder, "型材采购.xlsx"), "--sets", str(args.qty)]))

    # ===== 步骤2.2 (2026-11-16 皇上钦定): 型材采购表按零件加列 =====
    # 每个零件行：每套支数 / 一根6m可开零件数 / 一根6m可开套数(小数) / 余料mm
    step("2.2", "型材采购加列(每零件可开套数)", lambda: py_script(
        os.path.join(S9, "xingcai_purchase_enrich.py"),
        [os.path.join(folder, "型材采购.xlsx"), "--sets", str(args.qty)]))

    # 2026-08-11 内存优化: BOM 打开总装聚合后 SW 内存高, 截图前重启
    step("2.5", "重启SW(供截图)", restart_sw)

    # ===== 步骤5 总外购件带截图, ×qty =====
    step(5, "总外购件 BOM+截图", lambda: py_script(
        os.path.join(S4, "sw_buyout_export.py"),
        ["--dir", folder, "-o", folder, "--template", "总外购件",
         "--filter-value", "外购件"] + qty_args))

    # 2026-08-11 内存优化: 96 张截图后 SW 内存 6-7GB, 标准件聚合打开总装必失败, 重启
    step("5.5", "重启SW(供标准件)", restart_sw)

    # ===== 步骤6 标准件采购带截图, ×qty =====
    step(6, "标准件采购 BOM+截图", lambda: py_script(
        os.path.join(S4, "sw_buyout_export.py"),
        ["--dir", folder, "-o", folder, "--template", "标准件采购",
         "--filter-value", "标准件"] + qty_args))

    # ===== 步骤6.5/6.6/6.7 (2026-08-11 新增): 焊接部件清单 + 装配部件清单(skill8, 含步路明细) =====
    # skill8 内嵌 SW 启动/弹窗守卫, 跑完自动退 SW; 6.6 与 6.7 各自会拉起新 SW
    step("6.6", "焊接部件清单(skill8)", lambda: py_script(
        os.path.join(S8, "sw_weld_export.py"),
        ["--dir", folder, "-o", folder, "--template", "焊接部件清点清单",
         "--filter-value", "焊接", "--cache-dir", SHOT_CACHE] + qty_args + asm_args))
    step("6.7", "装配部件清单(skill8)", lambda: py_script(
        os.path.join(S8, "sw_weld_export.py"),
        ["--dir", folder, "-o", folder, "--template", "焊接部件清点清单",
         "--filter-value", "装配", "--cache-dir", SHOT_CACHE] + qty_args + asm_args))

    # ===== 步骤7 关闭装配体(关键2): 步骤2-6 BOM 已重新打开装配体并结束,
    #     关闭它以释放锁定的零件, 否则钣金展开/合并读分类副本失败 =====
    # 2026-08-11: skill8(6.6/6.7) 跑完已退出 SW, 先重启 SW 再关总装
    step("6.9", "重启SW(供关总装)", restart_sw)
    if not getattr(args, "no_asm_close", False):
        step(7, "关闭装配体(供展开)", lambda: close_asm(asm))

    # 2026-08-11 内存优化: 截图后 SW 内存高, 展开前重启
    step("7.5", "重启SW(供展开)", restart_sw)

    # ===== 步骤8.0 (2026-08-29 新增): 焊接装配体分组收集 =====
    # 开总装递归, 按「焊接或装配标记=焊接」分组 -> 焊接分组.json
    # 归属规则: 焊接装配体->独立组; 非焊接->非焊接组(总组); 嵌套焊接各自成组(零件归最近焊接祖先)
    step("8.0", "焊接分组收集", lambda: py_script(
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "sw_weld_group.py"),
        ["collect", "--dir", folder] + asm_args))

    # ===== 步骤8 钣金展开 =====
    # AutoBaiJinZhanKai.exe 返回码不可靠(统计bug: 成功0但实际生成9个), 用产物校验替代退出码判断
    _has_ban = {}  # 记录是否有钣金件, 供 8.8/9 决定是否跳过
    def _do_zhankai():
        # 2026-10-12 修复: 无钣金件时(钣金件.csv 空/目录无 .SLDPRT) AutoBaiJinZhanKai 产出 0 个 dxf,
        # 之前无条件 raise 导致整个 skill7 中止。应跳过展开(无件可展), 不报错。
        parts = [f for f in os.listdir(ban_folder)
                 if f.lower().endswith(".sldprt")] if os.path.isdir(ban_folder) else []
        if not parts:
            log(f"[展开] 无钣金零件({len(parts)} 个 .SLDPRT), 跳过钣金展开")
            _has_ban["ok"] = False
            return
        _has_ban["ok"] = True
        run([ZHANKAI_EXE, "批量", ban_folder, "dxf"], check=False)
        time.sleep(2)
        dxf_list = [f for f in os.listdir(ban_folder) if f.lower().endswith(".dxf")]
        # 合并 DXF 会在本目录, 剔除合并产物再统计单个展开件
        real = [f for f in dxf_list if not f.startswith("合并图纸")]
        log(f"[展开] 生成 DXF {len(real)} 个")
        if not real:
            raise RuntimeError(f"钣金展开未产出任何 dxf: {ban_folder}")
    step(8, "钣金展开 DXF", _do_zhankai)

    # ===== 步骤8.8 (2026-08-29 新增): 钣金分组 + 各焊接组独立合并图纸 =====
    # 读 焊接分组.json: 每焊接组建 钣金件\<组名>\, 移入该组 DXF + 写过滤配料csv + 跑 dwg_merge
    # 根目录剩非焊接 DXF + 非焊接过滤csv, 供步骤9 总合并
    # 2026-10-12: 无钣金件时跳过(否则空目录合并报错)
    if _has_ban.get("ok", False):
        step("8.8", "钣金分组+分组合并", lambda: py_script(
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "sw_weld_group.py"),
            ["merge-ban", "--dir", folder]))

    # ===== 步骤9 合并 DXF(配料表=钣金件.csv, 数量已含qty) =====
    if _has_ban.get("ok", False):
        step(9, "合并 DXF", lambda: py_script(
            os.path.join(S6, "dwg_merge.py"), [ban_folder]))

    # ===== 步骤10 型材合成宏(读型材.csv, 数量已含qty) =====
    # 2026-08-11 内存优化: 展开/合并后重启 SW 再跑型材宏
    step("9.5", "重启SW(供型材宏)", restart_sw)

    # ===== 步骤10.0 (2026-08-29 新增): 型材分组准备 =====
    # 每组建 型材\_weld_tmp\<组名>\ (复制组型材件 + 写过滤 型材1.CSV), 供逐组跑宏
    step("10.0", "型材分组准备", lambda: py_script(
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "sw_weld_group.py"),
        ["prep-profile", "--dir", folder]))

    # ===== 步骤10 型材合成宏(2026-08-29 改: 逐组跑, 读各组 _weld_tmp\<组名>\ 的型材1.CSV) =====
    def _do_profile_groups():
        tmp_root = os.path.join(xing_folder, "_weld_tmp")
        if not os.path.isdir(tmp_root):
            log("[型材宏] 无 _weld_tmp 目录, 跳过型材合成宏(装配无型材件)")
            return
        gdirs = [os.path.join(tmp_root, d) for d in sorted(os.listdir(tmp_root))
                 if os.path.isdir(os.path.join(tmp_root, d))]
        if not gdirs:
            log("[型材宏] _weld_tmp 下无组目录, 跳过型材合成宏(装配无型材件)")
            return
        for gdir in gdirs:
            log(f"[型材宏] 处理组目录: {gdir}")
            py_script(os.path.join(S5, "sw_run_macro.py"),
                      ["--macro", MACRO_PROFILE, "--module", PROFILE_MODULE,
                       "--method", PROFILE_METHOD, "--dir", gdir])
    step(10, "型材合成宏(逐组)", _do_profile_groups)

    # ===== 步骤10.9 (2026-08-29 新增): 型材分组汇总 =====
    # 非焊接组 -> 型材\带数量零件\ (原位置); 焊接组 -> 型材\带数量零件\<组名>\; 删 _weld_tmp
    step("10.9", "型材分组汇总", lambda: py_script(
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "sw_weld_group.py"),
        ["collect-profile", "--dir", folder]))

    total = time.time() - t_total0
    log("\n===== 步骤耗时统计 (共 %.1fs) ====" % total)
    for n, t, dt in _timings:
        pct = dt / total * 100 if total else 0
        log(f"  步骤{n:<2} {t:<14} {dt:7.1f}s  {pct:4.1f}%")
    log("\n===== 全部完成 =====")
    stop_popup_guard()

    # ★内存保护(2026-08-07): 跑完自动退出 SW 释放大装配体占的内存(默认),
    #   除非 --keep-sw 显式要求保留 SW 供人工继续操作。
    if not getattr(args, "keep_sw", False) and FREE_SW_ON_EXIT:
        release_sw_memory(force_exit=True)
        log("[MEM] SW 已退出, 大装配体占用的内存已释放")


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        # 守卫清理兜底: 任何退出路径(含校验失败/步骤抛错/人工中断)都确保不留 popup 守卫残留
        try:
            stop_popup_guard()
        except Exception:
            pass
        raise
    except Exception:
        try:
            stop_popup_guard()
        except Exception:
            pass
        raise
