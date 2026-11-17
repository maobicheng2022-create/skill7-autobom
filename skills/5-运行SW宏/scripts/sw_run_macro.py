# -*- coding: utf-8 -*-
"""
sw_run_macro.py — 驱动 SolidWorks 运行 .swp VBA 宏，并自动清理宏的 MsgBox 弹窗。

已实测验证（2026-08-02）：
- SolidWorks RunMacro2 签名 = RunMacro2(FilePathName, ModuleName, ProcedureName, Options, Error)
  5 参数。误传 4 参 (path, progid, method, arg) 会报"类型不匹配"。
- 宏第一步检查 ActiveDoc，SW 必须先打开一个 sldprt 零件才能运行宏。
- 宏内 MsgBox（如"批量另存完成"）是 SW 模态弹窗，会让 RunMacro2 阻塞不返回。
  关键：点弹窗必须在【后台线程】与 RunMacro2 并行跑，不能用顺序调用（RunMacro2 会
  一直阻塞直到弹窗被点，顺序代码永远执行不到点弹窗的逻辑）。

用法:
  python sw_run_macro.py --macro <宏.swp> --dir <零件目录> [--module Macro51] [--method main]

依赖: pywin32（win32com / win32gui）。SW 需已开启，Python 需 64 位。
"""

import argparse
import ctypes
import ctypes.wintypes
import glob
import os
import sys
import threading
import time

import win32com.client
import win32con
import win32gui

DEFAULT_MODULE = "Macro51"
DEFAULT_METHOD = "main"


def log(msg):
    print(msg, flush=True)


def connect_sw():
    """连接 SW：优先复用现有实例(GetActiveObject)，失败则 Dispatch 兼容(OpenClaw exec 场景可连桌面实例)。"""
    try:
        return win32com.client.GetActiveObject("SldWorks.Application")
    except Exception:
        # exec 后台进程 COM 上下文读不到 ROT 时退回 Dispatch（实测可连到桌面 PID 12452）
        return win32com.client.Dispatch("SldWorks.Application")


def find_first_part(directory):
    """找目录下第一个 .sldprt 作为激活前提（宏要求 ActiveDoc 非空）。
    2026-08-30 修复: glob.glob 在含中文目录路径下匹配失效(返回空),
    导致 _weld_tmp + 中文组名 + 斜杠 + 里明明有 sldprt 却报 FileNotFoundError。
    改 os.listdir + endswith 对中文目录稳定。
    """
    try:
        names = sorted(os.listdir(directory))
    except OSError:
        names = []
    for n in names:
        low = n.lower()
        if low.endswith(".sldprt") or low.endswith(".sldasm"):
            return os.path.join(directory, n)
    raise FileNotFoundError(f"目录下未找到 .sldprt/.sldasm 文件: {directory}")


def _enum_cb(hwnd, results, wanted_kw):
    if win32gui.IsWindowVisible(hwnd):
        if win32gui.GetClassName(hwnd) == "#32770":
            title = win32gui.GetWindowText(hwnd)
            # 匹配任一关键词；wanted_kw 为空则匹配任何标题非空对话框
            if not wanted_kw or any(kw in title for kw in wanted_kw):
                results.append(hwnd)
    return True


def _find_msgbox(wanted_kw):
    results = []
    win32gui.EnumWindows(lambda h, l: _enum_cb(h, results, wanted_kw), None)
    return results


def _click_ok(dlg, user32):
    """在对话框 dlg 内找"确定"按钮并点击；无按钮则发 WM_CLOSE。返回是否处理了。"""
    btns = []
    win32gui.EnumChildWindows(
        dlg, lambda h, l: btns.append(h) if win32gui.GetClassName(h) == "Button" else True, None
    )
    for b in btns:
        txt = win32gui.GetWindowText(b)
        if "确" in txt or "确定" in txt or "OK" in txt.upper():
            user32.SendMessageW(b, 0x00F5, 0, 0)  # BM_CLICK
            return True
    try:
        win32gui.PostMessage(dlg, win32con.WM_CLOSE, 0, 0)
        return True
    except Exception:
        return False


def run_dismiss_forever(wanted_kw=None, stop_flag=None, poll=0.8):
    """后台线程：持续找 SW 的模态对话框点确定，直到 stop_flag 置位。
    必须在调用 RunMacro2 前启动本线程。返回已处理弹窗计数。
    wanted_kw: 匹配弹窗标题的关键词列表(任一词命中即点)；默认匹配 SOLIDWORKS/AutoBOM 等。
    重要: AutoBOM 型材宏的完成弹窗标题是 'AutoBOM 型材批量处理'(不含 SOLIDWORKS),
    故必须默认匹配 AutoBOM, 否则 RunMacro2 会扣住永不返回。
    """
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    user32.SendMessageW.argtypes = [ctypes.wintypes.HWND, ctypes.wintypes.UINT,
                                    ctypes.wintypes.WPARAM, ctypes.wintypes.LPARAM]
    user32.SendMessageW.restype = ctypes.wintypes.LPARAM
    user32.PostMessageW.argtypes = [ctypes.wintypes.HWND, ctypes.wintypes.UINT,
                                    ctypes.wintypes.WPARAM, ctypes.wintypes.LPARAM]
    user32.PostMessageW.restype = ctypes.wintypes.BOOL
    handled = 0
    while not stop_flag.is_set():
        dlgs = _find_msgbox(wanted_kw)
        for dlg in dlgs:
            if _click_ok(dlg, user32):
                handled += 1
                log(f"[UI] 已点掉弹窗 hwnd={dlg} (累计 {handled})")
        time.sleep(poll)
    return handled


def _count_sldprt_recursive(root):
    """递归统计目录下所有 .sldprt 数量(含嵌套子目录)。
    修复统计bug: 型材合成宏输出带数量零件/<规格>/*.sldprt 是两级结构,
    旧版只数直接子文件导致 带数量零件=0 误报。"""
    n = 0
    for dirpath, _dirnames, filenames in os.walk(root):
        n += len([f for f in filenames if f.lower().endswith(".sldprt")])
    return n


def collect_result(part_dir):
    """列出生成的分了类子文件夹及其中 sldprt 数量(递归统计, 防嵌套目录漏报)。"""
    subdirs = []
    for name in sorted(os.listdir(part_dir)):
        full = os.path.join(part_dir, name)
        if os.path.isdir(full):
            n = _count_sldprt_recursive(full)
            subdirs.append((name, n))
    return subdirs


def main():
    ap = argparse.ArgumentParser(description="驱动 SW 运行 .swp 宏 + 自动点弹窗")
    ap.add_argument("--macro", required=True, help="宏文件路径 .swp")
    ap.add_argument("--dir", required=True, help="要处理的零件目录")
    ap.add_argument("--module", default=DEFAULT_MODULE)
    ap.add_argument("--method", default=DEFAULT_METHOD)
    ap.add_argument("--timeout", type=int, default=900, help="宏执行超时秒数")
    ap.add_argument("--poll", type=float, default=0.8)
    ap.add_argument("--popup-extra", default="AutoBOM",
                    help="额外匹配的弹窗标题关键词(逗号分隔), 追加到默认关键词后")
    args = ap.parse_args()

    macro_path = os.path.abspath(args.macro)
    if not os.path.exists(macro_path):
        sys.exit(f"宏文件不存在: {macro_path}")
    part_dir = os.path.abspath(args.dir)
    if not os.path.isdir(part_dir):
        sys.exit(f"目录不存在: {part_dir}")

    sw = connect_sw()
    log(f"[SW] 已连接: {sw.RevisionNumber}")

    part = find_first_part(part_dir)
    log(f"[SW] 打开零件作为激活前提: {os.path.basename(part)}")
    md = sw.OpenDoc6(part, 1, 1, "", 0, 0)
    if not md:
        sys.exit(f"无法打开零件: {part}")

    # 关键：先启动后台弹窗清理线程，再调 RunMacro2（RunMacro2 会阻塞）
    # 默认关键词覆盖 SOLIDWORKS 和 AutoBOM（型材宏完成弹窗标题是 'AutoBOM 型材批量处理'）
    popup_kw = ["SOLIDWORKS", "AutoBOM"] + [k.strip() for k in args.popup_extra.split(",") if k.strip()]
    stop_flag = threading.Event()
    dismiss = threading.Thread(target=run_dismiss_forever, args=(popup_kw, stop_flag, args.poll), daemon=True)
    dismiss.start()

    log(f"[MACRO] 运行 {os.path.basename(macro_path)} (module={args.module}, method={args.method})")
    log("[MACRO] RunMacro2 会阻塞直到宏完成；后台线程负责点掉弹窗...")
    try:
        result = sw.RunMacro2(macro_path, args.module, args.method, 0, 0)
        log(f"[MACRO] RunMacro2 返回: {result}")
    finally:
        # 宏返回后再等一小段确保最后一个弹窗被点掉，然后停止后台线程
        time.sleep(3)
        stop_flag.set()
        dismiss.join(timeout=5)

    subdirs = collect_result(part_dir)
    log("\n=== 分类结果 ===")
    if subdirs:
        for name, n in subdirs:
            log(f"  [{name}] : {n} 个零件")
        log(f"共 {len(subdirs)} 个分类文件夹")
    else:
        log("未生成分类文件夹（检查零件是否都有'采购类别'属性）")


if __name__ == "__main__":
    main()
