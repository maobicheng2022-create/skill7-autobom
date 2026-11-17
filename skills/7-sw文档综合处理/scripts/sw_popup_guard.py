# -*- coding: utf-8 -*-
"""SW 自动化弹窗全面守卫：自动处理所有会阻塞 skill7 流程的模态弹窗。
类型 & 处理：
  1. 'SOLIDWORKS PDM' 未检出询问 -> 点'否'（不检出只读继续）
  2. '另存为' PDM 对话框 -> 直接 WM_CLOSE(点'取消'有 PDM 弹窗无效风险, 2026-08-30 四转确认)
  3. 'SOLIDWORKS' 完成提示 -> 点'确定'
  4. '覆盖提示！' (AutoBaiJinZhanKai 钣金展开批量) -> 点'是'(覆盖, 重新生成正确图号DXF)
  5. 任意含'未检出'/'批量另存完成'的框 -> 相应按钮
后台常驻轮询。用法: python sw_popup_guard.py

2026-10-17 修复(真实卡死根因): 旧版只扫 sldworks.exe 进程的窗口(见下注释), 但
skill7 步骤8 调用的 C:\\AutoBOM\\AutoBaiJinZhanKai.exe 批量展开时, 若目标目录已存在
同名 DXF 会弹「覆盖提示！」(#32770, 属 AutoBaiJinZhanKai.exe), 不在扫描范围内
→ 无人处理, 批处理永久卡死(实测 15:13 卡到被人工发现)。故扫描进程扩展为
sldworks.exe + AutoBaiJinZhanKai.exe。
"""
import ctypes, ctypes.wintypes as wt, time, sys, signal

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

# —— 2026-08-30 二转根因修复 ——
# 旧版守卫全系统扫描 + 同步 GetWindowTextW/SendMessageW 跨进程调用,
# 一旦某个忙/卡窗口让一次调用无限阻塞, 守卫就永久卡死(CPU 归零不再轮询),
# 后续 PDM 另存为弹窗无人处理 → 分类再次卡死。
# 新版: ①只扫 sldworks.exe 进程窗口(PDM 对话框本来就是 SW 进程内 #32770, 不影响覆盖)
#       ②读标题/点按钮一律 SendMessageTimeoutW 带 200ms 超时 → 永不卡死。
SMTO_ABORTIFHUNG = 0x0002
WM_GETTEXT = 0x000D
WM_GETTEXTLENGTH = 0x000E
BM_CLICK = 0x00F5
WM_CLOSE = 0x0010

# 需要监视弹窗的进程(小写)。sldworks=Pdm/另存为/完成提示; AutoBaiJinZhanKai=钣金展开「覆盖提示！」(2026-10-17)
WATCH_EXES = ('sldworks.exe', 'autobaijinzhankai.exe')


def _get_text(hwnd, timeout_ms=200):
    """SendMessageTimeoutW 带超时读窗口标题, 避免对忙/卡窗口无限阻塞。"""
    ln = ctypes.c_size_t()
    if not user32.SendMessageTimeoutW(hwnd, WM_GETTEXTLENGTH, 0, 0, SMTO_ABORTIFHUNG, timeout_ms, ctypes.byref(ln)):
        return ''
    n = ln.value
    if n <= 0:
        return ''
    buf = ctypes.create_unicode_buffer(n + 1)
    res = ctypes.c_size_t()
    if not user32.SendMessageTimeoutW(hwnd, WM_GETTEXT, n + 1, buf, SMTO_ABORTIFHUNG, timeout_ms, ctypes.byref(res)):
        return ''
    return buf.value.strip()


def sw_pids():
    """返回所有可能弹出阻塞对话框的进程 PID 集合。
    2026-10-17: 除 sldworks.exe 外, 加入 AutoBaiJinZhanKai.exe(钣金展开批量的「覆盖提示！」)。"""
    pids = set()
    TH32CS_SNAPPROCESS = 0x00000002
    class PROCESSENTRY32W(ctypes.Structure):
        _fields_ = [("dwSize", wt.DWORD), ("cntUsage", wt.DWORD), ("th32ProcessID", wt.DWORD),
                    ("th32DefaultHeapID", ctypes.POINTER(wt.ULONG)), ("th32ModuleID", wt.DWORD),
                    ("cntThreads", wt.DWORD), ("th32ParentProcessID", wt.DWORD),
                    ("pcPriClassBase", wt.LONG), ("dwFlags", wt.DWORD),
                    ("szExeFile", ctypes.c_wchar * 260)]
    snap = kernel32.CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    if snap == ctypes.c_void_p(-1).value:
        return pids
    pe = PROCESSENTRY32W()
    pe.dwSize = ctypes.sizeof(PROCESSENTRY32W)
    if kernel32.Process32FirstW(snap, ctypes.byref(pe)):
        while True:
            if pe.szExeFile.lower() in WATCH_EXES:
                pids.add(pe.th32ProcessID)
            if not kernel32.Process32NextW(snap, ctypes.byref(pe)):
                break
    kernel32.CloseHandle(snap)
    return pids


def find_poped_windows():
    """返回所有可能阻塞的 SW 弹窗列表 [(hwnd, title)] —— 只扫 sldworks.exe 进程窗口。"""
    out = []
    pids = sw_pids()
    if not pids:
        return out
    @ctypes.WINFUNCTYPE(wt.BOOL, wt.HWND, wt.LPARAM)
    def cb(h, lp):
        if not user32.IsWindowVisible(h):
            return True
        pid = wt.DWORD()
        user32.GetWindowThreadProcessId(h, ctypes.byref(pid))
        if pid.value not in pids:
            return True
        t = _get_text(h)
        if not t:
            return True
        if t in ('SOLIDWORKS PDM', '另存为', 'SOLIDWORKS', '保存', '保存确认', '覆盖提示！', '检出'):
            out.append((h, t))
        return True
    user32.EnumWindows(cb, 0)
    return out

# 退出标志：收到 SIGTERM/KeyboardInterrupt 时置位，主循环据此优雅退出。
# (2026-08-08) skill7 结束时会 terminate 本进程；若不能优雅退出则变孤儿进程，
# 残留后会自动点掉/关闭 SW 弹窗，干扰皇上正常用 SW。故必须响应外部终止信号。
_STOP = False

def _on_signal(signum, frame):
    global _STOP
    _STOP = True
    print(f"[SW-GUARD] 收到信号 {signum}，准备退出", flush=True)

def _norm_btn(s):
    """归一化按钮文本: 去掉 (&X) 快捷键后缀。
    2026-08-11 修复: '取消(&C)' 之前匹配不上策略里的 '取消', 导致虚拟件另存为弹窗永不处理。"""
    import re
    return re.sub(r'\(&.\)', '', s).strip()


def click_button(dlg, button_texts):
    """在对话框中按优先级点按钮。返回是否点了某按钮。"""
    res = [None]
    @ctypes.WINFUNCTYPE(wt.BOOL, wt.HWND, wt.LPARAM)
    def cb(h, lp):
        if res[0]:
            return True
        label = _get_text(h)
        for want in button_texts:
            if label == want or _norm_btn(label) == _norm_btn(want):
                # 带 200ms 超时点按钮, 防同步阻塞
                user32.SendMessageTimeoutW(h, BM_CLICK, 0, 0, SMTO_ABORTIFHUNG, 200, ctypes.byref(ctypes.c_size_t()))
                res[0] = want
                return True
        return True
    user32.EnumChildWindows(dlg, cb, 0)
    return res[0]

# 对话框类型 -> 优先点击的按钮顺序
STRATEGY = {
    'SOLIDWORKS PDM': ['否(&N)', '否', '取消', '是(&Y)'],   # 不检出只读
    '另存为':         ['取消', '否(&N)', '否'],              # 取消另存
    'SOLIDWORKS':     ['确定', '否(&N)', '否', '取消'],      # 完成提示
    '保存':           ['取消', '否', '是'],                  # 保存提示
    '保存确认':       ['取消', '否', '是'],
    '覆盖提示！':      ['是(&Y)', '是', '确定'],              # 2026-10-17 钣金展开: 覆盖旧DXF(重新生成)
    # 2026-10-21: PDM「检出」对话框(标题就叫'检出', 属 sldworks, #32770)。
    # 实测 26102101\双口变径自动闸 步骤8.0 collect 被它卡死 7 分钟(要检出 100 Files)。
    # 一律点『取消』: 检出会在 PDM 里锁住上百个文件、影响同事, 绝不能自动点『检出』。
    # 取消后该文件保持只读, 后续写操作会失败(可接受, 由人决定是否检出)。
    '检出':           ['取消', '否(&N)', '否'],
}

def main():
    interval = float(sys.argv[1]) if len(sys.argv) > 1 else 1.5
    print(f"[SW-GUARD] 启动 轮询间隔={interval}s (处理 PDM未检出/另存为/完成提示)", flush=True)
    # 2026-10-03 彻底修复: 记录父进程 PID, 父进程消失(含被强杀/崩溃)则自动退出,
    # 杜绝主脚本未调用 stop_popup_guard 时守卫残留后台干扰人工使用 SW。
    _parent_pid = None
    try:
        _parent_pid = int(sys.argv[2]) if len(sys.argv) > 2 else None
    except Exception:
        _parent_pid = None
    if _parent_pid is None:
        try:
            import ctypes as _c
            _parent_pid = _c.windll.kernel32.GetCurrentProcessId()  # 占位, 下面改真实父进程
        except Exception:
            _parent_pid = None
    # 真正取父进程 PID (Python 侧: 用 psutil 不可靠/未装, 用 Toolhelp 按当前进程找父)
    if _parent_pid is None or _parent_pid == 0:
        try:
            _pid_me = kernel32.GetCurrentProcessId()
            _ppid = 0
            TH32CS_SNAPPROCESS = 0x00000002
            class _PE(ctypes.Structure):
                _fields_ = [("dwSize", wt.DWORD), ("cntUsage", wt.DWORD), ("th32ProcessID", wt.DWORD),
                            ("th32DefaultHeapID", ctypes.POINTER(wt.ULONG)), ("th32ModuleID", wt.DWORD),
                            ("cntThreads", wt.DWORD), ("th32ParentProcessID", wt.DWORD),
                            ("pcPriClassBase", wt.LONG), ("dwFlags", wt.DWORD),
                            ("szExeFile", ctypes.c_wchar * 260)]
            _snap = kernel32.CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
            if _snap != ctypes.c_void_p(-1).value:
                _pe = _PE(); _pe.dwSize = ctypes.sizeof(_PE)
                if kernel32.Process32FirstW(_snap, ctypes.byref(_pe)):
                    while True:
                        if _pe.th32ProcessID == _pid_me:
                            _ppid = _pe.th32ParentProcessID
                            break
                        if not kernel32.Process32NextW(_snap, ctypes.byref(_pe)):
                            break
                kernel32.CloseHandle(_snap)
            _parent_pid = _ppid
        except Exception:
            _parent_pid = None
    if _parent_pid:
        print(f"[SW-GUARD] 父进程 PID={_parent_pid}, 父进程消失将自动退出", flush=True)
    else:
        print("[SW-GUARD] 未取得父进程 PID, 将依赖外部 stop 清理", flush=True)
    # 注册信号：外部 terminate/taskkill 时优雅退出
    try:
        signal.signal(signal.SIGTERM, _on_signal)
    except Exception:
        pass
    try:
        signal.signal(signal.SIGINT, _on_signal)
    except Exception:
        pass
    handled = 0
    while not _STOP:
        # 父进程守护: 父进程已消失 → 退出 (防残留)
        if _parent_pid:
            try:
                _h = kernel32.OpenProcess(0x0400, False, _parent_pid)  # PROCESS_QUERY_LIMITED_INFORMATION
                if not _h:
                    print(f"[SW-GUARD] 父进程 {_parent_pid} 已退出, 自动退出 (累计处理 {handled})", flush=True)
                    break
                _ex = wt.DWORD()
                _ok = kernel32.GetExitCodeProcess(_h, ctypes.byref(_ex))
                kernel32.CloseHandle(_h)
                if _ok and _ex.value != 259:  # 259=STILL_ACTIVE
                    print(f"[SW-GUARD] 父进程 {_parent_pid} 已结束(exit={_ex.value}), 自动退出 (累计处理 {handled})", flush=True)
                    break
            except Exception:
                pass
        try:
            for h, title in find_poped_windows():
                if title in ('另存为', '保存', '保存确认'):
                    # PDM 另存为对话框: 直接 WM_CLOSE, 不点任何按钮。
                    # 2026-08-30 四转根因修复: 此类弹窗有的带「取消」按钮但点了无效,
                    # 守卫 click_button 点到「取消」返回 clicked=True, 永远走不到 WM_CLOSE → 弹窗挂死。
                    # WM_CLOSE 已实测 100% 有效(SZ.Z3.0/SZ.Z4.0 均验证), 直接关掉让阻塞的 Save() 返回。
                    user32.PostMessageW(h, WM_CLOSE, 0, 0)
                    handled += 1
                    print(f"[SW-GUARD] 弹窗'{title}' hwnd={h} -> WM_CLOSE 关闭 累计{handled}", flush=True)
                    continue
                strategy = STRATEGY.get(title, ['取消', '否(&N)', '确定'])
                clicked = click_button(h, strategy)
                if clicked:
                    handled += 1
                    print(f"[SW-GUARD] 弹窗'{title}' hwnd={h} -> 点'{clicked}' 累计{handled}", flush=True)
        except Exception:
            pass
        time.sleep(interval)
    print(f"[SW-GUARD] 已退出 (累计处理 {handled} 个弹窗)", flush=True)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("[SW-GUARD] 退出")
