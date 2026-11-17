# -*- coding: utf-8 -*-
"""skill7 路径中枢（开源部署版）

把所有"写死在某台电脑上"的绝对路径集中到一处解析，优先级从高到低：

    1. 环境变量  SKILL7_<KEY>
    2. 配置文件  config/config.json（仓库内，已 gitignore）
                 或 %USERPROFILE%\\.skill7\\config.json
                 或 SKILL7_CONFIG 指向的文件
    3. 自动探测（SolidWorks 注册表、AutoCAD 注册表、C:\\AutoBOM、仓库自带模板）
    4. 内置默认值

任何脚本只要 `from skill7_paths import ...` 即可拿到本机正确路径，
不再需要改源码。用法见 docs/02-配置说明.md。
"""
import json
import os
import shutil
import sys

# ---------------------------------------------------------------- 仓库根目录
# shared/skill7_paths.py -> 仓库根
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SKILLS_DIR = os.path.join(REPO_ROOT, "skills")
BUNDLED_MACROS = os.path.join(REPO_ROOT, "macros")
BUNDLED_TEMPLATES = os.path.join(REPO_ROOT, "templates", "micDatas")


def _load_json(path):
    try:
        with open(path, encoding="utf-8-sig") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _config_candidates():
    env = os.environ.get("SKILL7_CONFIG")
    if env:
        yield env
    yield os.path.join(REPO_ROOT, "config", "config.json")
    yield os.path.join(os.path.expanduser("~"), ".skill7", "config.json")


CFG = {}
CONFIG_FILE = None
for _p in _config_candidates():
    if _p and os.path.isfile(_p):
        CFG = _load_json(_p)
        CONFIG_FILE = _p
        break


def _get(key, default=None):
    """环境变量 SKILL7_<KEY大写> > 配置文件 key > 默认值。"""
    env = os.environ.get("SKILL7_" + key.upper())
    if env:
        return env
    val = CFG.get(key)
    if val not in (None, ""):
        return val
    return default


def _norm(path):
    if not path:
        return path
    return os.path.normpath(os.path.expandvars(os.path.expanduser(str(path))))


def _first_existing(*paths):
    for p in paths:
        if p and os.path.isfile(p):
            return os.path.normpath(p)
    return None


# ---------------------------------------------------------------- 注册表探测
def _sw_from_registry():
    """从 HKLM\\SOFTWARE\\SolidWorks\\SOLIDWORKS 20xx\\Setup 的 SolidWorks Folder 取安装目录。"""
    if not sys.platform.startswith("win"):
        return None
    try:
        import winreg
    except Exception:
        return None
    best = None
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\SolidWorks") as root:
            versions = []
            i = 0
            while True:
                try:
                    name = winreg.EnumKey(root, i)
                except OSError:
                    break
                i += 1
                if name.upper().startswith("SOLIDWORKS 20"):
                    versions.append(name)
            versions.sort(key=lambda s: int("".join(ch for ch in s if ch.isdigit()) or 0),
                          reverse=True)  # 按版本号数值降序，避免 "SolidWorks 2022" 排到 "SOLIDWORKS 2025" 前面
            for name in versions:
                try:
                    with winreg.OpenKey(root, name + r"\Setup") as setup:
                        folder, _ = winreg.QueryValueEx(setup, "SolidWorks Folder")
                except OSError:
                    continue
                exe = os.path.join(folder, "sldworks.exe")
                if os.path.isfile(exe):
                    best = os.path.normpath(exe)
                    break
    except OSError:
        return None
    return best


def _acad_from_registry():
    """从 HKLM\\SOFTWARE\\Autodesk\\AutoCAD\\Rxx\\ACAD-xxx\\AcadLocation 取 acad.exe。"""
    if not sys.platform.startswith("win"):
        return None
    try:
        import winreg
    except Exception:
        return None
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Autodesk\AutoCAD") as root:
            releases = []
            i = 0
            while True:
                try:
                    releases.append(winreg.EnumKey(root, i))
                except OSError:
                    break
                i += 1
            releases.sort(reverse=True)
            for rel in releases:
                with winreg.OpenKey(root, rel) as rk:
                    i = 0
                    while True:
                        try:
                            prod = winreg.EnumKey(rk, i)
                        except OSError:
                            break
                        i += 1
                        try:
                            with winreg.OpenKey(rk, prod + r"\AcadLocation") as ak:
                                folder, _ = winreg.QueryValueEx(ak, "")
                        except OSError:
                            continue
                        exe = os.path.join(folder, "acad.exe")
                        if os.path.isfile(exe):
                            return os.path.normpath(exe)
    except OSError:
        return None
    return None


# ---------------------------------------------------------------- AutoBOM
AUTOBOM_DIR = _norm(_get("autobom_dir", r"C:\AutoBOM"))

# AutoBOM 模板/配置文件目录：AutoBOM 本体默认在 C:\micDatas（与 C:\AutoBOM 并列）
MICDATAS_DIR = _norm(_get("micdatas_dir", r"C:\micDatas"))
if not os.path.isdir(MICDATAS_DIR) and os.path.isdir(BUNDLED_TEMPLATES):
    # 没装 AutoBOM 模板时退回仓库自带模板，保证脚本能读到 setting.json
    MICDATAS_DIR = BUNDLED_TEMPLATES
AUTOBOM_CFG = os.path.join(MICDATAS_DIR, "pluginSettings", "BOM", "setting.json")

def _detect_shot_cache():
    """截图库自动探测：AutoBOM 目录下 + 历史遗留的 D:\\AutoBOM\\ShotCache。

    多份库并存时选 shot_index.db 最新的那份（缓存命中率高）；都没有则用 AutoBOM 目录下的。
    """
    candidates = [os.path.join(AUTOBOM_DIR, "ShotCache"), r"D:\AutoBOM\ShotCache"]
    best, best_mtime = None, -1
    for d in candidates:
        db = os.path.join(d, "shot_index.db")
        if os.path.isfile(db):
            m = os.path.getmtime(db)
            if m > best_mtime:
                best, best_mtime = d, m
    return best or os.path.join(AUTOBOM_DIR, "ShotCache")


SHOT_CACHE = _norm(_get("shot_cache_dir") or _detect_shot_cache())
ZHANKAI_EXE = _norm(_get("zhankai_exe", os.path.join(AUTOBOM_DIR, "AutoBaiJinZhanKai.exe")))

# ---------------------------------------------------------------- 宏
MACRO_DIR = _norm(_get("macro_dir", BUNDLED_MACROS))
MACRO_CLASSIFY = os.path.join(MACRO_DIR, "零件分类.swp")
MACRO_PROFILE = os.path.join(MACRO_DIR, "型材零件名称数量合成另存文件夹.swp")

# ---------------------------------------------------------------- 弹窗守卫
POPUP_GUARD = _norm(_get(
    "popup_guard",
    os.path.join(SKILLS_DIR, "7-sw文档综合处理", "scripts", "sw_popup_guard.py")))

# ---------------------------------------------------------------- SolidWorks
SOLIDWORKS_EXE = _norm(_first_existing(
    _get("solidworks_exe"),
    _sw_from_registry(),
    r"E:\SOLIDWORKS2025\SOLIDWORKS\sldworks.exe",
    r"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\sldworks.exe",
    r"D:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\sldworks.exe",
    r"E:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\sldworks.exe",
) or shutil.which("sldworks.exe") or "sldworks.exe")

# 兼容原脚本里的 SLDWORKS_* 命名
SLDWORKS_EXE = SOLIDWORKS_EXE
SLDWORKS_CANDIDATES = [
    SOLIDWORKS_EXE,
    r"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\sldworks.exe",
    r"D:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\sldworks.exe",
    r"E:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\sldworks.exe",
    "sldworks.exe",  # 保底: 靠 PATH 找
]

# ---------------------------------------------------------------- AutoCAD（仅 dwg2dxf 用）
ACAD_EXE = _norm(_get("acad_exe") or _acad_from_registry() or r"D:\CAD\AutoCAD 2026\acad.exe")

# ---------------------------------------------------------------- PDM 库标记
# 总装引用若指向 PDM 只读库，保存会弹"另存为"阻塞。填这里的子串会被识别为 PDM 库路径。
PDM_MARKERS = [str(x).lower() for x in _get("pdm_library_markers", []) or []]


# ---------------------------------------------------------------- 自检报告
def describe():
    """返回 (key, value, state) 列表，供 verify_env.py 打印。"""
    files = {
        "AUTOBOM_DIR": (AUTOBOM_DIR, os.path.isdir(AUTOBOM_DIR)),
        "AUTOBOM_CFG": (AUTOBOM_CFG, os.path.isfile(AUTOBOM_CFG)),
        "ZHANKAI_EXE": (ZHANKAI_EXE, os.path.isfile(ZHANKAI_EXE)),
        "SHOT_CACHE": (SHOT_CACHE, os.path.isdir(SHOT_CACHE)),
        "MACRO_DIR": (MACRO_DIR, os.path.isdir(MACRO_DIR)),
        "MACRO_PROFILE": (MACRO_PROFILE, os.path.isfile(MACRO_PROFILE)),
        "POPUP_GUARD": (POPUP_GUARD, os.path.isfile(POPUP_GUARD)),
        "SOLIDWORKS_EXE": (SOLIDWORKS_EXE, os.path.isfile(SOLIDWORKS_EXE)),
        "ACAD_EXE": (ACAD_EXE, os.path.isfile(ACAD_EXE)),
        "SKILLS_DIR": (SKILLS_DIR, os.path.isdir(SKILLS_DIR)),
    }
    return [(k, v[0], v[1]) for k, v in files.items()]


if __name__ == "__main__":
    print("配置文件: %s" % (CONFIG_FILE or "(未使用，走自动探测)"))
    for key, val, ok in describe():
        print("%-4s %-16s %s" % ("OK" if ok else "!!", key, val))