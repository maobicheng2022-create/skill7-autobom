# -*- coding: utf-8 -*-
"""skill7 环境自检 — 装完先跑这一条，把"能不能跑"一次问清。

用法:
    python -X utf8 tools/verify_env.py

退出码 0 = 关键项全过；1 = 有关键项不过（会打印怎么修）。
"""
import json
import os
import platform
import sys

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(HERE, "shared"))

import skill7_paths as P  # noqa: E402

GUID = "{E781EAD9-A4FA-4560-B0A2-994D349CB64A}"
REQUIRED_TEMPLATES = ["钣金", "型材", "型材采购", "总外购件", "标准件采购", "焊接部件清点清单"]
REQUIRED_MACROS = [
    "零件分类.swp",
    "型材零件名称数量合成另存文件夹.swp",
    "焊接或装配分类(焊接）.swp",
    "焊接或装配分类(装配).swp",
    "焊接或装配分类(清空）.swp",
]

results = []  # (level, name, detail)  level: OK / WARN / FAIL


def add(level, name, detail=""):
    results.append((level, name, detail))


def check_python():
    v = sys.version_info
    if v >= (3, 10):
        add("OK", "Python 版本", "%d.%d.%d" % v[:3])
    else:
        add("FAIL", "Python 版本", "%d.%d.%d（需要 3.10+）" % v[:3])
    if platform.architecture()[0] == "64bit":
        add("OK", "Python 位数", "64bit")
    else:
        add("FAIL", "Python 位数", "32bit（SolidWorks COM 需要 64 位 Python）")


def check_deps():
    seen = set()
    for mod, pip_name in [("win32com.client", "pywin32"), ("pythoncom", "pywin32"),
                          ("openpyxl", "openpyxl"), ("PIL", "Pillow"), ("ezdxf", "ezdxf")]:
        if pip_name in seen:
            continue
        seen.add(pip_name)
        try:
            __import__(mod)
            add("OK", "依赖 %s" % pip_name, "")
        except Exception as e:
            add("FAIL", "依赖 %s" % pip_name, "pip install %s（%s）" % (pip_name, repr(e)[:60]))


def check_paths():
    for key, val, ok in P.describe():
        if key == "ACAD_EXE":
            continue  # AutoCAD 仅合并 dwg 时才需要
        level = "OK" if ok else ("WARN" if key in ("SOLIDWORKS_EXE", "SHOT_CACHE") else "FAIL")
        add(level, key, val)


def check_autobom():
    exe = os.path.join(P.AUTOBOM_DIR, "AutoBOM.exe")
    dll = os.path.join(P.AUTOBOM_DIR, "AutoBOM.SW.dll")
    for f in (exe, dll):
        if os.path.isfile(f):
            add("OK", "AutoBOM 文件", f)
        else:
            add("FAIL", "AutoBOM 文件", "缺少 %s（先跑 tools/install.ps1）" % f)


def check_addin():
    if not sys.platform.startswith("win"):
        return
    try:
        import winreg
    except Exception:
        return
    # SW 的插件列表只读机器级 HKLM；HKCU 只用于“启动时加载”开关。
    try:
        winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"Software\SolidWorks\Addins\%s" % GUID)
        add("OK", "SW 插件注册(HKLM\\Addins)", GUID)
    except OSError:
        add("FAIL", "SW 插件注册(HKLM\\Addins)",
            "未注册。SW 只读 HKLM，必须用管理员运行 autobom/portable/AutoBOM-SW-register.ps1")
    try:
        winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\SolidWorks\Addins\%s" % GUID)
        add("WARN", "HKCU 同名注册(会盖住机器级)",
            "存在 HKCU\\Software\\SolidWorks\\Addins 同名项；SW 可能忽略 HKLM。请删掉它后重启 SW")
    except OSError:
        pass
    try:
        winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"CLSID\%s\InprocServer32" % GUID)
        add("OK", "COM 类注册(CLSID)", GUID)
    except OSError:
        add("WARN", "COM 类注册(CLSID)", "未注册（同上，管理员跑注册脚本）")


def check_templates():
    path = P.AUTOBOM_CFG
    if not os.path.isfile(path):
        add("FAIL", "BOM 模板 setting.json", path)
        return
    try:
        with open(path, encoding="utf-8-sig") as f:
            cfg = json.load(f)
    except Exception as e:
        add("FAIL", "BOM 模板 setting.json", "解析失败: %r" % (e,))
        return
    names = [t.get("Name") for t in cfg.get("Templates", [])]
    missing = [n for n in REQUIRED_TEMPLATES if n not in names]
    if missing:
        add("FAIL", "BOM 模板", "缺模板: %s（当前: %s）" % (", ".join(missing), ", ".join(names)))
    else:
        add("OK", "BOM 模板", "%d 套: %s" % (len(names), ", ".join(names)))


def check_macros():
    if not os.path.isdir(P.MACRO_DIR):
        add("FAIL", "宏目录", P.MACRO_DIR)
        return
    have = set(os.listdir(P.MACRO_DIR))
    macro_count = len([m for m in have if m.lower().endswith(".swp")])
    missing = [m for m in REQUIRED_MACROS if m not in have]
    if missing:
        add("FAIL", "SW 宏", "缺 %s（宏目录 %s）" % (", ".join(missing), P.MACRO_DIR))
    else:
        add("OK", "SW 宏", "%d 个 .swp（%s）" % (macro_count, P.MACRO_DIR))


def check_skills():
    need = ["1-图号名称更新", "3-导出BOM", "4-导出BOM带截图", "5-运行SW宏",
            "6-合并DXF", "7-sw文档综合处理", "8-焊接装配标志分类输出BOM带截图",
            "9-整理型材", "10-虚拟件外部化与Pipe标记", "11-PVC管采购计算"]
    missing = [n for n in need if not os.path.isdir(os.path.join(P.SKILLS_DIR, n))]
    if missing:
        add("FAIL", "skills 目录", "缺: %s" % ", ".join(missing))
    else:
        add("OK", "skills 目录", "%d 个配套 skill" % len(need))


def check_sw_running():
    try:
        import win32com.client
    except Exception:
        return
    try:
        app = win32com.client.GetActiveObject("SldWorks.Application")
        add("OK", "SolidWorks 连接", "已连上运行中的 SW %s" % getattr(app, "GetBuildNumbers", lambda: "")())
    except Exception:
        add("WARN", "SolidWorks 连接", "当前没有可连接的 SW（脚本会自动拉起，可忽略）")


def main():
    check_python()
    check_deps()
    check_paths()
    check_autobom()
    check_addin()
    check_templates()
    check_macros()
    check_skills()
    check_sw_running()

    width = max(len(n) for _, n, _ in results) + 2
    print("=" * 78)
    print("skill7 环境自检   配置文件: %s" % (P.CONFIG_FILE or "(未使用，走自动探测)"))
    print("=" * 78)
    for level, name, detail in results:
        mark = {"OK": "[OK]  ", "WARN": "[警告]", "FAIL": "[失败]"}[level]
        print("%s %-*s %s" % (mark, width, name, detail))

    fails = [r for r in results if r[0] == "FAIL"]
    warns = [r for r in results if r[0] == "WARN"]
    print("-" * 78)
    print("通过 %d 项，警告 %d 项，失败 %d 项"
          % (len(results) - len(fails) - len(warns), len(warns), len(fails)))
    if fails:
        print("\n关键项没过，先按上面 [失败] 行的提示修；修完重跑本脚本。")
        return 1
    print("\n环境就绪。跑一条线试试：")
    print('  python -X utf8 "skills\\7-sw文档综合处理\\scripts\\sw_doc_process.py" "<图纸文件夹>" --qty 1')
    return 0


if __name__ == "__main__":
    sys.exit(main())