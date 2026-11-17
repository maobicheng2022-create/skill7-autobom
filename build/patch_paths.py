# -*- coding: utf-8 -*-
"""把仓库里 12 个 skill 的硬编码绝对路径改成 shared/skill7_paths.py 解析。

只做两件事：
  1. 在文件 import 区末尾插入一段 bootstrap（把 shared/ 加进 sys.path 并导入路径常量）
  2. 删除/改写原来写死的常量

每个替换都断言"恰好命中一次"，命中不了就报错，避免静默改坏。
跑法：  python -X utf8 build/patch_paths.py [--check]
"""
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SK = os.path.join(ROOT, "skills")

BOOTSTRAP = '''
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
'''

# 每个文件：bootstrap=True 表示插入；edits = [(old, new), ...]
FILES = {
    "0-skill管理规范/scripts/autobom_backup.py": {
        "edits": [
            ('AUTOBOM_DIR = r"C:\\AutoBOM"\n', ""),
            ('CONFIG_DIR = r"C:\\micDatas\\pluginSettings\\BOM"\n',
             'CONFIG_DIR = os.path.join(MICDATAS_DIR, "pluginSettings", "BOM")\n'),
            ('BACKUP_ROOT = r"C:\\AutoBOM\\_backup"\n',
             'BACKUP_ROOT = os.path.join(AUTOBOM_DIR, "_backup")\n'),
            ('s = r"C:\\AutoBOM\\_src\\AutoBOM.exe"',
             's = os.path.join(AUTOBOM_DIR, "_src", "AutoBOM.exe")'),
        ],
    },
    "3-导出BOM/scripts/autobom_export.py": {
        "edits": [
            ('AUTOBOM_CFG = r"C:\\micDatas\\pluginSettings\\BOM\\setting.json"\n', ""),
            ('cache_dir = r"D:\\AutoBOM\\ShotCache"', 'cache_dir = SHOT_CACHE'),
        ],
    },
    "3-导出BOM/scripts/prop_cache.py": {
        "edits": [
            ('DEFAULT_CACHE_DIR = r"C:\\AutoBOM\\ShotCache"  # 与截图库同一目录',
             'DEFAULT_CACHE_DIR = SHOT_CACHE  # 由 shared/skill7_paths.py 解析'),
        ],
    },
    "4-导出BOM带截图/scripts/prop_cache.py": {
        "edits": [
            ('DEFAULT_CACHE_DIR = r"D:\\AutoBOM\\ShotCache"  # 与截图库同一目录',
             'DEFAULT_CACHE_DIR = SHOT_CACHE  # 由 shared/skill7_paths.py 解析'),
        ],
    },
    "8-焊接装配标志分类输出BOM带截图/scripts/prop_cache.py": {
        "edits": [
            ('DEFAULT_CACHE_DIR = r"D:\\AutoBOM\\ShotCache"  # 与截图库同一目录',
             'DEFAULT_CACHE_DIR = SHOT_CACHE  # 由 shared/skill7_paths.py 解析'),
        ],
    },
    "4-导出BOM带截图/scripts/shot_cache.py": {
        "edits": [
            ('DEFAULT_CACHE_DIR = r"D:\\AutoBOM\\ShotCache"',
             'DEFAULT_CACHE_DIR = SHOT_CACHE  # 由 shared/skill7_paths.py 解析'),
        ],
    },
    "8-焊接装配标志分类输出BOM带截图/scripts/shot_cache.py": {
        "edits": [
            ('DEFAULT_CACHE_DIR = r"D:\\AutoBOM\\ShotCache"',
             'DEFAULT_CACHE_DIR = SHOT_CACHE  # 由 shared/skill7_paths.py 解析'),
        ],
    },
    "4-导出BOM带截图/scripts/sw_buyout_export.py": {
        "edits": [
            ('AUTOBOM_CFG = r"C:\\micDatas\\pluginSettings\\BOM\\setting.json"\n', ""),
            ('exe = _sw_exe_from_registry() or r"E:\\Program Files\\SOLIDWORKS Corp\\SOLIDWORKS\\sldworks.exe"',
             'exe = _sw_exe_from_registry() or SOLIDWORKS_EXE'),
            ('help="截图存放目录(默认 C:\\\\AutoBOM\\\\ShotCache 全局截图库; 旧版 IMG 子目录已废弃)"',
             'help="截图存放目录(默认 %s)" % SHOT_CACHE'),
            ('help="截图缓存库目录(默认 C:\\\\AutoBOM\\\\ShotCache)"',
             'help="截图缓存库目录(默认 %s)" % SHOT_CACHE'),
        ],
    },
    "4-导出BOM带截图/scripts/sw_preshot.py": {
        "edits": [
            ('SLDWORKS_EXE = r"E:\\Program Files\\SOLIDWORKS Corp\\SOLIDWORKS\\sldworks.exe"\n', ""),
        ],
    },
    "6-合并DXF/scripts/dwg2dxf.py": {
        "edits": [
            ('ACAD_EXE = r"D:\\CAD\\AutoCAD 2026\\acad.exe"\n', ""),
        ],
    },
    "7-sw文档综合处理/scripts/sw_doc_process.py": {
        "edits": [
            ('POPUP_GUARD = r"C:\\Users\\mao\\.openclaw\\workspace\\turing-workspace\\scripts\\sw_popup_guard.py"\n', ""),
            ('SLDWORKS_EXE = r"E:\\Program Files\\SOLIDWORKS Corp\\SOLIDWORKS\\sldworks.exe"\n', ""),
            ('SLDWORKS_CANDIDATES = [\n    SLDWORKS_EXE,\n    r"C:\\Program Files\\SOLIDWORKS Corp\\SOLIDWORKS\\sldworks.exe",\n    r"D:\\Program Files\\SOLIDWORKS Corp\\SOLIDWORKS\\sldworks.exe",\n    "sldworks.exe",  # 保底: 靠 PATH 找\n]\n', ""),
            ('GLOBAL_SKILLS = r"C:\\Users\\mao\\.openclaw\\workspace\\skills"\n', ""),
            ('MACRO_CLASSIFY = r"Z:\\solidworks information\\macro\\零件分类.swp"\n', ""),
            ('MACRO_PROFILE = r"Z:\\solidworks information\\macro\\型材零件名称数量合成另存文件夹.swp"\n', ""),
            ('ZHANKAI_EXE = r"C:\\AutoBOM\\AutoBaiJinZhanKai.exe"\n', ""),
            ('r"D:\\AutoBOM\\ShotCache"', 'SHOT_CACHE', "all"),
        ],
    },
    "7-sw文档综合处理/scripts/sw_tuhao_categorize.py": {
        "edits": [
            ('SLDWORKS_EXE = r"E:\\Program Files\\SOLIDWORKS Corp\\SOLIDWORKS\\sldworks.exe"\n', ""),
            ('if "d:\\\\feedinghero" not in low:',
             'if not any(m in low for m in PDM_MARKERS):'),
        ],
    },
    "7-sw文档综合处理/scripts/sw_weld_group.py": {
        "edits": [
            ('GLOBAL_SKILLS = r"C:\\Users\\mao\\.openclaw\\workspace\\skills"\n', ""),
            ('POPUP_GUARD = r"C:\\Users\\mao\\.openclaw\\workspace\\turing-workspace\\scripts\\sw_popup_guard.py"\n', ""),
            ('SLDWORKS_CANDIDATES = [\n    r"E:\\Program Files\\SOLIDWORKS Corp\\SOLIDWORKS\\sldworks.exe",\n    r"C:\\Program Files\\SOLIDWORKS Corp\\SOLIDWORKS\\sldworks.exe",\n    r"D:\\Program Files\\SOLIDWORKS Corp\\SOLIDWORKS\\sldworks.exe",\n    "sldworks.exe",\n]\n', ""),
        ],
    },
    "7-sw文档综合处理/scripts/sw_batch_categorize.py": {
        "edits": [
            ('"Start-Process \'E:\\\\Program Files\\\\SOLIDWORKS Corp\\\\SOLIDWORKS\\\\sldworks.exe\'"],',
             '"Start-Process \'%s\'" % SOLIDWORKS_EXE],'),
        ],
    },
    "8-焊接装配标志分类输出BOM带截图/scripts/sw_weld_export.py": {
        "edits": [
            ('AUTOBOM_CFG = r"C:\\micDatas\\pluginSettings\\BOM\\setting.json"\n', ""),
            ('POPUP_GUARD = r"C:\\Users\\mao\\.openclaw\\workspace\\turing-workspace\\scripts\\sw_popup_guard.py"\n', ""),
            ('SLDWORKS_CANDIDATES = [\n    r"E:\\Program Files\\SOLIDWORKS Corp\\SOLIDWORKS\\sldworks.exe",\n    r"C:\\Program Files\\SOLIDWORKS Corp\\SOLIDWORKS\\sldworks.exe",\n    r"D:\\Program Files\\SOLIDWORKS Corp\\SOLIDWORKS\\sldworks.exe",\n    "sldworks.exe",  # 保底: 靠 PATH 找\n]\n', ""),
            ('help="截图存放目录(默认 C:\\\\AutoBOM\\\\ShotCache 全局截图库; 旧版 IMG 子目录已废弃)"',
             'help="截图存放目录(默认 %s)" % SHOT_CACHE'),
            ('help="截图缓存库目录(默认 C:\\\\AutoBOM\\\\ShotCache)"',
             'help="截图缓存库目录(默认 %s)" % SHOT_CACHE'),
        ],
    },
    "10-虚拟件外部化与Pipe标记/scripts/sw_virtual_extract_and_mark.py": {
        "edits": [
            ('POPUP_GUARD = r"C:\\Users\\mao\\.openclaw\\workspace\\turing-workspace\\scripts\\sw_popup_guard.py"\n', ""),
        ],
    },
    "11-PVC管采购计算/scripts/sw_pvc_purchase.py": {
        "edits": [
            ('AUTOBOM_CFG = r"C:\\micDatas\\pluginSettings\\BOM\\setting.json"\n', ""),
            ('POPUP_GUARD = r"C:\\Users\\mao\\.openclaw\\workspace\\turing-workspace\\scripts\\sw_popup_guard.py"\n', ""),
            ('SLDWORKS_CANDIDATES = [\n    r"E:\\Program Files\\SOLIDWORKS Corp\\SOLIDWORKS\\sldworks.exe",\n    r"C:\\Program Files\\SOLIDWORKS Corp\\SOLIDWORKS\\sldworks.exe",\n    r"D:\\Program Files\\SOLIDWORKS Corp\\SOLIDWORKS\\sldworks.exe",\n    "sldworks.exe",\n]\n', ""),
        ],
    },
}


def insert_bootstrap(text):
    if "shared/skill7_paths.py" in text and "from skill7_paths import" in text:
        return text, False
    lines = text.splitlines(keepends=True)
    last_import = -1
    for i, line in enumerate(lines[:100]):
        s = line.rstrip("\r\n")
        if s.startswith("import ") or s.startswith("from "):
            last_import = i
    if last_import < 0:
        raise RuntimeError("找不到 import 区，无法插入 bootstrap")
    lines.insert(last_import + 1, BOOTSTRAP)
    return "".join(lines), True


def main():
    check = "--check" in sys.argv
    problems = 0
    for rel, spec in FILES.items():
        path = os.path.join(SK, rel)
        if not os.path.isfile(path):
            print("!! 缺文件: %s" % rel)
            problems += 1
            continue
        with open(path, encoding="utf-8-sig") as f:
            text = f.read()

        for old, new, *rest in spec["edits"]:
            n = text.count(old)
            want_all = bool(rest and rest[0] == "all")
            if (want_all and n < 1) or (not want_all and n != 1):
                print("!! %s: 期望命中 %s 次，实际 %d 次 -> %r"
                      % (rel, "≥1" if want_all else "1", n, old[:70]))
                problems += 1
                continue
            text = text.replace(old, new) if want_all else text.replace(old, new, 1)

        text, inserted = insert_bootstrap(text)
        if not check:
            with open(path, "w", encoding="utf-8", newline="") as f:
                f.write(text)
        print("%s %s%s" % ("[改]" if not check else "[查]", rel, "  (已插 bootstrap)" if inserted else "  (bootstrap 已存在)"))

    print("\n问题数: %d" % problems)
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())