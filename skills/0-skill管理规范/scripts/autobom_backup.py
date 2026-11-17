# -*- coding: utf-8 -*-
"""
autobom_backup.py — AutoBOM 版本更新前必备份工具（2026-08-06 钦定铁律）

用法:
  python -X utf8 autobom_backup.py            # 备份当前 AutoBOM 全套到带时间戳快照
  python -X utf8 autobom_backup.py --list     # 列出已有快照
  python -X utf8 autobom_backup.py --restore <快照名>    # 从指定快照回滚（先备份当前，再覆盖）

备份内容:
  C:\AutoBOM\ 下的 exe / dll / config(json) /
  C:\micDatas\pluginSettings\BOM\setting.json / CommonNode.json / SpecialNode.json / Template\
"""
import argparse
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

CONFIG_DIR = os.path.join(MICDATAS_DIR, "pluginSettings", "BOM")
BACKUP_ROOT = os.path.join(AUTOBOM_DIR, "_backup")
TS = time.strftime("%Y%m%d_%H%M%S")

# 要备份的 AutoBOM 根目录条目（文件）+ 整个 Plugin 很重要（含 addin/logo/插件配置）
AUTOBOM_FILES = [
    "AutoBOM.exe", "AutoBOM.exe.config",
    "AutoBOM.SW.dll", "AutoBOM.SW.dll.config",
    "AutoBaiJinZhanKai.exe", "AutoBaiJinZhanKai.exe.bak",
    "Addin_Config.json",
]
AUTOBOM_DIRS = ["Plugin", "icon"]
# 配置目录条目
CONFIG_FILES = ["setting.json", "CommonNode.json", "SpecialNode.json"]
CONFIG_DIRS = ["Template"]


def _copy_into(src, dst_dir):
    """复制文件或目录到目标目录下，返回是否成功"""
    if os.path.isfile(src):
        os.makedirs(dst_dir, exist_ok=True)
        shutil.copy2(src, os.path.join(dst_dir, os.path.basename(src)))
        return True
    if os.path.isdir(src):
        dst = os.path.join(dst_dir, os.path.basename(src))
        if os.path.exists(dst):
            shutil.rmtree(dst)
        shutil.copytree(src, dst)
        return True
    return False


def snapshot_name():
    return "backup_" + TS


def do_backup():
    name = snapshot_name()
    root = os.path.join(BACKUP_ROOT, name)
    os.makedirs(root, exist_ok=True)

    # 1) AutoBOM 根（扁平：文件直接放, Plugin/icon 放子目录）
    for f in AUTOBOM_FILES:
        _copy_into(os.path.join(AUTOBOM_DIR, f), root)
    for d in AUTOBOM_DIRS:
        _copy_into(os.path.join(AUTOBOM_DIR, d), root)

    # 2) AutoBOM 配置
    cfg_dst = os.path.join(root, "_config_BOM")
    for f in CONFIG_FILES:
        _copy_into(os.path.join(CONFIG_DIR, f), cfg_dst)
    for d in CONFIG_DIRS:
        _copy_into(os.path.join(CONFIG_DIR, d), cfg_dst)

    # 3) 备份源码工程（放到 _src_工程 子目录避免与扁平 exe 文件同名冲突）
    src_dst = os.path.join(root, "_src_工程")
    s = os.path.join(AUTOBOM_DIR, "_src", "AutoBOM.exe")
    if os.path.isdir(s):
        dst = os.path.join(src_dst, os.path.basename(s))
        if os.path.exists(dst):
            shutil.rmtree(dst)
        shutil.copytree(s, dst)

    print("✅ AutoBOM 备份完成:", root)
    # 回显关键文件清单确认
    for base, dirs, files in os.walk(root):
        for f in files:
            print("   ", os.path.join(base, f))
    return root


def list_backups():
    if not os.path.isdir(BACKUP_ROOT):
        print("无备份目录:", BACKUP_ROOT)
        return
    items = [os.path.join(BACKUP_ROOT, n) for n in os.listdir(BACKUP_ROOT)
             if os.path.isdir(os.path.join(BACKUP_ROOT, n))]
    items.sort(key=os.path.getmtime, reverse=True)
    print("已有 AutoBOM 快照（新→旧）:")
    for p in items:
        m = time.strftime("%Y-%m-%d %H:%M", time.localtime(os.path.getmtime(p)))
        # 统计大小
        sz = 0
        for b, d, fs in os.walk(p):
            for f in fs:
                sz += os.path.getsize(os.path.join(b, f))
        print(f"   {m}  {sz/1024:.0f}KB  {os.path.basename(p)}")


def do_restore(name):
    src = os.path.join(BACKUP_ROOT, name) if not os.path.isabs(name) else name
    if not os.path.isdir(src):
        print(f"❌ 快照不存在: {src}")
        sys.exit(1)
    # 备份当前，再覆盖
    print("先备份当前状态...")
    pre = do_backup()
    print("当前已备份:", pre)

    # 复制快照内容回 C:\AutoBOM\ 根
    for item in os.listdir(src):
        if item == "_config_BOM":
            continue
        s = os.path.join(src, item)
        _copy_into(s, AUTOBOM_DIR)
    # 配置
    cfg_src = os.path.join(src, "_config_BOM")
    if os.path.isdir(cfg_src):
        for item in os.listdir(cfg_src):
            _copy_into(os.path.join(cfg_src, item), CONFIG_DIR_FOR_CFG())
    print("✅ 已从快照回滚:", src)


def CONFIG_DIR_FOR_CFG():
    return CONFIG_DIR


def main():
    ap = argparse.ArgumentParser(description="AutoBOM 版本更新前备份/回滚")
    ap.add_argument("--list", action="store_true", help="列出已有快照")
    ap.add_argument("--restore", help="从指定快照名回滚")
    args = ap.parse_args()

    if args.list:
        list_backups()
        return
    if args.restore:
        do_restore(args.restore)
        return
    do_backup()


if __name__ == "__main__":
    main()
