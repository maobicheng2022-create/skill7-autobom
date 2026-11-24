# -*- coding: utf-8 -*-
"""打 AutoBOM 一键安装包（ZIP），用于发给别人直接解压安装到 SolidWorks。

产物： dist/AutoBOM_SW_安装包_<版本>.zip
包内结构：
    AutoBOM_SW_安装包/
        安装AutoBOM.bat
        install_autobom.ps1
        说明.txt
        payload/AutoBOM/*
        payload/micDatas/pluginSettings/*
        payload/macros/*.swp

用法：
    python -X utf8 tools/make_autobom_package.py [版本号]
"""
import os
import sys
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PORTABLE = os.path.join(ROOT, "autobom", "portable")
INSTALLER = os.path.join(ROOT, "autobom", "installer")
TEMPLATES = os.path.join(ROOT, "templates", "micDatas", "pluginSettings")
MACROS = os.path.join(ROOT, "macros")
DIST = os.path.join(ROOT, "dist")

SKIP_NAMES = {".bak", "log"}


def _skip(rel):
    low = rel.lower()
    return ("_src" in low or "backup" in low or low.endswith(".bak")
            or low.startswith("log\\") or "\\log\\" in low or "\\_dxf_tmp" in low)


def add_tree(zf, base_dir, prefix, zip_root):
    """把 base_dir 下全部文件加进 zip 的 zip_root/prefix/..."""
    n = 0
    for dirpath, dirnames, filenames in os.walk(base_dir):
        rel_dir = os.path.relpath(dirpath, base_dir)
        if rel_dir == ".":
            rel_dir = ""
        for fn in sorted(filenames):
            rel = os.path.join(rel_dir, fn) if rel_dir else fn
            if _skip(rel):
                continue
            src = os.path.join(dirpath, fn)
            arc = "/".join([zip_root, prefix.replace("\\", "/"), rel.replace("\\", "/")])
            zf.write(src, arc)
            n += 1
    return n


def main():
    ver = sys.argv[1] if len(sys.argv) > 1 else "v1.0.0"
    pkg_name = "AutoBOM_SW_安装包_%s" % ver
    os.makedirs(DIST, exist_ok=True)
    out = os.path.join(DIST, pkg_name + ".zip")

    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        top = {
            "install_autobom.bat": "安装AutoBOM.bat",   # 中文名，方便双击
            "install_autobom.ps1": "install_autobom.ps1",
            "说明.txt": "说明.txt",
        }
        for src_name, arc_name in top.items():
            src = os.path.join(INSTALLER, src_name)
            if not os.path.isfile(src):
                raise SystemExit("缺少 %s" % src)
            zf.write(src, "/".join([pkg_name, arc_name]))
        n_auto = add_tree(zf, PORTABLE, "payload/AutoBOM", pkg_name)
        n_tpl = add_tree(zf, TEMPLATES, "payload/micDatas/pluginSettings", pkg_name)
        n_mac = add_tree(zf, MACROS, "payload/macros", pkg_name)

    size_mb = os.path.getsize(out) / 1024.0 / 1024.0
    print("已生成: %s" % out)
    print("  体积: %.1f MB" % size_mb)
    print("  AutoBOM %d 个文件 / 模板 %d / 宏 %d" % (n_auto, n_tpl, n_mac))


if __name__ == "__main__":
    main()