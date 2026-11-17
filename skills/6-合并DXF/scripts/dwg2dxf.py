# -*- coding: utf-8 -*-
"""
dwg -> dxf 转换 (AutoCAD COM SaveAs)
把文件夹内所有 .dwg 批量另存为 .dxf (临时目录或同目录)。
用到 AutoCAD, 但仅后台批处理。
用法:
  python -X utf8 dwg2dxf.py <folder> [输出目录]
"""
import os
import sys
import time
import glob
import traceback

import pythoncom
import win32com.client

ACAD_APP_PROGID = "AutoCAD.Application"
ACAD_EXE = r"D:\CAD\AutoCAD 2026\acad.exe"

# AutoCAD DXF 版本常量(来自 acad.h)
ACAD_R2018_DXF = 64   # DXF 2018 (R2018)
# 各种可用: 0=R12, 1=R13, ... 用 acDb2000_dxf etc. 实际用版本字符串更稳


def _listdir_ext(folder, ext):
    """列出目录下指定扩展名(小写)的文件, 按名称排序。
    2026-08-30 修复: glob.glob 在含中文目录路径下匹配失效(返回空),
    导致中文图纸目录里的 dwg 扫不到。用 os.listdir + endswith 替代。
    """
    ext = ext.lower()
    try:
        names = os.listdir(folder)
    except OSError:
        return []
    return sorted(os.path.join(folder, n) for n in names if n.lower().endswith(ext))


def wait_quiescent(app, timeout=120):
    t0 = time.time()
    while time.time() - t0 < timeout:
        try:
            if app.GetAcadState().IsQuiescent:
                return True
        except Exception:
            pass
        time.sleep(0.3)
    return False


def get_acad():
    try:
        app = win32com.client.GetActiveObject(ACAD_APP_PROGID)
        return app
    except Exception:
        os.startfile(ACAD_EXE, "open")
        for _ in range(120):
            time.sleep(1.0)
            try:
                return win32com.client.GetActiveObject(ACAD_APP_PROGID)
            except Exception:
                continue
        raise RuntimeError("AutoCAD 启动超时")


def dwg_to_dxf(app, dwg_path, dxf_path):
    """单个 dwg -> dxf。返回 True/False。"""
    doc = None
    for _ in range(40):
        try:
            if app.Documents.Count:
                break
        except Exception:
            time.sleep(0.5)
    try:
        app.Documents.Open(dwg_path, False)
    except Exception as e:
        print("  open err:", repr(e))
        return False
    # 拿刚打开的文档
    doc = None
    for _ in range(60):
        try:
            if app.Documents.Count:
                doc = app.Documents.Item(app.Documents.Count - 1)
                _ = doc.Name
                break
        except Exception:
            time.sleep(0.5)
    if doc is None:
        return False
    try:
        # DXF 版本: 实测 61 = ac2013_dxf (完整103KB), 1 = R12(16KB不完整)
        # 用 61 (DXF R2013) 最稳。独立子目录下 SaveAs 只产生目标 .dxf
        doc.SaveAs(dxf_path, 61)
        wait_quiescent(app, 30)
        # 关闭
        for _ in range(10):
            try:
                doc.Close(False)
                break
            except Exception:
                time.sleep(0.5)
        # 确认目标存在
        return os.path.exists(dxf_path)
    except Exception as e:
        print("  saveas err:", repr(e))
        try:
            doc.Close(False)
        except Exception:
            pass
        return False



def dwg2dxf_folder(app, dwg_list, outdir, delay=2.0):
    """把传入的 dwg 列表批量转成 dxf 到 outdir。返回 (ok, fail)。
    delay: 每个文件转换间延时秒, 避免快速开关文档导致 CAD 不稳定/崩溃。
    """
    if not os.path.isdir(outdir):
        os.makedirs(outdir)
    # 先清空 CAD 里的文档
    for _ in range(40):
        try:
            if app.Documents.Count == 0:
                break
            app.Documents.Item(0).Close(False)
        except Exception:
            time.sleep(0.5)
    ok, fail = 0, 0
    for i, d in enumerate(dwg_list):
        name = os.path.basename(d)
        dxf = os.path.join(outdir, os.path.splitext(name)[0] + ".dxf")
        if os.path.exists(dxf):
            print("skip(已有):", name, flush=True)
            ok += 1
            continue
        print("convert(%d/%d):" % (i + 1, len(dwg_list)), name, "...", end=" ", flush=True)
        try:
            r = dwg_to_dxf(app, d, dxf)
        except Exception as e:
            r = False
            print("EXC", repr(e))
        if r and os.path.exists(dxf):
            print("OK", flush=True)
            ok += 1
        else:
            print("FAIL")
            fail += 1
        time.sleep(delay)
    print("转换完成: 成功 %d, 失败 %d" % (ok, fail))
    return (ok, fail)


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    folder = sys.argv[1].rstrip("\\")
    # 默认转换输出到独立子目录, 避免和源 dwg 混在一起产生冗余文件
    outdir = sys.argv[2] if len(sys.argv) > 2 else os.path.join(folder, "_dxf_tmp")
    if not os.path.isdir(folder):
        raise SystemExit("folder not exist: %s" % folder)
    dwgs = _listdir_ext(folder, ".dwg")
    print("dwg to convert:", len(dwgs))
    if not dwgs:
        return
    app = get_acad()
    dwg2dxf_folder(app, dwgs, outdir)


if __name__ == "__main__":
    main()
