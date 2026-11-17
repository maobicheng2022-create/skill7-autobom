# -*- coding: utf-8 -*-
"""
dwg-merge 主控脚本 (纯Python合并 + AutoCAD COM 转档)
====================================================
把目标文件夹内多份零件图纸(.dwg) + 配料表(.csv) 合并生成一张 DXF,
按厚度分层排布, 每件标注"材料-厚度-做N个", 数量取自 csv。

流程:
  1. 扫描文件夹: .dwg 列表 + .csv (取第一个作配料表)
  2. dwg -> dxf: 用 AutoCAD COM SaveAs(版本61) 批量转档到 _dxf_tmp/
     (若文件夹已有 .dxf 且无对应的 .dwg, 则直接复用; 优先用 dxf 源)
  3. 合并: ezdxf 读所有 dxf -> 解析文件名(代号_名称-材料-厚度) -> 查 csv 数量
     -> 按厚度分层排布 -> 插块 + 标注 -> 输出 合并图纸_日期.dxf
  4. 报告: 输出 合并报告.txt (成功数/分层/空文件错误日志/跳过原因)

输出(都在目标文件夹):
  - 合并图纸_YYYYMMDD_HHMMSS.dxf
  - 合并报告.txt
  - _dxf_tmp/  (转换中间产物, 保留以备复用, 可手动删)

用法:
  python -X utf8 dwg_merge.py <文件夹路径>
"""
import os
import sys
import glob
import shutil
from datetime import datetime

# 引入同目录模块
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from dwg2dxf import dwg2dxf_folder, get_acad
from merge_dxf_core import merge_dxf, write_report


def _listdir_ext(folder, ext):
    """列出目录下指定扩展名(小写)的文件, 按名称排序。
    2026-08-30 修复: glob.glob 在含中文的目录路径下对匹配失效(返回空),
    导致焊接组子目录(中文组名)的 csv/dxf 找不到, 合并图纸永远生成不出。
    用 os.listdir + endswith 替代, 对中文目录稳定。
    """
    ext = ext.lower()
    try:
        names = os.listdir(folder)
    except OSError:
        return []
    return sorted(os.path.join(folder, n) for n in names if n.lower().endswith(ext))


def _find_csv(folder):
    cs = _listdir_ext(folder, ".csv")
    return cs[0] if cs else None


def cleanup_after(app):
    """可选: 关闭 AutoCAD 文档(不影响已运行的CAD)。"""
    try:
        while app.Documents.Count:
            try:
                app.Documents.Item(0).Close(False)
            except Exception:
                break
    except Exception:
        pass


def run(folder):
    folder = folder.rstrip("\\")
    if not os.path.isdir(folder):
        raise SystemExit(u"文件夹不存在: %s" % folder)

    csv_path = _find_csv(folder)
    if not csv_path:
        raise SystemExit(u"文件夹里没有 .csv 配料表文件: %s" % folder)
    print(u"[1/3] 配料表: %s" % os.path.basename(csv_path))

    # --- 2. 确保有 dxf(转换或复用) ---
    dxf_dir = os.path.join(folder, "_dxf_tmp")
    dwgs = _listdir_ext(folder, ".dwg")
    dxfs_reuse = _listdir_ext(folder, ".dxf")

    # 需要转换的 dwg: 文件夹里的 dwg(排除下料/合并类命名, 那些往往无"_-材料-厚度"结构)
    to_convert = []
    for d in dwgs:
        base = os.path.splitext(os.path.basename(d))[0]
        # 跳过明显的输出/备份文件
        if base.startswith("_") or base.lower() in ("下料", "下料 "):
            continue
        to_convert.append(d)

    if not os.path.isdir(dxf_dir):
        os.makedirs(dxf_dir)

    print(u"[2/3] 转换 dwg -> dxf (共 %d 个)..." % len(to_convert))
    app = get_acad()
    try:
        synthesized = dwg2dxf_folder(app, to_convert, dxf_dir)
    finally:
        cleanup_after(app)

    # 合并时扫描 dxf 源 = 目录内已有 dxf + 转换生成的 dxf
    # 为统一, 把源文件夹里符合命名规范的 dxf 也复制进 _dxf_tmp 以便一起合并
    for dx in dxfs_reuse:
        target = os.path.join(dxf_dir, os.path.basename(dx))
        if not os.path.exists(target):
            try:
                shutil.copy2(dx, target)
            except Exception:
                pass

    # --- 3. 合并 ---
    out_name = "合并图纸_" + datetime.now().strftime("%Y%m%d_%H%M%S") + ".dxf"
    out_path = os.path.join(folder, out_name)
    report_path = os.path.join(folder, "合并报告.txt")
    print(u"[3/3] 合并 dxf -> %s" % out_name)
    result = merge_dxf(dxf_dir, csv_path, out_path)
    lines = write_report(result, report_path)

    # --- 摘要输出(agent 读stdout/返回) ---
    print()
    print(u"===== 合并完成 =====")
    for l in lines:
        print(l)
    print(u"输出文件: %s" % out_path)
    print(u"报告文件: %s" % report_path)
    return out_path


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    run(sys.argv[1])
