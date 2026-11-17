# -*- coding: utf-8 -*-
"""
ezdxf 核心合并模块: 读多个零件 dxf -> 按 csv 数量 + 厚度分层 -> 插块 + 标注 -> 输出合并 dxf
不依赖 AutoCAD。源 dwg 需先由外部转成 dxf。
"""
import os
import re
import csv
import glob
from datetime import datetime

import ezdxf

# 排布参数(横向间距已从 1500 缩小到 800, 既分隔又不浪费空间)
SPACING = 800.0       # 横向间距
ROW_GAP = 600.0       # 换行多留空间
TXT_HEIGHT = 50.0     # 标注文字高度
ROW_SPACING_MULT = 4.0 # 行高额外倍数

# 标注文字样式(用 TrueType 宋体, 避免 AutoCAD 打开时 Standard 样式中文乱码)
LABEL_STYLE = "MERGE_SONGTI"
LABEL_FONT = "simsun.ttf"


def read_csv_map(path):
    """读 csv: 第一列代号, 第二列数量。返回 {代号: 数量}。"""
    m = {}
    with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
        rdr = csv.reader(f)
        rows = list(rdr)
    # 跳过表头(含非数字第二列的)
    for row in rows:
        if len(row) < 2:
            continue
        code = row[0].strip()
        qty_s = row[1].strip()
        if not code:
            continue
        try:
            qty = int(float(qty_s))
        except ValueError:
            continue  # 表头行如 '$总数量' 会被跳过
        m[code] = qty
    return m


def get_code(fname):
    """取 '_' 前部分作代号。"""
    pos = fname.find("_")
    if pos >= 0:
        return fname[:pos]
    return fname.rsplit(".", 1)[0]  # 无下划线则去扩展名


def parse_filename(fname):
    """
    解析 `代号_名称-材料-厚度.dxf` -> (代号, 名称, 材料, 厚度).
    兼容: 代号_名称 之间可能空格(空格也是下划线变体, 取第一个空格/下划线前作代号)
    """
    base = os.path.splitext(os.path.basename(fname))[0]
    parts = base.split("-")
    if len(parts) < 3:
        return None
    # parts[0] 可能是 '代号_名称' 或 '代号 名称'
    first = parts[0]
    fpart = re.split(r"[_\s]", first, maxsplit=1)
    code = fpart[0]
    name = fpart[1] if len(fpart) > 1 else ""
    material = parts[1]
    thickness = parts[2]
    return (code, name, material, thickness)


def thickness_to_num(th):
    """提取厚度字符串里的数字。"""
    nums = re.sub(r"[^0-9.]", "", th)
    try:
        return float(nums) if nums else 0.0
    except ValueError:
        return 0.0


def dxf_bbox(msp):
    """用 ezdxf.bbox 计算边界, 支持 LINE/ARC/INSERT/TEXT 等。返回 (minx,miny,maxx,maxy) 或 None。"""
    from ezdxf import bbox as _bbox
    ents = list(msp)
    if not ents:
        return None
    try:
        box = _bbox.extents(ents, fast=True)
    except Exception:
        try:
            box = _bbox.extents(ents)
        except Exception:
            return None
    if box is None or box.has_data is False:
        return None
    try:
        mn = box.extmin
        mx = box.extmax
    except Exception:
        return None
    return (mn[0], mn[1], mx[0], mx[1])


def load_part_dxf(path):
    """读 dxf, 返回 (doc, msp, bbox)。空/无法读取返回 None。"""
    try:
        doc = ezdxf.readfile(path)
        msp = doc.modelspace()
    except Exception:
        return None
    ents = [e for e in msp]
    if not ents:
        return None
    bbox = dxf_bbox(msp)
    if bbox is None:
        return None
    return (doc, msp, bbox)


def txt_width(text, height):
    """粗估文字宽度(等宽近似)。"""
    return len(text) * height * 0.8


def _listdir_ext(folder, ext):
    """列出目录下指定扩展名(小写)的文件, 按名称排序。
    2026-08-30 修复: glob.glob 在含中文的目录路径下匹配失效(返回空),
    导致焊接组子目录(中文组名)里的 dxf/csv 扫不到, 合并图纸成功插入 0。
    """
    ext = ext.lower()
    try:
        names = os.listdir(folder)
    except OSError:
        return []
    return sorted(os.path.join(folder, n) for n in names if n.lower().endswith(ext))


def merge_dxf(folder, csv_path, out_path):
    """
    合并 folder 下所有 dxf。返回报告 dict。
    """
    # 排除输出/报告类文件(match 合并图纸_ 前缀 或 报告), 避免自扫进合并结果
    dxfs = []
    for f in _listdir_ext(folder, ".dxf"):
        base = os.path.basename(f).lower()
        if base.startswith("合并图纸_") or base.startswith("_"):
            continue
        dxfs.append(f)
    csvmap = read_csv_map(csv_path)

    result = {
        "total_dxf": len(dxfs),
        "inserted": 0,
        "skipped_empty": [],   # [(文件名, 原因)]
        "skipped_noqty": [],   # [(文件名, dxf代号, 疑似正确代号或None)]
        "skipped_parse": [],   # 文件名解析失败
        "missing_in_dxf": [],  # 2026-08-13: CSV 有配料但文件夹里没对应 dxf(如展开失败) -> [代号]
        "layers": {},          # 厚度 -> 零件数
    }

    # 先解析所有有效零件
    parts = []  # (th_num, th_str, material, quantity, doc_src, bbox, fname)
    all_dxf_codes = set()  # 2026-08-13: 所有 dxf 解析出的代号(含跳过的), 反查 CSV 缺件
    for f in dxfs:
        name = os.path.basename(f)
        parsed = parse_filename(name)
        if parsed is None:
            result["skipped_parse"].append(name)
            continue
        code, pname, material, thickness = parsed
        all_dxf_codes.add(code)
        qty = csvmap.get(code)
        if qty is None or qty <= 0:
            # 2026-08-13: 严格按图号合并 — 匹配不上就是图号没更新(skill1 没跑/没成功), 不猜疑似代号
            result["skipped_noqty"].append((name, code))
            continue
        loaded = load_part_dxf(f)
        if loaded is None:
            result["skipped_empty"].append((name, "dxf无实体或无边界"))
            continue
        part_doc, part_msp, bbox = loaded
        parts.append((thickness_to_num(thickness), thickness, material, qty,
                      part_doc, part_msp, bbox, name))
        result["layers"].setdefault(thickness, 0)
        result["layers"][thickness] += 1

    # 2026-08-13: 反查 — CSV 有配料但没有任何 dxf 对应(如钣金展开失败/文件缺失)
    for code in csvmap:
        if code not in all_dxf_codes:
            result["missing_in_dxf"].append(code)

    # 按厚度排序 -> 薄到厚
    parts.sort(key=lambda p: p[0])

    # 建输出文档
    out_doc = ezdxf.new("R2010", setup=True)
    out_msp = out_doc.modelspace()

    # 创建标注文字样式(宋体 TrueType, 支持中文显示)
    try:
        out_doc.styles.new(LABEL_STYLE, dxfattribs={"font": LABEL_FONT})
    except Exception:
        pass

    # 第一步: 把所有源图实体导入为块定义(用 Importer), 便于按位置 add_blockref
    block_names = {}  # fname -> blockname
    for idx, (th_num, th_str, material, qty, part_doc, part_msp, bbox, fname) in enumerate(parts):
        bn = "BLK_%03d" % idx
        blk = out_doc.blocks.new(name=bn)
        from ezdxf.addons import Importer
        imp = Importer(part_doc, out_doc)
        imp.import_entities(part_msp, target_layout=blk)
        imp.finalize()
        block_names[fname] = bn

    # 第二步: 按厚度分层排布插入块 + 标注
    cur_y = 0.0
    cur_th = None
    shift_x = 0.0
    row_max_h = 0.0

    for (th_num, th_str, material, qty, part_doc, part_msp, bbox, fname) in parts:
        minx, miny, maxx, maxy = bbox
        w = maxx - minx
        h = maxy - miny

        # 换厚度 -> 换行
        if cur_th is not None and th_num != cur_th:
            cur_y -= (row_max_h + ROW_GAP)
            shift_x = 0.0
            row_max_h = 0.0
        cur_th = th_num

        bn = block_names[fname]
        # 插入点: 块原点即源图 min 点, 平移到 (shift_x, cur_y) => insert = (shift_x - minx, cur_y - miny)
        insert = (shift_x - minx, cur_y - miny, 0)
        out_msp.add_blockref(bn, insert)

        # 标注文字: 底部中心下移, 用宋体样式避免中文乱码
        label = "%s-%s-做%d个" % (material, th_str, qty)
        tx = shift_x + w / 2
        ty = cur_y - TXT_HEIGHT * 1.5
        txt = out_msp.add_text(label, dxfattribs={"height": TXT_HEIGHT, "style": LABEL_STYLE})
        txt.set_placement((tx, ty, 0))

        # 更新行高/横移
        need_h = h + TXT_HEIGHT * ROW_SPACING_MULT
        if need_h > row_max_h:
            row_max_h = need_h
        shift_x += w + SPACING
        result["inserted"] += 1

    # 修正 DIMLFAC: ezdxf 合并时默认标注样式的 dimlfac=100, 会导致 CAD 测距/标注值放大 100 倍
    # 统一重置 header $DIMLFAC 与所有 DIMSTYLE 的 dimlfac 为 1.0 (治本, 2026-08-09)
    _reset_dimlfac(out_doc)

    # 保存
    out_doc.saveas(out_path)
    result["out_path"] = out_path
    result["layers_sorted"] = sorted(result["layers"].keys(),
                                     key=lambda t: thickness_to_num(t))
    return result


def _reset_dimlfac(doc):
    """把 header $DIMLFAC 和所有 DIMSTYLE 的 dimlfac 重置为 1.0, 避免 CAD 测距放大。"""
    try:
        hd = doc.header
        cur = hd.get("$DIMLFAC")
        if cur is not None and abs(cur - 1.0) > 1e-9:
            hd["$DIMLFAC"] = 1.0
    except Exception:
        pass
    if doc.dimstyles:
        for s in doc.dimstyles:
            try:
                cur = s.dxf.dimlfac
                if cur is not None and abs(cur - 1.0) > 1e-9:
                    s.dxf.dimlfac = 1.0
            except Exception:
                pass


def write_report(result, report_path):
    """写合并报告 txt(agent 可读回)。"""
    lines = []
    lines.append("=== 合并报告 %s ===" % datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    lines.append("扫描 dxf 总数: %d" % result["total_dxf"])
    lines.append("成功插入: %d" % result["inserted"])
    lines.append("厚度分层: %s" % ", ".join(result["layers_sorted"]))
    if result["skipped_empty"]:
        lines.append("")
        lines.append("[错误] 空文件(无图形, 已跳过):")
        for n, r in result["skipped_empty"]:
            lines.append("  - %s  (%s)" % (n, r))
    if result["skipped_noqty"]:
        lines.append("")
        lines.append("[警告] 代号不在CSV, 已跳过 — 无法合并说明图号未更新(先跑 skill1 图号更新再重导CSV/重合并):")
        for n, code in result["skipped_noqty"]:
            lines.append("  - %s  (dxf代号=%s)" % (n, code))
    if result["missing_in_dxf"]:
        lines.append("")
        lines.append("[警告] CSV有配料但无对应dxf (展开失败/文件缺失, 产线图会缺件):")
        for c in result["missing_in_dxf"]:
            lines.append("  - %s" % c)
    if result["skipped_parse"]:
        lines.append("")
        lines.append("[警告] 文件名无法解析, 已跳过:")
        for n in result["skipped_parse"]:
            lines.append("  - %s" % n)
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return lines


if __name__ == "__main__":
    import sys
    folder = sys.argv[1] if len(sys.argv) > 1 else r"C:\Users\mao\Desktop\桌面整理_20260727\CAD图纸"
    csv_path = sys.argv[2] if len(sys.argv) > 2 else ""
    out = sys.argv[3] if len(sys.argv) > 3 else os.path.join(folder, "_合并图纸_test.dxf")
    if not csv_path or not os.path.exists(csv_path):
        # 找文件夹里的 csv
        cs = _listdir_ext(folder, ".csv")
        csv_path = cs[0] if cs else ""
    if not csv_path:
        print("无 csv, 仅测试读取 dxf")
        dxfs = _listdir_ext(folder, ".dxf")
        print("dxf 文件:", len(dxfs))
        for f in dxfs[:20]:
            l = load_part_dxf(f)
            print(" ", os.path.basename(f)[:50], "->", ("有图形" if l else "空/无法读"))
        sys.exit(0)
    r = merge_dxf(folder, csv_path, out)
    for l in write_report(r, os.path.join(folder, "_合并报告_test.txt")):
        print(l)
    print("输出:", r["out_path"])
