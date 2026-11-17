# -*- coding: utf-8 -*-
"""型材采购表加计算列（原地更新，可重复跑）。

对 型材采购.xlsx 的每个零件行，按"一根 6m 先尽量切满该零件"算：
  每套支数            = 总数量 ÷ 套数
  一根6m可开零件数    = ⌊6000 ÷ 零件长⌋
  一根6m可开套数(小数)= 可开零件数 ÷ 每套支数
  余料mm              = 6000 − 可开零件数 × 零件长

例：零件长 424、每套 2 支 → 一根切 14 段（14×424=5936），余 64mm，
    可开 14÷2 = 7 套（不是 7.5；15 段要 6360mm，超 6m）。
"""
import argparse
import math
import os
import re
import sys

import openpyxl
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

BAR_LEN = 6000.0
NEW_COLS = ['每套支数', '一根6m可开零件数', '一根6m可开套数(小数)', '余料mm']


def detect_sets(xlsx_path):
    """从订单文件夹名（含上级 2 层）里的 `做N套` 推断套数；推断不到按 1 套。"""
    d = os.path.dirname(os.path.abspath(xlsx_path))
    for _ in range(3):
        m = re.search(r'做\s*(\d+)\s*套', os.path.basename(d) or '')
        if m:
            n = int(m.group(1))
            if n >= 1:
                return n
        parent = os.path.dirname(d)
        if parent == d:
            break
        d = parent
    return 1


def safe_float(v):
    if v is None:
        return 0.0
    s = str(v).replace('，', '.').replace(',', '.').strip()
    if s == '':
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def num(v):
    if v is None:
        return None
    v = float(v)
    return int(round(v)) if abs(v - round(v)) < 1e-9 else round(v, 2)


def main():
    ap = argparse.ArgumentParser(description='型材采购表加列：每套支数/一根6m可开零件数/可开套数(小数)/余料mm')
    ap.add_argument('xlsx')
    ap.add_argument('--sets', type=int, default=0,
                    help='本订单做几套（默认从文件夹名 做N套 推断，推断不到按 1）')
    args = ap.parse_args()

    if not os.path.isfile(args.xlsx):
        sys.exit(f'[错误] 找不到文件: {args.xlsx}')

    sets = args.sets if args.sets and args.sets > 0 else detect_sets(args.xlsx)
    if sets < 1:
        sets = 1

    wb = openpyxl.load_workbook(args.xlsx)
    ws = wb.worksheets[0]
    headers = [str(c.value).strip().lstrip('$') if c.value is not None else '' for c in ws[1]]

    def find(*names):
        for i, h in enumerate(headers):
            if h in names:
                return i + 1
        return None

    i_qty = find('总数量', '数量', 'qty', 'num')
    i_len = find('长', '长度', 'length')
    i_name = find('名称', 'name')
    if i_qty is None or i_len is None:
        sys.exit(f'[错误] 缺必需列（数量/长）。表头: {headers}')

    # 新列：已有就复用，没有就追加到末尾
    col = {}
    next_col = ws.max_column + 1
    for name in NEW_COLS:
        if name in headers:
            col[name] = headers.index(name) + 1
        else:
            col[name] = next_col
            next_col += 1
            c = ws.cell(1, col[name], name)
            c.font = Font(name='Calibri', size=11)
            c.alignment = Alignment(horizontal='center', vertical='center')
            ws.column_dimensions[get_column_letter(col[name])].width = 20 if '套数' in name else 16

    print('=' * 96)
    print(f"输入: {os.path.basename(args.xlsx)}   本单 {sets} 套   一根管 = {int(BAR_LEN)}mm")
    print('=' * 96)
    print(f"{'零件':<28}{'长mm':>8}{'每套支数':>9}{'可开零件':>9}{'可开套数':>10}{'余料mm':>9}")
    print('-' * 96)

    warns = []
    for r in range(2, ws.max_row + 1):
        qty = safe_float(ws.cell(r, i_qty).value)
        length = safe_float(ws.cell(r, i_len).value)
        name = str(ws.cell(r, i_name).value) if i_name else f'第{r}行'

        per_set_qty = qty / sets
        if abs(per_set_qty - round(per_set_qty)) > 1e-6 and qty > 0:
            warns.append(f'{name}: 总数量 {qty:g} 不是 {sets} 套的整数倍')

        if length > 0:
            pieces = int(math.floor((BAR_LEN + 1e-6) / length))
        else:
            pieces = 0

        if length > BAR_LEN + 1e-9:
            sets_per_bar = 0
            leftover = None
            warns.append(f'{name}: 单段 {length:g}mm 超过 6m，单根装不下')
        elif pieces > 0 and per_set_qty > 0:
            sets_per_bar = pieces / per_set_qty
            leftover = BAR_LEN - pieces * length
        else:
            sets_per_bar = 0
            leftover = None

        vals = {
            '每套支数': num(per_set_qty),
            '一根6m可开零件数': pieces,
            '一根6m可开套数(小数)': None if sets_per_bar is None else num(sets_per_bar),
            '余料mm': None if leftover is None else round(leftover, 1),
        }
        for name_c, v in vals.items():
            c = ws.cell(r, col[name_c], v)
            c.alignment = Alignment(horizontal='center', vertical='center')
            if name_c == '一根6m可开套数(小数)':
                c.number_format = '0.##'
            elif name_c == '余料mm':
                c.number_format = '0.#'

        spb = '-' if vals['一根6m可开套数(小数)'] is None else f"{vals['一根6m可开套数(小数)']:g}"
        left = '-' if vals['余料mm'] is None else f"{vals['余料mm']:g}"
        print(f"{name:<28}{length:>8.1f}{vals['每套支数'] if vals['每套支数'] is not None else '-':>9}"
              f"{pieces:>9}{spb:>10}{left:>9}")

    wb.save(args.xlsx)
    print('-' * 96)
    print(f"已写: {args.xlsx}")
    if warns:
        print('\n[提醒]')
        for w in warns[:10]:
            print(' -', w)


if __name__ == '__main__':
    main()