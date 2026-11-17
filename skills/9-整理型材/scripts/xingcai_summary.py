# -*- coding: utf-8 -*-
"""整理型材 - 从型材采购表归并规格，折算 6m 管采购条数，
并给出每根 6m 管最多能开几整套（按本订单每套用量）及每根余料。
"""
import argparse
import math
import os
import re
import sys
from collections import OrderedDict

import openpyxl

BAR_LEN = 6000.0  # 一根 6m = 6000mm


def load_rows(path):
    """读取 excel，返回 (headers, rows)。表头可能带 $ 前缀，统一去掉 $ 归一。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        sys.exit('[\u9519\u8bef] \u8868\u683c\u4e3a\u7a7a')
    headers = [str(c).lstrip('$').strip() if c is not None else '' for c in rows[0]]
    data = []
    for r in rows[1:]:
        if all(c is None or str(c).strip() == '' for c in r):
            continue
        data.append(r)
    return headers, data


def normalize(h):
    return (h or '').replace('\uff0c', '.').replace(',', '.').strip()


def safe_float(v):
    if v is None:
        return 0.0
    s = normalize(str(v))
    if s == '':
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def safe_int(v):
    return int(round(safe_float(v)))


def col_idx(headers, *names):
    for h in headers:
        for n in names:
            if h == n:
                return headers.index(h)
    return None


def detect_sets(xlsx_path):
    """从订单文件夹名（含上级 2 层）里的 `做N套` 推断套数；推断不到按 1 套。"""
    d = os.path.dirname(os.path.abspath(xlsx_path))
    for _ in range(3):
        m = re.search(r'\u505a\s*(\d+)\s*\u5957', os.path.basename(d) or '')
        if m:
            n = int(m.group(1))
            if n >= 1:
                return n
        parent = os.path.dirname(d)
        if parent == d:
            break
        d = parent
    return 1


def build_key(htype, w, h, t):
    """归一规格 key。方/长方管：长宽取 lo/hi 顺序无关 + 厚度；圆管：直径 + 厚度。"""
    htype = (htype or '').strip()
    if '\u5706\u7ba1' in htype:
        dia = safe_float(w) or safe_float(h)
        return ('\u5706\u7ba1', round(dia, 3), round(safe_float(t), 3))
    a, b = safe_float(w), safe_float(h)
    lo, hi = min(a, b), max(a, b)
    return (htype or '\u578b\u6750', round(lo, 3), round(hi, 3), round(safe_float(t), 3))


def s(x):
    x = float(x)
    return str(int(round(x))) if abs(x - round(x)) < 1e-9 else str(round(x, 2))


def num(n):
    """2 位小数内取整显示：4.0 -> 4，4.5 -> 4.5。"""
    n = float(n)
    return int(round(n)) if abs(n - round(n)) < 1e-9 else round(n, 2)


def fmt_key(key):
    kind = key[0]
    t = key[-1]
    if kind == '\u5706\u7ba1':
        return f'\u5706\u7ba1 \u03c6{s(key[1])} \u539a{s(t)}'
    return f'{kind} {s(key[1])}\u00d7{s(key[2])} \u539a{s(t)}'


def ffd_pack(items):
    """First-Fit Decreasing 一维下料（近似最优）。返回桶列表，每桶为段长列表。"""
    bins = []
    for seg in sorted(items, key=lambda v: -v):
        for b in bins:
            if sum(b) + seg <= BAR_LEN + 1e-6:
                b.append(seg)
                break
        else:
            bins.append([seg])
    return bins


def main():
    ap = argparse.ArgumentParser(
        description='\u6574\u7406\u578b\u6750\uff1a\u5f52\u5e76\u89c4\u683c\u3001\u6298\u7b97 6m \u7ba1\u91c7\u8d2d\u6761\u6570\uff0c'
                    '\u5e76\u7b97\u6bcf\u6839 6m \u7ba1\u6700\u591a\u80fd\u5f00\u51e0\u6574\u5957 + \u6bcf\u6839\u4f59\u6599')
    ap.add_argument('xlsx')
    ap.add_argument('-o', '--out', help='\u8f93\u51fa xlsx \u8def\u5f84\uff08\u9ed8\u8ba4 \u539f\u76ee\u5f55/\u6574\u7406\u578b\u6750_YYYYMMDD_HHMMSS.xlsx\uff09')
    ap.add_argument('--sets', type=int, default=0,
                    help='\u672c\u8ba2\u5355\u505a\u51e0\u5957\uff08\u9ed8\u8ba4\u4ece\u6587\u4ef6\u5939\u540d'
                         ' \u505aN\u5957 \u63a8\u65ad\uff0c\u63a8\u4e0d\u5230\u6309 1\uff09')
    ap.add_argument('--optimize', action='store_true',
                    help='\u517c\u5bb9\u4fdd\u7559\uff08\u9ed8\u8ba4\u5df2\u6309 FFD \u5b9e\u9645\u62fc\u7ba1\uff09')
    args = ap.parse_args()

    if not os.path.isfile(args.xlsx):
        sys.exit(f'[\u9519\u8bef] \u627e\u4e0d\u5230\u6587\u4ef6: {args.xlsx}')

    sets = args.sets if args.sets and args.sets > 0 else detect_sets(args.xlsx)
    if sets < 1:
        sets = 1

    headers, rows = load_rows(args.xlsx)
    i_name = col_idx(headers, '\u540d\u79f0', 'name')
    i_qty = col_idx(headers, '\u603b\u6570\u91cf', '\u6570\u91cf', 'qty', 'num')
    i_w = col_idx(headers, '\u5bbd', 'width')
    i_h = col_idx(headers, '\u9ad8', 'height')
    i_t = col_idx(headers, '\u539a\u5ea6', '\u539a', 'thickness')
    i_type = col_idx(headers, '\u578b\u6750\u7c7b\u578b', '\u7c7b\u578b', 'type')
    i_len = col_idx(headers, '\u957f', '\u957f\u5ea6', 'length')
    for nm, idx in [('\u6570\u91cf', i_qty), ('\u5bbd\u5ea6', i_w), ('\u9ad8\u5ea6', i_h), ('\u539a\u5ea6', i_t), ('\u957f\u5ea6', i_len)]:
        if idx is None:
            sys.exit(f'[\u9519\u8bef] \u8868\u4e2d\u7f3a\u5fc5\u9700\u5217: {nm}\uff08\u8868\u5934: {headers}\uff09')

    groups = OrderedDict()
    rows_by_key = OrderedDict()
    skipped = []
    for r in rows:
        try:
            qty = safe_float(r[i_qty])
            length = safe_float(r[i_len])
            wv = r[i_w] if i_w < len(r) else None
            hv = r[i_h] if i_h < len(r) else None
            tv = r[i_t] if i_t < len(r) else None
            typev = r[i_type] if i_type is not None and i_type < len(r) else None
            key = build_key(typev if typev else '', wv, hv, tv)
            if key not in groups:
                groups[key] = {'total_len': 0.0, 'qty': 0, 'names': set()}
                rows_by_key[key] = []
            g = groups[key]
            g['total_len'] += length * qty
            g['qty'] += safe_int(qty)
            nm = r[i_name] if i_name is not None and i_name < len(r) else ''
            if nm:
                g['names'].add(str(nm))
            rows_by_key[key].append((length, safe_int(qty), str(nm)))
        except Exception as e:
            skipped.append((r, str(e)))

    print('=' * 104)
    print(f"\u8f93\u5165: {os.path.basename(args.xlsx)}    \u4e00\u6839\u7ba1 = {int(BAR_LEN)}mm ({BAR_LEN/1000:.1f}m)"
          f"    \u672c\u5355 {sets} \u5957")
    print('=' * 104)
    print(f"{'\u89c4\u683c':<30}{'\u7c7b\u578b':<6}{'\u6570\u91cf':>6}{'\u6bcf\u5957\u652f\u6570':>9}{'\u603b\u957fmm':>10}"
          f"{'\u6bcf\u5957mm':>10}{'6m\u53ef\u5f00\u5957':>10}{'\u6bcf\u6839\u4f59\u6599':>10}{'6m\u6839\u6570':>7}{'\u5269\u4f59mm':>9}")
    print('-' * 104)

    out_rows = []
    uneven = []
    oversize = []
    for key, g in groups.items():
        total_len = g['total_len']
        fmt = fmt_key(key)
        per_set_len = total_len / sets
        per_set_qty = g['qty'] / sets
        if abs(per_set_qty - round(per_set_qty)) > 1e-6:
            uneven.append((fmt, g['qty'], sets))

        if per_set_len > 0:
            sets_per_bar = int(math.floor((BAR_LEN + 1e-6) / per_set_len))
        else:
            sets_per_bar = 0
        bar_left = BAR_LEN - sets_per_bar * per_set_len if sets_per_bar > 0 else None

        # 采购根数 = 实际拼管数（FFD，可达），不能只用 ceil(总长/6000)：
        # 长件时每根按整根只能切 N 段，余料拼不成段，按总长折算会少买。
        segs = []
        for length, qty, _ in rows_by_key[key]:
            segs.extend([length] * qty)
        segs = [seg for seg in segs if seg > 0]
        for seg in segs:
            if seg > BAR_LEN + 1e-9:
                oversize.append((fmt, seg))
        bins = ffd_pack(segs)
        n_bars = len(bins)
        waste = sum(BAR_LEN - sum(b) for b in bins)

        spb_txt = str(sets_per_bar) if sets_per_bar else '-'
        bl_txt = f'{bar_left:.1f}' if bar_left is not None else '-'
        print(f"{fmt:<30}{key[0]:<6}{g['qty']:>6}{num(per_set_qty):>9}{total_len:>10.1f}"
              f"{per_set_len:>10.1f}{spb_txt:>10}{bl_txt:>10}{n_bars:>7}{waste:>9.1f}")
        out_rows.append({
            '\u89c4\u683c': fmt, '\u7c7b\u578b': key[0], '\u5957\u6570': sets, '\u6570\u91cf': g['qty'],
            '\u6bcf\u5957\u652f\u6570': num(per_set_qty),
            '\u603b\u957fmm': round(total_len, 1),
            '\u6bcf\u5957\u7528\u91cfmm': round(per_set_len, 1),
            '\u6bcf\u6839\u53ef\u5f00\u5957\u6570': sets_per_bar if sets_per_bar else None,
            '\u6bcf\u6839\u4f59\u6599mm': round(bar_left, 1) if bar_left is not None else None,
            '6m\u6839\u6570': n_bars,
            '\u5269\u4f59mm': round(waste, 1),
            '\u5b9e\u4f8b': '\u3001'.join(sorted(g['names'])[:3]),
        })

    total_bars = sum(r['6m\u6839\u6570'] for r in out_rows)
    print('-' * 104)
    print(f"\u5408\u8ba1\u9700\u91c7\u8d2d {len(groups)} \u79cd\u89c4\u683c\uff0c\u5171 {total_bars} \u6839 6m \u7ba1"
          f"\uff08\u6309 {sets} \u5957\u8ba1\uff09")
    print('\u6ce8: 6m\u6839\u6570 = \u5b9e\u9645\u62fc\u7ba1\u6839\u6570\uff08FFD\uff0c\u4e70\u8fd9\u4e48\u591a\u6839\u4e00\u5b9a\u591f\u5207\uff09\uff1b'
          '6m\u53ef\u5f00\u5957/\u6bcf\u6839\u4f59\u6599 = \u5355\u6839 6m \u7ba1\u6ee1\u8db3\u6574\u5957\u7684\u4e0a\u9650\u53c2\u8003\u3002')

    if uneven:
        print(f"\n[\u63d0\u9192] {len(uneven)} \u4e2a\u89c4\u683c\u7684\u6570\u91cf\u4e0d\u662f\u5957\u6570\u7684\u6574\u6570\u500d"
              f"\uff08\u6bcf\u5957\u652f\u6570\u5df2\u6309\u5e73\u5747\u53d6\u6574\uff09: {uneven[:3]}")

    if oversize:
        print(f"\n[\u8b66\u544a] {len(oversize)} \u6bb5\u8d85\u8fc7 6m\uff08\u5355\u6839\u88c5\u4e0d\u4e0b\uff0c"
              f"\u9700\u62fc\u63a5/\u5b9a\u5236\uff09: {oversize[:3]}")

    if skipped:
        print(f"\n[\u8b66\u544a] {len(skipped)} \u884c\u65e0\u6cd5\u89e3\u6790\u5df2\u8df3\u8fc7: {skipped[:3]}")

    import datetime
    out_path = args.out
    if not out_path:
        ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        out_path = os.path.join(os.path.dirname(os.path.abspath(args.xlsx)) or '.', f'\u6574\u7406\u578b\u6750_{ts}.xlsx')
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '\u578b\u6750\u91c7\u8d2d\u6c47\u603b'
    cols = ['\u89c4\u683c', '\u7c7b\u578b', '\u5957\u6570', '\u6570\u91cf', '\u6bcf\u5957\u652f\u6570',
            '\u603b\u957fmm', '\u6bcf\u5957\u7528\u91cfmm', '\u6bcf\u6839\u53ef\u5f00\u5957\u6570',
            '\u6bcf\u6839\u4f59\u6599mm', '6m\u6839\u6570', '\u5269\u4f59mm', '\u5b9e\u4f8b']
    ws.append(cols)
    for r in out_rows:
        ws.append([r[c] for c in cols])
    widths = {'\u89c4\u683c': 26, '\u7c7b\u578b': 8, '\u5957\u6570': 6, '\u6570\u91cf': 6,
              '\u6bcf\u5957\u652f\u6570': 11, '\u603b\u957fmm': 10, '\u6bcf\u5957\u7528\u91cfmm': 12,
              '\u6bcf\u6839\u53ef\u5f00\u5957\u6570': 14, '\u6bcf\u6839\u4f59\u6599mm': 12,
              '6m\u6839\u6570': 9, '\u5269\u4f59mm': 9, '\u5b9e\u4f8b': 40}
    for c, name in enumerate(cols, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = widths.get(name, 18)
    wb.save(out_path)
    print(f"\n\u5df2\u5199: {out_path}")


if __name__ == '__main__':
    main()