# -*- coding: utf-8 -*-
"""
autobom_export.py — 从装配体按 AutoBOM 模板导出分类汇总表 (路线3, 纯后台无GUI)
- 读 AutoBOM 模板定义(setting.json): 列 + 每列公式
- 遍历 SW 装配体拿所有零件 + 计算基础量($宽/$高包络, $体积/$表面积/$单重, $总数量汇总)
- 用公式引擎按模板公式算"厚度""型材类型"等
- 按"采购类别"过滤(如=型材) + 汇总
- 导出 CSV/xlsx

用法(单模板):
  python autobom_export.py --asm <装配体.sldasm路径> --template 型材 \
      --filter-prop 采购类别 --filter-value 型材 -o out.csv [--fmt csv|xlsx]

用法(多模板一次扫装配体, ★省时间):
  python autobom_export.py --asm <装配体> --qty 3 \
      --templates-multi "钣金:钣金件:.../钣金件.csv;型材:型材:.../型材.csv;型材采购:型材:.../型材采购.xlsx"
  格式: 分号分隔多个 "模板名:过滤值:输出路径", 只扫装配体一次 + 算系统属性一次, 大幅省时。
"""
import argparse
import json
import os
import sys
import io
import collections

import win32com.client

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import sw_bom_extract as sbe
import prop_cache
from formula_engine import evaluate

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



def load_template(name):
    with open(AUTOBOM_CFG, encoding="utf-8") as f:
        cfg = json.load(f)
    for t in cfg.get("Templates", []):
        if t.get("Name") == name:
            props, formulas = [], {}
            for p in t.get("Props", []):
                props.append(p["Name"])
                formulas[p["Name"]] = p.get("Formula", "") or ""
            return props, formulas
    raise ValueError("AutoBOM 配置中未找到模板: " + name + " (可选: "
                     + ", ".join(t.get("Name", "?") for t in cfg.get("Templates", [])) + ")")


def compute_sys_props(doc):
    out = {}
    try:
        if sbe.mtype(doc) == 1:
            pd = sbe.gen().IPartDoc(doc._oleobj_)
            box = pd.GetPartBox(False)
            x = round(box[3] - box[0], 2)
            y = round(box[4] - box[1], 2)
            z = round(box[5] - box[2], 2)
        elif sbe.mtype(doc) == 2:
            ad = sbe.gen().IAssemblyDoc(doc._oleobj_)
            box = ad.GetBox(2)
            x = round((box[3] - box[0]) * 1000.0, 2)
            y = round((box[4] - box[1]) * 1000.0, 2)
            z = round((box[5] - box[2]) * 1000.0, 2)
        else:
            x = y = z = ""
        out["$长"] = y if y != "" else ""
        out["$宽"] = x if x != "" else ""
        out["$高"] = z if z != "" else ""
    except Exception:
        out["$长"] = out["$宽"] = out["$高"] = ""
    try:
        mp = doc.Extension.CreateMassProperty()
        out["$体积"] = round(mp.Volume * 1e9, 2)
        out["$表面积"] = round(mp.SurfaceArea * 1e6, 2)
        out["$单重"] = round(mp.Mass * 1000.0, 2)
    except Exception:
        out["$体积"] = out["$表面积"] = out["$单重"] = ""
    return out


def _open_asm(sw, asm):
    model = sbe.open_doc(sw, asm, 2)
    print(f"[SW] 打开装配体: {os.path.basename(asm)}")
    if model is None or sbe.mtype(model) != 2:
        print("错误: 需要一个装配体(.sldasm)")
        sys.exit(1)
    try:
        sbe.gen().IAssemblyDoc(model._oleobj_).ResolveAllLightWeightComponents(False)
    except Exception:
        pass
    return model


def _scan(sw, filter_rpops, cache_dir=None, asm_path=None, asm_cache_hit=None):
    """扫装配体一次 + 聚合所有零件 + 计算系统属性。返回 [OrderedDict(key->row)]
    asm_cache_hit 给定时直接用属性缓存(零 SW, sysp 缺的补算)。"""
    import re as _re
    # 收集所有可能引用的非$列
    all_ref = set()
    for c in filter_rpops:
        all_ref.update(c.req_keys)
    plain_cols = list(dict.fromkeys(c for c in all_ref if not c.startswith("$")))
    # 2026-08-12: 属性缓存全量字段 — 模板没引用也读(网址/采购链接/焊接或装配标记 供 skill4/8 用, 否则缓存缺字段)
    for _x in ("网址", "采购链接", "单位", "焊接或装配标记"):
        if _x not in plain_cols:
            plain_cols.append(_x)
    # 2026-08-11: 模板含"型号规格"列时按(路径,配置)拆行, 配置名写入该列(如 ST_UPVC90°弯头 的 DN15/DN20 分开行)
    use_cfg = "型号规格" in plain_cols
    if asm_cache_hit:
        agg = collections.OrderedDict()
        for kk, e in asm_cache_hit.items():
            if kk == "_v" or not isinstance(e, dict):
                continue  # 2026-08-12: 跳过版本号字段
            if e.get("is_asm"):
                continue
            if "\x00" in kk:
                fp, cfg = kk.split("\x00", 1)
            else:
                fp, cfg = kk, (e.get("config", "") or "")
            key = fp + ("\x00" + cfg if use_cfg else "")
            props = dict(e.get("props", {}))
            if use_cfg:
                props["型号规格"] = cfg
            agg[key] = {"props": props, "name": e.get("name", os.path.splitext(os.path.basename(fp))[0]),
                        "count": e.get("count", 1), "path": key,
                        "sysp": e.get("sysp", {})}
        print(f"[BOM] 聚合零件种类: {len(agg)} (属性缓存命中)")
        _miss = [k for k, e in agg.items() if not e.get("sysp") and os.path.exists(k)]
        if _miss and sw is not None:
            print(f"[PROPCACHE] sysp 补算 {len(_miss)} 件...")
            for key in _miss:
                try:
                    pm = sbe.open_doc(sw, key, 1)
                    if pm:
                        agg[key]["sysp"] = compute_sys_props(pm)
                except Exception:
                    agg[key]["sysp"] = {}
        return agg
    root, rows = sbe.extract_components(sw, plain_cols)
    print(f"[BOM] 组件实例总数: {len(rows)}")
    agg = collections.OrderedDict()
    for r in rows:
        if r["is_asm"]:
            continue
        cfg = r.get("config", "") or ""
        key = (r["path"] or r["name"]) + ("\x00" + cfg if use_cfg else "")
        if key not in agg:
            props = dict(r["props"])
            if use_cfg:
                props["型号规格"] = cfg
            agg[key] = {"props": props, "name": r["name"], "count": 0, "path": key}
        agg[key]["count"] += 1
    print(f"[BOM] 聚合零件种类: {len(agg)}")
    # 一次性算每个零件的系统属性
    meta = {}
    for key, e in agg.items():
        if os.path.exists(key):
            try:
                pm = sbe.open_doc(sw, key, 1)
                if pm:
                    e["sysp"] = compute_sys_props(pm)
            except Exception:
                e["sysp"] = {}
            _st = os.stat(key)
            meta[key] = {"size": _st.st_size, "mtime": _st.st_mtime}
        else:
            e["sysp"] = {}
    if cache_dir and asm_path:
        # 2026-08-12: payload 按 (路径,配置) 粒度写, config 用真实配置名
        # (原实现 config 取"型号规格"列, 钣金/型材模板无此列→config 全空→下游截图缓存 key 错)
        # 2026-08-12 午: 子装配也入缓存(带 is_asm 标记) — skill8 焊接/装配组件是子装配, 缺了筛不出
        payload = {}
        for r in rows:
            fp = r.get("path") or r.get("name")
            if not fp:
                continue
            cfg = r.get("config", "") or ""
            kk = fp + "\x00" + cfg
            if kk in payload:
                # 2026-09-19 修复: 双重计数 bug — 首写已取 agg 全量 count, 后续实例又 +1
                # (如称重传感器 8 实例 → 缓存 8+7=15)。现在首写 count=1, 这里每实例 +1。
                payload[kk]["count"] += 1
                continue
            base = agg.get(kk) or agg.get(fp) or {}
            props = dict(r.get("props", {}))
            if use_cfg:
                props["型号规格"] = cfg
            _st = None
            if os.path.exists(fp):
                try:
                    _st = os.stat(fp)
                except Exception:
                    _st = None
            payload[kk] = {"config": cfg, "count": 1,
                           "name": base.get("name", os.path.splitext(os.path.basename(fp))[0]),
                           "props": props, "sysp": base.get("sysp", {}),
                           "is_asm": bool(r.get("is_asm")),
                           "size": (_st.st_size if _st else 0),
                           "mtime": (_st.st_mtime if _st else 0)}
        # 2026-08-13: 总装自身也入缓存 — 焊接/装配标记可能打在总装根文档上(如 150KG 整机标焊接),
        # skill8 缓存命中时也要能筛出总装行, 否则焊接清单空表
        try:
            _rm = None
            if sw is not None:
                try:
                    _rm = sbe.wrap_model(sw.ActiveDoc)
                except Exception:
                    _rm = None
            if _rm is not None:
                _rcfg = sbe.config_name(_rm)
                _rprops = sbe.read_props(_rm, plain_cols, _rcfg)
                _rst = None
                if os.path.exists(asm_path):
                    _rst = os.stat(asm_path)
                payload[asm_path] = {"config": _rcfg, "count": 1,
                                     "name": os.path.splitext(os.path.basename(asm_path))[0],
                                     "props": _rprops, "sysp": {},
                                     "is_asm": True,
                                     "size": (_rst.st_size if _rst else 0),
                                     "mtime": (_rst.st_mtime if _rst else 0)}
        except Exception:
            pass
        prop_cache.put_asm_cache(cache_dir, asm_path, payload)
    return agg


def _build_row(e, cols, formulas, ref_names, filter_prop, qty):
    row = {}
    all_cols = list(dict.fromkeys(list(cols) + list(ref_names) + [filter_prop]))
    for c in all_cols:
        if c == "$总数量":
            row[c] = e["count"] * qty
        elif c.startswith("$"):
            row[c] = e.get("sysp", {}).get(c, "")
        else:
            row[c] = e["props"].get(c, "")
    for c in cols:
        if formulas.get(c, ""):
            row[c] = evaluate(formulas[c], lambda n: str(row.get(n, "") or ""))
    return row


def _write(out, cols, data_rows, ws_title):
    os.makedirs(os.path.dirname(os.path.abspath(out)) or ".", exist_ok=True)
    fmt = "xlsx" if out.lower().endswith(".xlsx") else "csv"
    if fmt == "csv":
        import csv as _csv
        with open(out, "w", encoding="utf-8-sig", newline="") as f:
            w = _csv.writer(f)
            w.writerow(cols)
            for dr in data_rows:
                w.writerow([dr.get(c, "") for c in cols])
        print(f"[OK] CSV 已导出: {out}")
    else:
        from openpyxl import Workbook
        from openpyxl.styles import Alignment
        from openpyxl.utils import get_column_letter
        wb = Workbook(); ws = wb.active; ws.title = ws_title[:31] or "Sheet"
        ws.append(cols)
        for dr in data_rows:
            ws.append([dr.get(c, "") for c in cols])
        wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)
        header = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for r in ws.iter_rows(min_row=1, max_row=ws.max_row):
            for cell in r:
                cell.alignment = header if cell.row == 1 else wrap
        if ws.max_column:
            for col in range(1, ws.max_column + 1):
                ws.column_dimensions[get_column_letter(col)].width = 18
        wb.save(out)
        print(f"[OK] XLSX 已导出: {out}")


class _Tpl:
    """解析单个模板输出规格"""
    def __init__(self, name, fv, out, filter_prop, qty):
        self.name = name
        self.fv = fv
        self.out = out
        self.filter_prop = filter_prop
        self.qty = qty
        self.cols, self.formulas = load_template(name)
        self.ref_names = set()
        import re as _re
        for f in self.formulas.values():
            self.ref_names.update(_re.findall(r"\{([^{}]+)\}", f))
        self.req_keys = set(
            [c for c in self.cols if not c.startswith("$")]
            + [r for r in self.ref_names if not r.startswith("$")]
            + [r for r in self.ref_names if r.startswith("$")]
            + [self.filter_prop])


def _run_multi(sw, asm, qty, tpl_specs, filter_prop):
    """多模板一次扫装配体: tpl_specs = [(模板名,过滤值,输出路径), ...]"""
    cache_dir = SHOT_CACHE
    _hit, _cdata = False, None
    if os.path.exists(asm):
        _hit, _cdata = prop_cache.get_asm_cache(cache_dir, asm)
        if _hit:
            print("[PROPCACHE] 命中属性缓存: (零 SW, 未打开总装)")
    if not _hit and sw is None:
        try:
            sw = win32com.client.GetActiveObject("SldWorks.Application")
        except Exception:
            sw = win32com.client.Dispatch("SldWorks.Application")
    specs = []
    for name, fv, out in tpl_specs:
        specs.append(_Tpl(name, fv, out, filter_prop, qty))
        print(f"[BOM] 模板 {name}: 过滤={fv} -> {os.path.basename(out)}")
    if not _hit:
        _open_asm(sw, asm)
    agg = _scan(sw, specs, cache_dir, asm, asm_cache_hit=(_cdata if _hit else None))
    for t in specs:
        data_rows = []
        for key, e in agg.items():
            row = _build_row(e, t.cols, t.formulas, t.ref_names, t.filter_prop, t.qty)
            fv = row.get(t.filter_prop, "")
            if str(fv).strip() == t.fv:
                data_rows.append(row)
        print(f"[BOM] [{t.fv}] {len(data_rows)} 行")
        _write(t.out, t.cols, data_rows, t.name)


def main():
    ap = argparse.ArgumentParser(description="从装配体按 AutoBOM 模板导出分类汇总表")
    ap.add_argument("--asm", default="", help="装配体路径(可选,缺省用SW当前激活装配体)")
    ap.add_argument("--qty", type=int, default=1, help="套数/数量, 导出表格的$总数量会乘以此值(默认1)")
    ap.add_argument("--template", default="", help="AutoBOM模板名(单模板模式), 如 型材/钣金")
    ap.add_argument("--filter-prop", default="采购类别", help="过滤属性列")
    ap.add_argument("--filter-value", default="", help="过滤值(单模板模式), 如 型材")
    ap.add_argument("-o", "--output", default="", help="输出文件(单模板模式)")
    ap.add_argument("--fmt", choices=["csv", "xlsx"], default="")
    ap.add_argument("--templates-multi", default="",
                    help="多模板一次导出(推荐): '模板名:过滤值:输出路径' 分号分隔, 只扫装配体一次")
    args = ap.parse_args()
    if args.qty < 1:
        print("错误: --qty 必须 >= 1"); sys.exit(1)
    if not args.asm:
        print("错误: 必须提供 --asm <装配体.sldasm路径> (或由综合处理 skill 传入)"); sys.exit(1)

    try:
        sw = win32com.client.GetActiveObject("SldWorks.Application")
    except Exception:
        sw = win32com.client.Dispatch("SldWorks.Application")  # exec 场景兼容

    # ===== 多模板模式 =====
    if args.templates_multi:
        specs = []
        for seg in args.templates_multi.split(";"):
            seg = seg.strip()
            if not seg:
                continue
            parts = seg.split(":")
            if len(parts) < 3:
                sys.exit(f"错误: --templates-multi 段格式应为 '模板名:过滤值:输出路径': {seg}")
            specs.append((parts[0], parts[1], ":".join(parts[2:])))
        if not specs:
            sys.exit("错误: --templates-multi 无有效段")
        _run_multi(sw, args.asm, args.qty, specs, args.filter_prop)
        return

    # ===== 单模板模式(兼容旧调用) =====
    if not args.template or not args.filter_value or not args.output:
        sys.exit("错误: 需提供 --templates-multi 或 (--template + --filter-value + -o)")
    _run_multi(sw, args.asm, args.qty, [(args.template, args.filter_value, args.output)],
               args.filter_prop)


if __name__ == "__main__":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    main()
