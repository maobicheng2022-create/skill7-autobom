# -*- coding: utf-8 -*-
"""
sw_buyout_export.py — 从装配体按 AutoBOM 模板导出 BOM 并对每零件生成三维等轴测截图嵌入 Excel(区别于 sw-autobom-export 的文字表格, 本技能含截图嵌入能力)
纯后台无GUI, 但截图需临时激活零件(复刻 AutoBOM PanoramaShot 渲染模式 B)。

流程:
1. 扫描文件夹找总装(代号最短 .sldasm, 排除 ST_) 并打开
2. 遍历装配体聚合所有零件(按文件路径去重, 累计数量)
3. 过滤 采购类别 = 外购件
4. 对每个外购件零件: 打开->等轴测视图->缩放适应->SaveAs2 存 <输出目录>/IMG/<零件名>_等轴测.jpg
5. 用 openpyxl 生成 .xlsx: 列 = AutoBOM "总外购件"模板(名称 $总数量 单位 $三维截图 采购截图参考 网址)
   $三维截图列嵌入每行零件的等轴测图片(列宽16 行高50, 仿 AutoBOM)
6. 保存到 指定输出目录/<filter_value>.xlsx (不放下级)

用法:
  python sw_buyout_export.py --dir <图纸文件夹> -o <输出目录> [--template 总外购件]
"""
import argparse
import json
import os
import subprocess
import sys
import io
import time
import collections

import win32com.client
import pythoncom

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from sw_categorize import ensure_purchasing_category
import shot_cache
import prop_cache

from openpyxl.utils import column_index_from_string

AUTOBOM_CFG = r"C:\micDatas\pluginSettings\BOM\setting.json"
FILTER_PROP = "采购类别"
FILTER_VALUE_DEFAULT = "外购件"

# ---- 截图内存保护 (2026-08-11): 连续截图 SW 内存只增不减(实测每张~300MB), 16GB机器必须分批重启 ----
SHOT_BATCH = 15            # 每截 N 张重启一次 SW
SHOT_MEM_TRIGGER = 12.0    # 剩余内存低于此百分比立即重启(防 OOM 被杀进程, 实测 0.8GB 时被系统 SIGKILL)


def _mem_free_pct():
    import ctypes
    class MEMORYSTATUSEX(ctypes.Structure):
        _fields_ = [("dwLength", ctypes.c_ulong), ("dwMemoryLoad", ctypes.c_ulong),
                    ("ullTotalPhys", ctypes.c_ulonglong), ("ullAvailPhys", ctypes.c_ulonglong),
                    ("ullTotalPageFile", ctypes.c_ulonglong), ("ullAvailPageFile", ctypes.c_ulonglong),
                    ("ullTotalVirtual", ctypes.c_ulonglong), ("ullAvailVirtual", ctypes.c_ulonglong),
                    ("ullAvailExtendedVirtual", ctypes.c_ulonglong)]
    try:
        ms = MEMORYSTATUSEX()
        ms.dwLength = ctypes.sizeof(MEMORYSTATUSEX)
        ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(ms))
        return ms.ullAvailPhys / ms.ullTotalPhys * 100.0
    except Exception:
        return 100.0


def _sw_exe_from_registry():
    import winreg
    try:
        base = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\SolidWorks")
        i = 0
        while True:
            try:
                sub = winreg.EnumKey(base, i)
            except OSError:
                break
            i += 1
            if not sub.lower().startswith("solidworks"):
                continue
            try:
                k = winreg.OpenKey(base, sub + r"\Setup")
                folder, _ = winreg.QueryValueEx(k, "SolidWorks Folder")
                exe = os.path.join(folder, "sldworks.exe")
                if os.path.exists(exe):
                    return exe
            except OSError:
                continue
    except OSError:
        pass
    return None


def _restart_sw(sw):
    """截图中途重启 SW 释放内存, 返回新 sw 对象。
    ExitApp 在总装/文档未关时可能不退(实测旧SW 8.6GB挂住) → 20s 后强制 Stop-Process 兑底。"""
    try:
        sw.ExitApp()
    except Exception:
        pass
    # 等旧 SW 进程退出(最多20s), 没退则强制杀
    import subprocess as _sp
    for _i in range(20):
        time.sleep(1)
        r = _sp.run(["powershell", "-NoProfile", "-Command",
                     "if (Get-Process sldworks -ErrorAction SilentlyContinue) { 'ALIVE' } else { 'GONE' }"],
                    capture_output=True, text=True)
        if "GONE" in (r.stdout or ""):
            break
    else:
        print("[MEM] ExitApp 未退出 SW, 强制 Stop-Process sldworks")
        _sp.run(["powershell", "-NoProfile", "-Command",
                 "Stop-Process -Name sldworks -Force -ErrorAction SilentlyContinue"], capture_output=True)
        time.sleep(3)
    exe = _sw_exe_from_registry() or r"E:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\sldworks.exe"
    if os.path.exists(exe):
        subprocess.Popen([exe])
    for _i in range(120):
        time.sleep(1)
        try:
            sw2 = win32com.client.GetActiveObject("SldWorks.Application")
            time.sleep(3)
            print(f"[MEM] SW 已重启并可连接(等了 {_i+1}s)")
            return sw2
        except Exception:
            continue
    raise RuntimeError("SW 重启后 120s 未连上")



# ---------- SW 类型库 ----------
def _gen():
    import glob as _glob
    import win32com.client.gencache as gencache
    base = gencache.GetGeneratePath()
    hits = [os.path.join(base, n) for n in os.listdir(base) if n.startswith("83A33D31-27C5-11CE-BFD4-00400513BB57x0x") and n.endswith(".py")]
    if not hits:
        raise RuntimeError("未找到 SW 类型库缓存,请先运行: python -m win32com.client.makepy sldworks.tlb")
    modname = "win32com.gen_py." + os.path.splitext(os.path.basename(hits[0]))[0]
    return __import__(modname, fromlist=['IModelDoc2', 'IAssemblyDoc', 'IComponent2'])


_GEN = None
def gen():
    global _GEN
    if _GEN is None:
        _GEN = _gen()
    return _GEN


def _wrap(m):
    if m is None:
        return None
    try:
        return gen().IModelDoc2(m._oleobj_)
    except Exception:
        return m


def _mtype(m):
    t = m.GetType
    return t() if callable(t) else t


def _mtitle(m):
    t = m.GetTitle
    return t() if callable(t) else t


def _mpath(m):
    t = m.GetPathName
    return t() if callable(t) else t


def _config_name(m):
    try:
        n = m.ConfigurationManager.ActiveConfiguration.Name
        return n() if callable(n) else n
    except Exception:
        return ""


def open_doc(sw, path, dtype=1):
    r = sw.OpenDoc6(path, dtype, 1, "", 0, 0)
    if isinstance(r, tuple):
        r = r[0]
    return _wrap(r) if r else None

def read_props(model, names, config):
    out = {}
    try:
        cpm = model.Extension.CustomPropertyManager(config)
        if not cpm:
            return {n: "" for n in names}
        for n in names:
            try:
                r = cpm.Get3(n, False, "", "")
                if isinstance(r, (tuple, list)) and len(r) >= 3:
                    out[n] = (r[2] if r[2] else r[1]) or ""
                elif isinstance(r, (tuple, list)) and len(r) == 2:
                    out[n] = r[1] or ""
                else:
                    out[n] = r or ""
            except Exception:
                out[n] = ""
    except Exception:
        return {n: "" for n in names}
    return out


def _bom_excluded(c2, doc):
    """2026-08-13 皇上钦定: 是否"不包括在材料明细表中"(Exclude from BOM)。
    组件级 IComponent2.ExcludeFromBOM 或 文档配置属性 "Exclude from bill of materials"=TRUE → True。"""
    try:
        v = c2.ExcludeFromBOM
        if callable(v):
            v = v()
        if v:
            return True
    except Exception:
        pass
    try:
        cn = _config_name(doc)
        cpm = doc.Extension.CustomPropertyManager(cn)
        if cpm:
            for pn in ("Exclude from bill of materials", "不包括在材料明细表中"):
                try:
                    r = cpm.Get(pn)
                    if isinstance(r, (tuple, list)):
                        v = r[1] if len(r) > 1 else ""
                    else:
                        v = r
                    if str(v or "").strip().upper() in ("TRUE", "1", "是", "YES"):
                        return True
                except Exception:
                    continue
    except Exception:
        pass
    return False


# ---------- 模板 ----------
def load_template(name):
    with open(AUTOBOM_CFG, encoding="utf-8") as f:
        cfg = json.load(f)
    for t in cfg.get("Templates", []):
        if t.get("Name") == name:
            return [p["Name"] for p in t.get("Props", [])]
    raise ValueError("未找到 AutoBOM 模板: " + name + " (可选: " +
                     ", ".join(t.get("Name", "?") for t in cfg.get("Templates", [])) + ")")


# ---------- 找总装 ----------
def find_top_assembly(directory):
    cands = []
    for fn in os.listdir(directory):
        if not fn.lower().endswith(".sldasm"):
            continue
        if fn.startswith("~") or fn.startswith("ST_") or fn.lower().startswith("st_"):
            continue  # 排除 ~$ 锁文件 / ST_ 标准件
        if fn.upper().startswith("PIPE_"):
            continue  # 2026-08-11: 排除 Pipe_ 管道外购装配(会干扰找总装)
        code = fn.split("_")[0]
        cands.append((len(code), code, os.path.join(directory, fn)))
    if not cands:
        raise ValueError("目录下未找到装配体(.sldasm, 排除 ST_): " + directory)
    cands.sort(key=lambda x: (x[0], x[1]))
    return cands[0][2]


# ---------- 遍历聚合 ----------
def collect_parts(sw, asm_path, cache_dir=None, asm_cache_hit=None):
    """打开总装, 遍历所有组件, 按文件路径聚合 (去重+累计数量)。
    cache_dir 给定时: 扫描结束把结果写属性缓存; asm_cache_hit 为预取缓存时直接返回。"""
    if asm_cache_hit:
        agg = collections.OrderedDict()
        for k, e in asm_cache_hit.items():
            if k == "_v" or not isinstance(e, dict):
                continue  # 2026-08-12: 跳过版本号字段
            if e.get("is_asm"):
                continue  # 2026-08-12: 子装配不入外购件(缓存已含 is_asm 标记)
            if "\x00" in k:
                p, cfg = k.split("\x00", 1)
            else:
                p, cfg = k, e.get("config", "")
            key = p + "\x00" + cfg
            agg[key] = {"count": e.get("count", 1), "props": e.get("props", {}),
                        "config": cfg, "path": p,
                        "name": e.get("name", os.path.splitext(os.path.basename(p))[0])}
        return agg
    asm = open_doc(sw, asm_path, 2)
    if asm is None or _mtype(asm) != 2:
        raise RuntimeError("无法打开总装配体: " + asm_path)
    try:
        gen().IAssemblyDoc(asm._oleobj_).ResolveAllLightWeightComponents(False)
    except Exception:
        pass
    agg = collections.OrderedDict()  # path -> {count, props:{...}, name}
    meta = {}  # path -> {size, mtime} 供属性缓存
    IComponent2 = gen().IComponent2

    def _walk(adoc, excluded=False):
        """2026-08-13: 递归 GetComponents(True) 逐层; 父排除 → 整棵子树跳过(排除标记在父实例上)。"""
        try:
            comps = gen().IAssemblyDoc(adoc._oleobj_).GetComponents(True)
        except Exception:
            comps = []
        for comp in (comps or []):
            try:
                c2 = IComponent2(comp._oleobj_)
                doc = c2.GetModelDoc2()
                if doc is None:
                    continue
                # 强转 IModelDoc2, 失败则尝试用动态方式访问
                try:
                    doc = gen().IModelDoc2(doc._oleobj_)
                except Exception:
                    pass
                if excluded or _bom_excluded(c2, doc):
                    continue  # 2026-08-13: 不包括在材料明细表中 → 整棵子树跳过
                # 健壮性: 访问 GetType 失败则跳过该组件(可能是坏引用/未加载)
                t = _mtype(doc)
                if t == 2:
                    _walk(doc, False)  # 子装配递归(自身不在外购件, 内部零件才算)
                    continue
                p = _mpath(doc) or _mtitle(doc)
                if not p:
                    continue
                cn = _config_name(doc)
                # 2026-08-11: 补读 单位/网址/采购链接(模板列需要; 之前只读3个导致网址列全空)
                # 2026-08-12: 加 焊接或装配标记 — 写缓存时字段要全, 否则覆盖 skill3 缓存后 skill8 筛不出
                props = read_props(doc, [FILTER_PROP, "名称", "代号", "单位", "网址", "采购链接", "焊接或装配标记"], cn)
                if os.path.exists(p):
                    _st = os.stat(p)
                    meta[p] = {"size": _st.st_size, "mtime": _st.st_mtime}
            except Exception:
                continue
            # 2026-08-11: 按(路径,配置)聚合, 配置名供"型号规格"列(如 ST_UPVC90°弯头 的 DN15/DN20 分开行)
            key = p + "\x00" + cn
            if key not in agg:
                # 2026-09-19: 名称优先用属性「名称」(干净名), 缺省回退文件名去扩展名 — 与 skill3 一致
                _nm = str(props.get("名称", "") or "").strip()
                if not _nm:
                    _nm = os.path.splitext(os.path.basename(p))[0] if p else _mtitle(doc)
                agg[key] = {"count": 0, "props": props, "config": cn, "path": p,
                            "name": _nm}
            agg[key]["count"] += 1
            # 补全可能因配置不同缺失的名称
            for pk in ("名称", "代号"):
                if not agg[key]["props"].get(pk) and props.get(pk):
                    agg[key]["props"][pk] = props[pk]

    _walk(asm)
    # 2026-08-12: 不写属性缓存 — skill4 扫不到子装配, 写缓存会覆盖 skill3 的完整缓存(含 is_asm/焊接标记), 下游 skill8 筛不出焊接组件。
    # 缓存统一由 skill3 步骤2 维护; skill4 未命中时只扫不写。
    return agg


# ---------- 等轴测截图 (复刻 AutoBOM PanoramaShot 渲染模式 B) ----------
def shot_iso(sw, path, img_dir, config_name=None):
    """打开零件 -> 等轴测 -> 缩放适应 -> SaveAs2 存 jpg。返回图片路径。
    config_name: 指定截图配置(缓存 key 用), None 则用文档活动配置。"""
    os.makedirs(img_dir, exist_ok=True)
    if _path_exists(path) and path.lower().endswith((".sldprt", ".sldasm")):
        dtype = 2 if path.lower().endswith(".sldasm") else 1
        try:
            doc = open_doc(sw, path, dtype)
        except Exception:
            doc = None
    else:
        doc = None
    if doc is None:
        return None
    cfg = config_name or _config_name(doc)
    out = shot_cache.shot_path_of(img_dir, path, cfg)
    try:
        # 关键: OpenDoc6 打开后 SW 仍保持原文档为激活, 必须 ActivateDoc 把目标零件切到前台视口
        doc.Visible = True
        name = os.path.splitext(os.path.basename(path))[0]
        try:
            sw.ActivateDoc(name)
        except Exception:
            pass
        try:
            doc.ShowConfiguration2(cfg)
        except Exception:
            pass
        doc.ShowNamedView2("*等轴测", 7)
        doc.ViewZoomtofit2()
        try:
            doc.ActiveView.DisplayMode = 5  # 着色显示(与 AutoBOM 一致)
        except Exception:
            pass
        try:
            sw.Refresh(0)
        except Exception:
            pass
        doc.SaveAs2(out, 0, True, True)
        if os.path.exists(out):
            # 缩小超大截图(默认视口分辨率可能高达上亿像素), 控内存/体积; 最长边 1000px, 保比例
            try:
                from PIL import Image as _PIL
                _PIL.MAX_IMAGE_PIXELS = None
                _im = _PIL.open(out)
                if max(_im.size) > 1000:
                    _im.thumbnail((1000, 1000), _PIL.LANCZOS)
                    _im.convert("RGB").save(out, "JPEG", quality=88)
                _im.close()
            except Exception as _e:
                print("    [shot] 缩图失败(保留原图):", repr(_e)[:80])
            return out
    except Exception as e:
        print("    [shot] 失败:", path, "|", repr(e)[:80])
    finally:
        # 必须用文档标题关闭(CloseDoc 参数是标题而非文件路径), 否则文档残留导致 SW 堆文档/资源耗尽
        try:
            if name:
                sw.CloseDoc(name)
        except Exception:
            try:
                if os.path.exists(path):
                    sw.CloseDoc(os.path.basename(path))
            except Exception:
                pass
    return None


def _path_exists(p):
    return bool(p) and os.path.exists(p)


# ---------- 导出 xlsx(WPS 嵌入单元格图片: cellimages.xml + DISPIMG 公式) ----------
# 参考 AutoBOM 同款: 图片存 xl/media/ + xl/cellimages.xml 定义 <etc:cellImage>(WPS 专有命名空间)
# + 单元格填 =DISPIMG("ID_XXX",1) 公式 => 图片成为单元格的一部分, 随单元格过滤/排序/拷贝
import uuid as _uuid


def _col_start_emu(colw_chars):
    """各列起点 EMU 坐标(1字符≈7px, 1px≈9525EMU, 72dpi)"""
    x = 0; out = {}
    for i, w in enumerate(colw_chars, start=1):
        out[i] = x
        x += int(w * 7.0 * 9525)
    return out


def export_wps_xlsx(cols, rows_with_img, out_path, img_col_key="$三维截图", header_h_pt=24, row_h_pt=70):
    """生成真正嵌入单元格图片的 xlsx(WPS/Excel365 识别)。
    rows_with_img 每项含 {各列值..., '_img': 图片路径}
    """
    from PIL import Image as PILImage
    PILImage.MAX_IMAGE_PIXELS = None

    # 缺省列宽(字符): 非图片列 12, 图片列 20
    colw = [20 if c == img_col_key else 12 for c in cols]
    col_start = _col_start_emu(colw)
    img_col_idx = None
    for i, c in enumerate(cols, start=1):
        if c == img_col_key:
            img_col_idx = i
            break
    if img_col_idx is None:
        raise ValueError("模板列缺少图片列: " + img_col_key)

    # 1) 准备 media png 和 cellimages 条目, 分配 ID_ GUID
    tmpdir = os.path.join(os.path.dirname(os.path.abspath(out_path)), "_wps_tmp")
    os.makedirs(tmpdir, exist_ok=True)
    gids = []
    media = []          # (缩放png源路径, zip内路径)
    cellimg_pics = []   # cellimages.xml 的 <etc:cellImage> 片段
    cell_rels = []
    for rid, item in enumerate(rows_with_img, start=1):
        src = item.get("_img")
        if not src or not os.path.exists(src):
            gids.append(None)
            continue
        gid = "ID_" + _uuid.uuid4().hex.upper()
        gids.append(gid)
        im = PILImage.open(src)
        im.thumbnail((120, 90), PILImage.LANCZOS)
        wpx, hpx = im.size
        pp = os.path.join(tmpdir, f"e{rid}.png")
        im.save(pp, "PNG")
        cy = int((header_h_pt + (rid - 1) * row_h_pt) * 12700)
        cx = col_start[img_col_idx]
        ecx, ecy = int(wpx * 9525), int(hpx * 9525)
        media.append((pp, f"xl/media/image{rid}.png"))
        cellimg_pics.append(
            f'<etc:cellImage><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{rid}" name="{gid}"/><xdr:cNvPicPr/></xdr:nvPicPr>'
            f'<xdr:blipFill><a:blip r:embed="rId{rid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>'
            f'<xdr:spPr><a:xfrm><a:off x="{cx}" y="{cy}"/><a:ext cx="{ecx}" cy="{ecy}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic></etc:cellImage>')
        cell_rels.append(
            f'<Relationship Id="rId{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image{rid}.png"/>')

    def xml_esc(s):
        return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                .replace('"', "&quot;"))

    # 2) sheet1.xml: 表头 + 数据行(D列写 DISPIMG 公式)。s=0表头样式, s=1数据居中换行样式
    rows = ['<row r="1">']
    for c, col in enumerate(cols, start=1):
        L = chr(64 + c)
        rows.append(f'<c r="{L}1" s="0" t="inlineStr"><is><t>{xml_esc(col)}</t></is></c>')
    rows.append('</row>')
    for r_off, item in enumerate(rows_with_img):
        r = r_off + 2
        gid = gids[r_off]
        row = [f'<row r="{r}" ht="{row_h_pt}" customHeight="1">']
        for c, col in enumerate(cols, start=1):
            L = chr(64 + c)
            if col == img_col_key and gid:
                row.append(f'<c r="{L}{r}" s="1" t="str"><f>_xlfn.DISPIMG(&quot;{gid}&quot;,1)</f>'
                           f'<v>=DISPIMG(&quot;{gid}&quot;,1)</v></c>')
            else:
                v = item.get(col, "")
                if isinstance(v, (int, float)):
                    row.append(f'<c r="{L}{r}" s="1"><v>{v}</v></c>')
                else:
                    row.append(f'<c r="{L}{r}" s="1" t="inlineStr"><is><t>{xml_esc(v)}</t></is></c>')
        row.append('</row>')
        rows.append("".join(row))
    sheet1 = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<dimension ref="A1:' + chr(64 + len(cols)) + str(len(rows_with_img) + 1) + '"/>'
        '<sheetData>' + "".join(rows) + '</sheetData></worksheet>')

    # 样式: 0=表头(加粗居中换行), 1=数据(居中换行)
    styles_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="2"><font><sz val="11"/><name val="宋体"/></font>'
        '<font><b/><sz val="11"/><name val="宋体"/></font></fonts>'
        '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="2">'
        '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
        '</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '</styleSheet>')

    # 3) 其余部分 XML
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="png" ContentType="image/png"/>'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/cellimages.xml" ContentType="application/vnd.wps-officedocument.cellimage+xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '</Types>')
    workbook = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="总外购件" sheetId="1" r:id="rId1"/></sheets></workbook>')
    wb_rels = ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '<Relationship Id="rId3" Type="http://www.wps.cn/officeDocument/2020/cellImage" Target="cellimages.xml"/>'
        '</Relationships>')
    sheet_rels = ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>')
    cellimages_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<etc:cellImages xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:etc="http://www.wps.cn/officeDocument/2017/etCustomData">'
        + "".join(cellimg_pics) + '</etc:cellImages>')
    cellimages_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(cell_rels) + '</Relationships>')
    root_rels = ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>')

    # 4) 打包
    import zipfile
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", workbook)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/worksheets/sheet1.xml", sheet1)
        z.writestr("xl/styles.xml", styles_xml)
        z.writestr("xl/worksheets/_rels/sheet1.xml.rels", sheet_rels)
        z.writestr("xl/cellimages.xml", cellimages_xml)
        z.writestr("xl/_rels/cellimages.xml.rels", cellimages_rels)
        for sp, dp in media:
            z.write(sp, dp)
    # 清理临时 png
    try:
        for fn in os.listdir(tmpdir):
            os.remove(os.path.join(tmpdir, fn))
        os.rmdir(tmpdir)
    except Exception:
        pass
    return out_path


export_xlsx = export_wps_xlsx  # 主入口走 WPS 嵌入方案


def col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _border():
    from openpyxl.styles import Border, Side
    thin = Side(style="thin")
    return Border(left=thin, right=thin, top=thin, bottom=thin)


def AlignmentHelper_center():
    from openpyxl.styles import Alignment
    return Alignment(horizontal="center", vertical="center")


def FontHelper_bold():
    from openpyxl.styles import Font
    return Font(bold=True)


def BorderHelper_thin():
    return _border()


def main():
    ap = argparse.ArgumentParser(description="从装配体按 AutoBOM 模板导出 BOM 并对每零件生成等轴测截图嵌入 Excel(截图嵌入能力为区别于 sw-autobom-export 的核心)")
    ap.add_argument("--dir", required=True, help="图纸文件夹(自动找总装)")
    ap.add_argument("-o", "--output", required=True, help="输出目录(不放下级)")
    ap.add_argument("--template", default="总外购件", help="AutoBOM模板名, 可任意切换(默认 总外购件; 如 型材/钣金 等从 setting.json 读取)")
    ap.add_argument("--img-dir", default="", help="截图存放目录(默认 C:\\AutoBOM\\ShotCache 全局截图库; 旧版 IMG 子目录已废弃)")
    ap.add_argument("--cache-dir", default=shot_cache.DEFAULT_CACHE_DIR, help="截图缓存库目录(默认 C:\\AutoBOM\\ShotCache)")
    ap.add_argument("--no-cache", action="store_true", help="禁用截图缓存(强制全部重新截图)")
    ap.add_argument("--no-prop-cache", action="store_true", help="禁用属性缓存(强制重新打开总装扫属性)")
    ap.add_argument("--limit", type=int, default=0, help="仅处理前N个外购件(测试用)")
    ap.add_argument("--filter-value", required=True, help="采购类别过滤值(必填), 如 外购件/标准件/型材/钣金件")
    ap.add_argument("--qty", type=int, default=1, help="套数/数量, 导出表格的$总数量会乘以此值(默认1)")
    args = ap.parse_args()
    if args.qty < 1:
        print("错误: --qty 必须 >= 1"); sys.exit(1)
    if not args.dir or not os.path.isdir(args.dir):
        print("错误: --dir 图纸文件夹必填且需存在"); sys.exit(1)
    FILTER_VALUE = args.filter_value  # 全局覆盖(本脚本内引用)

    cols = load_template(args.template)
    if FILTER_PROP not in cols:
        cols = cols + [FILTER_PROP]
    # 2026-08-12: AutoBOM 模板已去掉 $三维截图(防 AutoBOM 自动截图), 本脚本自动补列出图
    if "$三维截图" not in cols:
        cols = cols + ["$三维截图"]

    os.makedirs(args.output, exist_ok=True)
    cache_dir = args.cache_dir or shot_cache.DEFAULT_CACHE_DIR
    os.makedirs(cache_dir, exist_ok=True)
    from datetime import datetime as _dt
    out_file = os.path.join(args.output, f"{args.template}_{_dt.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    print(f"[BOM] 带时间戳输出: {os.path.basename(out_file)}")
    print(f"[CACHE] 截图库: {cache_dir}")

    asm = find_top_assembly(args.dir)
    print("[SCAN] 总装配体:", os.path.basename(asm))
    sw = None  # 懒连接: 只有需要截图时才碰 SW
    # ★ 属性缓存优先: 命中则连 SW 都不需要 (2026-08-12 钦定)
    agg = None
    if not args.no_prop_cache:
        _hit, _cdata = prop_cache.get_asm_cache(cache_dir, asm)
        if _hit:
            agg = collect_parts(None, asm, cache_dir, asm_cache_hit=_cdata)
            print(f"[PROPCACHE] 命中属性缓存: {len(agg)} 项 (零 SW, 未打开总装)")
    if agg is None:
        try:
            sw = win32com.client.GetActiveObject("SldWorks.Application")
        except Exception:
            sw = win32com.client.Dispatch("SldWorks.Application")  # exec 场景兼容
        agg = collect_parts(sw, asm, cache_dir)
        print(f"[PROPCACHE] 未命中, 打开总装重扫并更新缓存: {len(agg)} 项")

    # ===== 自动判定采购类别 (2026-08-05 钦定): 已有属性优先, 空则按命名规则写入 =====
    auto_filled = 0
    for p, e in agg.items():
        existing = e["props"].get(FILTER_PROP, "").strip()
        if existing:
            continue
        if sw is not None and os.path.exists(e["path"]):
            c = ensure_purchasing_category(sw, e["path"], log=lambda *a: None)
            e["props"][FILTER_PROP] = c
            auto_filled += 1
    if auto_filled:
        print(f"[CATEG] 自动判定并写入采购类别: {auto_filled} 个零件")

    buy = collections.OrderedDict()
    for p, e in agg.items():
        if str(e["props"].get(FILTER_PROP, "")).strip() == FILTER_VALUE:
            buy[p] = e
    print(f"[BOM] 过滤 {FILTER_PROP}={FILTER_VALUE}: {len(buy)} 行")

    rows = []
    cache_hit = 0
    cache_new = 0
    for i, (p, e) in enumerate(buy.items()):
        if args.limit and i >= args.limit:
            break
        name = e["name"]
        cfg = e.get("config", "")
        st = None
        try:
            st = os.stat(e["path"])
        except Exception:
            st = None
        img = None
        # 缓存优先: 命中则完全不开 SW
        if st and not args.no_cache:
            img = shot_cache.get_shot(cache_dir, e["path"], cfg, st.st_size, st.st_mtime)
            if img:
                cache_hit += 1
                print(f"    [缓存] {os.path.basename(e['path'])} ({cfg})")
        if img is None:
            # 需要截图 → 懒连接 SW (缓存全命中时全程不碰 SW)
            if sw is None:
                try:
                    sw = win32com.client.GetActiveObject("SldWorks.Application")
                except Exception:
                    sw = win32com.client.Dispatch("SldWorks.Application")
            # 2026-08-11: 截图中按张数/剩余内存触发重启 SW 防内存堆积(16GB机器, 0.8GB空闲时系统会杀进程)
            if i > 0 and (i % SHOT_BATCH == 0 or _mem_free_pct() < SHOT_MEM_TRIGGER):
                print(f"[MEM] 截图 {i}/{len(buy)}, 剩余内存 {_mem_free_pct():.0f}%, 重启 SW 释放内存...")
                sw = _restart_sw(sw)
            img = shot_iso(sw, e["path"], cache_dir, config_name=cfg)
            if img:
                cache_new += 1
                if st:
                    shot_cache.put_shot(cache_dir, e["path"], cfg, name, cfg, img,
                                        st.st_size, st.st_mtime)
                print(f"    [图] {os.path.basename(img)}")
        row = {
            "名称": name,
            "型号规格": e.get("config", ""),  # 2026-08-11: 配置名写型号规格列
            "$总数量": e["count"] * args.qty,  # 套数倍数
            "单位": e["props"].get("单位", ""),
            "$三维截图": img or "",
            "采购截图参考": "",
            "网址": e["props"].get("网址") or e["props"].get("采购链接", ""),  # 2026-08-11: 网址/采购链接兼容
            "_img": img,
        }
        rows.append(row)

    export_cols = [c for c in cols if c != FILTER_PROP]
    export_xlsx(export_cols, rows, out_file)
    n, mb = shot_cache.stats(cache_dir)
    print(f"[CACHE] 本次: 命中 {cache_hit} 张, 新截 {cache_new} 张; 截图库共 {n} 条记录 / {mb:.1f} MB")
    print(f"[OK] 已导出: {out_file} (含 {len(buy)} 张等轴测截图嵌入, 截图库: {cache_dir})")


if __name__ == "__main__":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    main()
