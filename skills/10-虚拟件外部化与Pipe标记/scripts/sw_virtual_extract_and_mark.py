# -*- coding: utf-8 -*-
"""
sw_virtual_extract_and_mark.py — Skill10 虚拟件外部化 + Pipe_/ST_ 外购件标记
功能(按 2026-08-11 钦定, 嵌入 skill7 首位):
  1. 打开总装 → 找出所有虚拟零部件(无磁盘路径) → 外部化保存到图纸文件夹
     (文件名取 '虚拟文件夹^名称' 的 ^ 后部分, 如 DN15TK23^Pipe_2-xxx.SLDPRT -> Pipe_2-xxx.SLDPRT)
     → 保存总装(引用持久化) → 关闭总装
  2. 遍历文件夹: Pipe_/ST_ 前缀的 .sldprt/.sldasm(含外部化新生成的文件)
     → 写 采购类别=外购件 → Save → CloseDoc
依赖: sw_popup_guard.py (默认自带启动/停止; --no-guard 时复用外部守卫, skill7 内嵌用)
用法:
  python sw_virtual_extract_and_mark.py --dir <图纸文件夹> [--asm <总装路径>] [--no-guard]
"""
import argparse, io, os, re, sys, time

import win32com.client

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "..", "3-导出BOM", "scripts"))
import sw_bom_extract as sbe

POPUP_GUARD = r"C:\Users\mao\.openclaw\workspace\turing-workspace\scripts\sw_popup_guard.py"
PY = sys.executable


def log(msg):
    print(msg, flush=True)


def find_top_assembly(folder, asm_arg=None):
    """找总装: 排除 ~$ / ST_ / Pipe_ 前缀"""
    if asm_arg:
        return asm_arg
    cands = [os.path.join(folder, f) for f in os.listdir(folder)
             if f.lower().endswith(".sldasm") and not f.startswith("~")
             and not f.upper().startswith("ST_") and not f.upper().startswith("PIPE_")]
    if not cands:
        return None
    return sorted(cands, key=lambda p: len(os.path.splitext(os.path.basename(p))[0].split("_")[0]))[0]


def connect_sw():
    for _ in range(30):
        try:
            return win32com.client.GetActiveObject("SldWorks.Application")
        except Exception:
            time.sleep(3)
    raise SystemExit("SW 连接失败")


def _is_virtual_path(p, title):
    """判断组件是否为虚拟零部件:
    ①路径为空 ②路径在 SW 虚拟件临时目录(%TEMP%\swx*\VC~~\) ③标题/名称含 ^
    2026-08-11 修复: 之前只判路径空, 漏检了 SW 给虚拟件分配的临时文件路径。"""
    if not p:
        return True
    pl = p.lower()
    if '\\vc~~\\' in pl or '\\vc~~' in pl:
        return True
    if '\\temp\\swx' in pl or pl.startswith('/tmp/swx'):
        return True
    if title and '^' in str(title):
        return True
    return False


def externalize_virtuals(sw, asm_path, folder):
    """打开总装 → 虚拟零部件外部化。返回 (外部化列表, 是否有虚拟件)"""
    model = sbe.open_doc(sw, asm_path, 2)
    if model is None:
        raise RuntimeError(f"打开总装失败: {asm_path}")
    log(f"[A] 打开总装: {os.path.basename(asm_path)}")
    time.sleep(2)
    gen = sbe.gen()
    ad = gen.IAssemblyDoc(model._oleobj_)
    try:
        ad.ResolveAllLightWeightComponents(False)
        log("[A] 轻化组件已解析")
    except Exception as e:
        log(f"[A] Resolve 跳过: {repr(e)[:60]}")
    time.sleep(2)
    comps = ad.GetComponents(False)
    log(f"[A] 组件实例: {len(comps)}")
    IComponent2 = gen.IComponent2
    virtual = {}
    for comp in comps:
        try:
            c2 = IComponent2(comp._oleobj_)
            doc = c2.GetModelDoc2()
            if doc is None:
                continue
            doc = gen.IModelDoc2(doc._oleobj_)
            p = sbe.mpath(doc)
            t = sbe.mtitle(doc)
            if _is_virtual_path(p, t):
                if t not in virtual:
                    virtual[t] = c2
        except Exception as e:
            log(f"[A] 组件遍历 err: {repr(e)[:60]}")
    log(f"[A] 虚拟零部件种类: {len(virtual)}")
    saved = []
    for title, c2 in virtual.items():
        base = title.split("^")[-1]
        if not base.lower().endswith(".sldprt"):
            base += ".SLDPRT"
        target = os.path.join(folder, base)
        try:
            # 2026-08-11: IComponent2 无 SaveAs, 用 SaveVirtualComponent(保存虚拟件并转为外部引用)
            r = c2.SaveVirtualComponent(target)
            if r:
                saved.append((title, target))
                log(f"[A] 外部化: {title} -> {base}")
            else:
                log(f"[A] 外部化返回False {title}")
        except Exception as e:
            log(f"[A] 外部化失败 {title}: {repr(e)[:80]}")
    if virtual:
        try:
            model.Save()
            log("[A] 总装已保存(引用持久化)")
        except Exception as e:
            log(f"[A] 总装保存 err: {repr(e)[:80]}")
    # 关闭总装
    try:
        sw.CloseDoc(sbe.mtitle(model))
        log("[A] 总装已关闭")
    except Exception as e:
        log(f"[A] 关闭 err: {repr(e)[:60]}")
    return saved, bool(virtual)


def externalize_in_assembly(sw, model, folder):
    """对已打开的装配体: 外部化其内部虚拟零部件(否则 Save 会弹'另存为')。返回外部化数量。"""
    gen = sbe.gen()
    ad = gen.IAssemblyDoc(model._oleobj_)
    try:
        ad.ResolveAllLightWeightComponents(False)
    except Exception:
        pass
    time.sleep(1)
    comps = ad.GetComponents(False)
    IComponent2 = gen.IComponent2
    virtual = {}
    for comp in comps:
        try:
            c2 = IComponent2(comp._oleobj_)
            doc = c2.GetModelDoc2()
            if doc is None:
                continue
            doc = gen.IModelDoc2(doc._oleobj_)
            p = sbe.mpath(doc)
            t = sbe.mtitle(doc)
            if _is_virtual_path(p, t):
                if t not in virtual:
                    virtual[t] = c2
        except Exception as e:
            log(f"[B] 遍历 err: {repr(e)[:60]}")
    for title, c2 in virtual.items():
        base = title.split("^")[-1]
        if not base.lower().endswith(".sldprt"):
            base += ".SLDPRT"
        target = os.path.join(folder, base)
        try:
            r = c2.SaveVirtualComponent(target)
            if r:
                log(f"[B] 装配内虚拟件外部化: {title} -> {base}")
            else:
                log(f"[B] 外部化返回False {title}")
        except Exception as e:
            log(f"[B] 外部化失败 {title}: {repr(e)[:80]}")
    if virtual:
        try:
            model.Save()
            log(f"[B] 装配体已保存(引用持久化): {sbe.mtitle(model)}")
        except Exception as e:
            log(f"[B] 保存 err: {repr(e)[:80]}")
    return len(virtual)


def mark_buyout(sw, folder):
    """Pipe_/ST_ 前缀文件 → 采购类别=外购件 (装配体先外部化内部虚拟件)"""
    pattern = re.compile(r"^(pipe|ST)[_＿]", re.IGNORECASE)
    files = [f for f in os.listdir(folder)
             if pattern.match(f) and f.lower().endswith((".sldprt", ".sldasm"))]
    log(f"[B] 待标记 Pipe_/ST_ 文件: {len(files)}")
    ok = fail = skip = 0
    todo = sorted(files)
    i = 0
    while i < len(todo):
        f = todo[i]
        i += 1
        p = os.path.join(folder, f)
        dtype = 2 if f.lower().endswith(".sldasm") else 1
        doc = None
        for attempt in range(3):
            r = sw.OpenDoc6(p, dtype, 0, "", 0, 0)
            d = r[0] if isinstance(r, tuple) else r
            if d is not None:
                doc = d
                break
            time.sleep(2)
        if doc is None:
            log(f"[B] 打开失败: {f}")
            fail += 1
            continue
        try:
            if dtype == 2:
                # 装配体: 先外部化内部虚拟件, 否则 Save 弹'另存为'卡死
                ext = externalize_in_assembly(sw, doc, folder)
                if ext:
                    # 新外部化文件若为 Pipe_/ST_ 前缀, 加入待标记队列
                    for t in os.listdir(folder):
                        if pattern.match(t) and t.lower().endswith((".sldprt", ".sldasm")) and t not in todo:
                            todo.append(t)
                            log(f"[B] 新外部化文件入队: {t}")
            # 2026-08-13: 已标记(采购类别=外购件)的文件跳过, 不重写不Save(省时间+不破坏截图缓存mtime)
            try:
                cpm0 = doc.GetActiveConfiguration().CustomPropertyManager
                v0 = cpm0.Get("采购类别")
                if v0 is not None and str(v0).strip() == "外购件":
                    skip += 1
                    continue
            except Exception:
                pass
            cpm = doc.GetActiveConfiguration().CustomPropertyManager
            cpm.Add2("采购类别", 30, "外购件")
            doc.Save()
            ok += 1
        except Exception as e:
            log(f"[B] 标记失败 {f}: {repr(e)[:80]}")
            fail += 1
        finally:
            try:
                sw.CloseDoc(os.path.splitext(f)[0])
            except Exception:
                pass
    log(f"[B] 完成: 成功 {ok}, 跳过已标记 {skip}, 失败 {fail}")
    return ok, fail


def main():
    ap = argparse.ArgumentParser(description="Skill10 虚拟件外部化 + Pipe_/ST_ 外购件标记")
    ap.add_argument("--dir", required=True, help="图纸文件夹")
    ap.add_argument("--asm", default="", help="总装路径(可选, 缺省自动找)")
    ap.add_argument("--no-guard", action="store_true", help="复用外部弹窗守卫(skill7 内嵌用)")
    args = ap.parse_args()

    if not os.path.isdir(args.dir):
        raise SystemExit(f"文件夹不存在: {args.dir}")

    guard_proc = None
    if not args.no_guard:
        if not os.path.exists(POPUP_GUARD):
            raise SystemExit(f"缺依赖弹窗守卫: {POPUP_GUARD}")
        import subprocess
        guard_proc = subprocess.Popen([PY, "-X", "utf8", POPUP_GUARD, "1.0"])
        log("[A] 已启动弹窗守卫")

    t0 = time.time()
    try:
        sw = connect_sw()
        asm = find_top_assembly(args.dir, args.asm or None)
        if asm:
            ext_saved, has_virtual = externalize_virtuals(sw, asm, args.dir)
        else:
            log("[A] 未找到总装, 跳过外部化")
            ext_saved, has_virtual = [], False
        mark_buyout(sw, args.dir)
        log(f"===== Skill10 完成 ({time.time()-t0:.1f}s): 外部化 {len(ext_saved)} 个虚拟件, Pipe_/ST_ 标记完成 =====")
    finally:
        if guard_proc is not None:
            import subprocess
            guard_proc.terminate()
            try:
                guard_proc.wait(timeout=3)
            except Exception:
                guard_proc.kill()
            log("[A] 弹窗守卫已停止")


if __name__ == "__main__":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    main()
