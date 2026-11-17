# -*- coding: utf-8 -*-
"""
sw_batch_categorize.py — 分批零件分类(替代易爆内存的零件分类.swp 宏)
读"采购类别"属性(Extension.CustomPropertyManager(config).Get3) → 复制到分类文件夹。
每批(默认60)后自动重启 SW 释放内存(CloseDoc 不还给 OS, 连续打开几百文件必爆)。
用法:
  python sw_batch_categorize.py --dir <图纸文件夹> [--batch 60]
依赖: SW 需已打开(GetActiveObject); 内部自动 ExitApp+重启+等 ROT。
"""
import argparse, io, os, shutil, subprocess, sys, time

import win32com.client

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "..", "3-导出BOM", "scripts"))
import sw_bom_extract as sbe

PY = sys.executable
VALID = {"钣金件", "型材", "外购件", "标准件", "机加件", "机加工"}


def log(msg):
    print(msg, flush=True)


def connect_sw():
    for _ in range(30):
        try:
            return win32com.client.GetActiveObject("SldWorks.Application")
        except Exception:
            time.sleep(3)
    raise SystemExit("SW 连接失败")


def get_categ(sw, path):
    """打开零件读 采购类别。返回 (值, 成功)"""
    try:
        r = sw.OpenDoc6(path, 1, 1, "", 0, 0)
        doc = r[0] if isinstance(r, tuple) else r
    except Exception as e:
        log(f"  [ERR] OpenDoc6 {os.path.basename(path)}: {repr(e)[:60]}")
        return "", False
    if doc is None:
        return "", False
    try:
        doc = sbe.wrap_model(doc)
        cfg = sbe.config_name(doc)
        cpm = doc.Extension.CustomPropertyManager(cfg)
        if cpm:
            res = cpm.Get3("采购类别", False, "", "")
            v = res[1] if isinstance(res, tuple) and len(res) >= 2 else res
            val = (v or "").strip()
        else:
            val = ""
    except Exception as e:
        log(f"  [ERR] 读属性 {os.path.basename(path)}: {repr(e)[:60]}")
        val = ""
    finally:
        try:
            sw.CloseDoc(os.path.splitext(os.path.basename(path))[0])
        except Exception:
            pass
    return val, True


def restart_sw(sw):
    """退出 SW 并重启(释放内存), 返回新 sw 连接"""
    try:
        sw.ExitApp()
    except Exception:
        pass
    time.sleep(5)
    subprocess.run(["powershell", "-NoProfile", "-Command",
                    "Start-Process 'E:\\Program Files\\SOLIDWORKS Corp\\SOLIDWORKS\\sldworks.exe'"],
                   capture_output=True, text=True)
    time.sleep(8)
    return connect_sw()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dir", required=True)
    ap.add_argument("--batch", type=int, default=60)
    args = ap.parse_args()

    folder = args.dir
    # 只处理根目录直接 .sldprt(不含分类子目录, 排除 ~$ 锁文件残留)
    parts = sorted(f for f in os.listdir(folder)
                   if f.lower().endswith(".sldprt") and not f.startswith("~"))
    total = len(parts)
    log(f"零件总数: {total}, 每批 {args.batch}")

    sw = connect_sw()
    ok = fail = 0
    i = 0
    while i < total:
        batch = parts[i:i + args.batch]
        log(f"===== 批次 {i // args.batch + 1} [{i},{i + len(batch)}) =====")
        for f in batch:
            p = os.path.join(folder, f)
            val, success = get_categ(sw, p)
            if not success:
                log(f"  [FAIL] 打开失败 {f}")
                fail += 1
                continue
            if not val or val not in VALID:
                val = "未分类"
            td = os.path.join(folder, val)
            os.makedirs(td, exist_ok=True)
            try:
                shutil.copy2(p, os.path.join(td, f))
                ok += 1
                log(f"  [{val}] {f}")
            except Exception as e:
                log(f"  [FAIL] 复制失败 {f}: {repr(e)[:60]}")
                fail += 1
        i += len(batch)
        if i < total:
            log(f"批次完成, 重启 SW 释放内存...")
            sw = restart_sw(sw)
            log("SW 已重启")
    log(f"===== 分类完成: 成功 {ok}, 失败 {fail} =====")


if __name__ == "__main__":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    main()
