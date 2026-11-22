# -*- coding: utf-8 -*-
"""装配体属性缓存库 (2026-08-12 凌晨钦定, skill3/4/8 共用)
思路: 与截图缓存同款 — 总装文件没变 → 直接复用上次扫出的组件清单+全部属性, 零 SW。
命中条件: asm 文件 size+mtime 一致, 且每个零件的 size+mtime 与磁盘一致(零件改了→整体重扫)。
2026-08-12 午: 加 CACHE_VERSION = 6 — 旧缓存缺 网址/采购链接/焊接或装配标记/is_asm 字段, 必须失效重扫。

用法:
    hit, data = prop_cache.get_asm_cache(cache_dir, asm_path)   # data = {path: {config, count, name, props, size, mtime, is_asm}}
    if not hit: ... 打开总装重扫 ...; prop_cache.put_asm_cache(...)
"""
import json
import os
import sqlite3
import time

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

DEFAULT_CACHE_DIR = SHOT_CACHE  # 由 shared/skill7_paths.py 解析

# 2026-08-12: 缓存结构版本。payload 结构变化(新增字段/属性)时必须 +1, 旧缓存自动失效。
CACHE_VERSION = 6


def _db_path(cache_dir):
    return os.path.join(cache_dir, "prop_index.db")


def _connect(cache_dir):
    os.makedirs(cache_dir, exist_ok=True)
    db = sqlite3.connect(_db_path(cache_dir), timeout=30)
    db.execute(
        """CREATE TABLE IF NOT EXISTS asm_cache(
            asm_path TEXT PRIMARY KEY,
            asm_size INTEGER DEFAULT 0,
            asm_mtime REAL DEFAULT 0,
            created_at REAL DEFAULT 0,
            payload TEXT DEFAULT '')""")
    db.commit()
    return db


def _stat(p):
    try:
        st = os.stat(p)
        return st.st_size, st.st_mtime
    except Exception:
        return None, None


def get_asm_cache(cache_dir, asm_path):
    """命中返回 (True, data); 未命中返回 (False, None)。
    data = {file_path: {"config":..,"count":..,"name":..,"props":{..},"size":..,"mtime":..,"is_asm":..}}"""
    try:
        db = _connect(cache_dir)
        row = db.execute("SELECT asm_size, asm_mtime, payload FROM asm_cache WHERE asm_path=?",
                         (asm_path,)).fetchone()
        db.close()
    except Exception:
        return False, None
    if not row:
        return False, None
    size, mtime = _stat(asm_path)
    if size is None:
        return False, None
    if row[0] != size or abs(row[1] - mtime) > 2:
        return False, None
    try:
        data = json.loads(row[2])
    except Exception:
        return False, None
    # 2026-08-12: 版本不匹配 → 旧结构缓存失效, 强制重扫
    if not isinstance(data, dict) or data.get("_v") != CACHE_VERSION:
        return False, None
    # 校验每个零件文件没变(变了→整体重扫); payload key 可能为 path\x00config
    for k, e in list(data.items()):
        if k == "_v" or not isinstance(e, dict):
            continue
        fp = k.split("\x00")[0]
        sz, mt = _stat(fp)
        if sz is None:
            return False, None
        if e.get("size") != sz or abs((e.get("mtime") or 0) - mt) > 2:
            return False, None
    return True, data


def put_asm_cache(cache_dir, asm_path, data):
    size, mtime = _stat(asm_path)
    if size is None:
        return False
    try:
        db = _connect(cache_dir)
        payload = dict(data)
        payload["_v"] = CACHE_VERSION
        db.execute("INSERT OR REPLACE INTO asm_cache VALUES(?,?,?,?,?)",
                   (asm_path, size, mtime, time.time(), json.dumps(payload, ensure_ascii=False)))
        db.commit()
        db.close()
        return True
    except Exception:
        return False
