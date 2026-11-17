# -*- coding: utf-8 -*-
"""截图缓存库 (skill4/skill8 共用, 2026-08-11 深夜钦定方案)
原理: 截图一次存 ShotCache 库, SQLite 索引 (文件路径+配置名+大小+修改时间),
      命中直接复用 → BOM 导出不再打开 SW 截图, 零渲染零内存风险。
用法:
    img = shot_cache.get_shot(cache_dir, file_path, config)          # 命中返回 jpg 路径
    if not img: img = shot_iso(...); shot_cache.put_shot(...)        # 未命中才截
"""
import hashlib
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


def _db_path(cache_dir):
    return os.path.join(cache_dir, "shot_index.db")


def _connect(cache_dir):
    os.makedirs(cache_dir, exist_ok=True)
    db = sqlite3.connect(_db_path(cache_dir), timeout=30)
    db.execute(
        """CREATE TABLE IF NOT EXISTS shots(
            file_path TEXT NOT NULL,
            config_name TEXT NOT NULL DEFAULT '',
            name TEXT DEFAULT '',
            spec TEXT DEFAULT '',
            file_size INTEGER DEFAULT 0,
            file_mtime REAL DEFAULT 0,
            shot_path TEXT DEFAULT '',
            created_at REAL DEFAULT 0,
            PRIMARY KEY(file_path, config_name))""")
    db.commit()
    return db


def shot_filename(file_path, config_name):
    """缓存图文件名: md5(路径|配置).jpg"""
    key = f"{file_path}|{config_name or ''}"
    return hashlib.md5(key.encode("utf-8", "ignore")).hexdigest()[:24] + ".jpg"


def shot_path_of(cache_dir, file_path, config_name):
    return os.path.join(cache_dir, shot_filename(file_path, config_name))


def get_shot(cache_dir, file_path, config_name, file_size=None, file_mtime=None):
    """命中条件: 记录存在 + 图文件存在 + (若给 size/mtime) 与磁盘一致(图纸改了自动重截)。
    返回 shot 图路径; 未命中返回 None。"""
    try:
        db = _connect(cache_dir)
        row = db.execute(
            "SELECT shot_path, file_size, file_mtime FROM shots "
            "WHERE file_path=? AND config_name=?",
            (file_path, config_name or "")).fetchone()
        db.close()
    except Exception:
        return None
    if not row:
        return None
    if file_size is not None and row[1] is not None and row[1] != file_size:
        return None
    if file_mtime is not None and row[2] is not None and abs(row[2] - file_mtime) > 2:
        return None
    sp = row[0]
    if not sp or not os.path.exists(sp):
        return None
    return sp


def put_shot(cache_dir, file_path, config_name, name, spec, shot_path, file_size=None, file_mtime=None):
    """登记/更新一条截图记录。shot_path 通常 = shot_path_of(...)。"""
    try:
        db = _connect(cache_dir)
        db.execute(
            "INSERT OR REPLACE INTO shots VALUES(?,?,?,?,?,?,?,?)",
            (file_path, config_name or "", name or "", spec or "",
             int(file_size or 0), float(file_mtime or 0),
             shot_path or "", time.time()))
        db.commit()
        db.close()
        return True
    except Exception:
        return False


def stats(cache_dir):
    """返回 (记录数, 占用MB) 供日志。"""
    try:
        db = _connect(cache_dir)
        n = db.execute("SELECT COUNT(*) FROM shots").fetchone()[0]
        db.close()
    except Exception:
        n = 0
    total = 0.0
    try:
        for f in os.listdir(cache_dir):
            if f.lower().endswith(".jpg"):
                total += os.path.getsize(os.path.join(cache_dir, f))
    except Exception:
        pass
    return n, total / 1024 / 1024


def refresh_stats(cache_dir, paths=None):
    """把缓存记录的 size/mtime 刷成磁盘当前值(仅属性写入后调用, 图内容未变)。
    防止 图号/分类 等写属性步骤导致截图缓存批量失效重截。返回刷新条数。"""
    n = 0
    try:
        db = _connect(cache_dir)
        if paths is None:
            rows = db.execute("SELECT file_path FROM shots").fetchall()
        else:
            rows = [(p,) for p in paths]
        for (fp,) in rows:
            try:
                st = os.stat(fp)
                db.execute("UPDATE shots SET file_size=?, file_mtime=? WHERE file_path=? AND shot_path IS NOT NULL",
                           (st.st_size, st.st_mtime, fp))
                n += 1
            except Exception:
                pass
        db.commit()
        db.close()
    except Exception:
        pass
    return n
