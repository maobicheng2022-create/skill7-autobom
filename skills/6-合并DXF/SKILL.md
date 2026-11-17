---
name: "6-合并DXF"
description: "合并一个文件夹内多份零件图纸(DWG/DXF)，按CSV配料表(代号,数量)确定每种做几份，按厚度分层插入同一张图并标注'材料-厚度-做N个'，输出产线DXF+报告。源为DXF时纯Python不碰AutoCAD。一条命令跑通。"
---

# 批量合并 DWG 图纸 → 输出产线 DXF（按 CSV 数量 + 厚度分层）

> ⚡ **唯一权威入口，一条命令跑通。** 不要自行写 CAD 转换/合并/分层逻辑——下面命令的脚本已全部内置。

## 怎么用（唯一入口）

```
python -X utf8 "skills\6-合并DXF\scripts\dwg_merge.py" "<输入文件夹>"
```

- `<输入文件夹>`：含若干零件图（`.dwg` 或 `.dxf`）+ 一张配料表 `.csv` 的文件夹。
- 脚本自动：取第一个 `.csv` 作配料表(代号,数量) → 把所有零件图合并到一张 DXF → 每种零件插 1 份图块 + 标注 `材料-厚度-做N个` → 按厚度分层排列 → 空文件自动跳过并记报告。
- 输出：`合并图纸_YYYYMMDD_HHMMSS.dxf` + `合并报告.txt`（在输入文件夹内）。
- 源是 `.dxf` 时纯 Python 不需要 AutoCAD；源含 `.dwg` 时需 AutoCAD(COM) 转 dxf。

## 输入文件命名约定（必须匹配，否则无法解析进报告警告）

- 零件图：`代号_名称-材料-厚度.dwg`（或 `.dxf`）——`_`前=代号(对应csv第一列)、第一个`-`后=材料、最后一个`-`后=厚度(可有小数)
  - 示例 `SZ.0507.05.01.H02.05.0_盖板-SUS201-0.7.dwg`
- CSV：至少两列 `代号,数量`（表头 `代号,$总数量` 或 `代号,数量` 均可）
- 文件夹可混放 `.SLDPRT` 等非 dwg/dxf 文件，会被忽略；输出文件(`合并图纸_*`)自动排除

## 跑完怎么验证

- 确认 `合并图纸_*.dxf` 存在且实体数>0；标注数 = 成功插入数，且每标注与 csv 数量一致；厚度分层正确；空文件在报告 `[错误]` 段。

## 维护参考（以下是人维护用，执行无需读）

### 设计要点
- 用 ezdxf 纯 Python 生成 DXF（和 LISP 二进制插入等价，完全可程序化）；dwg→dxf 用 AutoCAD COM `SaveAs`(版本号 **61**=DXF R2013)，一次性后台批处理，输出到独立 `_dxf_tmp/` 子目录防污染。

### 环境与依赖
- AutoCAD 2026 (COM `AutoCAD.Application`, exe `D:\CAD\AutoCAD 2026\acad.exe`)——仅 dwg→dxf 转换；Python 3.14 + pywin32 + ezdxf(1.4.4)
- 建议 `OPENBLAS_NUM_THREADS=1`（脚本内已设）；约 12 张图 ≈ 40~60 秒，CAD 崩易则重跑（已转 dxf 会跳过复用）

### 排布/样式
- SPACING=800、ROW_GAP=600、TXT_HEIGHT=50；标注样式 `MERGE_SONGTI`(宋体 simsun.ttf) 避免 AutoCAD 中文乱码

### 坑
- SaveAs 版本号必须 **61**；转换输出到独立子目录；ezdxf `bbox.extents` 正确处理实体；`合并图纸_*` 自动排除免二次自扫

---

