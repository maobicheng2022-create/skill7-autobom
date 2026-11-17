---
name: "3-导出BOM"
description: "从SW装配体按AutoBOM模板导出分类汇总表(CSV/Excel)，纯后台无GUI。一条命令跑通, 唯一入口。--asm/--template/--filter-value/-o 必填。"
status: proposal
version: "v2"
date: "2026-08-02T10:04:44.977Z"
---

# AutoBOM 分类汇总表导出

> ⚡ **唯一权威入口，一条命令跑通。** 不要自行连 SW、不要复制 AutoBOM 计算逻辑——下面命令的脚本已内置全部数据提取与公式计算。

## 怎么用（唯一入口，必填 4 个参数）

前置：SolidWorks 已开启。

```
python -X utf8 "skills\3-导出BOM\scripts\autobom_export.py" --asm "<装配体.sldasm路径>" --template <模板名> --filter-prop 采购类别 --filter-value <分类值> -o "<输出.csv或.xlsx>"
```

参数：
- `--asm`：装配体路径（缺省用 SW 当前激活装配体）
- `--template`：AutoBOM 模板名，如 `型材`（从 setting.json 读取，新增模板自动适配）
- `--filter-prop` / `--filter-value`：过滤属性列和值，如 `采购类别`/`型材`
- `-o`：输出文件路径（按扩展名或 `--fmt` 决定 csv/xlsx）；可选 `--fmt csv|xlsx`
- 可选：`--qty 套数`（`$总数量`×套数）；多模板一次导出用 `--templates-multi "模板:过滤值:输出路径;..."`

输出：一份 Excel/CSV，列与 AutoBOM 对应模板一致，行只含过滤后的零件，已汇总数量。

## 跑完怎么验证

- 确认输出文件存在、行列正确：列=模板列，行=过滤后的零件且已汇总数量。

## 维护参考（以下是人维护用，执行无需读）

### 底层机制（实测验证，勿改）
本技能通过 SW API 直接提取数据并复刻 AutoBOM 的计算，不依赖 GUI：
1. `GetComponents(False)` 枚举装配体所有组件实例（扁平化，含子装配）
2. 逐组件读「配置特定属性」（须带配置名），`CreateMassProperty()` 算体积/表面积/单重，`GetPartBox(False)` 算包络（`$长=Y $宽=X $高=Z`）
3. 模板公式读 `C:\micDatas\pluginSettings\BOM\setting.json`，用移植的 formula_engine.py 算公式列

### 模板（当前）
- **型材**：`代号 $总数量 $宽 $高 厚度 型材类型 采购类别`（厚度=ROUND({$体积}/({$表面积}/2),0)；型材类型=IF 公式判长方/圆/方管）
- 新增模板只需在 AutoBOM 里加，本技能自动适配

### 依赖
- Python 64 位 + pywin32 + openpyxl（xlsx 时）；SW 类型库缓存 `python -m win32com.client.makepy sldworks.tlb`；`python -X utf8`
