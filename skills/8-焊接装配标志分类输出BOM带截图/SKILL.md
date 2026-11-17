---
name: "8-焊接装配标志分类输出BOM带截图"
description: "从装配体按AutoBOM模板以「焊接或装配标记」属性分类导出BOM(每零件等轴测截图WPS嵌入Excel)。一条命令跑通。--template与--filter-value必填。"
---

# 焊接/装配标志分类 BOM 导出（含三维截图嵌入）

> ⚡ **唯一权威入口，一条命令跑通。** 不要自行写截图/嵌入/属性分类逻辑——下面命令的脚本已全部内置。

## 怎么用（唯一入口，--template 与 --filter-value 必填）

前置：SolidWorks 已开启。

```
python -X utf8 "skills\8-焊接装配标志分类输出BOM带截图\scripts\sw_weld_export.py" --dir "<图纸文件夹>" -o "<输出目录>" --template "<模板名>" --filter-value "<焊接或装配标记值>" [--asm "<总装.sldasm>"] [--qty 套数]
```

参数：
- `--dir`：图纸文件夹，自动扫 .sldasm 找总装配体（代号最短，排除 ST_ 开头）
- `-o`：输出目录（xlsx 平铺在此目录本身；截图收进 `<输出目录>/IMG/` 子文件夹）
- **`--template`：AutoBOM 模板名，必填**（如 焊接部件清点清单 / 出货清单）
- **`--filter-value`：焊接或装配标记过滤值，必填**（如 焊接 / 装配）
- `--asm`：可选，显式指定总装配体 .sldasm 完整路径（自动识别错时用）
- `--qty`：套数（`$总数量`×qty，默认1，>=1）；`--limit N` 调试用

输出：`<输出目录>/<模板名>_YYYYMMDD_HHMMSS.xlsx`（文件名带时间戳，图片 WPS 嵌入单元格）+ `<输出目录>/IMG/<零件名>_等轴测.jpg`。

> 与 skill4 区别：skill4 按「采购类别」分类且跳过子装配；本 skill 按「**焊接或装配标记**」属性分类，且**不跳过子装配**（焊接/装配标记常直接标在焊接组件 .SLDASM 上，且不自动判定采购类别）。

## 跑完怎么验证

- 确认 xlsx 生成、文件名带时间戳、每行有截图、IMG 子目录有对应等轴测 jpg；焊接组件（若是子装配）被正确纳入而非空表。

## 维护参考（以下是人维护用，执行无需读）

### 技术要点（实测，勿改）
- 遍历不跳过子装配：`GetComponents(False)` 扁平化后，焊接组件(.SLDASM) 的 ModelDoc2 也返回，须检查其「焊接或装配标记」属性并纳入
- 等轴测截图 = 复刻 AutoBOM PanoramaShot 模式 B；**必须 `sw.ActivateDoc(零件名)` 切前台**（OpenDoc6 后 SW 仍保持原文档激活，不切会截到总装配体）
- 图片 WPS 嵌入单元格（写 xl/cellimages.xml + `=DISPIMG("ID_xxx",1)` 公式），非 openpyxl 浮动图；Excel 365/WPS 兼容
- 截图缩小：SaveAs2 原生视口极高分辨率，必须 PIL 缩到最长边 1000px/quality=88；PIL.MAX_IMAGE_PIXELS=None 防警告
- CloseDoc 传标题（不带路径扩展名）；居中+换行用手写 sheet1.xml + styles.xml
- 模板从 `C:\micDatas\pluginSettings\BOM\setting.json` 读取，模板列不含过滤属性时自动补列再剔除（不污染输出列）

### 常用模板
`焊接部件清点清单`：列 = `名称 代号 $总数量 单位 $三维截图`

### 依赖
- SW 已开；Python 64 位 + pywin32/openpyxl/Pillow；SW 类型库缓存已生成；`python -X utf8`
