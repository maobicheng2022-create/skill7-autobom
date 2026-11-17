---
name: "4-导出BOM带截图"
description: "从装配体按AutoBOM模板导出BOM(每零件等轴测截图WPS嵌入Excel)，后台无GUI。一条命令跑通。--filter-value必填。"
status: proposal
version: "v2"
date: "2026-08-02T15:08:39.450Z"
---

# 总外购件 BOM 导出（含三维截图嵌入）

> ⚡ **唯一权威入口，一条命令跑通。** 不要自行写截图/嵌入逻辑——下面命令的脚本已内置全部渲染与 Excel 嵌入。

## 怎么用（唯一入口）

前置：SolidWorks 已开启。输入：**图纸文件夹路径**（自动找总装并打开）+ 输出目录。

```
python -X utf8 "C:\Users\mao\.openclaw\workspace\skills\4-导出BOM带截图\scripts\sw_buyout_export.py" --dir "<图纸文件夹>" -o "<输出目录>" [--template 总外购件] [--filter-value 外购件]
```

参数：
- `--dir`：图纸文件夹，自动扫 .sldasm 找总装配体（代号最短，排除 ST_ 标准件）
- `-o`：输出目录（xlsx 与 IMG 都平铺在此目录，不放下一级子文件夹）
- `--template`：AutoBOM 模板名（默认 总外购件；标准件采购 结构相同）
- `--filter-value`：过滤值，如 `外购件`（按采购类别过滤）
- `--qty 套数`（可选，`$总数量`×套数）；`--limit N`（调试，仅处理前 N 个）

输出：`<输出目录>/总外购件_YYYYMMDD_HHMMSS.xlsx`（Excel，图片嵌入单元格）+ `<输出目录>/IMG/`（每零件等轴测 jpg）。

## 跑完怎么验证

- 确认 xlsx 生成、每行有三维截图、IMG 子目录有对应的等轴测 jpg。

## 维护参考（以下是人维护用，执行无需读）

### 技术要点（实测验证，勿改）
- 等轴测截图 = 复刻 AutoBOM PanoramaShot 模式 B：OpenDoc6→ShowNamedView2("*等轴测",7)→ViewZoomtofit2→SaveAs2(图片路径,0,True,True) 存 jpg
- SaveAs2 是 4 参数 `(NewName,0,True,True)`；OpenDoc6 打开即激活，截图完 CloseDoc 关闭
- 图片嵌入用 openpyxl `ws.add_image(img,"D"+行号)`，width=100 height=60，列宽16 行高50
- 必须 `PIL.Image.MAX_IMAGE_PIXELS=None` 防 DecompressionBomb 警告
- 遍历 `GetComponents(False)` 扁平化，按文件路径去重累加数量；总装识别排除 ST_

### 依赖
- SW 已开；Python 64 位 + pywin32/openpyxl/Pillow；SW 类型库缓存已生成；`python -X utf8`
