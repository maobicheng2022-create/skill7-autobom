---
name: "5-运行SW宏"
description: "驱动SolidWorks运行.swp宏并自动点掉宏弹窗。一条命令跑通, 唯一入口。--macro与--dir必填。"
status: proposal
version: "v2"
date: "2026-08-02T07:37:07.301Z"
---

# SolidWorks 运行宏

> ⚡ **唯一权威入口，一条命令跑通。** 不要自行 COM 调用、不要写点弹窗线程——下面命令的脚本已全部内置。

## 怎么用（唯一入口，--macro 与 --dir 必填）

前置：SolidWorks 已开启；宏文件路径存在。

```
python -X utf8 "C:\Users\mao\.openclaw\workspace\skills\5-运行SW宏\scripts\sw_run_macro.py" --macro "<宏.swp路径>" --dir "<要处理的零件目录>"
```

参数：
- `--macro`：**必填**，宏文件 `.swp` 的完整路径（无默认值，必须显式传）
- `--dir`：**必填**，宏要处理的零件目录（多数宏会遍历其下所有 sldprt）
- `--module` / `--method`：可选，覆盖 VBA 模块/方法（默认 Macro51 / main）
- `--popup-extra`：可选，追加匹配弹窗标题的关键词（默认匹配 SOLIDWORKS/AutoBOM）
- `--timeout`：可选，宏执行超时秒数（默认 900）

脚本自动：连接 SW → 打开目录第一个 sldprt → 启动后台弹窗清理线程 → RunMacro2 运行宏 → 自动点掉宏弹窗 → 打印结果。

## 跑完怎么验证

- 看脚本打印的 `[MACRO] RunMacro2 返回` 结果和分类清单。若打印"未生成分类文件夹"，说明零件缺目标属性或宏逻辑不同。

## 维护参考（以下是人维护用，执行无需读）

### 底层机制（已实测，勿改）
- SW RunMacro2 签名为 5 参 `(FilePathName, ModuleName, ProcedureName, Options, Error)`；误传 4 参报"类型不匹配"
- 宏第一步常检查 ActiveDoc，必须先打开一个 sldprt；宏内 MsgBox 是模态弹窗 → 须后台线程预启动点弹窗（顺序调用会卡死）
- 点弹窗默认匹配标题含 SOLIDWORKS/AutoBOM（AutoBOM型材宏完成弹窗标题是 'AutoBOM 型材批量处理'，须匹配）

### 常用宏观路径（7 综合处理 orchestrator 实际用的）
`Z:\solidworks information\macro\零件分类.swp`（分类宏）、`Z:\solidworks information\macro\型材零件名称数量合成另存文件夹.swp`（型材宏，module=`AutoBOM_型材批量处理`）

### 注意事项
- 运行宏会**真实修改文件/创建子文件夹**，执行前确认目录正确
- SW 已开；Python 64 位 + pywin32；`python -X utf8`
