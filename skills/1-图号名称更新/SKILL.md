---
name: "1-图号名称更新"
description: "按最新文件名(代号_名称)用纯COM直写SW装配体全层级每个文件的代号/名称属性, 绝不增删其他属性。一条命令跑通, 唯一入口。"
---

# 文件名称代号批量更新（SW 纯COM直写）

> ⚡ **唯一权威入口，一条命令跑通。** 不要自行连接 SW、不要遍历组件、不要写辅助 COM 脚本——下面命令里的脚本全干了。

## 怎么用（唯一入口）

前置：SolidWorks 正在运行。

```
python -X utf8 "C:\Users\mao\.openclaw\workspace\skills\1-图号名称更新\scripts\sw_批量更新图号_v2.py"
```

- 脚本自动连接 SW、打开活动装配体、遍历全层级组件，按「最新文件名」的 `代号_名称`（下划线分隔）直写每个文件的「代号」「名称」属性。
- **含顶层装配体自身**（2026-10-19 修复：旧版只遍历组件，顶层被漏掉 → 顶层「代号」会一直停在旧值，实测炉具03 顶层是 `SZLJ.01.0` 而文件名已是 `SZLJ.03.0`）。汇总数字 = 组件数 + 1。
- 工程图（SLDDRW）不在装配体层级内，本 skill 不处理；本套图纸也不写「代号/名称」属性（标题栏引用模型属性），无需处理。
- **只写这两个属性，其他属性绝不增删改。**
- 可选：只处理文件名含某关键词的文件（小范围验证）：
  ```
  $env:SW_FILTER="关键词"; python -X utf8 "C:\Users\mao\.openclaw\workspace\skills\1-图号名称更新\scripts\sw_批量更新图号_v2.py"
  ```
- 可选：指定装配体路径（不必先手工打开；SKILL 文档旧版未记载，代码 `--asm` 参数支持）：
  ```
  python -X utf8 "C:\Users\mao\.openclaw\workspace\skills\1-图号名称更新\scripts\sw_批量更新图号_v2.py" --asm "D:\...\XXX.SLDASM"
  ```
- 空代号/空名称的文件自动跳过（格式不规范不硬写）。
- 注意：只处理「装配体里引用的文件」。同文件夹里不在装配体中的零件不会被更新（如需处理要单独指定）。

## 跑完怎么验证

- 脚本打印成功/跳过/失败数量。检查日志确认无一失败。
- 如需深度核对：对比部分文件的「代号/名称」属性 vs 文件名，应一致。

## 维护参考（以下是人维护用，执行无需读）

### 为什么不用宏（重要教训）
旧方案（调 `分离代号名称(*).swp` 宏）有致命坑：宏用 ActiveDoc.GetTitle() 取 SW 全局活动文档，脚本 OpenDoc6 后 ActiveDoc 仍停在总装配体 → 宏实际没执行 + 宏会增删质量/采购类别/设计日期等一堆属性 → 污染。正解：纯 COM 直写，不调宏、不依赖 ActiveDoc。

### 关键脚本
`scripts/sw_批量更新图号_v2.py`（本 skill 自带支持文件，同 workspace 根目录副本）。

### 校验（交付前）
- 逐文件比对 代号/名称 vs 文件名，确认不一致=0
- 抽查单个文件属性数跑前后不变（证明没增删）：如 `SZ.0704.06.0.01_夹具工装`=25

### API 注意
- `OpenDoc6` 返回**元组** `(doc, errors, warnings)`，取 `[0]`；文件类型 SLDPRT=1 / SLDASM=2；`Get3(name,False,"")` 读值；`GetNames()` 取属性名数组；gencache `SldWorks` 类绑定

### 遗留（用户已知）
7 个装配体残留 12 个 Admin 垃圾属性（v1宏历史污染，非v2造成）：Approval/AppDate/Auditing/Author/Standardize/AuditDate/Process/Date/ProofReader/PrfReadDate/StdDate/ProcessDate。清理需单独脚本（用户未拍板）。

## 遗留（用户已知，未处理）
- 7 个装配体残留 12 个 Admin 垃圾属性（v1宏历史污染，非v2造成）：Approval/AppDate/Auditing/Author/Standardize/AuditDate/Process/Date/ProofReader/PrfReadDate/StdDate/ProcessDate。清理需单独脚本（用户尚未拍板执行）。
