---
name: "0-skill管理规范"
description: "全局 Skill 管理公约：创建/修改/废弃/恢复 Skill 必须遵守的规则，所有 Agent 启动时自动加载。"
---

# Skill 管理公约

本规范是全局 skill 管理规则，所有 agent 在操作 skill 时必须遵守。
操作前先调用 `skill_workshop list` 查看当前提案状态。

## 一、Skill 存放位置

```
workspace\skills\          ← 全局共享（所有 agent 可见可用）
workspace\skills-disabled\ ← 禁用/备份（不加载，保留以备恢复）
```

每个 agent 的 workspace 不单独维护 skill 副本，统一用全局目录。
> 注：skill_workshop apply 新建/修改的 skill 会落盘到当前 agent 的 `agents\<id>\skills\{dir}`，**不会自动写全局**。需把该目录整体复制到全局 `workspace\skills`（覆盖同名），并用 `Get-FileHash` 校验一致，才算真正上线全局。

## 二、命名规范

- 目录名 = `{编号}-{中文名}`，如 `7-sw文档综合处理`、`8-焊接装配标志分类输出BOM带截图`
- SKILL.md 的 `name` 字段与目录名严格一致
- 编号从 0 递增，不跳号（0=本规范，1~13=已占用，14+ 空闲）
- 废弃的 skill 移到 `skills-disabled\`，不直接删除

## 三、创建/修改 Skill — 必须走 skill_workshop

**铁律：严禁手动创建/编辑/删除 `skills\` 目录下的任何文件。**

| 操作 | 正确方式 | 禁止方式 |
|---|---|---|
| 新建 skill | `skill_workshop create` → `revise`(补脚本) → `apply` | 手动 mkdir + write SKILL.md |
| 修改 skill | `skill_workshop update/revise` → `apply` | 直接 edit SKILL.md |
| 废弃 skill | 移到 `skills-disabled\` | 直接 rm -rf |
| 恢复 skill | 从 `skills-disabled\` 移回 | 重新 create |
| 上线全局 | 把 agent skills 下的目录整体复制到全局 `workspace\skills` 并 Hash 校验 | 只改 agent 副本不同步全局 |

## 四、Skill 目录结构

```
{编号}-{中文名}/
  SKILL.md       ← 必须有，含 name + description
  scripts/       ← Python 脚本放这里
  references/    ← 参考文档放这里
```

## 五、现有 Skill 索引（编号 0~13 已占用）

| 编号 | 目录/key | 名称 | 用途 | 依赖 |
|---|---|---|---|---|
| 0 | 0-skill管理规范 | 本规范 | Skill 管理公约 | 无 |
| 1 | 1-图号名称更新 | 图号名称更新 | 遍历 SW 装配体全层级，按文件名 `代号_名称` COM 直写代号/名称 | SW 运行中 |
| 2 | 2-钣金展开 | 钣金展开 | 批量 SW 钣金展开导出 DWG/DXF，卷圆去折弯线 | SW + AutoBOM 插件 |
| 3 | 3-导出BOM | 导出BOM | 装配体 → 分类汇总表(CSV/Excel)，按模板导出 | SW 运行中 |
| 4 | 4-导出BOM带截图 | 导出BOM带截图 | 装配体 → 总外购件表 + 三维截图嵌入 Excel | SW + pywin32/openpyxl/Pillow |
| 5 | 5-运行SW宏 | 运行SW宏 | 驱动 SW 运行 .swp 宏 + 自动点弹窗 | SW 运行中 |
| 6 | 6-合并DXF | 合并DXF | 合并 DWG/DXF 零件图，按 CSV 分层，输出产线 DXF+报告 | Python ezdxf |
| 7 | 7-sw文档综合处理 | sw文档综合处理 | 一键串行处理 SW 总装（图号→分类→多模板BOM→钣金展开→合并DXF→型材合成） | 依赖 1/3/4/5/6 |
| 8 | 8-焊接装配标志分类输出BOM带截图 | 焊接装配标志分类输出BOM带截图 | 从装配体按 AutoBOM 模板以「焊接或装配标记」属性分类导出 BOM，含每零件等轴测截图 WPS 嵌入 Excel；`--template` 与 `--filter-value` 必填；xlsx 平铺输出目录、截图进 IMG 子文件夹 | SW + pywin32/openpyxl/Pillow |
| 9 | 9-整理型材 | 整理型材 | 读型材采购 Excel 按规格归并(方/长方管长宽可调换+厚度、圆管直径)，算总长度，按 6m/根折算采购条数，可 `--optimize` 下料优化；输出终端+汇总 xlsx | openpyxl |
| 10 | 10-虚拟件外部化与Pipe标记 | 虚拟件外部化与Pipe标记 | （待补） | — |
| 11 | 11-PVC管采购计算 | PVC管采购计算 | 从 SW 总装读步路/线路零件(PVC管段)规格+长度，×套数后按规格做 6m 管 1D 下料，输出采购表 xlsx | openpyxl |
| 12 | 12-调度任务信息提取 | 调度任务信息提取 | 提取调度所需上游信息：SW总装装配层级(只装配体不遍历零件)+焊接/装配分类BOM+采购表与后台库存+售前工期，输出 调度中间数据.json + 调度信息摘要.txt（2026-08-17 钦定） | SW 运行中 + pywin32/openpyxl |
| 13 | 13-微信订单统计 | 微信订单统计 | 导出微信群聊记录、统计"炉灶发单群"当日各型号订单分布、按型号统计下单台数；无需微信密码（进程内存只读提取密钥 + SQLCipher 解密 + ZSTD 解压），唯一入口 wechat_orders.py，支持 extract/export/stats/daily/schedule（2026-08-29 由 DeepSeek Harness 制作后编入） | C:\Python314 + zstandard + wechatauto 库 |
| 15 | 15-激光管机30天排产 | 激光管机30天排产 | 为唯一一台激光管机下料岗排未来30天管材开料计划（大锅灶铁/钢+商用炉具25/30/35），状态驱动贪心+逐日校验+内置A/B参考计划，输出逐日开管任务/工时/每日库存表，可输出CSV/Excel/桌面（2026-09 皇上钦定编入） | Python（标准库+openpyxl） |
| 16 | 16-多设备型材钣金汇总合并 | 多设备型材钣金汇总合并 | 多个设备图纸文件夹的型材带数量零件+钣金件DXF按严格图号合并数量，输出统一型材汇总+钣金件汇总文件夹（带数量零件改'做N个'、DXF归拢合并成一张、CSV/Excel）；唯一入口 merge_multi_devices.py（2026-09-08 皇上钦定编入） | Python + pywin32/openpyxl/ezdxf + 依赖 6-合并DXF |

## 六、新建 Skill 流程

## ★五之二、AutoBOM 版本备份铁律（2026-08-06 钦定）

**任何 AutoBOM 版本更新 / exe 替换 / dll 替换 / 配置变更，必须先备份，便于回滚。**

统一入口（一条命令）：

```
python -X utf8 "D:\deepseek harness\.dsh\skills\0-skill管理规范\scripts\autobom_backup.py"
```

- 默认：备份当前 AutoBOM 全套（exe/dll/config/setting.json/Template/Plugin）到 `C:\AutoBOM\_backup\backup_YYYYMMDD_HHMMSS\`
- `--list`：查已有快照
- `--restore <快照名>`：先备份当前再回滚到指定快照

**教训（2026-08-06）**：旧版 AutoBOM.exe（496,640B，7/31 构建）保存配置时会把 setting.json 的 `Templates`/`CurrentTemplate` 字段整个丢掉（该类无此属性），导致多套 BOM 模板全丢失。务必确认用的是 `C:\AutoBOM\AutoBOM.exe` 新版（514,560B）。更新前跑备份、回滚从快照恢复。

1. 确定下一个可用编号（从 14 开始）
2. 用 `skill_workshop create` 提 proposal
3. `revise` 补全脚本（scripts/ 支持文件）
4. apply 生效（落到 agent 目录）
5. 把目录整体复制到全局 `workspace\skills` 并 Hash 校验
6. 更新本规范"现有 Skill 索引"表
7. 改名/改编号时：agent + 全局两份目录都改，并同步索引表与脚本内路径引用

## 七、Proposal 清理

- 定期 `skill_workshop list` 查看 pending/rejected 提案
- 已 applied 的旧版本 proposal 可忽略（proposals.json 自动管理）
- 同名 skillKey 冲突时先查现有提案再操作
