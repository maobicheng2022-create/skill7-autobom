---
name: "0-skill管理规范"
description: "Skill 管理公约：目录结构、命名、依赖与 AutoBOM 版本备份铁律。"
---

# Skill 管理公约

本仓库把「SolidWorks 出 BOM」拆成一组可独立调用、也可串联的 skill。本文件规定它们怎么放、怎么改、怎么回滚。

## 一、Skill 存放位置

```
<仓库>\
  skills\{编号}-{中文名}\SKILL.md          ← 说明与调用示例
  skills\{编号}-{中文名}\scripts\*.py      ← 可执行脚本（唯一入口）
  shared\skill7_paths.py                  ← 所有路径解析（AutoBOM/宏/SW/截图库）
  config\config.json                      ← 本机配置（不入库）
```

- **依赖脚本只有一份**：`7-sw文档综合处理` 通过相对路径调用 `skills\1|3|4|5|6|8|9|10` 里的脚本，不再维护多副本。
- 想接进自己的 agent 框架（如 OpenClaw / DeepSeek Harness）：把整个 `skills\` 目录拷到其全局 skills 目录即可，脚本仍能相对找到 `shared\`。

## 二、命名规范

- 目录名 = `{编号}-{中文名}`，如 `7-sw文档综合处理`、`8-焊接装配标志分类输出BOM带截图`
- `SKILL.md` 的 `name` 字段与目录名严格一致
- 编号从 0 递增，不跳号；本仓库占用 0~11
- 要废弃某个 skill：从 `skills\` 移出到 `skills-disabled\`，不要直接删（便于回滚）

## 三、修改 Skill 的规矩

1. **先备份再改**：改脚本前 `git commit` 或复制一份 `*.bak_时间戳`。
2. **改完必须自测**：`python -X utf8 tools/verify_env.py` + 目标脚本 `--help`。
3. **路径不许写死**：新增路径一律加进 `shared/skill7_paths.py`，脚本里 `from skill7_paths import ...`。
4. **产物契约不许变**：`7-sw文档综合处理` 的产物清单（见其 SKILL.md）是对下游的承诺，改列名/目录结构要同步文档。

## 四、Skill 目录结构

```
{编号}-{中文名}/
  SKILL.md       ← 必须有，含 name + description + 用法 + 产物校验
  scripts/       ← Python 脚本放这里
  references/    ← 参考文档放这里（可选）
```

## 五、现有 Skill 索引（本仓库随附 0~11）

| 编号 | 目录 | 名称 | 用途 | 依赖 |
|---|---|---|---|---|
| 0 | 0-skill管理规范 | 本规范 | Skill 管理公约 + AutoBOM 备份铁律 | 无 |
| 1 | 1-图号名称更新 | 图号名称更新 | 遍历 SW 装配体全层级，按文件名 `代号_名称` COM 直写代号/名称 | SW 运行中 |
| 2 | 2-钣金展开 | 钣金展开 | 批量 SW 钣金展开导出 DWG/DXF，卷圆去折弯线 | SW + AutoBaiJinZhanKai |
| 3 | 3-导出BOM | 导出BOM | 装配体 → 分类汇总表(CSV/Excel)，按 AutoBOM 模板导出 | SW 运行中 |
| 4 | 4-导出BOM带截图 | 导出BOM带截图 | 装配体 → 总外购件表 + 三维等轴测截图嵌入 Excel | SW + pywin32/openpyxl/Pillow |
| 5 | 5-运行SW宏 | 运行SW宏 | 驱动 SW 运行 .swp 宏 + 自动点弹窗 | SW 运行中 |
| 6 | 6-合并DXF | 合并DXF | 合并 DWG/DXF 零件图，按配料 CSV 分层，输出产线 DXF + 报告 | Python ezdxf |
| 7 | 7-sw文档综合处理 | sw文档综合处理 | 一键串行：图号→分类→多模板 BOM×套数→截图→焊接分组→展开→合并→型材合成 | 依赖 1/3/4/5/6/8/9/10 |
| 8 | 8-焊接装配标志分类输出BOM带截图 | 焊接装配标志分类输出BOM带截图 | 按「焊接或装配标记」属性分类导出 BOM（含截图），`--template`/`--filter-value` 必填 | SW + pywin32/openpyxl/Pillow |
| 9 | 9-整理型材 | 整理型材 | 型材采购表按规格归并、算 6m 根数、每套支数/余料 | openpyxl |
| 10 | 10-虚拟件外部化与Pipe标记 | 虚拟件外部化与Pipe标记 | 虚拟件 SaveVirtualComponent 外部化 + Pipe_/ST_ 标外购件 | SW 运行中 |
| 11 | 11-PVC管采购计算 | PVC管采购计算 | 从总装读步路管段规格+长度，×套数后按 6m 管下料 | openpyxl |

> 12~18（调度信息提取、微信订单统计、炉具排产、激光管机排产、多设备汇总、采购库更新）属内部业务系统，不在本仓库。

## 六、AutoBOM 版本备份铁律（2026-08-06 立）

**任何 AutoBOM 版本更新 / exe 替换 / dll 替换 / 配置变更，必须先备份，便于回滚。**

```bash
python -X utf8 "skills\0-skill管理规范\scripts\autobom_backup.py"
```

- 默认：备份当前 AutoBOM 全套（exe/dll/config/setting.json/Template/Plugin）到 `<autobom_dir>\_backup\backup_YYYYMMDD_HHMMSS\`（默认 `C:\AutoBOM\_backup`）
- `--list`：查已有快照
- `--restore <快照名>`：先备份当前再回滚到指定快照
- `tools\install.ps1` 覆盖安装时也会自动先快照

**教训**：旧版 AutoBOM.exe（496,640B）保存配置时会把 setting.json 的 `Templates`/`CurrentTemplate` 字段整个丢掉，导致多套 BOM 模板全丢失。换版前先备份，出事从快照恢复。

## 七、新增 Skill 流程

1. 取下一个编号（本仓库下一个是 12）
2. 建 `skills\{编号}-{中文名}\`，写 `SKILL.md`（name/description/用法/产物）
3. 脚本放 `scripts\`，能单文件跑通；需要路径就加进 `skill7_paths.py`
4. `python -X utf8 tools/verify_env.py` 自检 + 目标脚本 `--help` 冒烟
5. 更新本文件「现有 Skill 索引」表
6. 若 `7-sw文档综合处理` 要串联它，改 `sw_doc_process.py` 的步骤表并同步更新 7 的 SKILL.md