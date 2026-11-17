# 04 · 运行 skill7

## 1. 一条命令

```powershell
python -X utf8 "skills\7-sw文档综合处理\scripts\sw_doc_process.py" "<图纸文件夹>" --qty <套数>
```

例：

```powershell
python -X utf8 "skills\7-sw文档综合处理\scripts\sw_doc_process.py" "D:\图纸\SZ.0203.0_手动切换三通阀_做10套" --qty 10
```

前提：文件夹里有总装配体 `.sldasm`（自动选代号最短的、排除 `ST_` 前缀；也可 `--asm` 指定）。

## 2. 参数

| 参数 | 说明 |
|---|---|
| `folder` | 图纸文件夹（必填） |
| `--qty N` | 做几套（必填语义，默认 1）。**只乘各 BOM 的 `$总数量`** |
| `--asm <路径>` | 手动指定总装（一般不用） |
| `--from <步骤>` | 断点续跑：从该步骤开始 |
| `--to <步骤>` | 只跑到该步骤 |
| `--skip 1.6,2.1` | 跳过指定步骤（逗号分隔） |
| `--keep-sw` | 跑完不退出 SW（默认跑完退出释放内存） |
| `--no-asm-close` | 跳过关闭装配体（调试用，正常别用） |

## 3. 内部步骤表

> 调用方**不要**手动逐条跑，这些是 `sw_doc_process.py` 内部串行完成的。

| 步骤 | 动作 | 关键脚本 |
|---|---|---|
| A | 虚拟件外部化 + `Pipe_`/`ST_` 标外购件 | `10-.../sw_virtual_extract_and_mark.py` |
| 0 | 图号名称更新（按 `代号_名称` 写属性并 Save） | `sw_tuhao_categorize.py` |
| 0.5 / 0.8 | 关装配体 / 重启 SW 释放内存 | 内部 |
| 1 | 分批零件分类（每批 60 件 + 批间重启 SW） | `sw_batch_categorize.py` |
| 1.5 / 1.6 | 刷新截图缓存 / 全量预截图 | `4-.../sw_preshot.py` |
| 2 | 一次扫装配体导三表 ×qty | `3-.../autobom_export.py --templates-multi` |
| 2.1 / 2.2 | 型材整理 + 型材采购加计算列 | `9-.../xingcai_summary.py`、`xingcai_purchase_enrich.py` |
| 2.5 | 重启 SW | 内部 |
| 5 | 总外购件 + 每件等轴测截图 ×qty | `4-.../sw_buyout_export.py` |
| 5.5 | 重启 SW（截图很吃内存） | 内部 |
| 6 | 标准件 + 截图 ×qty | `4-.../sw_buyout_export.py` |
| 6.6 / 6.7 | 焊接部件清单 / 装配部件清单 | `8-.../sw_weld_export.py` |
| 7 / 7.5 | 关装配体 / 重启 SW | 内部 |
| 8.0 | 按「焊接或装配标记=焊接」收集焊接组 | `sw_weld_group.py collect` |
| 8 | 钣金展开 DXF | `AutoBaiJinZhanKai.exe 批量` |
| 8.8 | 每焊接组建 `钣金件\<组名>\` + 各组合并图 | `sw_weld_group.py merge-ban` |
| 9 | 非焊接件总合并 DXF | `6-合并DXF/dwg_merge.py` |
| 9.5 | 重启 SW | 内部 |
| 10.0 | 型材按组准备 `_weld_tmp` | `sw_weld_group.py prep-profile` |
| 10 | 型材合成宏（逐组） | `5-.../sw_run_macro.py` |
| 10.9 | 型材分组汇总 → `带数量零件\` | `sw_weld_group.py collect-profile` |

无钣金件 / 无型材件时会自动跳过对应步骤（不会中止整条流程）。

## 4. 大单怎么跑

单条 1 小时以上很常见。**不要**在前台用短超时跑，用封装脚本：

```powershell
.\tools\run_skill7.ps1 -Folder "D:\图纸\某设备_做10套" -Qty 10
```

它做的事：

- 日志写到 `<仓库>\logs\skill7_<文件夹名>_<时间>.log`
- 前台等待但每 10 秒回显最后 3 行
- 中断后按提示用 `-From` 续跑

## 5. 断点续跑

```powershell
# 中断在 8.8 之后（产物已全），从 9 接着跑
.\tools\run_skill7.ps1 -Folder "..." -Qty 10 -From 9

# 只跑 8~10
.\tools\run_skill7.ps1 -Folder "..." -Qty 10 -From 8 -To 10

# 跳过指定步骤
.\tools\run_skill7.ps1 -Folder "..." -Qty 10 -Skip "1.6,2.1"
```

续跑前确认前置步骤产物在（脚本每步会自行校验，缺了会直接报错指出来）。

## 6. 跑之前/之中要知道的事

- **SW 会自动拉起**：脚本连不上 `SldWorks.Application` 时会用 `Start-Process` 正常启动（不是 `-Embedding`），并轮询等待 ≤120 秒。
- **弹窗守卫**：PDM 库/另存为/完成提示等模态框会被自动点掉（`sw_popup_guard.py` 随父进程退出，异常也不残留）。
- **内存**：SW `CloseDoc` 不还内存给 OS，几百个件会吃 6–10 GB。流程在分类批次间、截图后、展开前自动重启 SW。16 GB 机器够用；建议跑之前关掉浏览器大页面。
- **别同时手动用 SW**：脚本会开关文档、重启 SW，你同时操作会互相打断。
- **产物都在 `<图纸文件夹>` 下**，不改动原始 `.SLDPRT/.SLDASM`（只写属性：代号/名称/采购类别）。

## 7. 失败了怎么定位

脚本每一步都会 `step("x.x", "名称", ...)` 打印；失败时最后一行是失败步骤名。

```powershell
# 看日志尾部
Get-Content "logs\skill7_xxx.log" -Tail 40
```

常见失败点：

| 现象 | 多半原因 |
|---|---|
| A 步骤失败 | 总装没找到 / 虚拟件保存被 PDM 弹窗拦住 |
| 0 步骤失败 | SW 未连上 / 属性写入被锁 |
| 2 步骤失败 | `setting.json` 模板缺失或列名对不上 |
| 5/6 截图失败 | 内存不足被系统杀掉 / 零件太大渲染超时 |
| 8 展开 0 个 DXF | `AutoBaiJinZhanKai.exe` 路径不对 / 分类目录没有 `.SLDPRT` |
| 9 合并数量不对 | 钣金件.csv 与 DXF 代号对不上（看 `合并报告.txt`） |
| 10 型材宏失败 | 宏目录不对 / `型材1.CSV` 缺件 / 宏被安全策略禁用 |

排障清单见 [06-常见问题](06-常见问题.md)。