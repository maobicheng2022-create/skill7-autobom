---
name: "7-sw文档综合处理"
description: "SW总装一键综合处理(图号→分类→多表BOM×qty→钣金展开→合并DXF→型材宏)的唯一权威入口。直接照抄执行下方命令, 禁止自行前置检查/开总装/关文档/重写逻辑。强制folder+qty。"
---

# 7-sw文档综合处理

> ⚡ **本 skill 是 SW 综合处理的唯一权威入口。** 调用方不要再做任何前置检查、不要自行打开/关闭装配体、不要自己写排障脚本、不要重写处理逻辑——**直接照抄下方```执行```一节的可运行命令跑通即可**。脚本内部已包含全部前置校验、装配体管理、产物统计。除非命令返回失败，否则不要介入。

当需要把一批 SW 零件从「图号更新 → 零件分类 → 各类 BOM 导出（数量×套数）→ 钣金展开 → DXF 合并 → 型材合成」串成一条线一次性跑完时使用本 skill。纯后台串行执行，逐步验证，任一步失败即停并报错点。

## 开始前的强制前提（缺一不可，缺就拒绝，不要自己补）

1. **文件夹路径 `folder`**：SW 图纸所在文件夹（必填，必须存在）。
2. **套数 `qty`**：这套图纸要做几套（必填，必须 >=1）。约定：qty 只乘各 BOM 表格的 **$总数量**；合并 DXF / 型材宏复用已乘过的 csv，不单独乘。
3. 可选 `--asm <总装.sldasm>`：如果不填，脚本会自动找文件夹里代号最短的 `.sldasm`（排除 `ST_`）作为主总装。**一般不用填。**

其余一切（SW 是否在跑、Z盘、AutoBOM、宏、锁文件、装配体开关）脚本内部都会自己检查处理，**你不需要、也不要去做前置检查。**

## 怎么用（唯一入口，一条命令跑通）

> 脚本已封装全部 12 步串行流程，**不需要你逐条执行下面的步骤表**。

运行下面这一条命令即可（用真实文件夹替换 `<图纸文件夹>`，用真实套数替换 `<套数>`）：

```
python -X utf8 "C:\Users\mao\.openclaw\workspace\skills\7-sw文档综合处理\scripts\sw_doc_process.py" "<图纸文件夹>" --qty <套数>
```

例如处理「测试」文件夹、做 1 套：

```
python -X utf8 "C:\Users\mao\.openclaw\workspace\skills\7-sw文档综合处理\scripts\sw_doc_process.py" "C:\Users\mao\Desktop\agent team\其他项目\测试" --qty 1
```

## ★执行姿势（2026-08-30 固化，防 exec SIGKILL 连坐）

> skill7 大单 1h+，exec 层默认短超时会**连坐杀死**整个 python 树（孤儿模式：exec shell 被杀但 python 还活着继续跑，容易卡在某个无守卫的中间态）。
> **长任务必须用「后台 + 长 yieldMs/轮询」姿势执行，禁止设短超时。**

推荐姿势（PowerShell）：

```powershell
# 1) 后台启动, 日志写文件
Start-Process -FilePath "C:\Python314\python.exe" -ArgumentList "-X","utf8","C:\Users\mao\.openclaw\workspace\skills\7-sw文档综合处理\scripts\sw_doc_process.py","<图纸文件夹>","--qty","<套数>" -RedirectStandardOutput "C:\Users\mao\.openclaw\workspace\turing-workspace\memory\skill7_run.log" -RedirectStandardError "C:\Users\mao\.openclaw\workspace\turing-workspace\memory\skill7_run.err" -WindowStyle Hidden
# 2) 轮询日志 (不要设 exec 短超时; 用长 yieldMs/process poll)
Get-Content -LiteralPath "C:\Users\mao\.openclaw\workspace\turing-workspace\memory\skill7_run.log" -Encoding UTF8 -Tail 20
```

铁律：
- exec 跑 skill7 **禁止** `timeout` 短超时（30min 必杀）；用后台 + 长 yieldMs + process poll。
- 网关重启会连坐杀 exec → 长任务期间别动网关。
- 若中断在步骤 8.8 之后（产物已全），用下方**断点续跑**补跑，不必整条重跑。

## ★断点续跑（2026-08-30 新增：--from/--to/--skip）

skill7 支持步骤级续跑，中断后不必整条重跑 2-3 小时：

```
# 从步骤 9 续跑(中断在 8.8 之后, 前面产物已全)
python -X utf8 "C:\Users\mao\.openclaw\workspace\skills\7-sw文档综合处理\scripts\sw_doc_process.py" "<图纸文件夹>" --qty <套数> --from 9

# 只跑 8 到 10 段
python -X utf8 "...sw_doc_process.py" "<图纸文件夹>" --qty <套数> --from 8 --to 10

# 跳过指定步骤(逗号分隔)
python -X utf8 "...sw_doc_process.py" "<图纸文件夹>" --qty <套数> --skip 1.6,2.1
```

步骤号：A(虚拟件), 0(图号+分类), 1.5(刷新截图库), 1.6(预截图), 2(BOM三表), 2.1(型材整理), 2.5/5.5/6.9/7.5/9.5(重启SW), 5(外购), 6(标准件), 6.6/6.7(焊接/装配清单), 7(关装配), 8.0(分组收集), 8(展开), 8.8(分组合并), 9(合并DXF), 10.0(型材准备), 10(型材宏), 10.9(型材汇总)。

> 断点续跑前**确认前面已完成步骤的产物存在**（脚本内部各步会自行校验缺产物报错）；若某步骤产物缺失会失败，按需先补跑该步骤。

## ★PDM 库引用本地化（2026-08-30 根治）

总装引用可能把零件指到 PDM 库只读副本（如 D:\feedingHero），对它 Save() 会弹 PDM 另存为框阻塞。
已内嵌自动处理：收集路径时若发现 PDM 库路径 + 桌面夹有同名可写副本 → **重定向到桌面副本**（属性可正常刷全）；桌面无可写副本的真库件 → **跳过不处理**（库件属性原本已对）。无需人工干预。

- 入口脚本**唯一**且满足你所需能力：`sw_doc_process.py`。**不需要读它源码、不需要在它之外再写任何辅助脚本。**
- 命令跑完会自动在末尾打印**每步耗时表**。
- 任一步失败脚本会自行停止并告诉你失败在哪一步，届时再针对那一步查。
- 若手动指定总装，追加 `--asm "<总装.sldasm>"` 。

## 运行后要验证哪些产物（用于确认成功）

全部落在 `<图纸文件夹>` 及子目录，按序检查存在：
- `<图纸文件夹>\钣金件\钣金件.csv`（钣金 BOM）
- `<图纸文件夹>\型材\型材.csv`（型材 BOM）
- `<图纸文件夹>\型材采购.xlsx`（型材采购 BOM）
- `<图纸文件夹>\总外购件_YYYYMMDD_HHMMSS.xlsx` + `\IMG\`（外购件带等轴测截图）
- `<图纸文件夹>\标准件采购_YYYYMMDD_HHMMSS.xlsx` + `\IMG\`（标准件带截图）
- `<图纸文件夹>\钣金件\` 内全部钣金展开 DXF + `合并图纸_*.dxf` + `合并报告.txt`（总图已剔除焊接组）
- **每焊接组一份（2026-08-29）**：`<图纸文件夹>\钣金件\<焊接组名>\合并图纸_*.dxf` + 组内 `钣金件.csv`
- `<图纸文件夹>\型材\带数量零件\`（非焊接组，原位置）+ **`型材\带数量零件\<焊接组名>\`（每焊接组一份）**

缺哪个 → 定位到对应那一步的脚本与报错。详细档案见 `memory/2026-08-05.md`。

## 内部分步流程（仅供参考，不要手动逐条执行）

> 下面的 12 步是入口脚本 `sw_doc_process.py` 内部自动串行完成的，**调用方不要手动去跑**。

**分类宏建的子目录名是 钣金件/型材/外购件/标准件（注意带"件"字）；对应过滤值是 钣金件 / 型材。**

> **2026-08-10 内存优化（已内嵌）**：步骤A(虚拟件外部化+标记) 为首位；步骤1 已从宏改为**分批分类**（`sw_batch_categorize.py`，每批60件+批间自动重启SW）；关键步骤间自动重启 SW（0.8/2.5/5.5/7.5/9.5）防内存爆炸。16GB 机器实测 256 文件分类 0 失败。

| # | 步骤 | 脚本内部动作（勿手动执行） |
|---|---|---|
| A | 虚拟件外部化+标记 (2026-08-10) | `sw_virtual_extract_and_mark.py`：虚拟件 SaveVirtualComponent 外部化 + Pipe_/ST_ 标记外购件 |
| 0 | 图号名称更新 | `sw_批量更新图号_v2.py --asm 总装`；写入后 Save() 落盘 |
| 0.5 | 关闭装配体 | `sw.CloseDoc(总装标题)`，供分类 |
| 0.8 | 重启SW | 释放内存供分类 |
| 1 | 零件分类(分批) | `sw_batch_categorize.py --dir folder`：每批60件+批间重启SW → 建 4 子目录 |
| 2 | BOM三表一次导出 ×qty | `autobom_export.py --templates-multi "钣金:钣金件:…;型材:型材:…;型材采购:型材:…"` |
| 2.5 | 重启SW | 释放内存供截图 |
| 5 | 总外购件+截图 ×qty | `sw_buyout_export.py --template 总外购件 --filter-value 外购件 --qty N`（网址列已修） |
| 5.5 | 重启SW | 外购截图后必重启（96张截图也吃内存） |
| 6 | 标准件+截图 ×qty | `sw_buyout_export.py --template 标准件采购 --filter-value 标准件 --qty N` |
| 7 | 关闭装配体 | 释放 BOM 打开的装配体 |
| 7.5 | 重启SW | 释放内存供展开 |
| 8.0 | 焊接分组收集 (2026-08-29) | `sw_weld_group.py collect`：开总装递归，按「焊接或装配标记=焊接」装配体分组 → `焊接分组.json`（焊接组/非焊接组；嵌套焊接各自成组） |
| 8 | 钣金展开 DXF | `AutoBaiJinZhanKai.exe 批量 钣金件 dxf`；用产物校验非退出码 |
| 8.8 | 钣金分组+分组合并 (2026-08-29) | `sw_weld_group.py merge-ban`：每焊接组建 `钣金件\<组名>\`，移入该组 DXF + 写过滤配料csv + 各跑 dwg_merge 出独立合并图；根目录剩非焊接 DXF（含非焊接过滤csv）供步骤9 |
| 9 | 合并 DXF | `dwg_merge.py 钣金件`（取第一个 csv 作配料表；此时根目录已剔除焊接组） |
| 9.5 | 重启SW | 释放内存供型材宏 |
| 10.0 | 型材分组准备 (2026-08-29) | `sw_weld_group.py prep-profile`：每组建 `型材\_weld_tmp\<组名>\`，复制该组型材件 + 写过滤 `型材1.CSV`（__ungrouped__ = 非焊接组） |
| 10 | 型材合成宏(逐组) | 对 `_weld_tmp` 下每组跑 `sw_run_macro.py`（宏读各组目录的 型材1.CSV，输出各组 `带数量零件\`） |
| 10.9 | 型材分组汇总 (2026-08-29) | `sw_weld_group.py collect-profile`：非焊接组 → `带数量零件\`（原位置）；焊接组 → `带数量零件\<组名>\`；删 `_weld_tmp` |

## ★关键坑与修复点（2026-08-05/06 实测）

1. **skill1 必须 Save()**：曾缺 Save 导致属性只存内存、CloseDoc 后丢失 → `SZ.0704.10.0.0` 代号没改 → 合并 DXF 2个"代号不在CSV"跳过。已加 `doc.Save()` 落盘，合并从 7/9 → 9/9。
2. **装配体挂着锁分类副本**：总装 `SZ.0704.0` 开着 → 分类宏复制零件失败("未生成分类文件夹")。**步骤0.5 必须关装配体**（锁文件 54→0）。
3. **~$ 锁文件两类**：活进程持有 vs 孤儿残留(无进程,可删)。找总装时排除 `~$` 前缀(已内嵌)，避免把临时锁文件当总装。跑分类宏前先清理 SW 文档(SW 里曾堆 54 文档)。
4. **skill3 多模板一次导出**：`--templates-multi` 一次扫装配体出多表，3次独立 79s → 1次 34.6s。
5. **skill3 filter-prop**：须加进 plain_cols+all_cols，否则读不到过滤 0 行。
6. **close_asm 传标题**：`sw.CloseDoc(os.path.splitext(os.path.basename(p))[0])`，传路径静默失败。
7. **型材宏模块 `AutoBOM_型材批量处理`**（非 Macro51）；完成弹窗标题是 `AutoBOM 型材批量处理`，前台点弹窗线程须匹配 AutoBOM（已修）。
8. **钣金展开退出码不可靠**：AutoBaiJinZhanKai.exe 批量虽 rc=1 但实际成功生成（统计bug）。步骤8 用**产物校验**(dxf数>0, 剔除 合并图纸_*)。
9. **skill 双副本陷阱**：脚本统一指向全局 `C:\Users\mao\.openclaw\workspace\skills`（GLOBAL_SKILLS），勿混 turing-workspace\skills 本地副本。
10. **★弹窗守卫残留干扰（2026-08-08 修复）**：旧版守卫 `sw_popup_guard.py` 是 `while True` 死循环且只响应 Ctrl+C；skill7 跑完 terminate 不一定能让它退出，且旧 `stop_popup_guard()` 只杀本次 Popen 的那个守卫、只在正常完成路径调用 → 中途异常/被打断就残留成孤儿进程，**一直自动点掉/关闭 SW 弹窗，干扰皇上正常用 SW**。已双层修复：
    - `sw_popup_guard.py` 加 SIGTERM/SIGINT 信号处理（terminate 能优雅退出）；
    - skill7 `stop_popup_guard()` 升级为**全量清理**所有 `sw_popup_guard.py` 进程（含历史孤儿，按命令行匹配逐个 taskkill /F，不碰其它 python），并在 `__main__` 用 try/except 兜底——**任何退出路径（成功/异常/SystemExit）都确保守卫被清空，0 残留**。
11. **★虚拟件检测必须查临时路径（2026-08-10）**：虚拟零部件不是"无路径"！SW 给未保存虚拟件分配 `%TEMP%\swx*\VC~~\<虚拟夹>\DN15TK23^Pipe_2-xxx.SLDPRT` 临时路径。判断条件：①路径空 ②含 `\VC~~\`/`\temp\swx` ③名字含 `^`。外部化用 `IComponent2.SaveVirtualComponent(target)`（IComponent2 没有 SaveAs 方法），保存后自动改引用，再 `model.Save()`。
12. **★skill4 网址列（2026-08-10）**：collect_parts 读属性必须含 `网址`（部分件用 `采购链接`），网址列=`网址 or 采购链接`。属性用 `Extension.CustomPropertyManager(config).Get3(name,False,'','')`（config="默认<按加工>"）。
13. **★内存爆炸（2026-08-10）**：SW CloseDoc 后内存不还 OS，连续开几百文件 → 6-10GB 假死。方案：分类分批(每批60+重启SW) + 步骤间 restart_sw()（0.8/2.5/5.5/7.5/9.5）。截图步骤本身也吃内存——外购截图后必须 5.5 重启，否则标准件聚合打不开总装。
14. **★焊接装配体分组输出（2026-08-29 新增）**：按「焊接或装配标记=焊接」装配体分组（标记只标在装配体上，实测无零件单独标）。组名=装配体标题；嵌套焊接各自成组（零件归最近焊接祖先，A 组不含 B 的零件）；焊接组已从总合并图/总带数量零件**剔除**。坑：①型材宏查不到代号默认数量=2 → 组内 `型材1.CSV` 必须覆盖该组全部型材件 ②DXF 按代号匹配（`文件名-材质-板厚` 取 `-`前再 `_`前）③`钣金件.csv`/`型材1.CSV` 均 utf-8-sig（宏/合并读 utf-8-sig）④`_weld_tmp` 跑完自动删，异常也不残留 ⑤组内 csv 缺 DXF 只警告不阻断（dwg_merge 会报 missing_in_dxf）

## ★2026-08-06 三大运行前置根因（已内嵌进脚本，无需手动做）

脚本启动时会自动处理，**你无需也不要在外部手动做前置**：

1. **SW 必须以「正常方式」启动**（`Start-Process sldworks.exe`，非 `-Embedding`）→ exec 后台进程 COM 上下文才读得到桌面 SW 的 ROT，`GetActiveObject` 才连得上。脚本 `ensure_sw_normal_start()` 启动时自动：连不上就用 `Start-Process` 拉起 sldworks.exe 并轮询等待注册(≤120s)。
2. **PDM 库图纸的模态弹窗阻塞**（打开 D:\feedingHero 等 PDM 库图纸弹「未检出是否检出」/「另存为」/「完成提示」）→ 脚本自动后台拉起 `sw_popup_guard.py`（路径 `turing-workspace/scripts/sw_popup_guard.py`）常驻轮询自动点：PDM未检出点"否"、另存为点"取消"、完成提示点"确定"。**跑完保证 0 残留**（见关键坑 #10）：脚本正常完成/异常/中断都会全量清理守卫进程，绝不残留孤儿来干扰皇上 SW 的日常使用。
3. **`~$` 锁文件被误当总装** → 找总装已排除 `not f.startswith("~")`（见关键坑 #3）。

> 依赖缺失时（缺 sw_popup_guard.py）脚本会**拒绝启动**并给明确报错，不会带病跑。

## 依赖与前置

- SolidWorks（脚本会自动正常启动/连上，见上「三大运行前置根因」）；Python 64 位；pywin32 + openpyxl + Pillow + ezdxf；AutoBOM 插件 `C:\AutoBOM\AutoBaiJinZhanKai.exe`
- 依赖 skill（全局）：1-图号/3-导出BOM/4-导出BOM带截图/5-运行SW宏/6-合并DXF
- 宏路径固定：`Z:\solidworks information\macro\零件分类.swp` / `...\型材零件名称数量合成另存文件夹.swp`
- 目标文件夹需含总装配体 `.sldasm`（排除 ST_，代号最短者为总装）

## 校验（交付前看产物；深度核对规则如下）

先按上方「运行后要验证哪些产物」确认文件都在。
深度核对：DXF 数 = 钣金件数；**合并图纸 插入数 = 展开数（10=9+1）**；型材带数量零件有内容。任一缺失 → 定位对应脚本与报错。详细档案见 `memory/2026-08-05.md`。
