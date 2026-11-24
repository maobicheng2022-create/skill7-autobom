# skill7-autobom

把 SolidWorks 总装配体一键变成**采购 / 生产能直接用的清单**：图号名称修正 → 零件分类 → 多模板 BOM（×套数）→ 三维截图 → 焊接/装配分组 → 钣金展开 → DXF 合并 → 型材合成。

一套图纸（一个总装 + 若干零件）丢进来，一条命令跑完，产出一批 xlsx/csv/dxf：

| 产物 | 内容 |
|---|---|
| `钣金件\钣金件.csv` | 钣金零件清单（代号 + 数量×套数） |
| `钣金件\*.dxf` + `合并图纸_*.dxf` + `合并报告.txt` | 每块钣金展开图、按配料合并后的产线切割图 |
| `型材\型材.csv`、`型材采购.xlsx` | 型材零件清单 + 采购表（含每套支数/一根6m可开数/余料） |
| `整理型材_*.xlsx` | 按规格归并 + 6m 管根数 |
| `总外购件_*.xlsx` + `IMG\` | 外购件清单，每行嵌三维等轴测截图 |
| `标准件采购_*.xlsx` + `IMG\` | 标准件清单（含截图） |
| `焊接部件清单_*.xlsx` / `装配部件清单_*.xlsx` | 按「焊接或装配标记」分组 |
| `型材\带数量零件\`（每焊接组一份） | 型材按组拆好的零件副本 |

这套东西的定位：**不是通用 BOM 插件**，而是一条经过真实订单打磨的「图纸 → 采购/生产」流水线，附带把第三方 SolidWorks BOM 插件（AutoBOM）便携化部署。

---

## 1. 环境要求

| 项 | 要求 | 说明 |
|---|---|---|
| 系统 | Windows 10/11 64 位 | 依赖 SolidWorks COM，仅 Windows |
| SolidWorks | 2019 及以上（实测 2022 / 2025） | 必须能正常启动、能跑 VBA 宏 |
| Python | **64 位 3.10+**（实测 3.14） | 32 位连不上 SolidWorks COM |
| Python 包 | `pywin32 openpyxl Pillow ezdxf` | `pip install pywin32 openpyxl Pillow ezdxf` |
| AutoBOM | 仓库自带便携版（第三方插件） | 见 [docs/03](docs/03-AutoBOM与模板部署.md) |
| 宏 | 仓库自带 14 个 `.swp` | SW 需要能加载宏（默认即可） |
| AutoCAD | 可选 | 只有把 `.dwg` 转 `.dxf` 时才用；纯 `.dxf` 流程不需要 |

> 不需要联网、不需要数据库、不需要装 Office（WPS/Excel 只在你自己打开产物时用）。

## 2. 五分钟部署

> **只需要 AutoBOM 装到 SolidWorks（不需要跑流程）**：直接下载一键安装包，解压后双击 `安装AutoBOM.bat` 即可，不用装 Python：
> **https://github.com/maobicheng2022-create/skill7-autobom/releases/latest**
> 包内含 AutoBOM 本体 + 8 套 BOM 模板 + 14 个 SW 宏 + 自动注册脚本（自动申请管理员权限）。

```powershell
git clone https://github.com/maobicheng2022-create/skill7-autobom.git
cd skill7-autobom

# 装 AutoBOM + BOM 模板 + 宏，注册 SW 插件（建议管理员 PowerShell）
powershell -ExecutionPolicy Bypass -File tools\install.ps1

# 自检：Python/依赖/AutoBOM/模板/宏/插件注册 一次问清
python -X utf8 tools\verify_env.py
```

`install.ps1` 做的事：

1. 若已装过 AutoBOM → 先快照到 `C:\AutoBOM\_backup\backup_时间戳\`（改版可回滚）
2. 部署 `autobom\portable\*` → `C:\AutoBOM`
3. 部署 `templates\micDatas\pluginSettings\*` → `C:\micDatas`
4. 部署 `macros\*.swp` → 宏目录（默认仓库 `macros\`，可 `-MacroDir` 改）
5. 注册 SW 插件（HKLM/HKCU/CLSID 三处）
6. 写 `config\config.json` 并跑 `verify_env.py`

换路径部署：

```powershell
.\tools\install.ps1 -AutoBomDir D:\AutoBOM -MicDatasDir D:\micDatas -MacroDir "D:\macros"
```

装完重启 SolidWorks → **工具 → 插件 → 勾选 AutoBOM**。

## 3. 跑一条线

```powershell
# 前台（小单，能看实时进度）
python -X utf8 "skills\7-sw文档综合处理\scripts\sw_doc_process.py" "D:\图纸\SZ.0203.0_手动切换三通阀_做10套" --qty 10
```

大单（1 小时以上）用封装脚本，日志落盘、可续跑：

```powershell
.\tools\run_skill7.ps1 -Folder "D:\图纸\某设备_做10套" -Qty 10
```

中断了不用重跑全链：

```powershell
# 只从步骤 9 接着跑
.\tools\run_skill7.ps1 -Folder "D:\图纸\某设备_做10套" -Qty 10 -From 9
```

**`--qty` 只乘各 BOM 表的 `$总数量`**；合并 DXF / 型材宏复用已乘过的 csv，不会二次乘。

详细步骤表、断点号、失败定位见 [docs/04-运行skill7.md](docs/04-运行skill7.md)。

## 4. 目录结构

```
skill7-autobom/
├─ README.md
├─ LICENSE                    # 自定义许可：保留所有权利，禁止商用
├─ NOTICE.md                  # 第三方组件声明（AutoBOM / 依赖库 / 宏）
├─ config/
│  ├─ config.example.json     # 配置模板
│  └─ config.json             # 本机配置（gitignore，install.ps1 生成）
├─ shared/
│  └─ skill7_paths.py         # 路径中枢：AutoBOM / 模板 / 宏 / SW / 截图库
├─ skills/
│  ├─ 0-skill管理规范/
│  ├─ 1-图号名称更新/          ├─ 7-sw文档综合处理/   ← 总入口
│  ├─ 2-钣金展开/              ├─ 8-焊接装配标志分类输出BOM带截图/
│  ├─ 3-导出BOM/               ├─ 9-整理型材/
│  ├─ 4-导出BOM带截图/          ├─ 10-虚拟件外部化与Pipe标记/
│  ├─ 5-运行SW宏/              └─ 11-PVC管采购计算/
│  └─ 6-合并DXF/
├─ autobom/portable/          # AutoBOM 便携版（exe/dll/Plugin/依赖）
├─ macros/                    # 14 个 SolidWorks 宏
├─ templates/micDatas/        # BOM 模板（8 套）+ setting.json
├─ tools/
│  ├─ install.ps1             # 一键部署
│  ├─ run_skill7.ps1          # 一键跑 + 日志 + 续跑
│  └─ verify_env.py           # 环境自检
├─ docs/                      # 部署与排障文档
└─ build/                     # 维护脚本（路径参数化、文档本地化）
```

## 5. 配置

所有路径集中在 `config\config.json`（也可用环境变量 `SKILL7_*` 覆盖）：

| 键 | 默认值 | 说明 |
|---|---|---|
| `solidworks_exe` | 空 = 自动探测 | 注册表里多版本并存时建议显式指定 |
| `autobom_dir` | `C:\AutoBOM` | AutoBOM 安装目录 |
| `micdatas_dir` | `C:\micDatas` | BOM 模板目录 |
| `shot_cache_dir` | 空 = 自动探测 | 截图/属性缓存库（多份时选索引最新） |
| `macro_dir` | 仓库 `macros\` | SW 宏目录；想用 Z 盘共享宏就改这里 |
| `acad_exe` | 空 = 自动探测 | 仅 `.dwg → .dxf` 时需要 |
| `pdm_library_markers` | `[]` | PDM 只读库路径关键字（如 `["d:\\pdm"]`），命中则走"重定向/跳过"逻辑 |

完整优先级与示例见 [docs/02-配置说明.md](docs/02-配置说明.md)。

## 6. 单跑某个 skill

每个 skill 都能独立调用，互不强绑：

| 想做的事 | 命令 |
|---|---|
| 修正图号/名称 | `python -X utf8 "skills\1-图号名称更新\scripts\sw_批量更新图号_v2.py"` |
| 导一张 BOM | `python -X utf8 "skills\3-导出BOM\scripts\autobom_export.py" --asm X.SLDASM --template 型材 --filter-prop 采购类别 --filter-value 型材 -o out.csv` |
| 带截图的采购表 | `python -X utf8 "skills\4-导出BOM带截图\scripts\sw_buyout_export.py" --dir <文件夹> -o <输出目录>` |
| 合并 DXF | `python -X utf8 "skills\6-合并DXF\scripts\dwg_merge.py" <钣金件目录>` |
| 整理型材 | `python -X utf8 "skills\9-整理型材\scripts\xingcai_summary.py" "<型材采购.xlsx>" --sets 10` |
| 虚拟件外部化 | `python -X utf8 "skills\10-虚拟件外部化与Pipe标记\scripts\sw_virtual_extract_and_mark.py" --dir <文件夹>` |
| PVC 管采购 | `python -X utf8 "skills\11-PVC管采购计算\scripts\sw_pvc_purchase.py" --dir <文件夹> --qty 10` |

各 skill 的完整参数见对应 `skills\<编号>\SKILL.md`。

## 7. 文档

- [01-环境要求](docs/01-环境要求.md)
- [02-配置说明](docs/02-配置说明.md)
- [03-AutoBOM 与 BOM 模板部署](docs/03-AutoBOM与模板部署.md)
- [04-运行 skill7](docs/04-运行skill7.md)
- [05-产物清单与校验](docs/05-产物清单与校验.md)
- [06-常见问题](docs/06-常见问题.md)

## 8. 许可与第三方

- **本仓库自有脚本与文档**：保留所有权利，**禁止商业使用/转卖**，详见 [LICENSE](LICENSE)。
- **AutoBOM（`autobom/portable`）**：第三方 SolidWorks BOM 插件，版权归原作者；本仓库仅做便携化打包与配置适配，**不包含反编译源码**。商用请自行向原厂取得授权，详见 [NOTICE.md](NOTICE.md)。
- **依赖库**（NPOI / EPPlus / CSkin / Newtonsoft.Json / SharpZipLib / SolidWorks Interop 等）：各自原许可证，见 NOTICE。
- **SolidWorks 宏**（`macros\*.swp`）：由本仓库作者编写，随本许可发布。

---

## English quick start

A Windows-only SolidWorks pipeline that turns an assembly into purchasable BOMs (sheet metal CSV/DXF, profiles, buyout/standard parts with isometric screenshots, weld/assembly groups). Requires 64-bit Python 3.10+, `pywin32 openpyxl Pillow ezdxf`, SolidWorks 2019+, and the bundled AutoBOM add-in.

```powershell
git clone https://github.com/maobicheng2022-create/skill7-autobom.git
cd skill7-autobom
powershell -ExecutionPolicy Bypass -File tools\install.ps1     # deploy AutoBOM + templates + macros
python -X utf8 tools\verify_env.py                              # self-check
python -X utf8 "skills\7-sw文档综合处理\scripts\sw_doc_process.py" "<drawing folder>" --qty 10
```

See `docs/` for details (Chinese).