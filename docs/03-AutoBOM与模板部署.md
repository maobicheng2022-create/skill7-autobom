# 03 · AutoBOM 与 BOM 模板部署

## 1. AutoBOM 是什么

AutoBOM 是一个 SolidWorks 插件：在装配体里按模板列清单，把 `名称/代号/$总数量/$宽/$高/厚度/型材类型/…` 这类属性汇总成表，并支持三维截图列。

本仓库里的 `autobom\portable\` 是**便携版二进制**（不含反编译源码），目的是让流程开箱即用：

| 文件 | 作用 |
|---|---|
| `AutoBOM.exe` | 主界面（模板管理 / 看数 / 导出） |
| `AutoBOM.SW.dll` | SW 插件本体（被 SW 加载） |
| `AutoBaiJinZhanKai.exe` | 钣金批量展开（skill7 步骤 8 调用） |
| `*.dll` | NPOI / EPPlus / CSkin / Newtonsoft.Json / SharpZipLib / SQLite / SolidWorks Interop 等依赖 |
| `Plugin\` | 插件菜单配置、图标、`Addin_Config.json` |
| `icon\` | 界面图标 |
| `AutoBOM-SW-register.ps1` | 注册脚本（支持 `-InstallDir`） |

> 版权：AutoBOM 内核为第三方商业插件。本仓库**不包含其反编译源码**，只做便携打包与配置适配；商用请向原厂取得授权。详见 [NOTICE.md](../NOTICE.md)。

## 2. 一键部署

```powershell
# 管理员 PowerShell（注册表需要）
powershell -ExecutionPolicy Bypass -File tools\install.ps1
```

默认落到：

```
C:\AutoBOM\            AutoBOM 本体 + 依赖 + Plugin + icon
C:\micDatas\           pluginSettings\BOM\（模板 + 配置）
<仓库>\macros\          14 个 .swp
```

自定义路径：

```powershell
.\tools\install.ps1 -AutoBomDir D:\AutoBOM -MicDatasDir D:\micDatas -MacroDir "D:\macros"
```

只装文件、稍后再注册（例如没有管理员权限时）：

```powershell
.\tools\install.ps1 -SkipRegister
```

## 3. 插件注册（注册表三处）

`tools\install.ps1` 会自动调用，也可单独跑：

```powershell
# 管理员
powershell -ExecutionPolicy Bypass -File "C:\AutoBOM\AutoBOM-SW-register.ps1" -InstallDir "C:\AutoBOM"
```

它写三处：

| 位置 | 作用 |
|---|---|
| `HKLM\Software\SolidWorks\Addins\{E781EAD9-…}` | SW 插件列表里出现 AutoBOM |
| `HKCU\Software\SolidWorks\AddInsStartup\{E781EAD9-…}` | SW 启动时自动加载 |
| `HKCR\CLSID\{E781EAD9-…}\InprocServer32` | .NET COM 类指向 `AutoBOM.SW.dll`（`CodeBase` 用 URI 形式，路径含空格/中文也能解析） |

验证：

```powershell
[Type]::GetTypeFromCLSID([guid]'{E781EAD9-A4FA-4560-B0A2-994D349CB64A}')
```

能返回类型即注册成功。

## 4. BOM 模板（setting.json）

模板目录：`C:\micDatas\pluginSettings\BOM\`（可用 `micdatas_dir` 改）

| 文件 | 作用 |
|---|---|
| `setting.json` | 模板定义：每套模板的列 + 每列公式 + `CurrentTemplate` |
| `Template\*.xlsx` | 排版模板（列宽/表头/分页样式） |
| `CommonNode.json` / `SpecialNode.json` | 列节点/特殊节点配置 |

仓库自带 8 套模板，skill7 直接用其中这些：

| 模板 | 用途 |
|---|---|
| `钣金` | 钣金件清单（代号 + 总数量） |
| `型材` | 型材清单（代号/数量/宽/高/厚度/型材类型/**材质**） |
| `型材采购` | 型材采购（含 **材质** 列 + 多 `$长` 列，供 skill9 加计算列） |
| `总外购件` | 外购件（名称/型号规格/数量/单位/截图/网址） |
| `标准件采购` | 标准件（同上） |
| `焊接部件清点清单` | 焊接/装配分组清单（skill8 用） |
| `默认模板` / `型材外购件` | 备用 |

模板用法两条路：

1. **界面改**：打开 `AutoBOM.exe` → 设置 → 模板下拉 → 改列/公式 → 保存（写回 `setting.json`）
2. **直接改 JSON**：适合批量/版本管理，改完记得备份

> 模板列的 `$宽/$高` 在 AutoBOM 里是**包络尺寸**，型材类型公式依赖它们（方管/长方管/圆管判定）。
> 仓库里 `setting.json` 的公式已按"宽≠高→长方管；宽=25/32/38/42/…→圆管；否则方管"配好。

## 5. 版本备份（铁律）

任何 AutoBOM 更新/换 exe/换 dll/改配置前，先备份：

```powershell
python -X utf8 "skills\0-skill管理规范\scripts\autobom_backup.py"          # 备份
python -X utf8 "skills\0-skill管理规范\scripts\autobom_backup.py" --list   # 看快照
python -X utf8 "skills\0-skill管理规范\scripts\autobom_backup.py" --restore backup_20261117_101500
```

快照落在 `C:\AutoBOM\_backup\backup_YYYYMMDD_HHMMSS\`（exe/dll/config/Plugin/模板）。

`install.ps1` 覆盖安装时也会自动先快照，所以升级流程就是：

```powershell
git pull
powershell -ExecutionPolicy Bypass -File tools\install.ps1
python -X utf8 tools\verify_env.py
```

## 6. 常见问题

**SW 里看不到 AutoBOM 页签**

1. **先看是不是"只注册到 HKCU"**：SW 的插件列表**只读机器级** `HKLM\Software\SolidWorks\Addins\{GUID}`。
   非管理员运行时只能写 `HKCU\Software\SolidWorks\Addins`，注册了也**不会出现在 SW 里**，而且这条 HKCU 同名项还会盖住机器级那条。
   → 必须用**管理员**跑注册脚本；脚本（v1.0.1 起）会自动删除 HKCU 同名项。
2. `verify_env.py` 看 `SW 插件注册(HKLM\Addins)` / `COM 类注册(CLSID)` 是否 OK
3. 重启 SolidWorks（注册后必须重启）
4. 工具 → 插件 → 勾选 AutoBOM / 启动时加载
5. 还不行：`AutoBOM.SW.dll` 被移动过 → 重跑注册脚本（CodeBase 要指向真实路径）

**打开 AutoBOM 主界面报错**

- 缺 .NET Framework 4.8 → 装 4.8
- 缺 VC++ 运行库 → 装 `vc_redist.x64.exe`
- `Plugin\Addin_Config.json` 丢失 → 从 `autobom\portable\Plugin` 重新拷

**模板列对不上 / 导出少列**

- `setting.json` 与 `Template\*.xlsx` 不匹配 → 用仓库自带模板整体覆盖 `pluginSettings\BOM`
- 覆盖前先备份（见上）

**钣金展开没反应**

- `AutoBaiJinZhanKai.exe` 必须和 `AutoBOM.SW.dll` 在同一目录
- 该 exe 返回码不可靠（成功也可能 rc=1），skill7 用"产物校验"判断，不用退出码