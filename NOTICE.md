# 第三方组件声明（NOTICE）

本仓库**不是**一个纯粹的原创项目：为了让流程开箱可用，打包了若干第三方组件。
下面按目录说明来源与归属；许可证以各发布方的正式文本为准。

## 1. `autobom/portable/` — AutoBOM（第三方 SolidWorks 插件）

| 内容 | 说明 |
|---|---|
| `AutoBOM.exe` / `AutoBOM.SW.dll` / `Plugin/` / `icon/` | AutoBOM 插件，命名空间 `BigEngineer.SW`，**版权归其原作者/原厂** |
| `AutoBaiJinZhanKai.exe` | 钣金批量展开工具，随 AutoBOM 一同发布 |
| 其余同名依赖 DLL | 见下表 |

说明：

- 本仓库只做**便携化打包与配置适配**（路径参数化、模板整理、注册脚本），**不包含 AutoBOM 的反编译源码**。
- 本仓库作者与 AutoBOM 原厂无隶属关系，不主张其版权。
- **商用请自行向原厂购买/取得授权**；本仓库不提供、也不转授该授权。
- 若原厂认为本打包方式不妥，请联系仓库作者，会立即移除相关二进制。

## 2. 依赖库（`autobom/portable/*.dll`）

打包时识别的版本如下（`FileVersion`）：

| 组件 | 版本 | 常见许可证 |
|---|---|---|
| NPOI / NPOI.OOXML / NPOI.OpenXml4Net / NPOI.OpenXmlFormats | 2.4.1 | Apache License 2.0 |
| EPPlus | 4.5.3.2 | **LGPL-3.0**（4.x 分支；5.x 起改为 Polyform Noncommercial） |
| Newtonsoft.Json | 12.0.1 | MIT |
| ICSharpCode.SharpZipLib | 0.86.0 | GPL-2.0 **with classpath exception**（旧版） |
| CSkin | 16.1.14.3 | 免费 UI 库，版权归 CSkin 作者 |
| System.Data.SQLite | 1.0.102.0 | Public Domain |
| System.Buffers / System.Memory / System.Numerics.Vectors / System.Resources.Extensions / System.Runtime.CompilerServices.Unsafe | — | MIT（Microsoft） |
| GetUserIn.dll / PluginTrackingPoint.dll / solidworkstools.dll | — | 随 AutoBOM 发布，归属原厂 |
| SolidWorks.Interop.sldworks / swconst / swpublished（33.5.0.x） | SW 2025 | © Dassault Systèmes SolidWorks；使用需合法 SolidWorks 授权 |

> 这些二进制是为"让 AutoBOM 在本机跑起来"而随包分发的依赖，未做修改。
> 若你二次分发，请自行核对各库许可证（尤其 EPPlus 4.x 的 LGPL 与 SharpZipLib 0.86 的 GPL+例外条款）。

## 3. `templates/micDatas/`

- `setting.json`、`CommonNode.json`、`SpecialNode.json`、`Template/*.xlsx`：
  由本仓库作者基于 AutoBOM 模板机制配置/制作，随本仓库许可发布。
- 其中的公式（如"型材类型"判定）为本项目业务规则，不属于第三方。

## 4. `macros/*.swp`

14 个 SolidWorks VBA 宏由本仓库作者编写/维护，随本仓库许可发布。
宏内不包含第三方商业宏代码。

## 5. SolidWorks / Windows

- SolidWorks、SolidWorks Interop、COM 接口：© Dassault Systèmes，需合法授权。
- .NET Framework 4.8、Windows API：© Microsoft，按微软许可使用。

## 6. Python 依赖（不随仓库分发）

`pywin32`（PSF）、`openpyxl`（MIT）、`Pillow`（HPND）、`ezdxf`（MIT）。
由使用者通过 `pip install` 自行安装，适用各自许可证。

## 7. 商标

"SolidWorks"、"AutoCAD"、"WPS"、"Excel" 等为其各自所有者的商标，本仓库仅作说明性引用。

---

如你是上述任一组件的权利人，认为本仓库的打包方式侵犯了你的权益，请通过仓库 Issue 联系，
我们会在核实后立即调整或移除相关内容。