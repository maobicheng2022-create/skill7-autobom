# AutoBOM 一键安装包 v1.0.1

给 **另一台 Windows 电脑**把 AutoBOM 装进 SolidWorks 用的独立安装包（不需要装 Python、不需要克隆仓库）。

## v1.0.1 更新（2026-11-25）

- **修复"注册了但在 SW 里看不到页签"**：SW 的插件列表**只读机器级** `HKLM\Software\SolidWorks\Addins`；只写 `HKCU` 同名项不仅无效，还会盖住机器级那条。
  注册脚本现在会**自动删除 HKCU 同名项**，并明确提示必须管理员运行。
- `verify_env.py`：`HKLM\Addins` 缺失从"警告"改为"**失败**"，并新增"HKCU 同名注册冲突"检查。
- 文档补充：**SW 版本 interop 兼容性**（随包 DLL 按 SW2025 interop 编译，SW2022 上能加载但个别功能可能报"方法不存在"）。

## 下载

点下面的 **AutoBOM_SW_Installer_v1.0.1.zip** 下载（约 6 MB）。

## 安装（三步）

1. 把 zip **解压**到任意目录（不要在压缩包里直接双击）
2. 双击 **`安装AutoBOM.bat`**（会自动申请管理员权限，点"是"）
3. 重启 **SolidWorks** → `工具` → `插件` → 勾选 **AutoBOM**（并勾"启动时加载"）

## 包里有什么

| 内容 | 说明 |
|---|---|
| `payload/AutoBOM/` | 插件本体：`AutoBOM.exe`、`AutoBOM.SW.dll`、依赖 DLL、`Plugin/`、`icon/`、注册脚本 |
| `payload/micDatas/` | BOM 模板 8 套（默认 / 型材 / 钣金 / 型材采购 / 总外购件 / 标准件采购 / 型材外购件 / 焊接部件清点清单），含「型材类型」与「材质」公式 |
| `payload/macros/` | 14 个 SolidWorks 宏（型材合成、零件分类、焊接/装配分类等） |

安装脚本做的事：先把已有 AutoBOM 快照到 `C:\AutoBOM\_backup\backup_时间戳\`（可回滚）→ 复制文件 → 部署模板 → 部署宏 → 注册 SW 插件（HKLM/HKCU/CLSID 三处）→ 验证 COM。

## 环境要求

- Windows 10 / 11 **64 位**
- SolidWorks 2019 及以上（实测 2022 / 2025）
- .NET Framework 4.8（Win10 1903+ / Win11 自带）
- 管理员权限（写注册表）

## 装到别的位置

管理员 PowerShell：

```powershell
.\install_autobom.ps1 -AutoBomDir D:\AutoBOM -MicDatasDir D:\micDatas
```

宏默认装到 `<AutoBomDir>\macros`；想放共享盘用 `-MacroDir "\\nas\sw-macros"`。

## 验证

```powershell
[Type]::GetTypeFromCLSID([guid]'{E781EAD9-A4FA-4560-B0A2-994D349CB64A}')
```

能打印出类型名即注册成功。

## 说明 / 版权

- AutoBOM 内核为第三方商业 SolidWorks 插件，版权归原作者；本包仅做便携部署与配置适配，**不含反编译源码**，商用请自行向原厂取得授权。
- 只装这个包 = 获得 SW 里的 BOM 导出能力。要跑「图纸 → 采购/生产清单」全自动流程（图号/分类/BOM×套数/截图/展开/合并/型材按材质），请另外克隆完整仓库。

完整流程仓库：https://github.com/maobicheng2022-create/skill7-autobom