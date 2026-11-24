<#
.SYNOPSIS
    一键安装 AutoBOM 到 SolidWorks（自带自动提权）。

.DESCRIPTION
    从本包 payload 目录把下列内容装到本机并注册 SW 插件：
      payload\AutoBOM\*          -> C:\AutoBOM（可用 -AutoBomDir 改）
      payload\micDatas\pluginSettings -> C:\micDatas\pluginSettings（可用 -MicDatasDir 改）
      payload\macros\*.swp       -> C:\AutoBOM\macros（可选，skill7 的型材/分类宏）
    安装前会先把已有 AutoBOM 快照到 <AutoBomDir>\_backup\backup_时间戳\。

    需要：Windows 10/11 64 位、.NET Framework 4.8、管理员权限（脚本会自动弹 UAC）。
#>
param(
    [string]$AutoBomDir = "C:\AutoBOM",
    [string]$MicDatasDir = "C:\micDatas",
    [string]$MacroDir = "",
    [switch]$NoPause,
    [switch]$SkipRegister
)

$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSCommandPath
$Payload = Join-Path $Root 'payload'

function Say($msg, $color = 'Gray') { Write-Host $msg -ForegroundColor $color }

# ---------------- 自动提权 ----------------
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Say "需要管理员权限，正在重新以管理员身份启动..." Yellow
    $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$PSCommandPath`"")
    if ($AutoBomDir -ne 'C:\AutoBOM') { $argList += @('-AutoBomDir', "`"$AutoBomDir`"") }
    if ($MicDatasDir -ne 'C:\micDatas') { $argList += @('-MicDatasDir', "`"$MicDatasDir`"") }
    if ($MacroDir) { $argList += @('-MacroDir', "`"$MacroDir`"") }
    if ($NoPause) { $argList += '-NoPause' }
    try {
        Start-Process -FilePath 'powershell.exe' -ArgumentList $argList -Verb RunAs -Wait
    } catch {
        Say "提权失败（可能点了“否”）。请右键 安装AutoBOM.bat -> 以管理员身份运行。" Red
    }
    exit
}

if (-not (Test-Path $Payload)) {
    Say "找不到 payload 目录，请在解压后的完整文件夹里运行本脚本。" Red
    if (-not $NoPause) { Read-Host "按回车退出" }
    exit 1
}

Say "=== AutoBOM 安装 ===" Cyan
Say "  安装目录   : $AutoBomDir"
Say "  模板目录   : $MicDatasDir"
Say ""

# ---------------- 1. 备份旧版 ----------------
$oldExe = Join-Path $AutoBomDir 'AutoBOM.exe'
if (Test-Path $oldExe) {
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $bak = Join-Path $AutoBomDir "_backup\backup_$stamp"
    New-Item -ItemType Directory -Force -Path $bak | Out-Null
    Copy-Item (Join-Path $AutoBomDir '*.exe') $bak -Force -ErrorAction SilentlyContinue
    Copy-Item (Join-Path $AutoBomDir '*.dll') $bak -Force -ErrorAction SilentlyContinue
    Copy-Item (Join-Path $AutoBomDir '*.config') $bak -Force -ErrorAction SilentlyContinue
    if (Test-Path (Join-Path $AutoBomDir 'Plugin')) { Copy-Item (Join-Path $AutoBomDir 'Plugin') $bak -Recurse -Force }
    Say "  旧版已备份 -> $bak" Green
}

# ---------------- 2. 装 AutoBOM ----------------
New-Item -ItemType Directory -Force -Path $AutoBomDir | Out-Null
Copy-Item (Join-Path $Payload 'AutoBOM\*') $AutoBomDir -Recurse -Force
Say "  AutoBOM 文件已复制（exe / dll / Plugin / icon）" Green

# ---------------- 3. 装模板 ----------------
$tplSrc = Join-Path $Payload 'micDatas\pluginSettings'
if (Test-Path $tplSrc) {
    $tplDst = Join-Path $MicDatasDir 'pluginSettings'
    $setting = Join-Path $tplDst 'BOM\setting.json'
    if (Test-Path $setting) {
        $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        Copy-Item $setting "$setting.bak_$stamp" -Force
        Say "  旧 setting.json 已备份" Green
    }
    New-Item -ItemType Directory -Force -Path $tplDst | Out-Null
    Copy-Item (Join-Path $tplSrc '*') $tplDst -Recurse -Force
    Say "  BOM 模板（8 套）已就位" Green
}

# ---------------- 4. 装宏（可选） ----------------
if (-not $MacroDir) { $MacroDir = Join-Path $AutoBomDir 'macros' }
$macroSrc = Join-Path $Payload 'macros'
if (Test-Path $macroSrc) {
    New-Item -ItemType Directory -Force -Path $MacroDir | Out-Null
    Copy-Item (Join-Path $macroSrc '*.swp') $MacroDir -Force
    Say ("  SW 宏 {0} 个 -> {1}" -f (@(Get-ChildItem (Join-Path $macroSrc '*.swp')).Count, $MacroDir)) Green
}

# ---------------- 5. 注册 SW 插件 ----------------
if (-not $SkipRegister) {
    Say "`n=== 注册 SolidWorks 插件 ===" Cyan
    $reg = Join-Path $AutoBomDir 'AutoBOM-SW-register.ps1'
    if (-not (Test-Path $reg)) { $reg = Join-Path $Payload 'AutoBOM\AutoBOM-SW-register.ps1' }
    & $reg -InstallDir $AutoBomDir -NoPause
}

# ---------------- 6. 验证 ----------------
$guid = [guid]'{E781EAD9-A4FA-4560-B0A2-994D349CB64A}'
$t = [Type]::GetTypeFromCLSID($guid)
Say ""
if ($t) {
    Say "✅ 安装完成：COM 已注册（$($t.Assembly.GetName().Name)）" Green
    Say "  下一步：重启 SolidWorks -> 工具 -> 插件 -> 勾选 AutoBOM" Cyan
} else {
    Say "!! 文件已装好，但 COM 注册验证失败。" Yellow
    Say "   请确认是以管理员身份运行，并检查 .NET Framework 4.8 是否安装。" Yellow
}
if (-not $NoPause) { Read-Host "`n按回车退出" }