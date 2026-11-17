<#
.SYNOPSIS
    skill7 一键部署：把 AutoBOM、BOM 模板、SW 宏装到本机，并注册 SW 插件。

.DESCRIPTION
    默认目标：
      AutoBOM   -> C:\AutoBOM      （可用 -AutoBomDir 改）
      模板/配置 -> C:\micDatas      （可用 -MicDatasDir 改）
      宏        -> 仓库 macros\     （可用 -MacroDir 改，例如 D:\macros）
    装完会写 config\config.json 并跑 tools\verify_env.py。

    升级覆盖前会自动把现有 AutoBOM 快照到 <AutoBomDir>\_backup\backup_时间戳\（钦定铁律：
    任何 exe/dll/配置更新先备份）。

.EXAMPLE
    # 默认路径部署（右键"使用 PowerShell 运行"，需管理员）
    .\tools\install.ps1

.EXAMPLE
    # 装到 D 盘，宏放 Z 盘
    .\tools\install.ps1 -AutoBomDir D:\AutoBOM -MicDatasDir D:\micDatas -MacroDir "D:\macros"
#>
param(
    [string]$AutoBomDir = "C:\AutoBOM",
    [string]$MicDatasDir = "C:\micDatas",
    [string]$MacroDir = "",
    [switch]$SkipRegister,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$RepoRoot = Split-Path -Parent $PSScriptRoot
$Portable = Join-Path $RepoRoot 'autobom\portable'
$Templates = Join-Path $RepoRoot 'templates\micDatas\pluginSettings'
$DefaultMacros = Join-Path $RepoRoot 'macros'

function Write-Step($msg) { Write-Host "`n=== $msg ===" -ForegroundColor Cyan }
function Write-Ok($msg) { Write-Host "  OK  $msg" -ForegroundColor Green }
function Write-Warn($msg) { Write-Host "  !   $msg" -ForegroundColor Yellow }

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not (Test-Path $Portable)) { throw "找不到 $Portable，请在仓库根目录运行本脚本" }

# ---------------------------------------------------------------- 0. 备份旧 AutoBOM
Write-Step "备份现有 AutoBOM（如存在）"
$existingExe = Join-Path $AutoBomDir 'AutoBOM.exe'
if ((Test-Path $existingExe) -and -not $Force) {
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $bak = Join-Path $AutoBomDir "_backup\backup_$stamp"
    New-Item -ItemType Directory -Force -Path $bak | Out-Null
    Copy-Item (Join-Path $AutoBomDir '*.exe') $bak -Force -ErrorAction SilentlyContinue
    Copy-Item (Join-Path $AutoBomDir '*.dll') $bak -Force -ErrorAction SilentlyContinue
    Copy-Item (Join-Path $AutoBomDir '*.config') $bak -Force -ErrorAction SilentlyContinue
    if (Test-Path (Join-Path $AutoBomDir 'Plugin')) {
        Copy-Item (Join-Path $AutoBomDir 'Plugin') $bak -Recurse -Force
    }
    Write-Ok "旧版快照 -> $bak"
} else {
    Write-Warn "未发现旧版 AutoBOM 或指定了 -Force，跳过备份"
}

# ---------------------------------------------------------------- 1. AutoBOM
Write-Step "部署 AutoBOM -> $AutoBomDir"
New-Item -ItemType Directory -Force -Path $AutoBomDir | Out-Null
Copy-Item (Join-Path $Portable '*') $AutoBomDir -Recurse -Force
Write-Ok "AutoBOM.exe / AutoBOM.SW.dll / 依赖 DLL / Plugin / icon 已就位"

# ---------------------------------------------------------------- 2. 模板
Write-Step "部署 BOM 模板 -> $MicDatasDir\pluginSettings"
New-Item -ItemType Directory -Force -Path (Split-Path -Parent $Templates) | Out-Null
if ((Test-Path (Join-Path $MicDatasDir 'pluginSettings\BOM\setting.json')) -and -not $Force) {
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $bak = Join-Path $MicDatasDir "pluginSettings\BOM\setting.json.bak_$stamp"
    Copy-Item (Join-Path $MicDatasDir 'pluginSettings\BOM\setting.json') $bak -Force
    Write-Ok "旧 setting.json 已备份 -> $bak"
}
Copy-Item $Templates (Join-Path $MicDatasDir 'pluginSettings') -Recurse -Force
Write-Ok "setting.json + CommonNode / SpecialNode + Template\*.xlsx 已就位"

# ---------------------------------------------------------------- 3. 宏
if ([string]::IsNullOrWhiteSpace($MacroDir)) { $MacroDir = $DefaultMacros }
Write-Step "部署 SW 宏 -> $MacroDir"
New-Item -ItemType Directory -Force -Path $MacroDir | Out-Null
Copy-Item (Join-Path $DefaultMacros '*.swp') $MacroDir -Force
$macroCount = (Get-ChildItem $MacroDir -Filter *.swp).Count
Write-Ok "$macroCount 个 .swp 宏已就位"

# ---------------------------------------------------------------- 4. 注册 SW 插件
if (-not $SkipRegister) {
    Write-Step "注册 AutoBOM SW 插件"
    if (-not $isAdmin) {
        Write-Warn "需要管理员权限才能写 HKLM/CLSID 注册表；请以管理员身份重跑："
        Write-Warn "  .\tools\install.ps1 -SkipRegister   # 先装文件，再单独跑注册脚本"
        Write-Warn "或直接右键 autobom\portable\AutoBOM-SW-register.ps1 -> 以管理员运行"
    } else {
        & (Join-Path $Portable 'AutoBOM-SW-register.ps1') -InstallDir $AutoBomDir
    }
}

# ---------------------------------------------------------------- 5. 写本机配置
Write-Step "生成本机配置 config\config.json"
$cfgPath = Join-Path $RepoRoot 'config\config.json'
$cfg = [ordered]@{
    solidworks_exe      = ''
    autobom_dir         = $AutoBomDir
    micdatas_dir        = $MicDatasDir
    shot_cache_dir      = (Join-Path $AutoBomDir 'ShotCache')
    macro_dir           = $MacroDir
    acad_exe            = ''
    pdm_library_markers = @()
}
if (Test-Path $cfgPath) {
    $old = Get-Content $cfgPath -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($k in @($cfg.Keys)) {
        if ($old.PSObject.Properties.Name -contains $k -and $old.$k) { $cfg[$k] = $old.$k }
    }
    $cfg['autobom_dir'] = $AutoBomDir
    $cfg['micdatas_dir'] = $MicDatasDir
    $cfg['macro_dir'] = $MacroDir
    Write-Ok "已保留原 config 中的自定义项，其余按本次部署更新"
}
New-Item -ItemType Directory -Force -Path (Split-Path -Parent $cfgPath) | Out-Null
$cfg | ConvertTo-Json -Depth 4 | Set-Content $cfgPath -Encoding UTF8
Write-Ok $cfgPath

# ---------------------------------------------------------------- 6. 自检
Write-Step "环境自检"
& python -X utf8 (Join-Path $RepoRoot 'tools\verify_env.py')
if ($LASTEXITCODE -ne 0) {
    Write-Warn "自检未全过，按上面 [失败] 行修；修完重跑：python -X utf8 tools\verify_env.py"
} else {
    Write-Host "`n部署完成。SolidWorks 里点开 AutoBOM 插件页签即可看到 BOM 工具。" -ForegroundColor Green
}