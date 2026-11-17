# ============================================================
#  AutoBOM SW 插件注册脚本
#
#  用法（需管理员）：
#    右键此文件 -> "使用 PowerShell 运行"
#    或： powershell -ExecutionPolicy Bypass -File .\AutoBOM-SW-register.ps1 -InstallDir "C:\AutoBOM"
#
#  只写三处注册表：
#    HKLM\Software\SolidWorks\Addins\{GUID}         SW 插件列表
#    HKCU\Software\SolidWorks\AddInsStartup\{GUID}  启动时自动加载
#    HKCR\CLSID\{GUID}\InprocServer32               .NET COM 类指向 AutoBOM.SW.dll
# ============================================================
param(
    [string]$InstallDir = "C:\AutoBOM",
    [switch]$NoPause
)

$ErrorActionPreference = 'Stop'

$g     = '{E781EAD9-A4FA-4560-B0A2-994D349CB64A}'
$class = 'BigEngineer.SW.SwAddin'
$asm   = 'AutoBOM.SW, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null'
$dll   = Join-Path $InstallDir 'AutoBOM.SW.dll'
$rv    = 'v4.0.30319'

if (-not (Test-Path $dll)) {
    Write-Host "找不到 $dll" -ForegroundColor Red
    Write-Host "请先部署 AutoBOM 文件，或用 -InstallDir 指定实际安装目录。" -ForegroundColor Yellow
    if (-not $NoPause) { Read-Host "按回车退出" }
    exit 1
}

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "需要管理员权限：右键 -> 使用 PowerShell 以管理员身份运行" -ForegroundColor Red
    if (-not $NoPause) { Read-Host "按回车退出" }
    exit 1
}

# file:///C:/Program%20Files/... 形式，含空格/中文也能被 COM 正确解析
$codebase = ([System.Uri]$dll).AbsoluteUri

Write-Host "=== 注册 AutoBOM SW 插件 ===" -ForegroundColor Cyan
Write-Host "  DLL: $dll" -ForegroundColor DarkGray

# --- 1. HKLM\Software\SolidWorks\Addins\{GUID} ---
$lk = "HKLM:\Software\SolidWorks\Addins\$g"
New-Item -Path $lk -Force | Out-Null
New-ItemProperty -Path $lk -Name '(default)'   -Value 0         -PropertyType DWord -Force | Out-Null
New-ItemProperty -Path $lk -Name 'Description' -Value 'AutoBOM' -Force | Out-Null
New-ItemProperty -Path $lk -Name 'Title'       -Value 'AutoBOM' -Force | Out-Null
Write-Host "  OK  HKLM\Software\SolidWorks\Addins\$g" -ForegroundColor Green

# --- 2. HKCU\Software\SolidWorks\AddInsStartup\{GUID} ---
$sk = "HKCU:\Software\SolidWorks\AddInsStartup\$g"
New-Item -Path $sk -Force | Out-Null
New-ItemProperty -Path $sk -Name '(default)' -Value 1 -PropertyType DWord -Force | Out-Null
Write-Host "  OK  HKCU\Software\SolidWorks\AddInsStartup\$g  (=1 启动加载)" -ForegroundColor Green

# --- 3. HKCR\CLSID\{GUID}\InprocServer32 ---
$clsid = "Registry::HKEY_CLASSES_ROOT\CLSID\$g"
New-Item -Path "$clsid"                        -Force | Out-Null
New-ItemProperty -Path $clsid -Name '(default)' -Value $class -Force | Out-Null
New-Item -Path "$clsid\InprocServer32"         -Force | Out-Null
New-ItemProperty -Path "$clsid\InprocServer32" -Name '(default)'      -Value 'mscoree.dll' -Force | Out-Null
New-ItemProperty -Path "$clsid\InprocServer32" -Name 'ThreadingModel' -Value 'Both'        -Force | Out-Null
New-ItemProperty -Path "$clsid\InprocServer32" -Name 'Class'          -Value $class        -Force | Out-Null
New-ItemProperty -Path "$clsid\InprocServer32" -Name 'Assembly'       -Value $asm          -Force | Out-Null
New-ItemProperty -Path "$clsid\InprocServer32" -Name 'RuntimeVersion' -Value $rv           -Force | Out-Null
New-ItemProperty -Path "$clsid\InprocServer32" -Name 'CodeBase'       -Value $codebase     -Force | Out-Null
New-Item -Path "$clsid\ProgId" -Force | Out-Null
New-ItemProperty -Path "$clsid\ProgId" -Name '(default)' -Value $class -Force | Out-Null
New-Item -Path "$clsid\Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}" -Force | Out-Null
Write-Host "  OK  HKCR\CLSID\$g" -ForegroundColor Green

# --- 验证 ---
$t = [Type]::GetTypeFromCLSID([guid]$g)
if ($t) {
    Write-Host ""
    Write-Host "注册成功：COM 激活 -> $($t.FullName)（$($t.Assembly.GetName().Name)）" -ForegroundColor Green
} else {
    Write-Host ""
    Write-Host "注册完成但 COM 验证失败，检查 DLL 是否被 SolidWorks 锁定/版本不符" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "下一步：重启 SolidWorks -> 工具 -> 插件 -> 勾选 AutoBOM" -ForegroundColor Cyan
if (-not $NoPause) { Read-Host "按回车退出" }