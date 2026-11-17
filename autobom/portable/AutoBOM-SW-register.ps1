# ============================================================
#  AutoBOM SW 插件注册脚本
#  使用方法: 右键此文件 -> "使用 PowerShell 运行" (需管理员)
#  或在管理员 PowerShell 中执行:  & .\AutoBOM-SW-register.ps1
# ============================================================
$ErrorActionPreference = 'Stop'

# 检查是否管理员
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "❌ 需要管理员权限！请右键 -> 使用 PowerShell 以管理员身份运行" -ForegroundColor Red
    Read-Host "按回车退出"
    exit 1
}

$g        = '{E781EAD9-A4FA-4560-B0A2-994D349CB64A}'
$class    = 'BigEngineer.SW.SwAddin'
$asm      = 'AutoBOM.SW, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null'
$dll      = 'C:/AutoBOM/AutoBOM.SW.dll'
$codebase = "file:///$dll"
$rv       = 'v4.0.30319'

Write-Host "=== 注册 AutoBOM SW 插件 (GUID $g) ===" -ForegroundColor Cyan

# --- 1. HKLM\Software\SolidWorks\Addins\{GUID}  (SW 插件列表) ---
$lk = "HKLM:\Software\SolidWorks\Addins\$g"
New-Item -Path $lk -Force | Out-Null
New-ItemProperty -Path $lk -Name '(default)'   -Value 0        -PropertyType DWord -Force | Out-Null
New-ItemProperty -Path $lk -Name 'Description' -Value 'AutoBOM' -Force | Out-Null
New-ItemProperty -Path $lk -Name 'Title'       -Value 'AutoBOM' -Force | Out-Null
Write-Host "  OK  HKLM\Software\SolidWorks\Addins\$g  (Description=AutoBOM, Title=AutoBOM)" -ForegroundColor Green

# --- 2. HKCU\Software\SolidWorks\AddInsStartup\{GUID}  (启动加载开关) ---
$sk = "HKCU:\Software\SolidWorks\AddInsStartup\$g"
New-Item -Path $sk -Force | Out-Null
New-ItemProperty -Path $sk -Name '(default)' -Value 1 -PropertyType DWord -Force | Out-Null
Write-Host "  OK  HKCU\Software\SolidWorks\AddInsStartup\$g  (=1 启动加载)" -ForegroundColor Green

# --- 3. HKCR\CLSID\{GUID} + InprocServer32  (.NET COM 类注册) ---
$clsid = "Registry::HKEY_CLASSES_ROOT\CLSID\$g"
New-Item -Path "$clsid"                       -Force | Out-Null
New-ItemProperty -Path $clsid -Name '(default)' -Value $class -Force | Out-Null
New-Item -Path "$clsid\InprocServer32"        -Force | Out-Null
New-ItemProperty -Path "$clsid\InprocServer32" -Name '(default)'      -Value 'mscoree.dll' -Force | Out-Null
New-ItemProperty -Path "$clsid\InprocServer32" -Name 'ThreadingModel' -Value 'Both'        -Force | Out-Null
New-ItemProperty -Path "$clsid\InprocServer32" -Name 'Class'          -Value $class        -Force | Out-Null
New-ItemProperty -Path "$clsid\InprocServer32" -Name 'Assembly'       -Value $asm          -Force | Out-Null
New-ItemProperty -Path "$clsid\InprocServer32" -Name 'RuntimeVersion' -Value $rv           -Force | Out-Null
New-ItemProperty -Path "$clsid\InprocServer32" -Name 'CodeBase'       -Value $codebase     -Force | Out-Null
New-Item -Path "$clsid\ProgId" -Force | Out-Null
New-ItemProperty -Path "$clsid\ProgId" -Name '(default)' -Value $class -Force | Out-Null
New-Item -Path "$clsid\Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}" -Force | Out-Null
Write-Host "  OK  HKCR\CLSID\$g  (AutoBOM.SW.dll COM 类)" -ForegroundColor Green

# --- 验证 ---
$t = [Type]::GetTypeFromCLSID([guid]$g)
if ($t) {
    Write-Host ""
    Write-Host "✅ 注册成功！COM 激活: $($t.FullName) (from $($t.Assembly.GetName().Name))" -ForegroundColor Green
} else {
    Write-Host "⚠️ 注册完成但 COM 验证失败，请检查 C:\AutoBOM\AutoBOM.SW.dll 是否存在" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "下一步: 打开(或重启) SOLIDWORKS, 在工具->插件 里勾选 AutoBOM" -ForegroundColor Cyan
Read-Host "按回车退出"
