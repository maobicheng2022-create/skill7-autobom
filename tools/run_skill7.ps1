<#
.SYNOPSIS
    skill7 一键跑一条线（后台 + 日志 + 计时），适合大单（1 小时以上）。

.DESCRIPTION
    只是 sw_doc_process.py 的薄封装：
      - 检查图纸文件夹/qty
      - 日志写到 logs\skill7_<文件夹名>_<时间>.log
      - 前台等待并实时显示最后几行（Ctrl+C 可中断，产物已生成的部分可 --from 续跑）

    真正的流程编排全部在 sw_doc_process.py 内部，这里不改任何业务逻辑。

.EXAMPLE
    .\tools\run_skill7.ps1 -Folder "D:\图纸\SZ.0203.0_手动切换三通阀_做10套" -Qty 10

.EXAMPLE
    # 中断后续跑步骤 9
    .\tools\run_skill7.ps1 -Folder "...\某图纸文件夹" -Qty 10 -From 9
#>
param(
    [Parameter(Mandatory = $true)][string]$Folder,
    [Parameter(Mandatory = $true)][int]$Qty,
    [string]$Asm = "",
    [string]$From = "",
    [string]$To = "",
    [string]$Skip = "",
    [string]$Python = "python"
)

$ErrorActionPreference = 'Stop'
$RepoRoot = Split-Path -Parent $PSScriptRoot
$Entry = Join-Path $RepoRoot 'skills\7-sw文档综合处理\scripts\sw_doc_process.py'
$LogDir = Join-Path $RepoRoot 'logs'

if (-not (Test-Path $Folder)) { throw "图纸文件夹不存在: $Folder" }
if ($Qty -lt 1) { throw "套数必须 >= 1" }
if (-not (Test-Path $Entry)) { throw "找不到入口脚本: $Entry" }

New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
$safe = [IO.Path]::GetFileNameWithoutExtension($Folder) -replace '[\\/:*?"<>|]', '_'
$log = Join-Path $LogDir ("skill7_{0}_{1:yyyyMMdd_HHmmss}.log" -f $safe, (Get-Date))

$argv = @('-X', 'utf8', $Entry, $Folder, '--qty', $Qty)
if ($Asm) { $argv += @('--asm', $Asm) }
if ($From) { $argv += @('--from', $From) }
if ($To) { $argv += @('--to', $To) }
if ($Skip) { $argv += @('--skip', $Skip) }

Write-Host "入口: $Entry" -ForegroundColor Cyan
Write-Host "日志: $log" -ForegroundColor Cyan
Write-Host "命令: $Python $($argv -join ' ')" -ForegroundColor DarkGray

# 后台跑 + 日志落盘；-Wait 让本脚本在前台等待，但 SW 长任务不会因终端超时被掐
$p = Start-Process -FilePath $Python -ArgumentList $argv -NoNewWindow -PassThru `
    -RedirectStandardOutput $log -RedirectStandardError "$log.err"

$lastLen = 0
while (-not $p.HasExited) {
    Start-Sleep -Seconds 10
    if (Test-Path $log) {
        $lines = Get-Content -LiteralPath $log -Tail 3 -ErrorAction SilentlyContinue
        if ($lines) { $lines | ForEach-Object { Write-Host "  | $_" -ForegroundColor DarkGray } }
    }
}
$code = $p.ExitCode

Write-Host "`n退出码: $code" -ForegroundColor ($(if ($code -eq 0) { 'Green' } else { 'Red' }))
Write-Host "完整日志: $log" -ForegroundColor Cyan
if ($code -ne 0) {
    Write-Host "`n失败点看日志最后 30 行：" -ForegroundColor Yellow
    Get-Content -LiteralPath $log -Tail 30 -ErrorAction SilentlyContinue
    if (Test-Path "$log.err") { Get-Content -LiteralPath "$log.err" -Tail 20 -ErrorAction SilentlyContinue }
    Write-Host "`n已生成的产物可续跑：-From <步骤号>，例如 -From 9" -ForegroundColor Yellow
}
exit $code