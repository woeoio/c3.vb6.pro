<#
.SYNOPSIS
    诊断并处置"打开文件 - 安全警告"（无法验证发布者 / 运行、保存）导致的 exe 启动阻塞。

.DESCRIPTION
    只操作本仓库路径下的编译产物，不修改 UAC、Defender、SmartScreen、Internet 区域等全局安全策略。

    会检查两类"启动拦截源"：
      1. 附件管理器标记：NTFS 备用数据流 Zone.Identifier（MOTW）。命中则用 Unblock-File 清除。
      2. NTFS 完整性标签：Low Mandatory Level。目录被显式标为 Low 时，其下所有产物继承 Low，
         ShellExecute 路径（资源管理器双击、cmd start、Windows PowerShell 5.1 的 Start-Process
         默认行为）会弹"打开文件 - 安全警告"，而 CreateProcess 路径（pwsh 7 的 Start-Process、
         .NET Process.Start(UseShellExecute=$false)、调用运算符 &）不会弹。

    默认只诊断、不做任何修改；加 -Fix 才处置，加 -Revert 回滚。

.NOTES
    2026-11-04 复核（更正一条错归因）：仓库根 SACL 曾被标 (OI)(CI)(NW) 低完整性，子目录产物
    全部继承 Low；Low IL 进程写不了 Medium IL 的 HKCU，于是 test_registry_settings(_x86)、
    b175_num_bstr(_x86)、b176_var_slot(_x86) 与 test_earlybound2 的 EB2-9/10 一起红。
    早期把它们记成"环境差：工作区二进制被拒写，移出工作区即通过"是错的 —— 与目录位置无关，
    是**进程完整性级别**。处置后全仓库 Low 残留 0（本脚本 -Fix 即把根目录改回 Medium）。
    复核命令：
        .\scripts\fix-exe-launch-warning.ps1            # 只诊断，列出命中项
        .\scripts\fix-exe-launch-warning.ps1 -Fix       # 改回 Medium 并清 MOTW
        icacls . /setintegritylevel Medium /T:C         # 等价的逐项处置（脚本内部用这条）

.PARAMETER Path
    要处理的根目录。默认取脚本所在目录的上一级（即仓库根）。

.PARAMETER Extension
    参与检查/处理的文件通配符，默认只覆盖二进制产物 *.exe/*.dll/*.ocx。

.PARAMETER Fix
    处置：对命中 MOTW 的产物执行 Unblock-File；把根目录完整性标签归一为 Medium（(OI)(CI) 继承）。

.PARAMETER Revert
    回滚 -Fix：把根目录完整性标签恢复为 Low（(OI)(CI) 继承）。

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File scripts\fix-exe-launch-warning.ps1
    仅诊断读数，不改任何东西。

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File scripts\fix-exe-launch-warning.ps1 -Fix
    处置本仓库产物（构建/测试脚本收尾调用这一条即可）。

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File scripts\fix-exe-launch-warning.ps1 -Revert
    恢复仓库根的 Low 完整性标签（回到处置前状态）。

.NOTES
    影响范围：仅 -Path 指定的目录树（含其子目录），可逆。
    回滚命令：icacls "<Path>" /setintegritylevel (OI)(CI)Low
#>
[CmdletBinding(DefaultParameterSetName = 'Diagnose')]
param(
    [string]$Path,
    [string[]]$Extension = @('*.exe', '*.dll', '*.ocx'),
    [Parameter(ParameterSetName = 'Fix')][switch]$Fix,
    [Parameter(ParameterSetName = 'Revert')][switch]$Revert
)

$ErrorActionPreference = 'Stop'
if (-not $Path) { $Path = Join-Path $PSScriptRoot '..' }
$Path = (Resolve-Path -LiteralPath $Path).Path
Write-Host "[path] $Path"

function Get-IntegrityLabel {
    param([string]$LiteralPath)
    $line = icacls $LiteralPath 2>&1 | Select-String -Pattern 'Mandatory Label'
    if (-not $line) { return 'None' }
    return $line.ToString().Trim()
}

function Get-Binaries {
    $ext = @{}
    foreach ($pat in $Extension) { $ext[($pat -replace '\*', '').ToLower()] = $true }
    Get-ChildItem -LiteralPath $Path -Recurse -File -Force -ErrorAction SilentlyContinue |
        Where-Object { $ext.ContainsKey($_.Extension.ToLower()) }
}

Write-Host '=== 1) 附件管理器标记 (Zone.Identifier / MOTW) ==='
$motw = @()
foreach ($f in Get-Binaries) {
    $streams = (Get-Item -LiteralPath $f.FullName -Stream * -ErrorAction SilentlyContinue).Stream
    if ($streams -contains 'Zone.Identifier') { $motw += $f.FullName }
}
Write-Host ("[motw] hits = {0}" -f $motw.Count)
$motw | Select-Object -First 5 | ForEach-Object { Write-Host "        $_" }

Write-Host '=== 2) 完整性标签 (Low 会触发 ShellExecute 的安全警告) ==='
Write-Host ("[label-root] " + (Get-IntegrityLabel -LiteralPath $Path))
$low = @()
foreach ($f in Get-Binaries) {
    if ((Get-IntegrityLabel -LiteralPath $f.FullName) -match 'Low Mandatory') { $low += $f.FullName }
}
Write-Host ("[label-low] binaries inherited Low = {0}" -f $low.Count)
$low | Select-Object -First 5 | ForEach-Object { Write-Host "        $_" }

if ($Fix) {
    Write-Host '=== 处置 (仅本仓库路径) ==='
    if ($motw.Count -gt 0) {
        $motw | ForEach-Object { Unblock-File -LiteralPath $_ }
        Write-Host ("[fix] Unblock-File = {0} 个文件" -f $motw.Count)
    } else {
        Write-Host '[fix] 无 MOTW，跳过 Unblock-File'
    }
    icacls $Path /setintegritylevel '(OI)(CI)Medium' | Out-Null
    Write-Host ('[fix] 完整性标签归一到 Medium: ' + (Get-IntegrityLabel -LiteralPath $Path))
    Write-Host '[fix] 回滚: 用 -Revert'
}

if ($Revert) {
    Write-Host '=== 回滚 ==='
    icacls $Path /setintegritylevel '(OI)(CI)Low' | Out-Null
    Write-Host ('[revert] 完整性标签恢复为 Low: ' + (Get-IntegrityLabel -LiteralPath $Path))
}

Write-Host ('[done] mode = ' + $PSCmdlet.ParameterSetName)
