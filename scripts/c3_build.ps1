#!/usr/bin/env powershell
# c3_build.ps1 - build a VB6 project with the current C3.exe (MSVC env set natively)
# Usage: c3_build.ps1 -Vbp <path> -OutDir <outdir> [-Extra "--emit-c"]
param(
    [Parameter(Mandatory=$true)][string]$Vbp,
    [Parameter(Mandatory=$true)][string]$OutDir,
    [string]$Extra = ""
)
# 工具集/SDK 发现 (2026-10-14): 原为写死本机 D:\ 路径 —— 换成 vswhere + 注册表发现
# (与 build.bat / ci.yml 同一条口径, 无硬编码), 本机已知路径仅作兜底。本机解析结果与旧
# 硬编码逐字相同 (vswhere 唯一命中 2019 Community; KitsRoot10=D:\Windows Kits\10)。
$msvc = ""; $kits = ""; $vcInstallDir = ""
$vswhere = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
if (Test-Path $vswhere) {
    $vsPath = & $vswhere -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath 2>$null
    if ($vsPath -and (Test-Path $vsPath)) {
        $vcInstallDir = Join-Path $vsPath "VC"
        $toolset = Get-ChildItem (Join-Path $vcInstallDir "Tools\MSVC") -Directory -ErrorAction SilentlyContinue |
                   Sort-Object { [version]$_.Name } -Descending | Select-Object -First 1
        if ($toolset) { $msvc = $toolset.FullName }
    }
}
$kitsReg = (Get-ItemProperty "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows Kits\Installed Roots" -ErrorAction SilentlyContinue).KitsRoot10
if ($kitsReg -and (Test-Path $kitsReg)) { $kits = $kitsReg.TrimEnd('\') }
if (-not $msvc) {
    # 兜底: 本机已知路径 (旧行为; 本机 14.29 口径不变)
    $msvc = "D:\Program Files (x86)\Microsoft Visual Studio\2019\Community\VC\Tools\MSVC\14.29.30133"
    if (-not $vcInstallDir) { $vcInstallDir = Split-Path -Parent (Split-Path -Parent $msvc) }
}
if (-not $kits) { $kits = "D:\Windows Kits\10" }
$kitVer = Get-ChildItem "$kits\Include" -Directory -ErrorAction SilentlyContinue |
          Where-Object { $_.Name -match '^10\.' } |
          Sort-Object { [version]$_.Name } -Descending | Select-Object -First 1
if (-not $kitVer) { $kitVer = [pscustomobject]@{ Name = "10.0.19041.0" } }
$kitinc = "$kits\Include\$($kitVer.Name)"
$kitlib = "$kits\Lib\$($kitVer.Name)"
Write-Host "== toolset: $msvc =="
Write-Host "== sdk    : $kitinc =="
$env:INCLUDE = "$msvc\include;$kitinc\um;$kitinc\ucrt;$kitinc\shared;$kitinc\winrt;$kitinc\cppwinrt"
$env:LIB     = "$msvc\lib\x64;$kitlib\um\x64;$kitlib\ucrt\x64"
$env:PATH    = "$msvc\bin\Hostx64\x64;$kits\bin\$($kitVer.Name)\x64;$env:PATH"
# T30-C 续 (ai/030 §六): 本脚本已把 INCLUDE/LIB/PATH 灌好, 再补上 P24-05 早退判据要的两枚
# (VCINSTALLDIR + VSCMD_ARG_TGT_ARCH) —— buildVcvarsPrefix 随即短路, 连那一遍 vcvarsall
# 捕获都省掉 (实测捕获本身 ~4.7s/次构建)。x86 时 TGT_ARCH 不匹配 ⇒ 原样走捕获重配 (旧行为)。
$env:VCINSTALLDIR = $vcInstallDir.TrimEnd('\') + "\"
$env:VSCMD_ARG_TGT_ARCH = "x64"

# 项目根目录: 环境变量 C3_PROJECT_DIR 优先, 缺省由脚本位置推导 (与 dev.ps1 同口径)。
# 原为硬编码 `Set-Location D:\c3.vb6.pro` —— 那是另一台机器/另一份 checkout 的路径,
# 在本机它**确实存在** (D:\c3.vb6.pro 是一份 9-30 的旧 checkout, 自带旧 .build\C3.exe),
# 于是脚本会静默跑那份**过期**的 C3.exe, 而不是当前工作树刚编出来的那份 —— 症状是
# "改了 cgen 代码、build.bat 重编了、但 emit 纹丝不动" (VbEclipse With-uc 修复被这份
# 旧二进制吞掉, 排查半天)。改成按脚本位置推导, 与 build.bat 的 C3_PROJECT_DIR 一致。
$ProjectDir = if ($env:C3_PROJECT_DIR) { $env:C3_PROJECT_DIR } else { Split-Path -Parent $PSScriptRoot }
Set-Location $ProjectDir
Write-Host "== building $Vbp (in $ProjectDir) =="
$argList = @($Vbp, "--output-dir", $OutDir)
# ai/030 续 (§10.6 同款约定): 本入口默认走 obj 级缓存 —— 84 个 RTL .c 只在 store 冷时
# 编一遍, 之后 RTL + 未变用户码全部命中 (实测 VBFlexGridDemo 默认路径里 RTL 批独占约 27s)。
# 编译器默认仍未动, 只在本入口显式传; C3_NO_OBJCACHE=1 可关 (复现 GA 那条路, 同 test.bat)。
if (-not $env:C3_NO_OBJCACHE -and $Extra -notmatch '--incremental') { $argList += "--incremental" }
if ($Extra -ne "") { $argList += ($Extra -split ' ') }
& .build\C3.exe @argList 2>&1 | ForEach-Object { "$_" }
Write-Host "== C3 exit=$LASTEXITCODE =="
