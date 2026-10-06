#!/usr/bin/env powershell
# c3_build.ps1 - build a VB6 project with the current C3.exe (MSVC env set natively)
# Usage: c3_build.ps1 -Vbp <path> -OutDir <outdir> [-Extra "--emit-c"]
param(
    [Parameter(Mandatory=$true)][string]$Vbp,
    [Parameter(Mandatory=$true)][string]$OutDir,
    [string]$Extra = ""
)
$msvc   = "D:\Program Files (x86)\Microsoft Visual Studio\2019\Community\VC\Tools\MSVC\14.29.30133"
$kits   = "D:\Windows Kits\10"
$kitinc = "$kits\Include\10.0.19041.0"
$kitlib = "$kits\Lib\10.0.19041.0"
$env:INCLUDE = "$msvc\include;$kitinc\um;$kitinc\ucrt;$kitinc\shared;$kitinc\winrt;$kitinc\cppwinrt"
$env:LIB     = "$msvc\lib\x64;$kitlib\um\x64;$kitlib\ucrt\x64"
$env:PATH    = "$msvc\bin\Hostx64\x64;$kits\bin\10.0.19041.0\x64;$env:PATH"

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
if ($Extra -ne "") { $argList += ($Extra -split ' ') }
& .build\C3.exe @argList 2>&1 | ForEach-Object { "$_" }
Write-Host "== C3 exit=$LASTEXITCODE =="
