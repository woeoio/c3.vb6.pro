param(
    [string]$C3 = ".build\C3.exe",
    [ValidateSet('x86','x64')][string]$Arch = 'x86',
    [string]$OutDir = ""
)
# C30-10 第三十刀：`As New <外部 ProgID>` 的判据（无 cscript 对照 —— 前期绑定语法
# VBScript 写不出来，两侧不可比）。判据是**VB6 语义的固定读数**，逐行断言：
#   n1 local_re=True typename=RegExp
#   n2 module_re=True typename=RegExp
#   n3 local_dict=1 typename=Dictionary
#   n4 module_dict=1 typename=Dictionary
# 修前：局部/模块级都只发 `void* r = NULL;`，四条全是空/NULL，stderr 一堆
#       [C3_LATEBIND_MISS ... kind=null]（取证见 ai/039 §11.7）。
# 第三十一刀：n3 的 typename 从 Long 收成 Dictionary（局部接口指针被当 LongPtr 装箱的
#       那一格），四行全部成为硬断言，原来那段「已知差异」报告分支已删。

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
if (-not $OutDir) { $OutDir = Join-Path $root (".build\newbind_" + $Arch) }
New-Item -ItemType Directory -Force $OutDir | Out-Null
$OutDir = (Resolve-Path $OutDir).Path

$src = Join-Path $PSScriptRoot 'NewBind.bas'
$dst = Join-Path $OutDir 'NewBind.bas'
$txt = [IO.File]::ReadAllText($src, [Text.Encoding]::UTF8)
[IO.File]::WriteAllText($dst, ($txt -replace "`r`n","`n" -replace "`n","`r`n"), (New-Object Text.UTF8Encoding($true)))

$vbp = @"
Type=Exe
Module=NewBind; NewBind.bas
Startup="Sub Main"
ExeName32="NewBind.exe"
Name="NewBind"
"@
[IO.File]::WriteAllText((Join-Path $OutDir 'NewBind.vbp'), ($vbp -replace "`n","`r`n"), (New-Object Text.UTF8Encoding($true)))

$exe = (Resolve-Path $C3).Path
$build = & $exe (Join-Path $OutDir 'NewBind.vbp') --arch $Arch --output-dir $OutDir 2>&1
$exePath = Join-Path $OutDir 'NewBind.exe'
if (-not (Test-Path $exePath)) {
    Write-Output ("FAIL newbinding: 编译失败`n" + ($build -join "`n"))
    exit 1
}
$p = Start-Process -FilePath $exePath -WorkingDirectory $OutDir -PassThru -Wait `
     -RedirectStandardOutput (Join-Path $OutDir 'out.txt') `
     -RedirectStandardError  (Join-Path $OutDir 'err.txt')
$out = @(Get-Content (Join-Path $OutDir 'out.txt') -Encoding Default)
$err = @(Get-Content (Join-Path $OutDir 'err.txt') -Encoding Default)

$want = @(
    'n1 local_re=True typename=RegExp',
    'n2 module_re=True typename=RegExp',
    'n3 local_dict=1 typename=Dictionary',
    'n4 module_dict=1 typename=Dictionary'
)
$bad = 0
for ($i = 0; $i -lt $want.Count; $i++) {
    $got = if ($i -lt $out.Count) { $out[$i].Trim() } else { '<缺行>' }
    if (-not $got.StartsWith($want[$i])) {
        if ($bad -eq 0) { Write-Output "FAIL newbinding" }
        Write-Output ("  第 {0} 行  期望前缀: {1}`n            实得: {2}" -f ($i+1), $want[$i], $got)
        $bad++
    }
}
if ($err.Count -gt 0) {
    Write-Output ("  stderr 非空（{0} 条）: {1}" -f $err.Count, ($err[0]))
    $bad++
}
if ($bad -gt 0) { exit 1 }

Write-Output ("PASS newbinding 4/4（reg_exp 局部+模块级、dict 局部+模块级；stderr 0 条）" )
exit 0
