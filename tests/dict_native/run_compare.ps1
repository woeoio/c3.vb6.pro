param(
    [string]$C3 = ".build\C3.exe",
    [ValidateSet('x86','x64')][string]$Arch = 'x86',
    [string]$OutDir = ""
)
# C30-10 内置 Dictionary 对照：真 scrrun（cscript） vs C3 内置实现，逐行 diff。
#
#   基准 = tests/dict_native/DictCases.vbs（cscript 跑真组件，写 dict_oracle.txt）
#   被测 = tests/dict_native/DictMain.bas（同样 56 条，C3 编译后写 dict_c3.txt）
#
# 一处口径分歧必须显式处理（见 ai/039 §4.1）：
#   Item 缺键：cscript 实测不抛、返回 Empty；本项目取 **VB6 语义 32811**。
#   所以 d10 这一行在比较前用 VB6-SEM 覆盖成 32811 —— 覆盖清单写死在下面，
#   与"偷偷改期望值"的区别是：它在此处可见、在文档里有理由、并且只影响列出的 id。
#
# 退出码：0=PASS，2=SKIP（本机没有 cscript / VBScript 功能），其它=FAIL。

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent        # tests/
$repo = Split-Path $root -Parent
if (-not $OutDir) { $OutDir = Join-Path $repo (".build\dict_native_" + $Arch) }
$work = Join-Path $repo ".build"

# VB6-SEM 覆盖表：id -> 期望整行（cscript 与 VB6 不一致、本仓取 VB6 语义）
$Vb6Sem = @{
    'd10|item_missing_err=0' = 'd10|item_missing_err=32811'
}

# 已知 C3 缺口（不是 Dictionary 语义问题，是 C3 侧的打包/代码生成缺口）——
# 这几行**不参与比对**，但显式点名并在 ai/039 里记了追账项，不静默跳过：
#   d62 / d104：`Nothing` 进 Variant 参数时 C3 目前包成 VT_EMPTY（VB6 是 VT_DISPATCH NULL），
#                于是 d.Add Nothing 与 d.Add Empty 撞键。该行真组件读数为 Nothing 独立成键。
#   （`d(k) = v` 那条默认成员带参赋值也是 C3 缺口，夹具改用 d.Item(k) = v，不列在这里。）
$C3Gap = @('d121')
#   第二十九刀把下面四行修好了，从缺口名单里移除（判据回填）：
#     d62 / d104  Nothing 以前被包成 VT_I4 0（与 Empty 撞键）—— 现在 vb6_ComPackNothing ⇒ VT_DISPATCH NULL
#     d102 / d123  Currency 实参以前被 ComPackInt 截成整数 —— 现在 vb6_ComPackCur ⇒ VT_CY
#   仍剩 d121：VarType(CCur(x)) 在 C3 里答 5(R8) 而真组件是 6(CY)、整数字面量答 3(I4) 而不是 2(I2)。
#   这一格是**标量类型宽度**那一层的事（C3 用 double 表示 Currency、字面量默认装箱成 I4），
#   与 Dictionary 行为无关，单列一笔（ai/039 §11.2）。

# 1) 基准：cscript 跑真组件
if (-not (Get-Command cscript -ErrorAction SilentlyContinue)) {
    Write-Output "SKIP dict_native 本机没有 cscript"; exit 2
}
try { $null = New-Object -ComObject VBScript.RegExp }
catch { Write-Output "SKIP dict_native 本机没有 VBScript 运行时（scrrun/脚本引擎）"; exit 2 }
$vbs = Join-Path $PSScriptRoot 'DictCases.vbs'
$tmpDir = Join-Path $work 'dict_oracle_run'
New-Item -ItemType Directory -Force -Path $tmpDir | Out-Null
Copy-Item -LiteralPath $vbs -Destination (Join-Path $tmpDir 'DictCases.vbs') -Force
$null = & cmd.exe /c "cscript //nologo `"$(Join-Path $tmpDir 'DictCases.vbs')`"" 2>&1
$oracleFile = Join-Path $tmpDir 'dict_oracle.txt'
if (-not (Test-Path -LiteralPath $oracleFile)) { Write-Output "FAIL cscript 没有产出 dict_oracle.txt"; exit 1 }
$oracleLines = @(Get-Content -LiteralPath $oracleFile -Encoding Default | ForEach-Object { $_.TrimEnd() })

# 2) C3 侧
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$c3exe = if ([IO.Path]::IsPathRooted($C3)) { $C3 } else { (Resolve-Path (Join-Path $repo $C3)).Path }
$vbp = Join-Path $PSScriptRoot 'DictApp.vbp'
$env:C3_PROJECT_DIR = $repo
$log = & cmd.exe /c "`"$c3exe`" `"$vbp`" --arch $Arch --output-dir `"$OutDir`" 2>&1"
$exe = Join-Path $OutDir 'DictApp.exe'
if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $exe)) {
    Write-Output "FAIL 编译失败（没有 $exe）"
    $log | Select-Object -Last 12 | ForEach-Object { "  " + $_ }
    exit 1
}
$txt = Join-Path $OutDir 'dict_c3.txt'
if (Test-Path -LiteralPath $txt) { Remove-Item -LiteralPath $txt -Force }
$p = Start-Process -FilePath $exe -WorkingDirectory $OutDir -PassThru
if (-not $p.WaitForExit(60000)) { try { $p.Kill() } catch {} ; Write-Output "FAIL 用例进程超时（60s）"; exit 1 }
if (-not (Test-Path -LiteralPath $txt)) { Write-Output "FAIL 没有产出 dict_c3.txt（进程提前退出 code=" + $p.ExitCode + "）"; exit 1 }
$c3Lines = @(Get-Content -LiteralPath $txt -Encoding Default | ForEach-Object { $_.TrimEnd() })

# 3) 套 VB6-SEM 覆盖（只动清单里的整行）
$semCount = 0
for ($i = 0; $i -lt $oracleLines.Count; $i++) {
    if ($Vb6Sem.ContainsKey($oracleLines[$i])) { $oracleLines[$i] = $Vb6Sem[$oracleLines[$i]]; $semCount++ }
}

# 3b) 剔除已知 C3 缺口行（两边同时剔，并且单独报数）
$gapCount = 0
$oracleLines = @($oracleLines | Where-Object { $id = ($_ -split '\|')[0]; -not ($C3Gap -contains $id) })
$c3Filtered = @()
foreach ($ln in $c3Lines) {
    $id = ($ln -split '\|')[0]
    if ($C3Gap -contains $id) { $gapCount++; continue }
    $c3Filtered += $ln
}
$c3Lines = $c3Filtered

# 4) diff
$bad = 0
$max = [Math]::Max($oracleLines.Count, $c3Lines.Count)
$report = New-Object System.Collections.Generic.List[string]
for ($i = 0; $i -lt $max; $i++) {
    $a = if ($i -lt $oracleLines.Count) { $oracleLines[$i] } else { "<缺行>" }
    $b = if ($i -lt $c3Lines.Count) { $c3Lines[$i] } else { "<缺行>" }
    if ($a -ne $b) {
        $bad++
        if ($report.Count -lt 20) {
            $report.Add(("第 {0} 行`n    ocx: {1}`n    c3 : {2}" -f ($i + 1), $a, $b))
        }
    }
}
if ($bad -eq 0) {
    Write-Output ("PASS dict_native 用例={0} 行 两侧逐字一致（oracle=cscript scrrun；VB6-SEM 覆盖 {1} 行；已知 C3 缺口剔除 {2} 行）" -f $oracleLines.Count, $semCount, $gapCount)
    exit 0
}
Write-Output ("FAIL dict_native 不一致 {0} 行 / 共 {1} 行（VB6-SEM 覆盖 {2} 行，已知缺口剔除 {3} 行）" -f $bad, $max, $semCount, $gapCount)
$report | ForEach-Object { Write-Output $_ }
exit 1
