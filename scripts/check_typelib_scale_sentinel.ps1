# check_typelib_scale_sentinel.ps1 - 规模哨兵：类型库符号只许"造一份"，不许随模块数线性长。
#
# 背景（ai/C3_PHOTODEMON_HANDOFF.md §25）：改造前每个模块的 SymbolTable 都深拷贝一份完整类型库符号集，
# PhotoDemon (539 模块 × 3042 符号) 吃到 ~12 GB。共享表落地后本哨兵盯两件事：
#   1. 峰值工作集随模块数的**斜率**（per-module 增量）—— 这是"有没有按模块重造全量符号"的直接读数；
#   2. 绝对峰值与墙钟 —— 兜住"斜率正常但基线暴涨"的另一种退化。
#
# 为什么用 slope 而不是单一阈值：机器的 RTL/CRT 基线在各 runner 上不一样 (几十 MB 级差)，
# 但"每多一个模块多花多少内存"是同一套代码的固有属性 —— 旧形状 ~7.5 MB/模块，共享表 ~0.2 MB/模块。
#
# 语料：N 份同构小类 + 与 PhotoDemon.vbp 相同的两条 Reference（shell32 / stdole2）——
# 不挂类型库的话，被保护的那条路径压根不会被走到，哨兵就成了空的。
#
# 用法:
#   powershell -File scripts/check_typelib_scale_sentinel.ps1 [-C3 <C3.exe>] [-N1 8] [-N2 64]
# 退出码: 0 通过 / 1 超标
param(
    [string]$C3 = "",
    [int]$N1 = 8,
    [int]$N2 = 64,
    [double]$MaxSlopeMB = 2.0,
    [double]$MaxPeakMB = 900.0,
    [int]$MaxSeconds = 120
)
$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot
if (-not $C3) { $C3 = Join-Path $Root ".build\C3.exe" }
if (-not (Test-Path $C3)) { Write-Host "MISSING C3.exe: $C3"; exit 2 }

# ---- S0: 跨模块注入的权威点 (2026-11 载荷共享) ----
# 注入只许共享 SymbolPayload, 不许再逐字段深拷贝 11 张类成员表 —— 这正是
# PhotoDemon 那 +2.7 GB / 378 s 的来源。斜率读数只能事后发现退步, 这里把它钉成结构断言:
#   · Driver::runCrossModuleResolution 必须调用 sharePayloadFrom()
#   · extSym 不许直接赋值/改写字面上的载荷表字段 (含 W() 写门形态)
$crossPath = Join-Path $Root "src\driver\driver_crossmod.cpp"
if (-not (Test-Path $crossPath)) { Write-Host "FAIL S0 missing src/driver/driver_crossmod.cpp"; exit 1 }
$cross = [System.IO.File]::ReadAllText($crossPath)
$payloadFields = 'memberNames|memberReturnTypes|memberProcKinds|memberParams|memberFieldTypes|memberFieldNames|publicFieldNames|memberFieldDispids|memberAccessLevels|memberLetParams|memberSetParams|comMethods|comSourceMethods|comModuleFunctions|delegateProcKind|delegateCallConv|delegateReturnType|delegateTargets|params'
$s0 = @()
if ($cross -notmatch 'sharePayloadFrom\s*\(') {
    $s0 += "S0 crossmod injection no longer shares the payload (sharePayloadFrom missing)"
}
foreach ($mm in [regex]::Matches($cross, 'extSym\s*->\s*(?:W\(\)\s*\.\s*)?(' + $payloadFields + ')\b')) {
    $s0 += ("S0 extSym still touches payload field '" + $mm.Groups[1].Value + "' (must share, not copy)")
}
if ($s0.Count -gt 0) {
    foreach ($m in $s0) { Write-Host ("FAIL " + $m) }
    exit 1
}
Write-Host "S0 ok: crossmod injection shares SymbolPayload (no field-by-field payload copy)"

$dir = Join-Path ([System.IO.Path]::GetTempPath()) ("c3_scale_" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Force -Path $dir | Out-Null

$clsBody = @'
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "@NAME@"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mValue As Long

Public Property Get Value() As Long
    Value = mValue
End Property

Public Property Let Value(ByVal v As Long)
    mValue = v
End Property

Public Sub Bump()
    mValue = mValue + 1
End Sub
'@
$vbpHead = @'
Type=Exe
Reference=*\G{50A7E9B0-70EF-11D1-B75A-00A0C90564FE}#1.0#0#C:\Windows\SysWOW64\shell32.dll#Microsoft Shell Controls And Automation
Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\SysWOW64\stdole2.tlb#OLE Automation
Startup=""
Command32=""
Name="ScaleSentinel"
HelpContextID="0"
CompatibleMode="0"
MajorVer=1
MinorVer=0
RevisionVer=0
AutoIncrementVer=0
ServerSupportFiles=0
VersionCompanyName=""
CompilationType=0
OptimizationType=0
FavorPentiumPro(tm)=0
CodeViewDebugInfo=0
NoAliasing=0
BoundsCheck=0
OverflowCheck=0
FlPointCheck=0
FDIVCheck=0
UnroundedFP=0
StartMode=0
Unattended=0
Retained=0
ThreadPerObject=0
MaxNumberOfThreads=1
'@
for ($i = 1; $i -le $N2; $i++) {
    $nm = "C{0:D3}" -f $i
    [System.IO.File]::WriteAllText((Join-Path $dir ($nm + '.cls')),
        $clsBody.Replace('@NAME@', $nm))
}
function Write-Vbp([string]$path, [int]$n) {
    $cls = @()
    for ($i = 1; $i -le $n; $i++) { $cls += ("Class=C{0:D3}; C{0:D3}.cls" -f $i) }
    [System.IO.File]::WriteAllText($path, ($vbpHead + ($cls -join "`r`n") + "`r`n"))
}
$vbp1 = Join-Path $dir "Scale$N1.vbp"
$vbp2 = Join-Path $dir "Scale$N2.vbp"
Write-Vbp $vbp1 $N1
Write-Vbp $vbp2 $N2

function Measure-Run([string]$vbp) {
    $log = [System.IO.Path]::GetTempFileName()
    $err = $log + ".err"
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $p = Start-Process -FilePath $C3 -ArgumentList ('"' + $vbp + '" --syntax-only') -WorkingDirectory $Root `
        -RedirectStandardOutput $log -RedirectStandardError $err -NoNewWindow -PassThru
    $peak = 0
    while (-not $p.HasExited) {
        try { $p.Refresh(); if ($p.PeakWorkingSet64 -gt $peak) { $peak = $p.PeakWorkingSet64 } } catch { }
        Start-Sleep -Milliseconds 50
    }
    $sw.Stop()
    try { if ($p.PeakWorkingSet64 -gt $peak) { $peak = $p.PeakWorkingSet64 } } catch { }
    # -PassThru 不等待时 ExitCode 可能还没回填，必须 WaitForExit 一次再读
    try { $p.WaitForExit() } catch { }
    $rc = $p.ExitCode
    Remove-Item $log, $err -Force -ErrorAction SilentlyContinue
    return @(($peak / 1MB), $sw.Elapsed.TotalSeconds, $rc)
}

try {
    $a = Measure-Run $vbp1
    $b = Measure-Run $vbp2
    $slope = ($b[0] - $a[0]) / [Math]::Max(1, ($N2 - $N1))
    Write-Host ("typelib_scale: N={0} peak={1:N1}MB {2:N1}s rc={3} | N={4} peak={5:N1}MB {6:N1}s rc={7} | slope={8:N2}MB/模块" -f `
        $N1, $a[0], $a[1], $a[2], $N2, $b[0], $b[1], $b[2], $slope)
    $bad = @()
    if ($a[2] -ne 0 -or $b[2] -ne 0) { $bad += "C3 退出码非 0" }
    if ($b[0] -gt $MaxPeakMB) { $bad += ("峰值 {0:N1}MB > {1:N0}MB" -f $b[0], $MaxPeakMB) }
    if ($slope -gt $MaxSlopeMB) { $bad += ("斜率 {0:N2}MB/模块 > {1:N1} (类型库符号按模块重造?)" -f $slope, $MaxSlopeMB) }
    if ($b[1] -gt $MaxSeconds) { $bad += ("墙钟 {0:N1}s > {1}s" -f $b[1], $MaxSeconds) }
    if ($bad.Count -gt 0) { foreach ($m in $bad) { Write-Host ("FAIL: " + $m) }; exit 1 }
    Write-Host "PASS"
    exit 0
} finally {
    Remove-Item $dir -Recurse -Force -ErrorAction SilentlyContinue
}
