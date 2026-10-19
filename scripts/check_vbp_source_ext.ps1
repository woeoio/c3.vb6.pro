# check_vbp_source_ext.ps1 - VB6 原生 .vbp 的源文件条目不带扩展名, C3 必须自己补 (账 #293)
#
# 起因: VB6 IDE 存盘写的是
#       Form=MainForm / Module=Module1 / Class=Cls1 / UserControl=Uc1 / PropertyPage=Pp1
# —— **一律不带扩展名**, 且源文件常常在子目录里 (`Module=Common\Common`)。
# ComCtlsDemo.vbp 能编过只是因为它手写了扩展名 (`Module=Common; Common\Common.bas`,
# 那是外部工具改出来的两段格式), 所以这个洞在仓里的夹具上完全测不到。
#
# C3 之前把 .vbp 的值原样当路径用, 于是 open("<proj>\Module1") 失败 => VB1006
# 「无法打开文件」。**任何 VB6 IDE 存出来的工程都编不出 exe。**
#
# 本哨兵钉两头, 缺一头都会红:
#   V1 正例: 一份完全照 VB6 原生写法 (条目不带扩展名) 的工程, 六种条目
#      (Form / Module / Class / UserControl / PropertyPage / 子目录 Module)
#      必须全部解析成功, 且每个源文件都在产物里留下自己的编译单元。
#   V2 反面证人: 已经带扩展名的条目不许被"顺手"改写 —— 放一条 `Module=Weird.bas`,
#      必须照样解析 (若无条件往后接, 会去找 Weird.bas.bas, 全部落空)。
#
# 断言打在**产物**上而不是会话目录上: `--emit-c` 把合并后的 C 写到 stdout, 每个源文件
# 一个 `VB6CGEN_<名>_H` 头守卫。第一版是去 %TEMP%\C3C 里挑"最近修改"的会话目录找 .c
# —— 那个目录在 `--emit-c` 下压根没有 .c, 于是要么假红, 要么**假绿**(撞上别的构建
# 留下的会话)。私有 TEMP + 断言产物, 两头都堵死。
#
# 用法: powershell -File scripts\check_vbp_source_ext.ps1   (PASS = exit 0)

param(
    [string]$Root = (Split-Path -Parent $PSScriptRoot),
    [string]$Exe  = ""      # 负控用: 拿另一枚 C3.exe 跑同一份夹具
)
$ErrorActionPreference = "Stop"
$root = $Root
$exe = if ($Exe) { $Exe } else { Join-Path $root ".build\C3.exe" }
if (-not (Test-Path -LiteralPath $exe)) {
    Write-Host ("FAIL missing " + $exe + "  (run scripts\build.bat first)")
    exit 1
}

$bad = @()

$work = Join-Path $env:TEMP ("c3_vbpext_" + [Guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Force -Path (Join-Path $work "Sub") | Out-Null

function Put([string]$rel, [string[]]$lines) {
    Set-Content -LiteralPath (Join-Path $work $rel) -Value $lines -Encoding ASCII
}

# ---- 夹具: VB6 原生写法, 条目一律不带扩展名 -------------------------------
Put "p.vbp" @(
    'Type=Exe',
    'Form=MainForm',
    'Module=Mod1',
    'Module=Sub\Deep',
    'Class=C1',
    'UserControl=U1',
    'PropertyPage=PP1'
)
Put "MainForm.frm" @(
    'VERSION 5.00',
    'Begin VB.Form MainForm ',
    '  Caption         =   "P"',
    'End',
    'Attribute VB_Name = "MainForm"',
    'Attribute VB_GlobalNameSpace = False',
    'Attribute VB_Creatable = False',
    'Attribute VB_PredeclaredId = True',
    'Attribute VB_Exposed = True'
)
Put "Mod1.bas" @('Attribute VB_Name = "Mod1"', 'Public Const M = 7')
Put "Sub\Deep.bas" @('Attribute VB_Name = "Deep"', 'Public Const D = 9')
Put "C1.cls" @('Attribute VB_Name = "C1"', 'Public Sub Go()', 'End Sub')
Put "U1.ctl" @(
    'VERSION 5.00', 'Begin VB.UserControl U1 ', 'End',
    'Attribute VB_Name = "U1"', 'Attribute VB_GlobalNameSpace = False',
    'Attribute VB_Creatable = False'
)
Put "PP1.pag" @(
    'VERSION 5.00', 'Begin VB.PropertyPage PP1 ', 'End',
    'Attribute VB_Name = "PP1"', 'Attribute VB_GlobalNameSpace = False'
)

# ---- V2 夹具: 条目已经带扩展名 ---------------------------------------------
Put "q.vbp" @('Type=Exe', 'Module=Weird.bas')
Put "Weird.bas" @('Attribute VB_Name = "Weird"', 'Public Const Z = 1')

# 私有 TEMP: 会话目录别跟同机其它 C3 抢 (编译 ComCtlsDemo / 跑别的哨兵都会建)。
$saveTEMP = $env:TEMP; $saveTMP = $env:TMP
$privTmp = Join-Path $work "_tmp"
New-Item -ItemType Directory -Force -Path $privTmp | Out-Null
$env:TEMP = $privTmp; $env:TMP = $privTmp

function Run-C3([string]$proj) {
    $out = Join-Path $work "emit.c"
    $err = Join-Path $work "emit.err"
    $p = Start-Process -FilePath $exe `
            -ArgumentList @("`"$proj`"", "--emit-c") `
            -NoNewWindow -Wait -PassThru `
            -RedirectStandardOutput $out -RedirectStandardError $err
    $txt = ""
    if (Test-Path -LiteralPath $out) { $txt = Get-Content -LiteralPath $out -Raw }
    $etxt = ""
    if (Test-Path -LiteralPath $err) { $etxt = Get-Content -LiteralPath $err -Raw }
    return @{ Exit = $p.ExitCode; Out = $txt; Err = $etxt }
}

# ---- V1 正例: 六种条目都要解析出来并各自留下编译单元 -----------------------
$r1 = Run-C3 (Join-Path $work "p.vbp")
if ($r1.Exit -ne 0) {
    $bad += ("V1 compiler exited " + $r1.Exit + " on an extensionless (VB6-native) vbp")
}
if ($r1.Err -match "VB1006") {
    $bad += ("V1 still reports VB1006 (cannot open file) on an extensionless vbp: " +
             (($r1.Err -split "`r?`n" | Where-Object { $_ -match "VB1006" } | Select-Object -First 1)))
}
foreach ($g in @('MAINFORM', 'MOD1', 'DEEP', 'C1', 'U1', 'PP1')) {
    if ($r1.Out -notmatch ("VB6CGEN_" + $g + "_H")) {
        $bad += ("V1 no compiled unit for " + $g + " (extensionless entry not resolved)")
    }
}

# ---- V2 反面证人 -----------------------------------------------------------
$r2 = Run-C3 (Join-Path $work "q.vbp")
if ($r2.Exit -ne 0) {
    $bad += ("V2 compiler exited " + $r2.Exit + " on an entry that already has an extension (Module=Weird.bas)")
    if ($r2.Err -match "VB1006") {
        $bad += "V2 reports VB1006 for Weird.bas - an extra extension was appended blindly"
    }
} elseif ($r2.Out -notmatch "VB6CGEN_WEIRD_H") {
    $bad += "V2 resolved Module=Weird.bas but emitted no unit for it"
}

$env:TEMP = $saveTEMP; $env:TMP = $saveTMP
Remove-Item -Recurse -Force $work -ErrorAction SilentlyContinue

if ($bad.Count -eq 0) {
    Write-Host "PASS"
    exit 0
}
foreach ($b in $bad) { Write-Host ("FAIL " + $b) }
exit 1