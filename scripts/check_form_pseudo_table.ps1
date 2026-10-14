# check_form_pseudo_table.ps1 - census: 窗体裸写属性的名字/读函数只在一张表里登记（账 #278 §B109）
#
# 起因（2026-10-09，实测 tests/czUI-main 与 tests/VbQRCodegen-master）：`.frm` 里裸写
# `ScaleWidth` / `ScaleHeight` / `WindowState` 是 VB6 的合法写法（等价于 `Me.` 打头），发码侧
# 一直答得对 —— `cgen_expr_ident_symbol.inc` 在窗体模块里拿 `getControlPropReadFn(Form, name)`
# 查名字表，命中就发 `vb6_Get…(hwnd)`。但那张表住在 backend，语义层问不到 ⇒
#   Option Explicit 的窗体：每条合法写法多配一条 VB3001（语料里 5 条，全是噪声）；
#   没写 Option Explicit 的窗体：更实 —— 裸名落进"隐式 Variant 局部"那一格，产物里多出
#     `vb6_VARIANT ScaleWidth = vb6_VariantEmpty();`，属性读被一枚空 Variant 挤掉。
# 与 host_pseudo.hpp（账 #159/#219）同族：同一个事实拷两份就是两个权威。本刀把表搬到
# `src/common/form_pseudo.hpp` (kFormPseudoRows)，backend 的通用段 + 窗体臂改问它，语义层用
# 同一个出口判"这名字窗体答不答"。
#
# 本哨兵钉五头，缺一头都会红：
#   V1 正例放行：一枚带 Option Explicit 的 .frm 裸读 ScaleWidth / ScaleHeight / WindowState /
#      CurrentX ⇒ diagnostics 里不许出现这几个名字的 VB3001。
#   V2 反面证人：同一枚窗体里放一枚真不存在的名 (`fpNoSuchNameAnywhere`) ⇒ 必须**仍然**报 VB3001
#      （只钉 V1 的哨兵会替"干脆什么都不报"背书 —— 那是同一个洞的另一种坏法）。
#   V3 宽松窗体：另一枚**没写** Option Explicit 的 .frm 裸读同名属性 ⇒ 产物里不许出现
#      `vb6_VARIANT ScaleWidth`（隐式局部），且必须出现 `vb6_GetScaleWidth(` / `vb6_GetWindowState(`
#      （真走属性读那条口）。
#   S1 单一登记处：backend 里 `formPseudoReadFn` 只许被调一次；老通用段那些硬编码行不许回来
#      （`return "vb6_GetControlLeft";` 在 src/backend 出现 0 次 —— 它现在只住在表里）。
#   S2 表本体：34 行、名字不重；语义层消费点 `formPseudoIsBare` 恰好 1 处。
#
# 输出 ASCII（控制台是 GBK，中文读数重定向后不可 grep）。文件必须 UTF-8 BOM + CRLF：
# PS 5.1 读无 BOM 的 .ps1 按 ANSI，行尾中文字节会吃掉换行 ⇒ ParserError 而退出码仍 0。
#
# 用法: powershell -File scripts\check_form_pseudo_table.ps1   (PASS = exit 0)

param(
    [string]$Root = (Split-Path -Parent $PSScriptRoot),
    [string]$Exe = ""      # 负控用: 拿另一枚 C3.exe 跑同一份夹具
)
$ErrorActionPreference = "Stop"
$root = $Root
$exe = if ($Exe) { $Exe } else { Join-Path $root ".build\C3.exe" }
if (-not (Test-Path -LiteralPath $exe)) {
    Write-Host ("FAIL missing " + $exe + "  (先跑 scripts/build.bat 产出 C3.exe)")
    exit 1
}

$bad = @()

# 只认**非注释行**里的登记点（账 #278 §B108 的负控实测：纯文本正则会连 `// addObj("X")` 一起算，
# 那种哨兵对"有人把这行注释掉了"完全无感）。这里返回 Match 对象，因为表行有两个组要取。
# 只认**非注释部分**里的登记点（账 #278 §B108 的负控实测：纯文本正则会连 `// addObj("X")` 一起
# 算进去，那种哨兵对"有人把这行注释掉了"完全无感）。做法 = 逐行先切掉 `//` 之后的部分再匹配。
# （别用 LastIndexOf("`n") 反推行首：实测那样会把上一整行当成前缀算进来，把好的行误判成注释。）
function Collect-Matches([string]$text, [string]$pattern) {
    $found = @()
    foreach ($line in ($text -split "`r?`n")) {
        $cut = $line.IndexOf('//')
        $body = if ($cut -ge 0) { $line.Substring(0, $cut) } else { $line }
        foreach ($m in [regex]::Matches($body, $pattern)) { $found += $m }
    }
    return $found
}

$work = Join-Path $env:TEMP ("c3_formpseudo_" + [Guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Path $work -Force | Out-Null

try {
    $frmExp = @(
        'VERSION 5.00',
        'Begin VB.Form FpExp',
        '   Caption         =   "FpExp"',
        '   ClientHeight    =   3000',
        '   ClientWidth     =   4500',
        '   ScaleHeight     =   3000',
        '   ScaleWidth      =   4500',
        'End',
        'Attribute VB_Name = "FpExp"',
        'Option Explicit',
        '',
        'Private Sub Form_Load()',
        '    Dim a As Long',
        '    a = ScaleWidth + ScaleHeight + WindowState + CLng(CurrentX)',
        '    Debug.Print a',
        '    Dim bogus As Long',
        '    bogus = fpNoSuchNameAnywhere',
        'End Sub',
        ''
    )
    $frmLoose = @(
        'VERSION 5.00',
        'Begin VB.Form FpLoose',
        '   Caption         =   "FpLoose"',
        '   ClientHeight    =   2000',
        '   ClientWidth     =   3000',
        '   ScaleHeight     =   2000',
        '   ScaleWidth      =   3000',
        'End',
        'Attribute VB_Name = "FpLoose"',
        '',
        'Private Sub Form_Load()',
        '    Debug.Print "W=" & ScaleWidth, "S=" & ScaleHeight',
        'End Sub',
        ''
    )
    $vbp = @(
        'Type=Exe',
        'Form=FpExp.frm',
        'Form=FpLoose.frm',
        'ExeName32="fpt.exe"',
        'Startup="FpExp"',
        ''
    )
    Set-Content -LiteralPath (Join-Path $work "FpExp.frm")   -Value $frmExp   -Encoding ASCII
    Set-Content -LiteralPath (Join-Path $work "FpLoose.frm") -Value $frmLoose -Encoding ASCII
    $proj = Join-Path $work "fpt.vbp"
    Set-Content -LiteralPath $proj -Value $vbp -Encoding ASCII

    $outC = Join-Path $work "emit.c"
    $errFile = Join-Path $work "emit.err"
    # 用 OS 级重定向取字节（--emit-c 的产物走 stdout, 诊断走 stderr）
    $p = Start-Process -FilePath $exe -ArgumentList @("`"$proj`"", "--emit-c") `
                       -NoNewWindow -Wait -PassThru `
                       -RedirectStandardOutput $outC -RedirectStandardError $errFile
    if ($p.ExitCode -ne 0) { $bad += ("V0 compiler exited " + $p.ExitCode + " on the probe project") }

    $emit = ""
    if (Test-Path -LiteralPath $outC) { $emit = Get-Content -LiteralPath $outC -Raw }
    $errTxt = ""
    if (Test-Path -LiteralPath $errFile) { $errTxt = Get-Content -LiteralPath $errFile -Raw }

    # ---------- V1: 四个真名字都不许报 VB3001 ----------
    foreach ($nm in @("ScaleWidth", "ScaleHeight", "WindowState", "CurrentX")) {
        if ($errTxt -match ("VB3001[^\r\n]*" + $nm)) {
            $bad += ("V1 form pseudo-property " + $nm + " still reports VB3001 (a .frm may write it bare)")
        }
    }

    # ---------- V2: 真不存在的那个名字必须继续报 ----------
    if ($errTxt -notmatch "VB3001[^\r\n]*fpNoSuchNameAnywhere") {
        $bad += "V2 NEGATIVE CONTROL FAILED: fpNoSuchNameAnywhere no longer reports VB3001 - the gate would be blanket-off"
    }

    # ---------- V3: 宽松窗体的产物两头钉 ----------
    if ($emit -match 'vb6_VARIANT\s+ScaleWidth\b') {
        $bad += "V3 emitted C still declares an implicit Variant local named ScaleWidth (the property read got pushed out)"
    }
    foreach ($fn in @("vb6_GetScaleWidth", "vb6_GetScaleHeight", "vb6_GetWindowState")) {
        if ($emit -notmatch ($fn + "\(")) {
            $bad += ("V3 emitted C has no " + $fn + "() call - the form's bare read did not go through the property route")
        }
    }

    # ---------- S1/S2: 单一登记处与表本体 ----------
    $tbl   = Join-Path $root "src\common\form_pseudo.hpp"
    $cutil = Join-Path $root "src\backend\cgen_util_ctrl.cpp"
    $exprf = Join-Path $root "src\semantics\semantic_analyzer_expr.cpp"
    foreach ($f in @($tbl, $cutil, $exprf)) {
        if (-not (Test-Path -LiteralPath $f)) { $bad += ("S missing file " + $f) }
    }
    if ((Test-Path -LiteralPath $tbl) -and (Test-Path -LiteralPath $cutil) -and (Test-Path -LiteralPath $exprf)) {
        $tblTxt  = Get-Content -LiteralPath $tbl -Raw
        $cutlTxt = Get-Content -LiteralPath $cutil -Raw
        $expTxt  = Get-Content -LiteralPath $exprf -Raw

        $rows = @()
        foreach ($m in (Collect-Matches $tblTxt '\{\s*"([a-z][a-z0-9]*)"\s*,\s*"([A-Za-z_0-9]+)"')) {
            $rows += $m.Groups[1].Value
        }
        if ($rows.Count -ne 34) {
            $bad += ("S2 kFormPseudoRows has " + $rows.Count + " rows (want exactly 34; a new row must be deliberate)")
        }
        if (@($rows | Sort-Object -Unique).Count -ne $rows.Count) {
            $bad += "S2 kFormPseudoRows has a duplicate name"
        }
        $fnNames = @()
        foreach ($m in (Collect-Matches $tblTxt '\{\s*"([a-z][a-z0-9]*)"\s*,\s*"([A-Za-z_0-9]+)"')) {
            $fnNames += $m.Groups[2].Value
        }
        if (@($fnNames | Sort-Object -Unique).Count -lt 30) {
            $bad += ("S2 kFormPseudoRows answers only " + (@($fnNames | Sort-Object -Unique).Count) +
                     " distinct read functions (want >= 30 - rows are not supposed to collapse)")
        }

        $calls = @(Collect-Matches $cutlTxt 'formPseudoReadFn\(').Count
        if ($calls -ne 1) {
            $bad += ("S1 backend calls formPseudoReadFn " + $calls + " times in cgen_util_ctrl.cpp (want exactly 1)")
        }
        # 老通用段那两条硬编码不许回来（它们现在只住在表里）
        $back = @(Collect-Matches $cutlTxt 'return "vb6_GetControlLeft"').Count
        if ($back -ne 0) {
            $bad += ("S1 the old hardcoded generic row is back in cgen_util_ctrl.cpp (" + $back +
                    " hits for return ""vb6_GetControlLeft"") - the table is the only place")
        }
        # 窗体那一臂必须是空的：读函数只从表里来。取 getControlPropReadFn 的函数体 → 里面的
        # `case FrmControlType::Form:` → 到它的 break; 为止，这段里不许再有 `propLower ==`。
        $fs = $cutlTxt.IndexOf('CCodeGen::getControlPropReadFn(')
        if ($fs -lt 0) { $bad += "S1 cannot find getControlPropReadFn in cgen_util_ctrl.cpp" }
        $fe = $cutlTxt.IndexOf("`nstd::string CCodeGen::", $fs + 1)
        if ($fe -lt 0) { $bad += "S1 cannot find the end of getControlPropReadFn" }
        if ($fs -ge 0 -and $fe -gt $fs) {
            $body = $cutlTxt.Substring($fs, $fe - $fs)
            $as = $body.IndexOf('case FrmControlType::Form:')
            if ($as -lt 0) { $bad += "S1 the Form arm is gone from getControlPropReadFn (the table needs a consumer there)" }
            $ae = $body.IndexOf('break;', $as)
            if ($as -ge 0 -and $ae -gt $as) {
                $arm = $body.Substring($as, $ae - $as)
                $inArm = @(Collect-Matches $arm 'propLower ==').Count
                if ($inArm -ne 0) {
                    $bad += ("S1 the Form arm of getControlPropReadFn still hardcodes " + $inArm +
                            " name(s) - kFormPseudoRows is the only place a name may be registered")
                }
            }
        }
        $cons = @(Collect-Matches $expTxt 'formPseudoIsBare\(').Count
        if ($cons -ne 1) {
            $bad += ("S2 semantics asks formPseudoIsBare " + $cons + " times (want exactly 1 consumer)")
        }
    }
} finally {
    Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
}

if ($bad.Count -gt 0) {
    foreach ($b in $bad) { Write-Host ("FAIL " + $b) }
    exit 1
}
Write-Host "PASS form_pseudo_table (V1 4 names silent, V2 bogus name still reported, V3 no implicit local + real read route, S1 one call site, S2 34 rows / 1 consumer)"
exit 0
