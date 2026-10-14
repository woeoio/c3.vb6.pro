# check_ctrl_prop_type_authority.ps1 - 账 #231 的结构性哨兵 (只扫编译器源码, 不起 cl)
#
# 为什么单开这一道: 控件属性的**类型**以前有两处权威 —— controlPropType 那张表, 以及
# inferExprType 里 C29-1a / C29-1b / C29-9 的三份自带名单 (kNumericFc / kStringFc3 /
# kStrFcCd / kNumFcCd)。同一个问句 ("Shape1.FillStyle 是什么类型") 两个地方各自答, 改一
# 处就会把另一处的旧答案留在原地; 两份名单重合的那四条 (Left/Top/Width/Height) 只是
# **恰好**一样才没出事 —— 账 #229 就是这条撞出来的: 数组元素那一形只有一处认识。
# 账 #231 把四份名单逐条搬进表里, inferExprType 只留一次问话。
#
# 规则 (改坏了会红, 不是装饰):
#   A1  那四个旧名单标识符与它的手写循环变量在 src/ 里回潮 = 0 (再开一份平行名单就红)
#   A2  controlPropType 全库提及数 = 3 (定义 + 声明 + **恰好一个**调用点)
#   A3  ctrlTypeOfMemberObject ("这枚对象是不是窗体控件") 同样 = 3, 且那唯一的调用点在
#       cgen_util_type.cpp —— 别处再判一次控件身份 = 又开一条平行路
#   A4  搬进来的那 30 条属性名必须逐条在表里答到 (`p == "<名>"`)。语料不覆盖的名字在
#       emit A/B 上是哑的 (#229 那轮的教训: changed 全是夹具自身新增行), 所以这条硬钉。
#   A5  表里那道 "不是 Unknown" 的闸必须还在、且只有 1 处 (自定义 OCX 的属性面归类型库)
#   A6 (256) defaultPropType = 3 (definition + declaration + **exactly 1** call site) and that
#       call site is inside ctrlDefaultPropOf - the box a late-bound argument gets has to follow the
#       value the emitter hands over, so the fold is the only place allowed to ask the default
#       property's type. Asking controlPropType from there again = two tables for one question.
#   A7 (256) ctrlDefaultPropOf = 4: definition + declaration + the TWO consumers (the identifier
#       emission and the late-bound argument packing). Before this knife those two each carried
#       their own copy of the decision, which is exactly how a BSTR ended up in a VT_DISPATCH slot.
#   A8 (230 §B137) the table's LongPtr answer must be consumed on BOTH roads, and the numeric box
#       it produces must not be re-flattened downstream:
#         a) boxToVariant carries a `case Vb6Type::LongPtr:` arm that boxes through `(intptr_t)`
#            (without it a handle read lands on _Generic's default arm = vb6_VariantObject, i.e.
#            an HWND stored as an IDispatch* - VarType 9 / TypeName from the host table / CStr empty);
#         b) that numeric-box emission exists in exactly ONE place in src/ (a second copy is a
#            parallel boxing road, which is the whole reason this account opened);
#         c) the RTL converter the vb6_VarCmp* family reads through (vb6_VarToDouble_internal)
#            has a VT_I8 arm - without it a 64-bit numeric box compares as 0.0 and §B136's
#            "always equal" returns on the Variant-slot face (measured on x64 through Fix 082's
#            own VarPtr box: `VarPtr(x) <> 0` answered False).
#
# 用法:  pwsh -File scripts\check_ctrl_prop_type_authority.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$bad = @()

# scan the whole src tree (.c/.h are RTL C sources and drop out on the extension filter)
$files = Get-ChildItem -LiteralPath (Join-Path $root "src") -Recurse -File |
          Where-Object { $_.Extension -in ".cpp", ".inc", ".hpp" }

# ---- A1: 旧名单不许回潮 ----
$old = @()
foreach ($f in $files) {
    $t = [System.IO.File]::ReadAllText($f.FullName)
    foreach ($m in [regex]::Matches($t, 'k(NumericFc|StringFc3|StrFcCd|NumFcCd)')) { $old += ($f.Name + ":" + $m.Value) }
    foreach ($m in [regex]::Matches($t, 'memFc\s*==')) { $old += ($f.Name + ":memFc==") }
}
if ($old.Count -ne 0) {
    $bad += ("A1 retired side-list back = " + ($old -join " | "))
}

# ---- A2 / A3: 一处定义、一处声明、恰好一个调用者 ----
$authRel = "src\backend\cgen_util_ctrl.cpp"
$auth = Join-Path $root $authRel
$callRel = "src\backend\cgen_util_type.cpp"
$callFile = Join-Path $root $callRel
if (-not (Test-Path -LiteralPath $auth)) { $bad += ("A2 authority file missing: " + $authRel) }
if (-not (Test-Path -LiteralPath $callFile)) { $bad += ("A3 caller file missing: " + $callRel) }

function Count-Mentions([string]$name) {
    $n = 0
    $where = @()
    foreach ($f in $files) {
        $c = @([regex]::Matches(([System.IO.File]::ReadAllText($f.FullName)), $name + '\s*\(')).Count
        if ($c -ne 0) { $n += $c; $where += ($f.Name + "=" + $c) }
    }
    return , @($n, ($where -join " "))
}

foreach ($pair in @(@("controlPropType", "A2"), @("ctrlTypeOfMemberObject", "A3"))) {
    $r = Count-Mentions $pair[0]
    if ($r[0] -ne 3) {
        $bad += ($pair[1] + " " + $pair[0] + " mentioned " + $r[0] +
                 " times (want definition + declaration + exactly 1 call site) -> " + $r[1])
    }
}
$callText = ""
if (Test-Path -LiteralPath $callFile) { $callText = [System.IO.File]::ReadAllText($callFile) }
if ($callText -notmatch 'controlPropType\s*\(') {
    $bad += "A3 the single call site of controlPropType is not in cgen_util_type.cpp (inference stopped asking the table)"
}
if ($callText -notmatch 'ctrlTypeOfMemberObject\s*\(') {
    $bad += "A3 the single call site of ctrlTypeOfMemberObject is not in cgen_util_type.cpp"
}

# ---- A4 / A5: 那张表本身 ----
$body = ""
if (Test-Path -LiteralPath $auth) {
    $h = [System.IO.File]::ReadAllText($auth)
    $mm = [regex]::Match($h, 'Vb6Type\s+CCodeGen::controlPropType\([\s\S]*?\r?\n\}')
    if (-not $mm.Success) { $bad += "A4 controlPropType body not found" } else { $body = $mm.Value }
}
if ($body -ne "") {
    $collapsed = @("shape", "fillstyle", "borderwidth", "borderstyle", "fillcolor",
                   "bordercolor", "x1", "y1", "x2", "y2",
                   "drive", "path", "pattern", "filename", "list",
                   "filter", "filetitle", "dialogtitle", "initdir", "defaultext", "fontname",
                   "flags", "cancelerror", "color", "min", "max", "copies", "fontsize",
                   "hwnd", "hdc")
    foreach ($p in $collapsed) {
        $n = @([regex]::Matches($body, 'p\s*==\s*"' + [regex]::Escape($p) + '"')).Count
        if ($n -lt 1) { $bad += ("A4 property dropped from the table: " + $p) }
    }
    $g = @([regex]::Matches($body, 'ctrlType\s*!=\s*FrmControlType::Unknown')).Count
    if ($g -ne 1) { $bad += ("A5 the not-Unknown gate is present " + $g + " times in the table (want 1)") }
}

# ---- A6 / A7 (256): the folded default property is asked from one place, by two consumers ----
$r6 = Count-Mentions 'defaultPropType'
if ($r6[0] -ne 3) {
    $bad += ('A6 defaultPropType mentioned ' + $r6[0] +
             ' times (want definition + declaration + exactly 1 call site) -> ' + $r6[1])
}
$foldBody = ''
if (Test-Path -LiteralPath $auth) {
    $h6 = [System.IO.File]::ReadAllText($auth)
    $m6 = [regex]::Match($h6, 'CCodeGen::ctrlDefaultPropOf\([\s\S]*?\r?\n\}')
    if (-not $m6.Success) { $bad += 'A6 ctrlDefaultPropOf body not found' } else { $foldBody = $m6.Value }
}
if ($foldBody -ne '') {
    if ($foldBody -notmatch 'valueType\s*=\s*defaultPropType\s*\(') {
        $bad += 'A6 ctrlDefaultPropOf no longer asks defaultPropType (the argument box stopped following the folded value)'
    }
    if ($foldBody -match 'controlPropType\s*\(') {
        $bad += 'A6 ctrlDefaultPropOf asks controlPropType again - two type tables for one question'
    }
}
$want7 = @{ 'cgen_util_ctrl.cpp' = 1; 'cgen_helpers.inc' = 1; 'cgen_util_com.cpp' = 1;
            'cgen_expr_ident_symbol.inc' = 1 }
$got7 = @{}
foreach ($f in $files) {
    $c7 = @([regex]::Matches(([System.IO.File]::ReadAllText($f.FullName)), 'ctrlDefaultPropOf\s*\(')).Count
    if ($c7 -ne 0) { $got7[$f.Name] = $c7 }
}
$tot7 = 0
foreach ($k in $got7.Keys) { $tot7 += $got7[$k] }
if ($tot7 -ne 4) {
    $bad += ('A7 ctrlDefaultPropOf mentioned ' + $tot7 +
             ' times (want definition + declaration + the two consumers) -> ' +
             (($got7.Keys | ForEach-Object { $_ + '=' + $got7[$_] }) -join ' '))
}
foreach ($k in $want7.Keys) {
    if (-not $got7.ContainsKey($k)) {
        $bad += ('A7 consumer disappeared: ' + $k + ' no longer asks ctrlDefaultPropOf')
    } elseif ($got7[$k] -ne $want7[$k]) {
        $bad += ('A7 ' + $k + ' asks ctrlDefaultPropOf ' + $got7[$k] + ' times (want ' + $want7[$k] + ')')
    }
}
if ($got7.Count -ne $want7.Count) {
    $bad += ('A7 a new file asks ctrlDefaultPropOf: ' + (($got7.Keys | Where-Object { -not $want7.ContainsKey($_) }) -join ' '))
}

# ---- A8 (230 §B137): the table's LongPtr answer must be consumed on BOTH roads ----
$boxBody = ''
if (Test-Path -LiteralPath $auth) {
    $h8 = [System.IO.File]::ReadAllText($auth)
    $m8 = [regex]::Match($h8, 'CCodeGen::boxToVariant\([\s\S]*?\r?\n\}')
    if (-not $m8.Success) { $bad += "A8 boxToVariant body not found" } else { $boxBody = $m8.Value }
}
if ($boxBody -ne '') {
    if ($boxBody -notmatch 'case\s+Vb6Type::LongPtr:') {
        $bad += ("A8 boxToVariant lost its LongPtr arm - a handle read falls to _Generic's default " +
                 "= vb6_VariantObject (an HWND parked in a VT_DISPATCH slot: VarType 9, TypeName from " +
                 "the host table, CStr empty, comparisons never equal)")
    }
    if ($boxBody -notmatch 'vb6_VariantFromValue\(\(intptr_t\)\(') {
        $bad += 'A8 the LongPtr arm no longer boxes through (intptr_t) - the pointer-width caliber (x86 VT_I4 / x64 VT_I8) is gone'
    }
}
$emit8 = @()
foreach ($f in $files) {
    $c8 = @([regex]::Matches(([System.IO.File]::ReadAllText($f.FullName)), 'vb6_VariantFromValue\(\(intptr_t\)\(')).Count
    if ($c8 -ne 0) { $emit8 += ($f.Name + "=" + $c8) }
}
if ($emit8.Count -ne 1 -or $emit8[0] -ne "cgen_util_ctrl.cpp=1") {
    $bad += ('A8 numeric handle box emitted from [' + ($emit8 -join " ") +
             '] (want exactly cgen_util_ctrl.cpp=1 - a second site is a parallel boxing road)')
}
$rtlCmp = Join-Path $root "src\rtl\core\vb6rtl\vb6rtl.c"
if (-not (Test-Path -LiteralPath $rtlCmp)) {
    $bad += 'A8 missing src\rtl\core\vb6rtl\vb6rtl.c'
} else {
    $m8c = [regex]::Match([System.IO.File]::ReadAllText($rtlCmp),
                          'static\s+double\s+vb6_VarToDouble_internal\([\s\S]*?\r?\n\}')
    if (-not $m8c.Success) {
        $bad += 'A8 vb6_VarToDouble_internal body not found'
    } elseif ($m8c.Value -notmatch 'case\s+\(vb6_vartype\)VT_I8:') {
        $bad += ('A8 vb6_VarToDouble_internal has no VT_I8 arm - a numeric 64-bit box compares as 0.0 ' +
                 'through the whole vb6_VarCmp* family (measured on x64 via Fix 082 VarPtr: `VarPtr(x) <> 0` answered False)')
    }
}

if ($bad.Count -eq 0) {
    Write-Host ("PASS ctrl-prop type authority: one table, one caller, " +
                "side-lists retired, " + $collapsed.Count + " collapsed names still answered, " +
                "LongPtr consumed by boxing + VarCmp converter") -ForegroundColor Green
    exit 0
}
foreach ($b in $bad) { Write-Host ("FAIL " + $b) -ForegroundColor Red }
exit 1
