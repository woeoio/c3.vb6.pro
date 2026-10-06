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
#   A4  搬进来的那 28 条属性名必须逐条在表里答到 (`p == "<名>"`)。语料不覆盖的名字在
#       emit A/B 上是哑的 (#229 那轮的教训: changed 全是夹具自身新增行), 所以这条硬钉。
#   A5  表里那道 "不是 Unknown" 的闸必须还在、且只有 1 处 (自定义 OCX 的属性面归类型库)
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
                   "flags", "cancelerror", "color", "min", "max", "copies", "fontsize")
    foreach ($p in $collapsed) {
        $n = @([regex]::Matches($body, 'p\s*==\s*"' + [regex]::Escape($p) + '"')).Count
        if ($n -lt 1) { $bad += ("A4 property dropped from the table: " + $p) }
    }
    $g = @([regex]::Matches($body, 'ctrlType\s*!=\s*FrmControlType::Unknown')).Count
    if ($g -ne 1) { $bad += ("A5 the not-Unknown gate is present " + $g + " times in the table (want 1)") }
}

if ($bad.Count -eq 0) {
    Write-Host ("PASS ctrl-prop type authority: one table, one caller, " +
                "side-lists retired, " + $collapsed.Count + " collapsed names still answered") -ForegroundColor Green
    exit 0
}
foreach ($b in $bad) { Write-Host ("FAIL " + $b) -ForegroundColor Red }
exit 1
