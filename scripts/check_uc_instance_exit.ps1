# check_uc_instance_exit.ps1 - 账 #191 的结构性哨兵 (只扫编译器源码, 不起 cl)
#
# 为什么单开这一道: 工程内 UserControl 的实例成员访问 (`ucX.AddSerie ...` / `ucX(i).SW()`) 必须
# **直发生成函数** (`vb6_ucX_SW((vb6_cls_ucX*)vb6_UC_InstanceOf(<宿主>))`), 不能落 `vb6_ComCall` 兜底
# —— HWND/宿主值不是 IDispatch, 兜底交回来是**空值**, 编得过、跑得起来、就是静默不干活
# (实测 Charts 2020/ucChartBar 的 `ucChartBar1(i).AddSerie` 整批这样哑掉)。
# Fix 112 当年只给"单枚控件"接了这条路, 控件数组的元素另抄一套 ⇒ 元素那形一路掉进兜底。
#
# 口径现在收在一处: CCodeGen::emitUcInstanceMemberExpr (src/backend/cgen_util_classcall.cpp),
# 单枚 (`vb6_hwnd_<名>`) 与数组元素 (`vb6_CtrlArr_GetAt(&vb6_arr_<名>, i)`) 两条路同调 —— 差的就是那一个串。
#
# 规则 (改坏了会红, 不是装饰):
#   U1  src/backend/detail/expr/ 里**非注释行**手拼 vb6_UC_InstanceOf( = 0 处
#       (值上下文的成员访问全在 expr/ 这一族; 要出新形状必须走出口, 别在 .inc 里再拼一份 this)
#   U2  权威在 cgen_util_classcall.cpp 恰好定义一次, 且函数体里 resolveClassMemberCall /
#       vb6_UC_InstanceOf( / pendingChainObj_ (调用形与值形两条协议) 三样都还在
#   U3  两条路各有一个调用点 (单枚在 cgen_expr_member_form_builtin.inc, 元素在 cgen_expr_member_precheck.inc)
#   U4  左值那一路 (cgen_util_comwrite.cpp 的 tryRewriteCOMLvalue) 对「prop_get_ 调用后面
#       还挂着一层成员」必须有独立的落点, 且排在字符串级折叠 (Pattern C/D2) **之前** ——
#       折叠只看 prop_get_ 那一段、把尾巴整段丢掉: `elem.Font.Size = 10` 会发成
#       prop_let_Font(elem, 10), 成员名没了、值槽还是 Font* (实测四条 C2440)。
#
# 用法:  pwsh -File scripts\check_uc_instance_exit.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$bad = @()

$authRel = "src\backend\cgen_util_classcall.cpp"
$auth = Join-Path $root $authRel
if (-not (Test-Path -LiteralPath $auth)) {
    Write-Host ("FAIL U0 authority file missing: " + $authRel) -ForegroundColor Red
    exit 1
}

$exprDir = Join-Path $root "src\backend\detail\expr"
if (-not (Test-Path -LiteralPath $exprDir)) {
    Write-Host ("FAIL U0 missing dir: src\backend\detail\expr") -ForegroundColor Red
    exit 1
}

# U1: expr/ 里不许再手拼 InstanceOf (注释行除外 —— 注释里提到它是为了说明别这么写)
$hand = 0
$handWhere = @()
$exprFiles = Get-ChildItem -LiteralPath $exprDir -Recurse -File |
             Where-Object { $_.Extension -in ".cpp", ".inc", ".hpp" }
foreach ($f in $exprFiles) {
    $ln = 0
    foreach ($line in ([System.IO.File]::ReadAllText($f.FullName) -split "`r?`n")) {
        $ln++
        $t = $line.Trim()
        if ($t.StartsWith("//") -or $t.StartsWith("*")) { continue }
        if ($line.Contains("vb6_UC_InstanceOf(")) {
            $hand++
            $handWhere += ($f.Name + ":" + $ln)
        }
    }
}
if ($hand -ne 0) {
    $bad += ("U1 hand-composed vb6_UC_InstanceOf in expr/ = " + $hand +
             " (must be 0; call emitUcInstanceMemberExpr) -> " + (($handWhere | Select-Object -First 6) -join " | "))
}

# U2: 权威的形状
$h = [System.IO.File]::ReadAllText($auth)
$defs = @([regex]::Matches($h, 'bool\s+CCodeGen::emitUcInstanceMemberExpr\s*\('))
if ($defs.Count -ne 1) {
    $bad += ("U2 emitUcInstanceMemberExpr defined " + $defs.Count + " times (want 1)")
}
$blk = [regex]::Match($h, 'emitUcInstanceMemberExpr[\s\S]{0,2600}?\r?\n\}')
if (-not $blk.Success) {
    $bad += "U2 authority body not found"
} else {
    foreach ($need in @('resolveClassMemberCall', 'vb6_UC_InstanceOf\(', 'pendingChainObj_', 'findClassMemberCallParams')) {
        if ($blk.Value -notmatch $need) {
            $bad += ("U2 the authority lost " + $need + " - that leg is why elements fell back to COM")
        }
    }
}

# U3: 两条路都在调它
$singleFile = Join-Path $exprDir "cgen_expr_member_form_builtin.inc"
$arrFile    = Join-Path $exprDir "cgen_expr_member_precheck.inc"
foreach ($pair in @(@("单枚控件", $singleFile), @("数组元素", $arrFile))) {
    if (-not (Test-Path -LiteralPath $pair[1])) { $bad += ("U3 missing " + $pair[0] + " file"); continue }
    $t = [System.IO.File]::ReadAllText($pair[1])
    if (@([regex]::Matches($t, 'emitUcInstanceMemberExpr\s*\(')).Count -lt 1) {
        $bad += ("U3 " + $pair[0] + " 那条路不再调用出口 (cgen 侧退回手拼/兜底了)")
    }
}

# U4: 左值改写那一路 —— 带尾巴的 prop_get_ 目标必须有独立落点, 且排在折叠前面
$comWrite = Join-Path $root "src\backend\cgen_util_comwrite.cpp"
if (-not (Test-Path -LiteralPath $comWrite)) {
    $bad += "U4 src/backend/cgen_util_comwrite.cpp is missing (the single LValue rewriter)"
} else {
    $cw = [System.IO.File]::ReadAllText($comWrite)
    $defsLhs = @([regex]::Matches($cw, 'bool\s+CCodeGen::tryRewriteCOMLvalue\s*\('))
    if ($defsLhs.Count -ne 1) {
        $bad += ("U4 tryRewriteCOMLvalue defined " + $defsLhs.Count + " times (want 1)")
    }
    $iTail = $cw.IndexOf("Pattern C-tail")
    $iFold = $cw.IndexOf("prop_get_ rewrite (Pattern C/D2)")
    if ($iTail -lt 0) {
        $bad += "U4 the chained object write (Pattern C-tail) is gone - prop_get_ LHS with a tail gets folded and loses the member name"
    } elseif ($iFold -lt 0) {
        $bad += "U4 the string-level fold anchor vanished (renamed? U4 pins the order between the two)"
    } elseif ($iTail -ge $iFold) {
        $bad += "U4 Pattern C-tail no longer runs BEFORE the string-level fold - the fold would keep eating the tail"
    } else {
        $seg = $cw.Substring($iTail - 2200, 2400)
        foreach ($need in @("vb6_ComSetProp(", "return false")) {
            if (-not $seg.Contains($need)) {
                $bad += ("U4 the chained-object branch lost " + $need + " (unrecognized tails must not be folded)")
            }
        }
    }
    $nTailSite = @([regex]::Matches($cw, "COM SetProp through prop_get_ object")).Count
    if ($nTailSite -ne 1) {
        $bad += ("U4 chained-object emit site count = " + $nTailSite + " (want exactly 1)")
    }
}

if ($bad.Count -eq 0) {
    Write-Host "PASS UC instance-member exit: expr/ 里 0 处手拼 InstanceOf, 单枚与数组元素同走 emitUcInstanceMemberExpr" -ForegroundColor Green
    exit 0
}
foreach ($b in $bad) { Write-Host ("FAIL " + $b) -ForegroundColor Red }
exit 1
