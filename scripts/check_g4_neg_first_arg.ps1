# check_g4_neg_first_arg.ps1 - 账 §B93 的结构性哨兵 (只扫编译器源码, 不起 cl)
#
# 为什么单开这一道: 语句级无括号调用 `obj.M -<首参>, ...` 的折叠点只有一处
# (src/parser/stmt/parser_stmt_assign.cpp 的 Fix <vbeclipse> 块)。它原先只认
# `binNeg.right` 是**字面数**, 于是 `obj.M -1 * w, h` (右操作数是二元乘) 与
# `obj.M -x, ...` (右操作数是标识符) 全部落空, 整句被当成"调用一个减法结果":
#   `(vb6_cls_M_Move(obj) - (1 * w))(h)`  ⇒ cl C2064
# PhotoDemon 实测 9 处 / 7 TU (Dialog_GradientEditor、FormLevels、Interface、PD2D、
# pdEdgeDetector、Tools_Clone、Tools_Gradient)。
#
# 语句级不存在"减法语句", 裸成员访问后跟 `-` 只能是调用语句的负数首参 —— 判据放宽到
# "任意非字符串/日期表达式"是**语义等价**的, 不是放宽断言。本哨兵钉住三件事:
#   N1  折叠点还在 (Marker 恰好 1 处, `negFirstArgCollapse = true` 恰 1 处)。
#   N2  判据是"negatable"而不是"right 必须是 LiteralExpr"; 字符串/日期排除仍在。
#   N3  折叠只对 MemberAccess 叶子 + Sub 生效 (别把赋值右侧的减法也折进来)。
#
# 用法:  powershell -NoProfile -ExecutionPolicy Bypass -File scripts\check_g4_neg_first_arg.ps1
#        -SrcRoot <dir> 可指向一棵只含 src/ 的副本树 (负测试用; 默认 = 仓库根)
# 退出码: 0 = 全绿; 1 = 红

param(
    [string]$SrcRoot = ""
)

$ErrorActionPreference = "Stop"
$root = if ($SrcRoot -ne "") { $SrcRoot } else { Split-Path -Parent $PSScriptRoot }
$bad = @()

function Read-Text([string]$rel) {
    $p = Join-Path $root $rel
    if (-not (Test-Path -LiteralPath $p)) { return $null }
    return [System.IO.File]::ReadAllText($p)
}

$authRel = "src\parser\stmt\parser_stmt_assign.cpp"
$authText = Read-Text $authRel
if ($null -eq $authText) {
    Write-Host ("FAIL B0 file missing: " + $authRel) -ForegroundColor Red
    exit 1
}

$srcDirs = @("src\backend", "src\semantics", "src\ast", "src\driver", "src\common", "src\parser") |
           ForEach-Object { Join-Path $root $_ }
$srcFiles = @()
foreach ($d in $srcDirs) {
    if (-not (Test-Path -LiteralPath $d)) { continue }
    $srcFiles += @(Get-ChildItem -LiteralPath $d -Recurse -File |
                   Where-Object { $_.Extension -in ".cpp", ".inc", ".hpp" })
}

# ---- N1: 折叠点唯一 ----
$markRx = [regex]::Escape('C3-SENTINEL(g4-neg-first-arg)')
$markAll = 0
foreach ($f in $srcFiles) {
    $markAll += @([regex]::Matches([System.IO.File]::ReadAllText($f.FullName), $markRx)).Count
}
if ($markAll -ne 1) {
    $bad += ("N1 C3-SENTINEL(g4-neg-first-arg) occurrences across src = " + $markAll + " (want 1)")
}
if (@([regex]::Matches($authText, $markRx)).Count -ne 1) {
    $bad += ("N1 the sentinel is not in " + $authRel + " (the collapse point moved)")
}
$collapseAll = 0
foreach ($f in $srcFiles) {
    $collapseAll += @([regex]::Matches([System.IO.File]::ReadAllText($f.FullName),
                     'negFirstArgCollapse\s*=\s*true')).Count
}
if ($collapseAll -ne 1) {
    $bad += ("N1 negFirstArgCollapse = true sites across src = " + $collapseAll + " (want 1: one collapse point)")
}

# ---- N2: 判据放宽到 negatable, 字符串/日期排除仍在 ----
$blk = [regex]::Match($authText,
    'C3-SENTINEL\(g4-neg-first-arg\)[\s\S]{0,4000}?negFirstArgCollapse\s*=\s*true')
if (-not $blk.Success) {
    $bad += "N2 cannot locate the collapse block (sentinel ... negFirstArgCollapse = true)"
} else {
    $body = $blk.Value
    if ($body -notmatch 'bool\s+negatable') {
        $bad += "N2 the collapse guard no longer has a 'negatable' decision (back to the literal-only gate)"
    }
    if ($body -notmatch 'negatable\s*=\s*false') {
        $bad += "N2 the 'not Sub / no right operand' rejection is gone (negatable must default off there)"
    }
    if ($body -notmatch 'if\s*\(\s*negatable\s*\)') {
        $bad += "N2 the collapse is no longer gated on negatable"
    }
    if ($body -notmatch 'LiteralExpr') {
        $bad += "N2 the literal special case is gone (strings/dates must still be excluded)"
    }
    if ($body -notmatch 'LiteralKind::String' -or $body -notmatch 'LiteralKind::Date') {
        $bad += "N2 string/date literal exclusion lost (a string minus something must not collapse)"
    }
    if ($body -notmatch 'binNeg\.op != BinaryOp::Sub') {
        $bad += "N2 the Sub-only guard is gone (Add/Mul statements must not collapse)"
    }
}

# ---- N3: 折叠仍只对 MemberAccess 叶子 ----
if ($authText -notmatch '\(\*leafSlot\)->kind\s*==\s*ASTNodeKind::MemberAccessExpr') {
    $bad += "N3 the leaf-slot MemberAccess guard is gone (bare identifier subtractions would collapse)"
}
if ($authText -notmatch 'calleeOverride110y\s*=\s*std::move\(\*leafSlot\)') {
    $bad += "N3 the collapsed leaf is no longer taken as the callee (MemberAccess 作 callee)"
}

if ($bad.Count -gt 0) {
    Write-Host "FAIL g4_neg_first_arg" -ForegroundColor Red
    $bad | ForEach-Object { Write-Host ("  " + $_) -ForegroundColor Red }
    exit 1
}
Write-Host ("PASS g4_neg_first_arg (authority=" + $authRel + ")") -ForegroundColor Green
exit 0
