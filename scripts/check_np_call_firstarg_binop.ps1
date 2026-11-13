# check_np_call_firstarg_binop.ps1 - 语句级无括号调用·首参括号分组 (§40) 的结构性哨兵
#
# 钉的问题 (PhotoDemon Controls/pdSliderStandalone.ctl:1135):
#     fHSVtoRGB (x - trackRadius) / hueSpread, 1, 1, tmpR, tmpG, tmpB
# VB6 语句级没有"表达式语句" ⇒ `f (...)` 只能是调用, 括号是**首参的分组**, `/ hueSpread`
# 继续把首参写下去。表达式层却把括号当实参表, 语句层再把整个二元表达式当 callee
# ⇒ 发 `vb6_Num_Div((double)(f(...)), (double)(hueSpread), vb6_VariantEmpty(), 1, 1, …)` ⇒ C2197。
# 折叠点有两格 (有二元运算符 / 没有二元运算符), 判据必须同时满足:
#   a) 最左叶是 IndexOrCallExpr, callee 是**裸标识符或 obj.Member**
#      (2026-11-11: 无 binop 格也收 obj.Member —— `m_Stream.SetPosition (a + b), 1`
#       原先整个落到"调用结果再调用", 发成 `vb6_...(me, expr, 0, 0, _own_1, 1)` ⇒ C2197);
#   b) 内层调用恰好 1 个位置实参、无命名实参;
#   c) 语句在逗号后还有实参 (`cur_.kind == TokenKind::Comma`) —— 单独成句的 `f (a)` 仍是正常调用;
#   d) 无二元运算符那一格, 首参**不得是可寻址左值** (Identifier/MemberAccess):
#      `f (v), x` 的括号在 VB6 是按值传临时, 折叠掉会变成按引用传原变量 —— 语义不同。
#
# 用法: powershell -NoProfile -ExecutionPolicy Bypass -File scripts\check_np_call_firstarg_binop.ps1
#       -SrcRoot <dir> 用于负测试 (只含 src/ 的副本树); 退出码 0=绿 1=红

param([string]$SrcRoot = "")

$ErrorActionPreference = "Stop"
$root = if ($SrcRoot -ne "") { $SrcRoot } else { Split-Path -Parent $PSScriptRoot }
$rel = "src\parser\stmt\parser_stmt_assign.cpp"
$p = Join-Path $root $rel

if (-not (Test-Path -LiteralPath $p)) {
    Write-Host ("FAIL S0 file missing: " + $rel) -ForegroundColor Red
    exit 1
}
$text = [System.IO.File]::ReadAllText($p)
$bad = @()

# ---- S1: 折叠点标记唯一 ----
$mark = [regex]::Escape('C3-SENTINEL(np-call-firstarg-binop)')
$nMark = @([regex]::Matches($text, $mark)).Count
if ($nMark -lt 1) {
    $bad += ("S1 sentinel marker missing in " + $rel)
}

# ---- S2: 两格折叠都必须挂 Comma 判据 ----
$nComma = @([regex]::Matches($text, 'cur_\.kind\s*==\s*TokenKind::Comma')).Count
if ($nComma -lt 4) {
    $bad += ("S2 'cur_.kind == TokenKind::Comma' occurrences = " + $nComma + " (want >= 4: 110y 两格 + 本条两格)")
}

# ---- S3: 裸标识符 callee + 恰 1 个位置实参 的判据在 (两格各一处) ----
$nIdentCallee = @([regex]::Matches($text, 'callee->kind\s*==\s*ASTNodeKind::IdentifierExpr')).Count
if ($nIdentCallee -lt 2) {
    $bad += ("S3 'callee->kind == ASTNodeKind::IdentifierExpr' occurrences = " + $nIdentCallee + " (want >= 2: binop 格 + 无 binop 格)")
}
$nArity1 = @([regex]::Matches($text, 'positional\.size\(\)\s*==\s*1')).Count
if ($nArity1 -lt 4) {
    $bad += ("S3 'positional.size() == 1' occurrences = " + $nArity1 + " (want >= 4: 110y 两格 + 本条两格)")
}

# ---- S4: 无 binop 格必须排除可寻址左值 ----
if ($text -notmatch 'argAddressable\s*=\s*\(k\s*==\s*ASTNodeKind::IdentifierExpr') {
    $bad += "S4 missing argAddressable guard (IdentifierExpr)"
}
if ($text -notmatch 'k\s*==\s*ASTNodeKind::MemberAccessExpr') {
    $bad += "S4 missing argAddressable guard (MemberAccessExpr)"
}
if ($text -notmatch '&&\s*!argAddressable') {
    $bad += "S4 missing '!argAddressable' in the fold condition"
}

if ($bad.Count -gt 0) {
    Write-Host "FAIL np_call_firstarg_binop" -ForegroundColor Red
    $bad | ForEach-Object { Write-Host ("  " + $_) -ForegroundColor Red }
    exit 1
}
Write-Host ("PASS np_call_firstarg_binop (authority=" + $rel + ")") -ForegroundColor Green
exit 0
