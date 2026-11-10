# check_g4_value_member_paren.ps1 - 账 §B94 的结构性哨兵 (只扫编译器源码, 不起 cl)
#
# 为什么单开这一道: 值成员被写成"带空括号"时 (`cmdOK.FontSize()`, `cboUpdates.UBound()`,
# `PDImages.GetActiveImage.Width()`), 成员读取本身已经产出**完整** C 表达式, 通用调用路
# 却又追加一对 `()` ⇒ "调用一个非函数" ⇒ cl C2064 (PhotoDemon 实测 9 条: options_Updates 4 /
# SelectionFilters 4 / pdCommandBarMini 1)。
#
# 收口点在 src/backend/detail/expr/cgen_expr_call_builtin_argtype.inc: 追加 `()` 的那一行
# 之前先认"读法尾注", 认到就当值访问直接返回。判据只认产出尾注, 不猜符号表:
#   · 有实参的调用不进这一格;  · 真方法调用没有尾注;  · 尾注是读法自己写的。
#
# 规则:
#   P1  Marker C3-SENTINEL(g4-value-member-empty-paren) 恰好 1 处, 且在权威 .inc 内。
#   P2  收口块先于 `lastExpr_ = callee + "(" + argList + ")";` (顺序 + 存在性)。
#   P3  三条读法尾注齐全 (`ctrl prop read` / `ctrl array UBound` / `ctrl array LBound`),
#       且类字段尾注 (`/* class var .`) 仍在。
#   P4  收口块必须有 `argList.empty()` 与 `lastExpr_ = callee;` + return (不许静默丢表达式)。
#
# 用法:  powershell -NoProfile -ExecutionPolicy Bypass -File scripts\check_g4_value_member_paren.ps1
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

$authRel = "src\backend\detail\expr\cgen_expr_call_builtin_argtype.inc"
$authText = Read-Text $authRel
if ($null -eq $authText) {
    Write-Host ("FAIL P0 file missing: " + $authRel) -ForegroundColor Red
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

# ---- P1: Marker 唯一 ----
$markRx = [regex]::Escape('C3-SENTINEL(g4-value-member-empty-paren)')
$markAll = 0
foreach ($f in $srcFiles) {
    $markAll += @([regex]::Matches([System.IO.File]::ReadAllText($f.FullName), $markRx)).Count
}
if ($markAll -ne 1) {
    $bad += ("P1 C3-SENTINEL(g4-value-member-empty-paren) occurrences across src = " + $markAll + " (want 1)")
}
if (@([regex]::Matches($authText, $markRx)).Count -ne 1) {
    $bad += ("P1 the sentinel is not in " + $authRel + " (the collection point moved)")
}

# ---- P2: 顺序 + 通用追加仍在 ----
$appendRx = 'lastExpr_\s*=\s*callee\s*\+\s*"\("\s*\+\s*argList\s*\+\s*"\)"\s*;'
$appendM = [regex]::Match($authText, $appendRx)
$markM   = [regex]::Match($authText, $markRx)
if (-not $appendM.Success) {
    $bad += "P2 the generic call append (lastExpr_ = callee + ""("" + argList + "")"";) is gone (real calls would break)"
} elseif (-not $markM.Success -or $markM.Index -ge $appendM.Index) {
    $bad += "P2 the value-read check is not before the generic append (the () would still be added)"
}

# ---- P3 / P4: 块形状 ----
$blkEnd = if ($appendM.Success) { $appendM.Index } else { $authText.Length }
$blkStart = if ($markM.Success) { $markM.Index } else { 0 }
$body = $authText.Substring($blkStart, [Math]::Max(0, $blkEnd - $blkStart))

if ($markM.Success) {
    foreach ($tail in @('/* ctrl prop read */', '/* ctrl array UBound */', '/* ctrl array LBound */')) {
        if ($body -notmatch [regex]::Escape($tail)) {
            $bad += ("P3 value-read tail missing from the collection block: " + $tail)
        }
    }
    if ($body -notmatch [regex]::Escape('/* class var .')) {
        $bad += "P3 class-field read tail (/* class var .) is no longer recognized"
    }
    if ($body -notmatch 'argList\.empty\(\)') {
        $bad += "P4 the guard lost argList.empty() (a real call with arguments must not be swallowed)"
    }
    if ($body -notmatch 'lastExpr_\s*=\s*callee\s*;') {
        $bad += "P4 the block no longer assigns lastExpr_ = callee (the value would be lost)"
    }
    if ($body -notmatch 'return\s*;') {
        $bad += "P4 the block no longer returns (it would fall through and still append ())"
    }
}

if ($bad.Count -gt 0) {
    Write-Host "FAIL g4_value_member_paren" -ForegroundColor Red
    $bad | ForEach-Object { Write-Host ("  " + $_) -ForegroundColor Red }
    exit 1
}
Write-Host ("PASS g4_value_member_paren (authority=" + $authRel + ")") -ForegroundColor Green
exit 0
