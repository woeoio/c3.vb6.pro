# check_currency_round_station.ps1 —— 「Currency 赋值圈 4 位小数」这一条规则的站点对账
# （账 #305 §11.158，第一百七十五刀）
#
# 为什么要有这一道：VB6 的 Currency 是**定点 4 位**（VT_CY = 值×10000），而 C3 用 double
# 承载它 —— 赋值站改前直接存原值，`c = 10.123456` 打 10.123456，真 VB6 存 10.1235
# （探针 .temp/k174/p1 的 P5 取证）。折数只许问 RTL 那一位主人 vb6_CCur
# （它也是 CCur() 的出口：round(v*10000)/10000），发码侧不另写一份 round。
#
# 为什么还得单独一张按名表：Currency 与 Date 同族 —— C 层都是 double，只按 C 类型串登记时
# inferExprType 只会回 Double（§11.158 探针实测：`Dim c As Currency` 的推断档 = 5 = Double），
# 于是 knownCurrencyVars_ 与 knownDoubleVars_ **并存**登记，口径照抄 Fix 175 的 knownDateVars_。
#
# 判据（任何一条不齐就红）：
#   R1 站点唯一：currencyAssignWrap 的定义（cgen_util_type.cpp）与声明（cgen_helpers.inc）各 1 处。
#   R2 Assignment 站经过它：narrowCheckAssign 体内调 currencyAssignWrap 恰好 1 处。
#   R3 Let 站经过它：cgen_setlet.cpp 里 currencyAssignWrap 恰好 2 处（转换分支 + 裸值分支）——
#      LetStmt 是另一条 AST 语句，改前它完全不经过 narrowCheckAssign，漏了这一站
#      就等于「显式 Let 不圈数」（探针 .temp/k175/p2 的 M6 就是它）。
#   R4 按名表唯一的家：knownCurrencyVars_ 在 cgen_state.inc 声明 1 处，三处过程入口各 clear 1 次。
#   R5 五处声明站都登记：形参 3（Function/Sub/Property）+ 模块级 1 + 局部 1 = insert 5 处。
#   R6 判据只问 RTL 主人：src/backend 里不许出现手写 `round(` 与 `10000` 同现的一行。
#   R7 夹具在册：tests/b178_currency_round 的三份输入 + run_tests.ps1 里 x64/x86 两格。
#   R8 反向证人还在：夹具里 Double 那一格必须钉 10.123456（不圈），否则「把 Currency 圈数
#      顺手做成了全局格式化」这种改法看不出来。
#   R9 折数只舍一次（§11.159）: 定标主人 vb6rtl_base.h 的 vb6_CurrencyScale 必须拆开缩放
#      (trunc + 只对小量 round)，vb6_CCur 只负责除回去，且不得再出现 `round(v * 10000` ——
#      探针 .temp/k176/p1 量到 A5/A8/A10/A12/A13 各偏一单位。
#   R10 CCur 家族与 VT_CY 三个消费站同一主人: vb6_CCurBSTR / vb6_ComPackCur /
#      vb6_VariantCur 都走 vb6_CurrencyScale，自己不许再写一份折数。
#   R11 全 RTL 只允许**一处** `round(... * 10000)`，且必须在那个主人体内（注释不算）。
#   R12 新夹具在册: tests/b179_currency_tie 三份输入 + runner x64/x86 两格。
$ErrorActionPreference = 'Continue'
$repo = Split-Path -Parent $PSScriptRoot

function Read-File([string]$rel) {
    $p = Join-Path $repo $rel
    if (-not (Test-Path -LiteralPath $p)) { return $null }
    return [IO.File]::ReadAllText($p)
}

$bad = @()

# ---------- R1 ----------
$def1 = Read-File 'src/backend/cgen_util_type.cpp'
$hdr1 = Read-File 'src/backend/detail/util/cgen_helpers.inc'
if ($null -eq $def1) { $bad += 'R1 cgen_util_type.cpp not found' }
else {
    $n = ([regex]::Matches($def1, 'std::string CCodeGen::currencyAssignWrap\(')).Count
    if ($n -ne 1) { $bad += ('R1 currencyAssignWrap definition = ' + $n + ', want 1') }
}
if ($null -eq $hdr1) { $bad += 'R1 cgen_helpers.inc not found' }
else {
    $n = ([regex]::Matches($hdr1, 'std::string currencyAssignWrap\(')).Count
    if ($n -ne 1) { $bad += ('R1 currencyAssignWrap declaration = ' + $n + ', want 1') }
}

# ---------- R2 ----------
if ($def1) {
    $i = $def1.IndexOf('std::string CCodeGen::narrowCheckAssign(')
    if ($i -lt 0) { $bad += 'R2 narrowCheckAssign body not found' }
    else {
        $tail = $def1.Substring($i, [Math]::Min(2000, $def1.Length - $i))
        $n = ([regex]::Matches($tail, 'currencyAssignWrap\(target, cValue\)')).Count
        if ($n -ne 1) { $bad += ('R2 narrowCheckAssign -> currencyAssignWrap = ' + $n + ', want 1') }
    }
}

# ---------- R3 ----------
$let3 = Read-File 'src/backend/stmt/cgen_setlet.cpp'
if ($null -eq $let3) { $bad += 'R3 cgen_setlet.cpp not found' }
else {
    $n = ([regex]::Matches($let3, 'currencyAssignWrap\(node\.target\.get\(\),')).Count
    if ($n -ne 2) { $bad += ('R3 Let station wraps = ' + $n + ', want 2') }
}

# ---------- R4 ----------
$st4 = Read-File 'src/backend/detail/util/cgen_state.inc'
if ($null -eq $st4) { $bad += 'R4 cgen_state.inc not found' }
else {
    $n = ([regex]::Matches($st4, 'std::unordered_set<std::string> knownCurrencyVars_;')).Count
    if ($n -ne 1) { $bad += ('R4 knownCurrencyVars_ declaration = ' + $n + ', want 1') }
}
$entries = @(
    'src/backend/decl/cgen_decl_func.cpp',
    'src/backend/decl/cgen_decl_proc.cpp',
    'src/backend/decl/cgen_decl_prop.cpp'
)
$clrTotal = 0
foreach ($e in $entries) {
    $t = Read-File $e
    if ($null -eq $t) { $bad += ('R4 missing ' + $e); continue }
    $n = ([regex]::Matches($t, 'knownCurrencyVars_\.clear\(\);')).Count
    if ($n -ne 1) { $bad += ('R4 ' + (Split-Path -Leaf $e) + ' clear = ' + $n + ', want 1') }
    $clrTotal += $n
}

# ---------- R5 ----------
$insTotal = 0
foreach ($e in $entries) {
    $t = Read-File $e
    if ($null -eq $t) { continue }
    $n = ([regex]::Matches($t, 'knownCurrencyVars_\.insert\(pLower\);')).Count
    if ($n -ne 1) { $bad += ('R5 ' + (Split-Path -Leaf $e) + ' param insert = ' + $n + ', want 1') }
    $insTotal += $n
}
$var5 = Read-File 'src/backend/decl/cgen_decl_var.cpp'
if ($null -eq $var5) { $bad += 'R5 cgen_decl_var.cpp not found' }
else {
    $n = ([regex]::Matches($var5, 'knownCurrencyVars_\.insert\(lower\);')).Count
    if ($n -ne 1) { $bad += ('R5 module-level insert = ' + $n + ', want 1') }
    $insTotal += $n
}
$loc5 = Read-File 'src/backend/decl/cgen_localdecl.cpp'
if ($null -eq $loc5) { $bad += 'R5 cgen_localdecl.cpp not found' }
else {
    $n = ([regex]::Matches($loc5, 'knownCurrencyVars_\.insert\(lower\);')).Count
    if ($n -ne 1) { $bad += ('R5 local Dim insert = ' + $n + ', want 1') }
    $insTotal += $n
}
if ($insTotal -ne 5) { $bad += ('R5 declaration stations = ' + $insTotal + ', want 5') }

# ---------- R6 ----------
$handRound = 0
Get-ChildItem -Path (Join-Path $repo 'src/backend') -Recurse -Include *.cpp,*.inc | ForEach-Object {
    $txt = [IO.File]::ReadAllText($_.FullName)
    foreach ($line in ($txt -split "`n")) {
        # 注释不算发码 —— 本刀的说明文字里就写着 round(v*10000)/10000 这个式子。
        $s = $line.Trim()
        if ($s.StartsWith('//') -or $s.StartsWith('*') -or $s.StartsWith('/*')) { continue }
        if ($line -match 'round\(' -and $line -match '10000') { $script:handRound++ }
    }
}
if ($handRound -ne 0) { $bad += ('R6 hand-written round/10000 in src/backend = ' + $handRound + ', want 0') }

# ---------- R7 ----------
foreach ($f in @('B178Main.bas', 'B178Box.cls', 'B178Res.vbp')) {
    $p = Join-Path $repo ('tests/b178_currency_round/' + $f)
    if (-not (Test-Path -LiteralPath $p)) { $bad += ('R7 fixture file missing: ' + $f) }
}
$run7 = Read-File 'tests/run_tests.ps1'
if ($null -eq $run7) { $bad += 'R7 tests/run_tests.ps1 not found' }
else {
    $n = ([regex]::Matches($run7, "Test-Vbp 'b178_currency_round")).Count
    if ($n -ne 2) { $bad += ('R7 b178 fixture cells = ' + $n + ', want 2 (x64 + x86)') }
}

# ---------- R8 ----------
$main8 = Read-File 'tests/b178_currency_round/B178Main.bas'
if ($null -eq $main8) { $bad += 'R8 B178Main.bas not found' }
else {
    if ($main8 -notmatch '"B178-11="\s*&\s*d') {
        $bad += 'R8 reverse witness missing: B178-11 must print the Double value unrounded'
    }
}

if ($bad.Count -gt 0) {
    Write-Host ('FAIL currency_round_station: ' + ($bad -join ' | '))
    exit 1
}Write-Host '  R1 currencyAssignWrap single station OK'
Write-Host '  R2 Assignment station goes through it OK'
Write-Host '  R3 Let station goes through it (2 emit sites) OK'
Write-Host '  R4 knownCurrencyVars_ declaration + 3 clears OK'
Write-Host '  R5 five declaration stations register OK'
Write-Host '  R6 no hand-written round/10000 in codegen OK'
Write-Host '  R7 fixture + runner cells OK'
Write-Host '  R8 Double reverse witness OK'

# ---------- R9 ----------
$base9 = Read-File 'src/rtl/core/vb6rtl/vb6rtl_base.h'
if ($null -eq $base9) { $bad += 'R9 vb6rtl_base.h not found' }
else {
    if ($base9 -notmatch 'static inline int64_t vb6_CurrencyScale\(double v\)') {
        $bad += 'R9 vb6_CurrencyScale owner not found in vb6rtl_base.h'
    }
    if ($base9 -notmatch 'const double ip = trunc\(v\);') {
        $bad += 'R9 vb6_CurrencyScale must split the integer part (trunc)'
    }
    if ($base9 -notmatch 'round\(fp \* 10000\.0\)') {
        $bad += 'R9 vb6_CurrencyScale must round only the small fractional product'
    }
}
$misc9 = Read-File 'src/rtl/core/vb6rtl/vb6rtl_misc.c'
if ($null -eq $misc9) { $bad += 'R9 vb6rtl_misc.c not found' }
else {
    $i = $misc9.IndexOf('double vb6_CCur(double v)')
    if ($i -lt 0) { $bad += 'R9 vb6_CCur definition not found' }
    else {
        $body = $misc9.Substring($i, [Math]::Min(900, $misc9.Length - $i))
        if ($body -notmatch 'vb6_CurrencyScale\(v\)') {
            $bad += 'R9 vb6_CCur must go through vb6_CurrencyScale'
        }
        if ($body -match 'round\(\s*v\s*\*\s*10000') {
            $bad += 'R9 vb6_CCur must not round the already-rounded product (double rounding)'
        }
    }
}

# ---------- R10 ----------
$conv10 = Read-File 'src/rtl/core/vb6rtl/vb6rtl_conv.c'
if ($null -eq $conv10) { $bad += 'R10 vb6rtl_conv.c not found' }
else {
    if ($conv10 -notmatch 'double\s+vb6_CCurBSTR\(BSTR s\)\s*\{\s*double d = vb6_NumVal\(s\); return vb6_CCur\(d\);\s*\}') {
        $bad += 'R10 vb6_CCurBSTR must go through vb6_CCur (single owner)'
    }
}
$pack10 = Read-File 'src/rtl/core/vb6com/vb6com_pack.c'
if ($null -eq $pack10) { $bad += 'R10 vb6com_pack.c not found' }
else {
    if ($pack10 -notmatch 'vb6_CurrencyScale\(val\)') {
        $bad += 'R10 vb6_ComPackCur must go through vb6_CurrencyScale'
    }
    $codePack = ($pack10 -split "`n" | Where-Object { $s = $_.Trim(); -not ($s.StartsWith('//') -or $s.StartsWith('*') -or $s.StartsWith('/*')) }) -join "`n"
    if ($codePack -match 'llround\(') {
        $bad += 'R10 vb6com_pack.c must not keep its own llround scaling'
    }
}
$var10 = Read-File 'src/rtl/core/vb6rtl/vb6rtl_variant.h'
if ($null -eq $var10) { $bad += 'R10 vb6rtl_variant.h not found' }
else {
    if ($var10 -notmatch 'v\.cyVal = vb6_CurrencyScale\(val\);') {
        $bad += 'R10 vb6_VariantCur must go through vb6_CurrencyScale'
    }
}

# ---------- R11 ----------
# 全 RTL 只允许一处 round(...*10000): 就是那位主人自己。
$ownerRound = 0
$strayRound = 0
Get-ChildItem -Path (Join-Path $repo 'src/rtl') -Recurse -Include *.c,*.h | ForEach-Object {
    $isOwner = ($_.Name -eq 'vb6rtl_base.h')
    $txt = [IO.File]::ReadAllText($_.FullName)
    foreach ($line in ($txt -split "`n")) {
        $s = $line.Trim()
        if ($s.StartsWith('//') -or $s.StartsWith('*') -or $s.StartsWith('/*')) { continue }
        if ($s -match 'round\(' -and $s -match '10000') {
            if ($isOwner) { $script:ownerRound++ } else { $script:strayRound++ }
        }
    }
}
if ($strayRound -ne 0) { $bad += ('R11 hand-written round(x*10000) outside the owner = ' + $strayRound + ', want 0') }
if ($ownerRound -ne 1) { $bad += ('R11 owner scaler lines = ' + $ownerRound + ', want 1') }

# ---------- R12 ----------
foreach ($f in @('B179Main.bas', 'B179Res.vbp')) {
    $p = Join-Path $repo ('tests/b179_currency_tie/' + $f)
    if (-not (Test-Path -LiteralPath $p)) { $bad += ('R12 fixture file missing: ' + $f) }
}
if ($null -ne $run7) {
    $n = ([regex]::Matches($run7, "Test-Vbp 'b179_currency_tie")).Count
    if ($n -ne 2) { $bad += ('R12 b179 fixture cells = ' + $n + ', want 2 (x64 + x86)') }
}

if ($bad.Count -gt 0) {
    Write-Host ('FAIL currency_round_station: ' + ($bad -join ' | '))
    exit 1
}
Write-Host '  R9 single rounding (integer part split, only the small fraction rounded) OK'
Write-Host '  R10 CCur family shares one owner OK'
Write-Host '  R11 no round(x*10000) left in RTL OK'
Write-Host '  R12 b179 fixture + runner cells OK'

# ---------- R13 ----------
# §11.160 (第一百七十七刀): ByVal Currency 形参在**调用点**折数 —— 出口唯一，两条实参路
# (位置 / 具名) 都走它；ByRef 不许圈 (交的是地址)，已折过的实参不许折第二次。
$def13 = Read-File 'src/backend/cgen_util_type.cpp'
if ($null -eq $def13) { $bad += 'R13 cgen_util_type.cpp not found' }
else {
    $n = ([regex]::Matches($def13, 'std::string CCodeGen::currencyByValArgWrap\(')).Count
    if ($n -ne 1) { $bad += ('R13 currencyByValArgWrap definition = ' + $n + ', want 1') }
    $i13 = $def13.IndexOf('std::string CCodeGen::currencyByValArgWrap(')
    if ($i13 -ge 0) {
        $body13 = $def13.Substring($i13, [Math]::Min(3000, $def13.Length - $i13))
        if ($body13 -notmatch '!isByVal || paramType != Vb6Type::Currency') {
            $bad += 'R13 the station must be ByVal-Currency only (ByRef hands over an address)'
        }
        if ($body13 -notmatch 'return "vb6_CCur\(" \+ cValue \+ "\)"') {
            $bad += 'R13 the station must fold through vb6_CCur (no second scaling formula)'
        }
        if ($body13 -notmatch 'knownCurrencyVars_\.count\(lower\)') {
            $bad += 'R13 Currency variables must not be folded twice (knownCurrencyVars_ gate)'
        }
    }
}
if ($null -ne $hdr1) {
    $n = ([regex]::Matches($hdr1, 'std::string currencyByValArgWrap\(')).Count
    if ($n -ne 1) { $bad += ('R13 declaration = ' + $n + ', want 1') }
}
$arg13 = Read-File 'src/backend/detail/expr/cgen_expr_call_arg_variant.inc'
if ($null -eq $arg13) { $bad += 'R13 cgen_expr_call_arg_variant.inc not found' }
else {
    $n = ([regex]::Matches($arg13, 'currencyByValArgWrap\(')).Count
    if ($n -ne 1) { $bad += ('R13 positional-arg station = ' + $n + ', want 1') }
}
$named13 = Read-File 'src/backend/detail/expr/cgen_expr_call_named_args.inc'
if ($null -eq $named13) { $bad += 'R13 cgen_expr_call_named_args.inc not found' }
else {
    $n = ([regex]::Matches($named13, 'currencyByValArgWrap\(')).Count
    if ($n -ne 1) { $bad += ('R13 named-arg station = ' + $n + ', want 1') }
}
foreach ($f in @('B180Main.bas', 'B180Res.vbp')) {
    $p = Join-Path $repo ('tests/b180_currency_byval/' + $f)
    if (-not (Test-Path -LiteralPath $p)) { $bad += ('R13 fixture file missing: ' + $f) }
}
if ($null -ne $run7) {
    $n = ([regex]::Matches($run7, "Test-Vbp 'b180_currency_byval")).Count
    if ($n -ne 2) { $bad += ('R13 b180 fixture cells = ' + $n + ', want 2 (x64 + x86)') }
}
$main13 = Read-File 'tests/b180_currency_byval/B180Main.bas'
if ($null -eq $main13) { $bad += 'R13 B180Main.bas not found' }
else {
    if ($main13 -notmatch 'CStr\(x\)') {
        $bad += 'R13 second column must print CStr(x): CLng alone hides the un-folded tail'
    }
    if ($main13 -notmatch 'TakeByVal2 1\.234567, "A1"') {
        $bad += 'R13 witness A1 (1.234567 -> 1.2346) missing'
    }
}
$main14 = Read-File 'tests/b181_currency_variant/B181Main.bas'
$def14 = Read-File 'src/backend/cgen_util_type.cpp'
if ($null -eq $def14) { $bad += 'R14 cgen_util_type.cpp not found' }
else {
    $n = ([regex]::Matches($def14, 'std::string CCodeGen::currencyVariantBox\(')).Count
    if ($n -ne 1) { $bad += ('R14 currencyVariantBox definition = ' + $n + ', want 1') }
    $i14 = $def14.IndexOf('std::string CCodeGen::currencyVariantBox(')
    if ($i14 -ge 0) {
        $body14 = $def14.Substring($i14, [Math]::Min(2500, $def14.Length - $i14))
        if ($body14 -notmatch 'knownCurrencyVars_\.count\(lower\)') {
            $bad += 'R14 the identity must be answered by knownCurrencyVars_ (one table with the fold stations)'
        }
        if ($body14 -notmatch 'return "vb6_VariantCur\(" \+ cExpr \+ "\)"') {
            $bad += 'R14 the box must be vb6_VariantCur (VT_CY), not a second packing formula'
        }
        if ($body14 -notmatch 'expr->kind != ASTNodeKind::IdentifierExpr') {
            $bad += 'R14 bare names only: members/array elements/arithmetic results are separate legs'
        }
    }
}
if ($null -ne $hdr1) {
    $n = ([regex]::Matches($hdr1, 'std::string currencyVariantBox\(')).Count
    if ($n -ne 1) { $bad += ('R14 declaration = ' + $n + ', want 1') }
}
$ctrl14 = Read-File 'src/backend/cgen_util_ctrl.cpp'
if ($null -eq $ctrl14) { $bad += 'R14 cgen_util_ctrl.cpp not found' }
else {
    $n = ([regex]::Matches($ctrl14, 'currencyVariantBox\(static_cast<Expr\*>\(valueNode\), cExpr\)')).Count
    if ($n -ne 1) { $bad += ('R14 Variant-assignment station = ' + $n + ', want 1') }
    $n = ([regex]::Matches($ctrl14, 'currencyVariantBox\(expr, cExpr\)')).Count
    if ($n -ne 1) { $bad += ('R14 argument station = ' + $n + ', want 1') }
}
foreach ($f in @('B181Main.bas', 'B181Res.vbp')) {
    $p = Join-Path $repo ('tests/b181_currency_variant/' + $f)
    if (-not (Test-Path -LiteralPath $p)) { $bad += ('R14 fixture file missing: ' + $f) }
}
if ($null -ne $run7) {
    $n = ([regex]::Matches($run7, "Test-Vbp 'b181_currency_variant")).Count
    if ($n -ne 2) { $bad += ('R14 b181 fixture cells = ' + $n + ', want 2 (x64 + x86)') }
}
if ($null -eq $main14) { $bad += 'R14 B181Main.bas not found' }
else {
    if ($main14 -notmatch 'P6') {
        $bad += 'R14 reverse witness P6 (Double must stay VT_R8) missing'
    }
    if ($main14 -notmatch 'VarType\(v\)' -or $main14 -notmatch 'TypeName\(v\)') {
        $bad += 'R14 the fixture must print both VarType and TypeName'
    }
    if ($main14 -match 'P4' -or $main14 -match 'P10') {
        $bad += 'R14 arithmetic results / array elements stay in the register: do not pin them here'
    }
}
$def15 = Read-File 'src/backend/cgen_util_type.cpp'
if ($null -eq $def15) { $bad += 'R15 cgen_util_type.cpp not found' }
else {
    $iCur15 = $def15.IndexOf('if (knownCurrencyVars_.count(lower)) return Vb6Type::Currency;')
    if ($iCur15 -lt 0) {
        $bad += 'R15 inferExprType must answer Currency for a registered Currency name'
    }
    $iDbl15 = $def15.IndexOf('if (knownDoubleVars_.count(lower)) return Vb6Type::Double;')
    if ($iCur15 -ge 0 -and $iDbl15 -ge 0 -and $iCur15 -gt $iDbl15) {
        $bad += 'R15 Currency must be asked before Double (a Currency is registered in both tables)'
    }
    $iMul15 = $def15.IndexOf('bin.op == BinaryOp::Mul')
    if ($iMul15 -lt 0) { $bad += 'R15 the * station is missing' }
    else {
        $win15 = $def15.Substring($iMul15, [Math]::Min(900, $def15.Length - $iMul15))
        if ($win15 -notmatch 'lCur15' -or $win15 -notmatch 'lt == Vb6Type::Single' -or $win15 -notmatch 'return Vb6Type::Double;') {
            $bad += 'R15 * with Single/Double must answer Double (probe A8/A9/B2), and only for a Currency x float pair'
        }
    }
}
foreach ($f in @('B182Main.bas', 'B182Res.vbp')) {
    $p = Join-Path $repo ('tests/b182_currency_arith/' + $f)
    if (-not (Test-Path -LiteralPath $p)) { $bad += ('R15 fixture file missing: ' + $f) }
}
if ($null -ne $run7) {
    $n = ([regex]::Matches($run7, "Test-Vbp 'b182_currency_arith")).Count
    if ($n -ne 2) { $bad += ('R15 b182 fixture cells = ' + $n + ', want 2 (x64 + x86)') }
}
$main15 = Read-File 'tests/b182_currency_arith/B182Main.bas'
if ($null -eq $main15) { $bad += 'R15 B182Main.bas not found' }
else {
    foreach ($w in @('A8', 'B3', 'E2', 'E3')) {
        if ($main15 -notmatch ('"' + $w + '"')) {
            $bad += ('R15 witness cell ' + $w + ' missing (floating side must stay Double/unrounded)')
        }
    }
    if ($main15 -notmatch 'B182-DONE') { $bad += 'R15 fixture sentinel B182-DONE missing' }
}
$bin16 = Read-File 'src/backend/expr/cgen_expr_binary.cpp'
if ($null -eq $bin16) { $bad += 'R16 cgen_expr_binary.cpp not found' }
else {
    $iBin16 = $bin16.IndexOf('inferExprType(node) == Vb6Type::Currency')
    if ($iBin16 -lt 0) {
        $bad += 'R16 the binary emit must round a Currency-typed + - * step'
    }
    else {
        $iWin16 = [Math]::Max(0, $iBin16 - 260)
        $win16 = $bin16.Substring($iWin16, [Math]::Min(900, $bin16.Length - $iWin16))
        if ($win16 -notmatch 'vb6_CCur\(') {
            $bad += 'R16 the step must fold through vb6_CCur (no second round formula)'
        }
        if ($win16 -notmatch 'BinaryOp::Add' -or $win16 -notmatch 'BinaryOp::Mul') {
            $bad += 'R16 the step must cover + - * (Div/Pow are always Double)'
        }
    }
    if ($bin16.Contains('vb6_CCur((" + left + " " + op + " " + right + "))') -eq $false) {
        $bad += 'R16 the fold shape is missing (vb6_CCur((left op right)))'
    }
}
foreach ($f in @('B183Main.bas', 'B183Res.vbp')) {
    $p = Join-Path $repo ('tests/b183_currency_step/' + $f)
    if (-not (Test-Path -LiteralPath $p)) { $bad += ('R16 fixture file missing: ' + $f) }
}
if ($null -ne $run7) {
    $n = ([regex]::Matches($run7, "Test-Vbp 'b183_currency_step")).Count
    if ($n -ne 2) { $bad += ('R16 b183 fixture cells = ' + $n + ', want 2 (x64 + x86)') }
}
$main16 = Read-File 'tests/b183_currency_step/B183Main.bas'
if ($null -eq $main16) { $bad += 'R16 B183Main.bas not found' }
else {
    foreach ($w in @('T4', 'T8', 'T3', 'T5')) {
        if ($main16 -notmatch ('"' + $w + '"')) {
            $bad += ('R16 witness cell ' + $w + ' missing (step vs Double-side)')
        }
    }
}
$arr17 = Read-File 'src/backend/expr/cgen_expr_array.cpp'
if ($null -eq $arr17) { $bad += 'R17 cgen_expr_array.cpp not found' }
else {
    if ($arr17.IndexOf('case Vb6Type::Currency: return "double";') -lt 0) {
        $bad += 'R17 a Currency array element must be stored as double (same model as a scalar)'
    }
    if ($arr17.IndexOf('case Vb6Type::Currency: return "int64_t";') -ge 0) {
        $bad += 'R17 the old int64_t Currency slot is still present (it truncates on store)'
    }
}
$ty17 = Read-File 'src/backend/cgen_util_type.cpp'
if ($null -eq $ty17) { $bad += 'R17 cgen_util_type.cpp not found' }
else {
    $i17 = $ty17.IndexOf('§11.164')
    if ($i17 -lt 0) { $bad += 'R17 the array-element Currency branch is missing' }
    else {
        $w17 = $ty17.Substring($i17, [Math]::Min(1600, $ty17.Length - $i17))
        if ($w17 -notmatch 'arrayElemTypes_' -or $w17 -notmatch 'Vb6Type::Currency') {
            $bad += 'R17 the element identity must be answered by arrayElemTypes_ as Currency'
        }
    }
}
foreach ($f in @('B184Main.bas', 'B184Res.vbp')) {
    $p = Join-Path $repo ('tests/b184_currency_array/' + $f)
    if (-not (Test-Path -LiteralPath $p)) { $bad += ('R17 fixture file missing: ' + $f) }
}
if ($null -ne $run7) {
    $n = ([regex]::Matches($run7, "Test-Vbp 'b184_currency_array")).Count
    if ($n -ne 2) { $bad += ('R17 b184 fixture cells = ' + $n + ', want 2 (x64 + x86)') }
}
$main17 = Read-File 'tests/b184_currency_array/B184Main.bas'
if ($null -eq $main17) { $bad += 'R17 B184Main.bas not found' }
else {
    foreach ($w in @('A4', 'A5', 'A7')) {
        if ($main17 -notmatch ('"' + $w + '"')) {
            $bad += ('R17 witness cell ' + $w + ' missing (ByRef / uninitialised / ReDim)')
        }
    }
}
$base18 = Read-File 'src/rtl/core/vb6rtl/vb6rtl_base.h'
if ($null -eq $base18) { $bad += 'R18 vb6rtl_base.h not found' }
else {
    if ($base18.IndexOf('vb6_CurrencyOverflows') -lt 0) {
        $bad += 'R18 there must be exactly one Currency range gate (vb6_CurrencyOverflows)'
    }
    if ($base18.IndexOf('922337203685477.5') -lt 0) {
        $bad += 'R18 the gate must use the Currency limit 922337203685477.5'
    }
    if ($base18 -match '\(int64_t\)\(ip \* 10000\.0 \+ round') {
        $bad += 'R18 an unguarded (int64_t) scale cast is still present (undefined behaviour)'
    }
}
$misc18 = Read-File 'src/rtl/core/vb6rtl/vb6rtl_misc.c'
if ($null -eq $misc18) { $bad += 'R18 vb6rtl_misc.c not found' }
else {
    $i18 = $misc18.IndexOf('double vb6_CCur(double v)')
    if ($i18 -lt 0) { $bad += 'R18 vb6_CCur not found' }
    else {
        $w18 = $misc18.Substring($i18, [Math]::Min(900, $misc18.Length - $i18))
        if ($w18 -notmatch 'vb6_CurrencyOverflows' -or $w18 -notmatch 'vb6_RaiseError\(6') {
            $bad += 'R18 vb6_CCur must raise Error 6 when the value is out of the Currency range'
        }
    }
}
foreach ($f in @('B185Main.bas', 'B185Res.vbp')) {
    $p = Join-Path $repo ('tests/b185_currency_overflow/' + $f)
    if (-not (Test-Path -LiteralPath $p)) { $bad += ('R18 fixture file missing: ' + $f) }
}
if ($null -ne $run7) {
    $n18 = ([regex]::Matches($run7, "Test-Vbp 'b185_currency_overflow")).Count
    if ($n18 -ne 2) { $bad += ('R18 b185 fixture cells = ' + $n18 + ', want 2 (x64 + x86)') }
}
$main18 = Read-File 'tests/b185_currency_overflow/B185Main.bas'
if ($null -eq $main18) { $bad += 'R18 B185Main.bas not found' }
else {
    foreach ($w in @('E1', 'E2', 'E3')) {
        if ($main18 -notmatch ('"' + $w + '="')) {
            $bad += ('R18 witness cell ' + $w + ' missing (overflow / clear)')
        }
    }
}
$ty19 = Read-File 'src/backend/cgen_util_type.cpp'
if ($null -eq $ty19) { $bad += 'R19 cgen_util_type.cpp not found' }
else {
    $i19 = $ty19.IndexOf('§11.166')
    if ($i19 -lt 0) { $bad += 'R19 the Abs/Int/Fix/Round Currency branch is missing' }
    else {
        $w19 = $ty19.Substring($i19, [Math]::Min(3000, $ty19.Length - $i19))
        foreach ($nm in @('"abs"', '"int"', '"fix"', '"round"', 'Vb6Type::Currency')) {
            if ($w19 -notmatch [regex]::Escape($nm)) {
                $bad += ('R19 the branch must name ' + $nm)
            }
        }
    }
}
$conv19 = Read-File 'src/rtl/core/vb6rtl/vb6rtl_conv.c'
if ($null -eq $conv19) { $bad += 'R19 vb6rtl_conv.c not found' }
else {
    # §11.207 (第二百二十四刀): Currency 那一格的身子提到共享的 vb6_VariantNumSameKind，
    # IntV/FixV 只负责传值 —— 所以「保 Currency」的判据跟着搬到新主人身上，
    # 同时仍钉住 IntV/FixV 自己不再另写一份（体内必须只有这一枚出口）。
    foreach ($fn in @('vb6_IntV', 'vb6_FixV')) {
        $j19 = $conv19.IndexOf('vb6_VARIANT ' + $fn + '(vb6_VARIANT v)')
        if ($j19 -lt 0) { $bad += ('R19 ' + $fn + ' not found') }
        else {
            $b19 = $conv19.Substring($j19, [Math]::Min(400, $conv19.Length - $j19))
            if ($b19 -notmatch 'vb6_VariantNumSameKind\(') {
                $bad += ('R19 ' + $fn + ' must route through vb6_VariantNumSameKind')
            }
        }
    }
    $k19 = $conv19.IndexOf('vb6_VariantNumSameKind(vb6_VARIANT v, double d)')
    if ($k19 -lt 0) { $bad += 'R19 vb6_VariantNumSameKind not found' }
    else {
        $c19 = $conv19.Substring($k19, [Math]::Min(900, $conv19.Length - $k19))
        foreach ($nm19 in @('vb6_vtCurrency', 'vb6_VariantCur')) {
            if ($c19 -notmatch [regex]::Escape($nm19)) {
                $bad += ('R19 the carrier wrapper must name ' + $nm19)
            }
        }
    }
}
foreach ($f in @('B191Main.bas', 'B191Res.vbp')) {
    $p = Join-Path $repo ('tests/b191_currency_math/' + $f)
    if (-not (Test-Path -LiteralPath $p)) { $bad += ('R19 fixture file missing: ' + $f) }
}
if ($null -ne $run7) {
    $n19 = ([regex]::Matches($run7, "Test-Vbp 'b191_currency_math")).Count
    if ($n19 -ne 2) { $bad += ('R19 b191 fixture cells = ' + $n19 + ', want 2 (x64 + x86)') }
}
$main19 = Read-File 'tests/b191_currency_math/B191Main.bas'
if ($null -eq $main19) { $bad += 'R19 B191Main.bas not found' }
else {
    foreach ($w in @('M11', 'M15')) {
        if ($main19 -notmatch ('"' + $w + '"')) {
            $bad += ('R19 witness cell ' + $w + ' missing (Variant carrier / ^ stays Double)')
        }
    }
}
# ---- R20 (§11.207 第二百二十四刀): 返回子类型 = 实参子类型（不只 Currency/Date）----
# 真机 .temp/k224p/P298/P300/P302: Integer 2 / Long 3 / Single 4 / Currency 6 /
# Byte 17 / Double 5 / Date 7 / Boolean→2，Int/Fix 的 Variant 载体同一条。
$ty20 = Read-File 'src/backend/cgen_util_type.cpp'
if ($null -eq $ty20) { $bad += 'R20 cgen_util_type.cpp not found' }
else {
    if ($ty20 -notmatch 'isNarrowIntExpr\(call\.positional\[0\]\.get\(\)\)') {
        $bad += 'R20 the numeric-kind branch no longer asks isNarrowIntExpr first (Round(i) would answer Long)'
    }
    foreach ($nm in @('Vb6Type::Byte', 'Vb6Type::Single', 'Vb6Type::Long')) {
        if ($ty20 -notmatch [regex]::Escape($nm)) { $bad += ('R20 the branch must name ' + $nm) }
    }
}
$ct20 = Read-File 'src/backend/cgen_util_ctrl.cpp'
if ($null -eq $ct20) { $bad += 'R20 cgen_util_ctrl.cpp not found' }
else {
    if ($ct20 -notmatch 'case Vb6Type::Integer:') {
        $bad += 'R20 boxToVariant lost its Integer arm (Int(i) would box as VT_R8)'
    }
}
$conv20 = Read-File 'src/rtl/core/vb6rtl/vb6rtl_conv.c'
if ($null -eq $conv20) { $bad += 'R20 vb6rtl_conv.c not found' }
else {
    if ($conv20 -notmatch 'vb6_VARIANT vb6_RoundV\(') { $bad += 'R20 vb6_RoundV definition missing' }
    $n20 = @([regex]::Matches($conv20, 'vb6_VariantNumSameKind\(')).Count
    if ($n20 -ne 4) { $bad += ('R20 vb6_VariantNumSameKind sites = ' + $n20 + ' (want 4 = def + Int/Fix/Round)') }
    if ($conv20 -notmatch 'vb6_vtByte:\s+return vb6_VariantByte') {
        $bad += 'R20 the carrier-kind wrapper lost its Byte arm'
    }
}
$runner20 = Read-File 'tests/run_tests.ps1'
if ($null -eq $runner20) { $bad += 'R20 run_tests.ps1 not found' }
else {
    $n20b = @([regex]::Matches($runner20, 'b231_num_kind\\B231Res\.vbp')).Count
    if ($n20b -ne 2) { $bad += ('R20 b231 run cells = ' + $n20b + ' (want 2 = x64 + x86)') }
}

if ($bad.Count -gt 0) {
    Write-Host ('FAIL currency_round_station: ' + ($bad -join ' | '))
    exit 1
}
Write-Host '  R13 ByVal Currency call-site station (one owner, both arg paths, no ByRef fold) OK'
Write-Host '  R14 Currency variable identity in Variant (VT_CY, both box stations) OK'
Write-Host '  R15 Currency arithmetic identity (+ - absorb, * with float is Double, / is Double) OK'
Write-Host '  R16 Currency intermediate rounds at every step (single fold exit) OK'
Write-Host '  R17 Currency array elements are doubles with the same identity (store/read/ByRef) OK'
Write-Host '  R18 Currency range gate raises Error 6 on overflow (no UB wrap) OK'
Write-Host '  R19 Abs/Int/Fix/Round keep the Currency identity (typed + Variant carrier) OK'
Write-Host '  R20 Abs/Int/Fix/Round answer the argument numeric kind (Integer/Long/Single/Byte/Date/Boolean) OK'
Write-Host 'PASS currency_round_station (R1..R20)'
exit 0
