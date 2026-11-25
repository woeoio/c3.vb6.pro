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
    if ($conv10 -notmatch 'double\s+vb6_CCurBSTR\(BSTR s\)\s*\{\s*return vb6_CCur\(vb6_NumVal\(s\)\);\s*\}') {
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
Write-Host 'PASS currency_round_station (R1..R12)'
exit 0
