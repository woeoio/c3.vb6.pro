# check_array_literal_variants.ps1 - Account #611 knife 348: Array() element kinds
#
# Why: Array(True/False/Null/Empty/Nothing/Array(...)) all used to be emitted with
# vb6_ArraySetLong (VT_I4), and Nothing / a nested array pushed a pointer into an
# int32 parameter. vb6_ArraySetVariant (vb6rtl_date.c) + the vb6_Variant* factories
# are the one authority for "this slot keeps its own vt". This sentinel pins the
# dispatch in cgen_expr_call_builtin_pre.inc at source level; the .bas fixture
# pins the runtime VarType readings (x64 + x86).
#
# Rules (each one is a real red when broken, not decoration):
#   R1  the Array() block can be located
#   R2  vb6_ArraySetVariant is emitted from exactly one place in that block
#   R3  all six LiteralKind cases are present (Boolean/Null/Empty/Nothing/Date/Single)
#   R4  the three VARIANT factories are used, with the exact occurrence counts
#   R5  nested Array() is detected by callee name + the _arr_ temp prefix
#   R6  the new branch precedes the legacy isLongArg arm (or it gets swallowed)
#   R7  the three legacy setters are still there (zero-regression guard)
#   R8  the runtime signature table still declares vb6_ArraySetVariant's 3rd arg
#   R9  the harness registers the fixture and its discriminating needles
#   R10 the fixture itself still contains those needles
#   R11 (第六十刀) the other four kind-fidelity outlets: ArraySetObject x3, Byte/Cur/Single x1
# Usage: powershell -File scripts\check_array_literal_variants.ps1  (PASS = exit 0)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$inc = Join-Path $root 'src\backend\detail\expr\cgen_expr_call_builtin_pre.inc'
$sig = Join-Path $root 'src\backend\cgen_util_type.cpp'
$harness = Join-Path $root 'tests\run_tests.ps1'
$fixture = Join-Path $root 'tests\jsonfast_611\ArrayLiteralVariants.bas'
foreach ($p in @($inc, $sig, $harness, $fixture)) {
    if (-not (Test-Path -LiteralPath $p)) { Write-Host "FAIL missing path: $p"; exit 1 }
}
$s = [System.IO.File]::ReadAllText($inc)
$g = [System.IO.File]::ReadAllText($sig)
$h = [System.IO.File]::ReadAllText($harness)
$f = [System.IO.File]::ReadAllText($fixture)
$bad = @()

# ---- R1: locate the Array() branch (ASCII anchors only, PS 5.1 safe) ----
$a = $s.IndexOf('arrLower == "array"')
$b = $s.IndexOf('P20-36', [Math]::Max($a, 0))
if ($a -lt 0 -or $b -le $a) {
    $bad += 'R1 cannot locate the Array() block'
    $blk = ''
} else {
    $blk = $s.Substring($a, $b - $a)

    # ---- R2: exactly one new setter outlet ----
    $nSetter = @([regex]::Matches($blk, 'vb6_ArraySetVariant\(')).Count
    if ($nSetter -ne 1) { $bad += "R2 vb6_ArraySetVariant call sites = $nSetter (want 1)" }

    # ---- R3: six literal kinds ----
    foreach ($k in @('LiteralKind::Boolean', 'LiteralKind::Null', 'LiteralKind::Empty',
                     'LiteralKind::Nothing', 'LiteralKind::Date', 'LiteralKind::Single')) {
        if ($blk.IndexOf($k) -lt 0) { $bad += "R3 missing case $k" }
    }

    # ---- R4: factories with exact counts ----
    # vb6_VariantDate( = 2 是 C30-11 第六十刀加的：字面量 `#1/2/2000#` 一格 +
    # `Array(<Date 值>)`（如 CDate(...) / Date 变量）一格，两条都是真出口。
    foreach ($pair in @(@('vb6_VariantBool(', 1), @('vb6_VariantDate(', 2),
                        @('vb6_VariantFromValue(', 3))) {
        $n = @([regex]::Matches($blk, [regex]::Escape($pair[0]))).Count
        if ($n -ne $pair[1]) { $bad += ("R4 " + $pair[0] + " count = $n (want " + $pair[1] + ")") }
    }

    # ---- R11 (第六十刀): 档保真的其余四个出口 ----
    # 对象/ Nothing 走 vb6_ArraySetObject（3 个调用点：Nothing 字面量、knownObjectVars_、
    # inferExprType==Object），Byte/Currency/Single 各一个 Variant 工厂。
    # 钉住理由：这四档曾经全落进"默认 vb6_ArraySetLong"（Array(obj) 答 TypeName=Long，
    # Array(CByte(65)) 答 VarType 3，Array(CDate(..)) 更因实参里含 L"..." 被猜成 BSTR ⇒ C2440）。
    foreach ($pair in @(@('vb6_ArraySetObject(', 3), @('vb6_VariantByte(', 1),
                        @('vb6_VariantCur(', 1), @('vb6_VariantSingle(', 1))) {
        $n = @([regex]::Matches($blk, [regex]::Escape($pair[0]))).Count
        if ($n -ne $pair[1]) { $bad += ("R11 " + $pair[0] + " count = $n (want " + $pair[1] + ")") }
    }

    # ---- R5: nested Array() detection ----
    foreach ($needle in @('ASTNodeKind::IndexOrCallExpr', '"_arr_"', 'nestedName == "array"')) {
        if ($blk.IndexOf($needle) -lt 0) { $bad += "R5 missing nested-array guard $needle" }
    }

    # ---- R6: new branch must come first ----
    $iNew = $blk.IndexOf('if (!variantSetter.empty())')
    $iOld = $blk.IndexOf('bool isLongArg = false;')
    if ($iNew -lt 0 -or $iOld -lt 0 -or $iNew -ge $iOld) {
        $bad += 'R6 variantSetter branch is not ahead of the legacy isLongArg arm'
    }

    # ---- R7: legacy setters survive ----
    foreach ($k in @('vb6_ArraySetLong(', 'vb6_ArraySetDouble(', 'vb6_ArraySetBSTR(')) {
        if ($blk.IndexOf($k) -lt 0) { $bad += "R7 legacy setter $k disappeared" }
    }
}

# ---- R8: signature table ----
$sigLine = @($g -split "`r?`n" | Where-Object { $_ -match 'vb6_ArraySetVariant' })
if ($sigLine.Count -ne 1 -or $sigLine[0] -notmatch 'vb6_VARIANT') {
    $bad += 'R8 runtime signature table lost vb6_ArraySetVariant / vb6_VARIANT third arg'
}

# ---- R9: harness registration ----
foreach ($needle in @('ArrayLiteralVariants.bas', 'AL-bt=11', 'AL-nothing=9',
                      'AL-nested=8204', 'AL-date=7', 'AL-obj=9', 'AL-byte=17',
                      'AL-cy=6', 'AL-sng=4', 'AL-date2=7')) {
    if ($h.IndexOf($needle) -lt 0) { $bad += "R9 harness missing $needle" }
}

# ---- R10: fixture needles ----
foreach ($needle in @('AL-bt=', 'AL-nothing=', 'AL-nested=', 'AL-obj=', 'AL-byte=',
                      'AL-cy=', 'AL-sng=', 'AL-date2=', 'AL-DONE')) {
    if ($f.IndexOf($needle) -lt 0) { $bad += "R10 fixture missing $needle" }
}

if ($bad.Count -gt 0) {
    $bad | ForEach-Object { Write-Host "FAIL $_" }
    exit 1
}
Write-Host 'PASS array_literal_variants'
exit 0
