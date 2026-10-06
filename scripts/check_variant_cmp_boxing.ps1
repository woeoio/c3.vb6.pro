# check_variant_cmp_boxing.ps1 - 账 #238 的结构性哨兵 (只扫源码, 不起 cl)
#
# 为什么单开这一道: vb6_VarCmp*(&A, &B) 这四个取址点以前只按"名字形状像左值"决定能不能 &。
# 于是 `Dim d As Double` 的地址被当 vb6_VARIANT* 递进 RTL —— 按 VARIANT 的布局读一个 8 字节
# 标量: 相等的两个数答 False (改前两架构实测), 而且读过头。形状上编译、运行、诊断都不响,
# 所以判据必须结构钉住: 取址那一问只许有一处答案 (CCodeGen::cmpOperandMayTakeAddr → 问
# isDefinitelyVariantExpr, 它读声明那几张表 + 符号表)。
#
# 规则 (改坏了会红, 不是装饰):
#   V1  判据只有一份: 声明 1 处 (cgen_helpers.inc) + 定义 1 处 (cgen_util_type.cpp),
#       且定义里必须真的问类型权威 (isDefinitelyVariantExpr) —— 换成自己再抄一张名单就是回归
#   V2  四个取址点全都问它 (兜底那一处对称地问左右两侧 => 调用次数 = 5)
#   V3  每一条把 "(&" + left / right 拼进发码的行, 自己或上面三行内必须有那次问话;
#       且不许出现 `isLvalue(...) ? ("&" + ...)` 那一形 (就是本刀撤掉的旧判据)
#
# 用法:  pwsh -File scripts\check_variant_cmp_boxing.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$bad = @()

$beDir = Join-Path $root "src\backend"
$binFile = Join-Path $beDir "expr\cgen_expr_binary.cpp"
$typeFile = Join-Path $beDir "cgen_util_type.cpp"
$helpersFile = Join-Path $beDir "detail\util\cgen_helpers.inc"
foreach ($p in @($binFile, $typeFile, $helpersFile)) {
    if (-not (Test-Path -LiteralPath $p)) {
        Write-Host ("FAIL V0 missing " + $p) -ForegroundColor Red
        exit 1
    }
}

$bin = [System.IO.File]::ReadAllText($binFile)
$typ = [System.IO.File]::ReadAllText($typeFile)
$hlp = [System.IO.File]::ReadAllText($helpersFile)

$PRED = "cmpOperandMayTakeAddr"

# ---- V1: 一处声明 + 一处定义, 定义里问权威 ----
$declCount = @([regex]::Matches($hlp, [regex]::Escape("bool " + $PRED))).Count
$defCount = @([regex]::Matches($typ, [regex]::Escape("bool CCodeGen::" + $PRED))).Count
if ($declCount -ne 1) {
    $bad += ("V1 declaration of " + $PRED + " in cgen_helpers.inc = " + $declCount + " (want exactly 1)")
}
if ($defCount -ne 1) {
    $bad += ("V1 definition of " + $PRED + " in cgen_util_type.cpp = " + $defCount + " (want exactly 1)")
} else {
    $mDef = [regex]::Match($typ, 'bool\s+CCodeGen::' + $PRED + '\([\s\S]*?\r?\n\}')
    if (-not $mDef.Success) {
        $bad += "V1 body of the predicate not found"
    } elseif ($mDef.Value.IndexOf("isDefinitelyVariantExpr") -lt 0) {
        $bad += ("V1 " + $PRED + " no longer asks isDefinitelyVariantExpr -> it answers the " +
                "Variant-ness of a name by itself again (that is how 238 got in)")
    }
}

# ---- V2: 四个取址点都问它 (兜底那一对称地问两侧 => 一共 5 次调用) ----
$calls = @([regex]::Matches($bin, [regex]::Escape($PRED) + "\s*\(")).Count
if ($calls -ne 5) {
    $bad += ("V2 call sites of " + $PRED + " in cgen_expr_binary.cpp = " + $calls +
             " (want exactly 5: variantAddr's own name leg + the two VarCmpLong legs +" +
             " both ends of the VarCmp fallback)")
}

# ---- V3: 每一条 & 拼接都被问过; 旧判据那一形不许回潮 ----
$binLines = $bin -split "`r?`n"
$ampPat = '"\(&" \+ (left|right)|"&" \+ (left|right)'
$unguarded = @()
foreach ($i in 0..($binLines.Count - 1)) {
    $t = $binLines[$i]
    if ($t -notmatch $ampPat) { continue }
    $seen = $false
    for ($k = $i; $k -ge [Math]::Max(0, $i - 3); $k--) {
        if ($binLines[$k].Contains($PRED)) { $seen = $true; break }
    }
    if (-not $seen) {
        $unguarded += ("line " + ($i + 1) + ": " + $t.Trim().Substring(0, [Math]::Min(64, $t.Trim().Length)))
    }
}
if ($unguarded.Count -ne 0) {
    $bad += ("V3 " + $unguarded.Count + " address-taking emit line(s) not gated by " + $PRED +
             " -> " + ($unguarded -join " | "))
}
$oldShape = @($binLines | Where-Object { $_.Contains("isLvalue(") -and $_.Contains('"&" +') })
if ($oldShape.Count -ne 0) {
    $bad += ("V3 the name-shape-only predicate is back (" + $oldShape.Count +
             " line(s)) -> a bare identifier is not proof of vb6_VARIANT storage")
}
$ampTotal = @($binLines | Where-Object { $_ -match $ampPat }).Count
if ($ampTotal -lt 4) {
    $bad += ("V3 only " + $ampTotal + " line(s) concatenate an address into a VarCmp call -" +
             " the four sites moved somewhere else, update this census on purpose")
}

if ($bad.Count -eq 0) {
    Write-Host ("PASS variant cmp boxing: one predicate, defined once, asks the type authority," +
                " 4 take-address sites all gated, no name-shape-only form left")
    exit 0
}
foreach ($b in $bad) { Write-Host ("FAIL " + $b) -ForegroundColor Red }
exit 1
