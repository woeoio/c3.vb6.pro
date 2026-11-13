# check_variant_unbox.ps1 —— 「Variant 表达式 → 具体型」的读腿唯一判据（账 §56.40，C2440 VARIANT 族）
#
# 为什么要有这一道：VB6 的 Variant 是**运行期带标签**的值。生成 C 里它是整只
# `vb6_VARIANT` 结构体；凡是「Variant 值流进具体型」的地方，都必须先按上下文
# 拆箱（vb6_VariantToString / ToLong / ToDouble），否则结构体直接递给 BSTR /
# double / int32_t 形参 = C2440。
#
# 这一族之所以集体漏掉，是因为「这枚表达式是不是整只 Variant 槽」这个问题有
# 两个实现：**文本投影**（isVariantSlot 的前缀表）与 **AST 权威**
# （isDefinitelyVariantExpr）。消费点过去一律只问文本那一半 ——
#   - UDT 字段声明为 Variant：C 文本是 `VB6_SA_AT(vb6_type_X, arr, i).member`，
#     不命中任何前缀表 ⇒ 答否 ⇒ 不拆箱
#     （PhotoDemon pdDictionary.c 239/248：`dItem As Variant` 读进 String 变量）；
#   - 类方法返回 Variant：C 文本是 `vb6_<Cls>_<Fn>(...)`，同理答否
#     （PhotoDemon Evaluator.c 474/494/514、pdSpinner.c 919、
#       TextSupport.c 357/370/386/398、toolpanel_TextBasic.c 1811）。
# AST 那一半本来就答得对，只是没被问到 —— 本刀把消费点改问 AST ∪ 文本。
#
# 判据面（任何一条不齐就红）：
#   R1 toLongIfVariant / toDoubleIfVariant / toStringIfVariant 三枚各定义 1 处，
#      且三枚内部都必须问 `isVariantSlot(cExpr, astExpr)`（不许再传 nullptr）。
#   R2 三个「值 → 具体型」的赋值消费点都必须把 AST 交上去：
#      赋值（cgen_assign_value_sem.inc）/ Let（cgen_setlet.cpp）/
#      BSTR 槽写入（cgen_util_ctrl.cpp 的 emitBstrSlotStore）。
#   R3 「类成员方法返回 Variant」这一问只有一个出口
#      （classMemberMethodReturnsVariant：定义 1 + 声明 1 + 消费 1）。
#   R4 isDefinitelyVariantExpr 的 UDT 档必须问 inferUdtFieldVb6Type
#      （字段声明档位的既有唯一出口；inferExprType 走 lookupModule，跨模块
#        Public Type 解析不到，会答 Unknown）。
#   R5 二元运算符三处（Like / ^ / Mod）必须走这一族出口，不许把操作数原样递进去。
#   R6 ParamArray 实参那一站必须问 AST。
#   R7 census 不许读空（文件长度 + 每条命中数）。
#
# 用法: powershell -File scripts\check_variant_unbox.ps1   (PASS = exit 0)

$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot
$pClasstype = Join-Path $Root 'src\backend\cgen_util_classtype.cpp'
$pHelpers   = Join-Path $Root 'src\backend\detail\util\cgen_helpers.inc'
$pType      = Join-Path $Root 'src\backend\cgen_util_type.cpp'
$pAssign    = Join-Path $Root 'src\backend\detail\stmt\cgen_assign_value_sem.inc'
$pSetlet    = Join-Path $Root 'src\backend\stmt\cgen_setlet.cpp'
$pCtrl      = Join-Path $Root 'src\backend\cgen_util_ctrl.cpp'
$pBinary    = Join-Path $Root 'src\backend\expr\cgen_expr_binary.cpp'
$pNamedArgs = Join-Path $Root 'src\backend\detail\expr\cgen_expr_call_named_args.inc'
foreach ($p in @($pClasstype, $pHelpers, $pType, $pAssign, $pSetlet, $pCtrl, $pBinary, $pNamedArgs)) {
    if (-not (Test-Path -LiteralPath $p)) { Write-Host ('FAIL: 读不到 ' + $p); exit 1 }
}

$cls = [IO.File]::ReadAllText($pClasstype)
$hlp = [IO.File]::ReadAllText($pHelpers)
$typ = [IO.File]::ReadAllText($pType)
$asg = [IO.File]::ReadAllText($pAssign)
$set = [IO.File]::ReadAllText($pSetlet)
$ctl = [IO.File]::ReadAllText($pCtrl)
$bin = [IO.File]::ReadAllText($pBinary)
$nam = [IO.File]::ReadAllText($pNamedArgs)

$viol = @()

# ---- R1: 三枚出口各定义 1 处, 且**函数体内**都问 AST ----
foreach ($fn in @('toLongIfVariant', 'toDoubleIfVariant', 'toStringIfVariant')) {
    $def = ([regex]::Matches($cls, ('std::string CCodeGen::' + $fn + '\s*\('))).Count
    if ($def -ne 1) { $viol += ('R1 ' + $fn + ' 定义 ' + $def + ' 处 (钉死 1 处)') }
    # 逐个函数体取出来看 —— 不能按整文件数 (同文件里别的函数如 isVariantCarrierExpr
    # 的签名里就没有 AST, 传 nullptr 是对的)。
    $m = [regex]::Match($cls, ('std::string CCodeGen::' + $fn + '\s*\([^)]*\)\s*\{'))
    if (-not $m.Success) {
        $viol += ('R1 ' + $fn + ' 找不到函数体 —— 判据读空')
        continue
    }
    $start = $m.Index
    $end = $cls.IndexOf("`n}", $start)
    if ($end -lt 0) { $end = [Math]::Min($start + 2000, $cls.Length) }
    $body = $cls.Substring($start, $end - $start)
    if ($body -notmatch 'isVariantSlot\(cExpr, astExpr\)') {
        $viol += ('R1 ' + $fn + ' 体内没问 AST (isVariantSlot(cExpr, astExpr)) —— 那一族仍然漏箱')
    }
    if ($body -match 'isVariantSlot\(cExpr, nullptr\)') {
        $viol += ('R1 ' + $fn + ' 体内还有只问文本那一半的调用 —— AST 丢了')
    }
}
$astAsk = ([regex]::Matches($cls, 'isVariantSlot\(cExpr, astExpr\)')).Count
if ($astAsk -ne 3) { $viol += ('R1 三枚出口里问 AST 的 ' + $astAsk + ' 处 (钉死 3 处 —— 少了就是把 AST 丢了)') }
# 只作读数 (同文件里 isVariantCarrierExpr 的签名里没有 AST, 那一处传 nullptr 是对的)
$nullAsk = ([regex]::Matches($cls, 'isVariantSlot\(cExpr, nullptr\)')).Count

# ---- R2: 三个赋值消费点必须交 AST ----
if ($asg -notmatch 'isVariantSlot\(value, node\.value\.get\(\)\)') {
    $viol += 'R2 赋值那一站没把 AST 交上去 (isVariantSlot(value, node.value.get()))'
}
if ($set -notmatch 'isVariantSlot\(value, node\.value\.get\(\)\)') {
    $viol += 'R2 Let 那一站没把 AST 交上去 (isVariantSlot(value, node.value.get()))'
}
if ($ctl -notmatch 'isVariantSlot\(value, rhs\)') {
    $viol += 'R2 BSTR 槽写入那一站 (emitBstrSlotStore) 没把 AST 交上去 (isVariantSlot(value, rhs))'
}

# ---- R3: 「类成员方法返回 Variant」只有一个出口 ----
$defR3 = ([regex]::Matches($typ, 'bool CCodeGen::classMemberMethodReturnsVariant\s*\(')).Count
if ($defR3 -ne 1) { $viol += ('R3 classMemberMethodReturnsVariant 定义 ' + $defR3 + ' 处 (钉死 1 处)') }
$declR3 = ([regex]::Matches($hlp, 'bool classMemberMethodReturnsVariant\s*\(')).Count
if ($declR3 -ne 1) { $viol += ('R3 classMemberMethodReturnsVariant 声明 ' + $declR3 + ' 处 (钉死 1 处)') }
$useR3 = ([regex]::Matches($typ, 'classMemberMethodReturnsVariant\s*\(')).Count - 1
if ($useR3 -ne 1) { $viol += ('R3 classMemberMethodReturnsVariant 消费 ' + $useR3 + ' 处 (钉死 1 处: isDefinitelyVariantExpr 的成员档)') }
if ($typ -notmatch 'memberReturnTypes\.find') {
    $viol += 'R3 那一问没读 memberReturnTypes —— 类方法返回档的权威在那张表里'
}

# ---- R4: UDT 字段档位问 inferUdtFieldVb6Type ----
if ($typ -notmatch 'inferUdtFieldVb6Type\(&expr\)\s*==\s*Vb6Type::Variant') {
    $viol += 'R4 isDefinitelyVariantExpr 的 UDT 档没问 inferUdtFieldVb6Type —— 跨模块 Public Type 的字段答不到'
}

# ---- R5: 二元运算符三处走这一族出口 ----
if ($bin -notmatch 'vb6_Like[^\n]*\n?[\s\S]{0,200}?toStringIfVariant') {
    if ($bin -notmatch 'toStringIfVariant\(left, node\.left\.get\(\)\)') {
        $viol += 'R5 Like 那一支没走 toStringIfVariant —— Variant 操作数整只结构体递给 BSTR 形参 (C2440)'
    }
}
if ($bin -notmatch 'toDoubleIfVariant\(left, node\.left\.get\(\)\)') {
    $viol += 'R5 ^ 那一支没走 toDoubleIfVariant —— Variant 操作数整只结构体递给 double 形参 (C2440)'
}
if ($bin -notmatch 'vb6_Num_Mod\(" \+ toLongIfVariant') {
    if ($bin -notmatch 'toLongIfVariant\(left, node\.left\.get\(\)\)') {
        $viol += 'R5 Mod 那一支没走 toLongIfVariant —— (int32_t)(Variant) 是 C2440'
    }
}
# 反钉: 旧形状不许复活
if ($bin -match 'vb6_Num_Mod\(\(int32_t\)\(') {
    $viol += 'R5 Mod 那一支还在用裸 (int32_t)() 强转 —— Variant 结构体不能这么转'
}

# ---- R6: ParamArray 实参那一站问 AST ----
if ($nam -notmatch 'isVariantSlot\(paArgExpr, node\.positional\[argIdx\]\.get\(\)\)') {
    $viol += 'R6 ParamArray 实参那一站没问 AST (isVariantSlot(paArgExpr, node.positional[argIdx].get()))'
}

# ---- R7: census ----
$census = [pscustomobject]@{
    astAsk = $astAsk; nullAsk = $nullAsk
    defR3 = $defR3; declR3 = $declR3; useR3 = $useR3
}
$minLen = @($cls, $hlp, $typ, $asg, $set, $ctl, $bin, $nam) | ForEach-Object { $_.Length } | Measure-Object -Minimum
if ($minLen.Minimum -lt 1000) {
    $viol += 'R7 源文件读出过短 —— 扫描读空, 不是一条绿着的判据'
}

Write-Host ('census: astAsk=' + $census.astAsk + ' nullAsk=' + $census.nullAsk +
            ' classMemberMethodReturnsVariant def=' + $census.defR3 +
            ' decl=' + $census.declR3 + ' use=' + $census.useR3)

if ($viol.Count -gt 0) {
    foreach ($v in $viol) { Write-Host ('  ' + $v) }
    Write-Host ('FAIL variant_unbox: ' + $viol.Count + ' 条违规 —— 恢复口径是让消费点改问 AST 那一半 (AST ∪ 文本), 不是放宽判据')
    exit 1
}
Write-Host 'PASS variant_unbox (R1..R7)'
exit 0
