# check_byref_arg_pack_authority.ps1 - 晚绑定实参「按引用发」只有一位主人
#   (账 #305 §B224 第二半，第二百三十九刀)
#
# 起因：真 VB6 客户端实测 (ai/probes/bool_string B4: `Dim o As Object` + `o.Bump(x)`，x As Long
# ⇒ x=11) —— 晚绑定对**带型变量**递的是 VT_BYREF|<该型>，服务端的写回落在调用方变量上；
# VBScript 的 Variant 实参递 VT_BYREF|VT_VARIANT 被 VB6 服务端判类型不匹配 ⇒ 差别在实参带不带型。
# 改前我们的客户端对**标量一律按值打包**，对外部服务器的 ByRef 写回因此落在虚空。
#
# 规则（改坏了会红）：
#   R1  五枚 Ref 打包器 (Int/Integer/Bool/Single/Double) 在 vb6com_pack.c 各定义 1 次、
#       在 vb6com.h 各声明 1 次 —— 按 C 存储宽度分标签，Integer/Boolean 都是 2 字节。
#   R2  「这枚实参该不该按引用」的唯一主人 comByRefArgExpr：定义 1 (cgen_util_com.cpp) +
#       声明 1 (cgen_helpers.inc) + 调用 2 (com_bind 的早绑定/晚绑定两支)。
#   R3  类型 → 打包器那张 switch 五档各 1 行；Integer 必须问 knownIntVars_ 而不是
#       inferExprType (后者对 Integer 答 Long ⇒ 按 I4 写回 = 越界写 4 字节)。
#   R4  早绑定那一支只对声明为 ByRef 的形参取址 (ByVal 一律不取址)：positional 与 named 两处闸。
#   R5  判据面：夹具两件 + runner 里五枚正针 / 两枚「仍按值」证人 / 三枚反面针。
#
# 用法: powershell -File scripts\check_byref_arg_pack_authority.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
if ($env:C3_CHECK_ROOT) { $root = $env:C3_CHECK_ROOT }

$packC   = Join-Path $root 'src\rtl\core\vb6com\vb6com_pack.c'
$comHdr  = Join-Path $root 'src\rtl\core\vb6com\vb6com.h'
$utilC   = Join-Path $root 'src\backend\cgen_util_com.cpp'
$helpers = Join-Path $root 'src\backend\detail\util\cgen_helpers.inc'
$bindInc = Join-Path $root 'src\backend\detail\expr\cgen_expr_call_com_bind.inc'
$harness = Join-Path $root 'tests\run_tests.ps1'
$fixBas  = Join-Path $root 'tests\b239_byrefarg\B239k.bas'
$fixVbp  = Join-Path $root 'tests\b239_byrefarg\B239k.vbp'

foreach ($p in @($packC, $comHdr, $utilC, $helpers, $bindInc, $harness, $fixBas, $fixVbp)) {
    if (-not (Test-Path -LiteralPath $p)) { Write-Host ('FAIL missing ' + $p) -ForegroundColor Red; exit 1 }
}

function Count-Pattern($path, $pattern) {
    $t = [System.IO.File]::ReadAllText($path)
    return ([regex]::Matches($t, $pattern)).Count
}

$fails = @()
function Expect($tag, $got, $want, $note) {
    if ($got -ne $want) {
        $script:fails += ('{0} = {1} (want {2}) -- {3}' -f $tag, $got, $want, $note)
        Write-Host ('FAIL {0} = {1}, want {2}  ({3})' -f $tag, $got, $want, $note) -ForegroundColor Red
    } else {
        Write-Host ('ok   {0} = {1}' -f $tag, $got) -ForegroundColor Green
    }
}

# R1 —— 五枚 Ref 打包器：定义各 1、声明各 1
foreach ($n in @('Int', 'Integer', 'Bool', 'Single', 'Double')) {
    Expect ("R1 vb6_ComPack${n}Ref defined")  (Count-Pattern $packC  ('(?m)^void\* vb6_ComPack' + $n + 'Ref\(')) 1 '定义只许一枚, 在 vb6com_pack.c'
    Expect ("R1 vb6_ComPack${n}Ref declared") (Count-Pattern $comHdr ('void\* vb6_ComPack' + $n + 'Ref\(')) 1 '声明只许一枚, 在 vb6com.h'
}

# R2 —— 唯一主人
Expect 'R2 comByRefArgExpr defined'  (Count-Pattern $utilC   '(?m)^std::string CCodeGen::comByRefArgExpr') 1 '定义只许一枚, 在 cgen_util_com.cpp'
Expect 'R2 comByRefArgExpr declared' (Count-Pattern $helpers 'std::string comByRefArgExpr\(Expr&') 1 '声明在 cgen_helpers.inc'
Expect 'R2 comByRefArgExpr called'   (Count-Pattern $bindInc 'comByRefArgExpr\(') 3 '三处消费者: 早绑定位置/命名 + 晚绑定'

# R3 —— 五档映射 + Integer 的宽度问句
Expect 'R3 IntRef mapping'      (Count-Pattern $utilC 'fn = "vb6_ComPackIntRef"') 1 'Long → VT_I4'
Expect 'R3 IntegerRef mapping'  (Count-Pattern $utilC 'fn = "vb6_ComPackIntegerRef"') 1 'Integer → VT_I2'
Expect 'R3 BoolRef mapping'     (Count-Pattern $utilC 'fn = "vb6_ComPackBoolRef"') 1 'Boolean → VT_BOOL'
Expect 'R3 SingleRef mapping'   (Count-Pattern $utilC 'fn = "vb6_ComPackSingleRef"') 1 'Single → VT_R4'
Expect 'R3 DoubleRef mapping'   (Count-Pattern $utilC 'fn = "vb6_ComPackDoubleRef"') 1 'Double → VT_R8'
Expect 'R3 Integer asks knownIntVars_' (Count-Pattern $utilC 'knownIntVars_\.count\(lower\) \? Vb6Type::Integer') 1 'inferExprType 对 Integer 答 Long, 宽度必须问专属表'

# R4 —— 早绑定那一支只对 ByRef 形参取址
Expect 'R4 positional ByRef gate' (Count-Pattern $bindInc 'i < sig\.params\.size\(\) && !sig\.params\[i\]\.isByVal') 1 'ByVal 不许取址'
Expect 'R4 named ByRef gate def'  (Count-Pattern $bindInc 'byRefParamNamedAny = ') 1 '命名实参那一道闸'
Expect 'R4 named ByRef gate call' (Count-Pattern $bindInc 'byRefParamNamedAny\(') 1 '命名实参只在一处问那道闸'

# R5 —— 判据面
$h = [System.IO.File]::ReadAllText($harness)
foreach ($n in @('vb6_ComPackIntRef(&n)', 'vb6_ComPackIntegerRef(&i)', 'vb6_ComPackDoubleRef(&d)',
                 'vb6_ComPackSingleRef(&s)', 'vb6_ComPackBoolRef(&b)')) {
    Expect ("R5 needle " + $n) ([regex]::Matches($h, [regex]::Escape($n))).Count 1 '五档 Ref 各一枚正针'
}
Expect 'R5 by-value witnesses' ([regex]::Matches($h, 'vb6_ComPackInt\((7|\(n \+ 1\))\)')).Count 2 '字面量/表达式仍按值'
Expect 'R5 absent lesion' ([regex]::Matches($h, 'vb6_ComPackInt\(&n\)')).Count 1 '改前那一形只许留在 Absent'

if ($fails.Count -gt 0) {
    Write-Host ('RED: ' + $fails.Count + ' guard(s) broken') -ForegroundColor Red
    exit 1
}
Write-Host 'PASS: 晚绑定实参按引用只有一位主人 (五档, 只认 lvalue)' -ForegroundColor Green
exit 0
