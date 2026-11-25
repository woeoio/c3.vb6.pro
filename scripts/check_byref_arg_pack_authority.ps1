# check_byref_arg_pack_authority.ps1 - 晚绑定实参「按引用发」只有一位主人
#   (账 #305 §B224 第二半，第二百三十九刀)
#
# 起因：真 VB6 客户端实测 (ai/probes/bool_string B4: `Dim o As Object` + `o.Bump(x)`，x As Long
# ⇒ x=11) —— 晚绑定对**带型变量**递的是 VT_BYREF|<该型>，服务端的写回落在调用方变量上；
# VBScript 的 Variant 实参递 VT_BYREF|VT_VARIANT 被 VB6 服务端判类型不匹配 ⇒ 差别在实参带不带型。
# 改前我们的客户端对**标量一律按值打包**，对外部服务器的 ByRef 写回因此落在虚空。
#
# 规则（改坏了会红）：
#   R1  六枚 Ref 打包器 (Int/Integer/Bool/Single/Double/BSTR) 在 vb6com_pack.c 各定义 1 次、
#       在 vb6com.h 各声明 1 次 —— 按 C 存储宽度分标签，Integer/Boolean 都是 2 字节。
#       (BSTR 那一枚是 §11.143 第一百六十一刀补的: 字符串实参改前一律按值，写回落在虚空。)
#   R2  「这枚实参该不该按引用」的唯一主人 comByRefArgExpr：定义 1 (cgen_util_com.cpp) +
#       声明 1 (cgen_helpers.inc) + 调用 2 (com_bind 的早绑定/晚绑定两支)。
#   R3  类型 → 打包器那张 switch 六档各 1 行；Integer 必须问 knownIntVars_ 而不是
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
$stateInc = Join-Path $root 'src\backend\detail\util\cgen_state.inc'
$callCpp = Join-Path $root 'src\backend\stmt\cgen_call.cpp'
$argEmit = Join-Path $root 'src\backend\detail\expr\cgen_expr_call_arg_emit.inc'
$harness = Join-Path $root 'tests\run_tests.ps1'
$fixBas  = Join-Path $root 'tests\b239_byrefarg\B239k.bas'
$fixVbp  = Join-Path $root 'tests\b239_byrefarg\B239k.vbp'

foreach ($p in @($packC, $comHdr, $utilC, $helpers, $bindInc, $harness, $fixBas, $fixVbp,
                $stateInc, $callCpp, $argEmit)) {
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

# R1 —— 九枚 Ref 打包器：定义各 1、声明各 1
foreach ($n in @('Int', 'Integer', 'Bool', 'Single', 'Double', 'BSTR', 'Byte', 'Cur', 'Date')) {
    Expect ("R1 vb6_ComPack${n}Ref defined")  (Count-Pattern $packC  ('(?m)^void\* vb6_ComPack' + $n + 'Ref\(')) 1 '定义只许一枚, 在 vb6com_pack.c'
    Expect ("R1 vb6_ComPack${n}Ref declared") (Count-Pattern $comHdr ('void\* vb6_ComPack' + $n + 'Ref\(')) 1 '声明只许一枚, 在 vb6com.h'
}

# R2 —— 唯一主人
Expect 'R2 comByRefArgExpr defined'  (Count-Pattern $utilC   '(?m)^std::string CCodeGen::comByRefArgExpr') 1 '定义只许一枚, 在 cgen_util_com.cpp'
Expect 'R2 comByRefArgExpr declared' (Count-Pattern $helpers 'std::string comByRefArgExpr\(Expr&') 1 '声明在 cgen_helpers.inc'
Expect 'R2 comByRefArgExpr called'   (Count-Pattern $bindInc 'comByRefArgExpr\(') 3 '三处消费者: 早绑定位置/命名 + 晚绑定'

# R3 —— 五档映射 + Integer 的宽度问句
Expect 'R3 IntRef mapping'      (Count-Pattern $utilC '"vb6_ComPackIntRef" \}.*// int32_t\*') 1 'Long → VT_I4'
Expect 'R3 IntegerRef mapping'  (Count-Pattern $utilC '"vb6_ComPackIntegerRef" \}.*// int16_t\*') 1 'Integer → VT_I2'
Expect 'R3 BoolRef mapping'     (Count-Pattern $utilC '"vb6_ComPackBoolRef" \}.*// int16_t\* \(VT_BOOL\)') 1 'Boolean → VT_BOOL'
Expect 'R3 SingleRef mapping'   (Count-Pattern $utilC '"vb6_ComPackSingleRef" \}.*// float\*') 1 'Single → VT_R4'
Expect 'R3 DoubleRef mapping'   (Count-Pattern $utilC '"vb6_ComPackDoubleRef" \}.*// double\*') 1 'Double → VT_R8'
Expect 'R3 BSTRRef mapping'     (Count-Pattern $utilC '"vb6_ComPackBSTRRef" \}.*// BSTR\*') 1 'String → VT_BYREF|VT_BSTR (§11.143)'
Expect 'R3 Integer asks knownIntVars_' (Count-Pattern $utilC 'knownIntVars_\.count\(lower\) \? Vb6Type::Integer') 1 'inferExprType 对 Integer 答 Long, 宽度必须问专属表'
# §11.146 (第一百六十四刀): 第七/八/九档 —— Byte / Currency / Date。
Expect 'R3 ByteRef mapping'     (Count-Pattern $utilC '"vb6_ComPackByteRef" \}.*// uint8_t\* \(VT_UI1\)') 1 'Byte → VT_UI1'
Expect 'R3 CurRef mapping'      (Count-Pattern $utilC '"vb6_ComPackCurRef" \}.*// int64_t\* \(VT_CY\)') 1 'Currency → VT_CY'
Expect 'R3 DateRef mapping'     (Count-Pattern $utilC '"vb6_ComPackDateRef" \}.*// double\*  \(VT_DATE\)') 1 'Date → VT_DATE'
# 货币**数组元素**仍按值: 标量 Currency 的 C 档是 double(值语义), 元素档是 int64_t, 两边不一个表示。
Expect 'R3 currency array stays by value' (Count-Pattern $utilC 'if \(et == Vb6Type::Currency\) return "";') 1 '在册欠项: 货币数组元素的存法'
Expect 'R3 nine ref tiers'      (Count-Pattern $utilC '\{ Vb6Type::\w+,\s+"vb6_ComPack\w+Ref" \}') 9 '表里正好九档'

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
# §11.143: 字符串那一档的判据面是 vbp 夹具两句 (b162_byref_string_arg / _x86)，不入这一张内联表。
Expect 'R5 bstr fixture cases' ([regex]::Matches($h, "Test-Vbp 'b162_byref_string_arg")).Count 2 '两格: x64 + x86'
# §11.144 (第一百六十二刀): 引用形不只裸标识符 —— 数组元素与成员访问两形的判据面。
Expect 'R7 array-elem shape'   (Count-Pattern $utilC 'static const std::regex arrElemRe\(') 1 '数组元素那一形只此一处判'
Expect 'R7 field shape asks decl' (Count-Pattern $utilC 'classFieldKind\(cls, ma\.memberName\)') 1 '类字段那一形问声明表, 不读尾注'
Expect 'R7 with shape asks decl'  (Count-Pattern $utilC 'classFieldKind\(wi\.className, wm\)') 1 'With 成员同一问'
Expect 'R7 addr fixture cases' ([regex]::Matches($h, "Test-Vbp 'b163_byref_addr")).Count 2 '两格: x64 + x86'
# §11.145 (第一百六十三刀): 按引用打包器那张表只许住一处, 三个消费点 (裸标识符/数组元素/类字段) 都问它。
Expect 'R8 ref-tier table'      (Count-Pattern $utilC 'kComPackRefTiers\[\]') 1 '表只住 cgen_util_com.cpp 一处'
Expect 'R8 ref-tier askers'     (Count-Pattern $utilC 'comByRefPackFnForVbType\(') 6 '定义 1 + 五处消费点 (裸标识符/数组元素/类字段/属性读那两问)'
Expect 'R8 numeric fixture cases' ([regex]::Matches($h, "Test-Vbp 'b164_byref_numeric_addr")).Count 2 '两格: x64 + x86'
# §11.147 (第一百六十五刀): 属性读当 ByRef 实参 —— 读-改-写。
# 真 VB6 不递属性本身 (属性没地址)，而是「拷临时量 → 递地址 → 回来经 Let 写回」。
# 登记只许住一处: 本语句的两行 (调用前 decl / 调用后 setter)，同收尾于 CallArgScope102。
Expect 'R9 prop-writeback ledger'   (Count-Pattern $stateInc 'std::vector<ComPropWriteBack> comPropWrites_;') 1 '台账只住 cgen_state.inc'
Expect 'R9 decl emitted before call' (Count-Pattern $callCpp 'for \(auto& pw147 : comPropWrites_\) c_\.emitLine\(pw147\.decl\);') 1 '读那一半: 调用前起手'
Expect 'R9 setter emitted in guard'  (Count-Pattern $callCpp 'for \(auto& pw147w : comPropWrites_\) c_\.emitLine\(pw147w\.setter\);') 1 '写回那一半: 调用后 (且在 102 守卫里)'
Expect 'R9 ledger cleared'           (Count-Pattern $callCpp 'g->comPropWrites_\.clear\(\);') 1 'CallArgScope102 统一收尾'
Expect 'R9 asks PropertyGet'         (Count-Pattern $utilC 'SymbolKind::PropertyGet') 1 '取型问声明侧'
Expect 'R9 asks PropertyLet'         (Count-Pattern $utilC 'SymbolKind::PropertyLet') 1 '只读属性写不回去'
Expect 'R9 currency prop by value' (Count-Pattern $utilC 'if \(propVt == Vb6Type::Currency\) return false;') 1 '在册: Currency 属性不开 (宽度)'
# §11.149 (第一百六十七刀): String 那一档开了，但必须多一行「退临时量」—— get 与 let
# 两头都是拷一份，所有权不接手。
Expect 'R9 string prop release'    (Count-Pattern $utilC 'vb6_BSTR_Free\(" \+ tmp') 1 'String 临时量归调用方退'
Expect 'R9 release field'         (Count-Pattern $stateInc 'std::string release;') 1 '台账多一格'
Expect 'R9 release emitted'       (Count-Pattern $callCpp 'if \(!pw149\.release\.empty\(\)\) c_\.emitLine\(pw149\.release\);') 1 '放在 102 守卫**外** (跳过调用也得退)'
Expect 'R9 prop fixture cases' ([regex]::Matches($h, "Test-Vbp 'b166_prop_byref")).Count 2 '两格: x64 + x86'
# §11.150 (第一百六十八刀): 带索引的属性 —— `b.Slot(2)` 在 AST 上是 IndexOrCallExpr
# (callee = MemberAccessExpr)。两处都得从 callee 取成员名：主人那一问与晚绑定那一族。
Expect 'R9 indexed prop recognized' (Count-Pattern $utilC 'ic\.callee->kind == ASTNodeKind::MemberAccessExpr') 2 '主人 + 晚绑定各一处'
Expect 'R9 index args split'        (Count-Pattern $utilC 'if \(ch == '','' && depth == 0\)') 1 '顶层逗号切 base / 索引实参'
Expect 'R9 index let order'        (Count-Pattern $utilC 'letFn \+ "\(" \+ baseOnly \+ idxArgs') 1 'Let 的取值形参在最后'
Expect 'R9 indexed prop fixture'   ([regex]::Matches($h, "Test-Vbp 'b168_indexed_prop_byref")).Count 2 '两格: x64 + x86'
# §11.151 (第一百六十九刀): ① With 块里的属性读 —— `.Tag` 是 WithMemberExpr，类名问
# withObjectInfoStack_；② 早绑定型不一致不得「假登记」—— 调用点用不上那枚临时量就不许台账里
# 多出两行 (wantCType 非空即调用点已定型参 C 型)。
Expect 'R9 with-prop asks stack'   (Count-Pattern $utilC 'wi\.kind == WithObjKind::ClassInstance\) cls = wi\.className;') 1 'With 主体的类名'
Expect 'R9 conv skips string'      (Count-Pattern $utilC 'if \(conv152 && propVt == Vb6Type::String\) return false;') 1 'BSTR 不换算'
Expect 'R9 want-ctype from caller' (Count-Pattern $argEmit 'pvt148, pvar148, cType\)') 1 '早绑定把形参 C 型递进去'
# §11.152 (第一百七十刀): 型不一致不再拒绝，改成「按形参 C 型造临时量 + 两向换算」。
# 三处: 判据里算出 conv152 / 声明用形参型 + 换入 / Let 换回。换算梯只住 narrowArgText。
Expect 'R9 narrow conv decided'   (Count-Pattern $utilC 'const bool conv152 = !wantCType\.empty\(\) && wantCType != propC;') 1 '型不一致 = 换算那一格'
Expect 'R9 narrow conv decl'      (Count-Pattern $utilC 'narrowArgText\(argText, wantCType, propC\)') 1 '换入'
Expect 'R9 narrow conv let'       (Count-Pattern $utilC 'narrowArgText\(tmp, propC, wantCType\)') 1 '换回'
Expect 'R9 narrow ladder'         (Count-Pattern $utilC 'vb6_ChkInt\(\(int32_t\)\(vb6_ChkLong\(vb6_FltToLng\(') 1 '浮点→Integer: 先取整再验界'
Expect 'R9 narrow fixture'        ([regex]::Matches($h, "Test-Vbp 'b170_prop_byref_narrow")).Count 2 '两格: x64 + x86'
Expect 'R9 with-prop fixture'      ([regex]::Matches($h, "Test-Vbp 'b169_with_prop_byref")).Count 2 '两格: x64 + x86'
# §11.148 (第一百六十六刀): 同一句问句给**两个**发射器用 —— 晚绑定实参那一族与早绑定
# 复合字面量那一族。定义 1 + 声明 1 + 调用 2；登记只许住台账一处，收尾只在语句尾巴。
Expect 'R9 prop-rw defined once'    (Count-Pattern $utilC '(?m)^bool CCodeGen::comPropReadWriteArg') 1 '定义只许一枚'
Expect 'R9 prop-rw declared'        (Count-Pattern $helpers 'bool comPropReadWriteArg\(Expr& expr') 1 '声明在 cgen_helpers.inc'
Expect 'R9 prop-rw two callers'     ((Count-Pattern $utilC 'comPropReadWriteArg\(') + (Count-Pattern $argEmit 'comPropReadWriteArg\(')) 3 '定义 1 + 两个消费点 (晚绑定/早绑定)'
Expect 'R9 early-bound site'        (Count-Pattern $argEmit 'pvt148, pvar148, cType\)\) \{') 1 '早绑定那一族把形参 C 型递进去 (型不一致也接 —— 主人内部换算)'
Expect 'R9 statement window open'   (Count-Pattern $callCpp 'comPropWriteWindow_ = true;') 1 '只在语句位置开窗'
Expect 'R9 statement window close'  (Count-Pattern $callCpp 'comPropWriteWindow_ = false;') 2 'emitExpr 之后 + CallArgScope102'
Expect 'R5 absent lesion' ([regex]::Matches($h, 'vb6_ComPackInt\(&n\)')).Count 1 '改前那一形只许留在 Absent'

if ($fails.Count -gt 0) {
    Write-Host ('RED: ' + $fails.Count + ' guard(s) broken') -ForegroundColor Red
    exit 1
}
Write-Host 'PASS: 晚绑定实参按引用只有一位主人 (五档, 只认 lvalue)' -ForegroundColor Green
exit 0
