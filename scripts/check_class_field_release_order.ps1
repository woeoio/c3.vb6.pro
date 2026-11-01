# check_class_field_release_order.ps1 - 类字段出厂那一圈只许先问「是不是数组」再问元素档
#   (账 #305 §B220，第二百二十二刀)
#
# 改前 vb6_cls_<C>_Destroy 的字段圈把两问排反了：先 `varType == Vb6Type::String`，后
# `isDynamicArray || !dimensions.empty()`，而 resolveArrayElemType() 把数组性剥掉 —— 于是一枚
# `Dim mK() As String` 答出 String，走进标量那一支发 `vb6_BSTR_Free(me->mK)`，也就是把整枚
# SafeArray 载体交给 SysFreeString。实测（同一枚夹具、同一根探针、x64 与 x86 两台）都死在
# Release 那一趟，RC=0xC0000374 = STATUS_HEAP_CORRUPTION；靶工程 VbEclipse 的 List/PopupMenu
# 就是这两枚（读数在 ai/022 第二百二十一刀那一节）。
#
# 规则（改坏了会红）：
#   R1  标量那一支的发码点全仓只许一枚（cgen_com.cpp 里 vb6_BSTR_Free(me-> 的 emitLine）。
#       命中 0 = 标量 String 字段没人放手了（漏 + 崩的另一种）；命中 2 = 有人又开了一处。
#   R2  问序钉死：那一枚 BSTR_Free 的行号必须**大于**数组那一问的行号（数组先问）。
#       这一条就是本刀的全部形状；把它改回前面，靶工程当场打死客户端。
#   R3  数组那一支仍把载体交回主人：cgen_com.cpp 里 vb6_SA_Destroy(me-> 命中 = 1。
#   R4  注册腿不许顺手改：同一文件里 knownBstrVars_.insert 命中 = 1（那一腿服务元素赋值，
#       产物里 `vb6_BSTR_Assign(&VB6_SA_AT(BSTR, …), …)` 是对的，与本刀的出厂问序无关）。
#   R5  夹具三种字段形状都在场：动态 mK() / 一维定长 mF(2) / 标量 mS 各 1 枚 —— 缺任一枚，
#       那一条分支就跑不到（定长那一枚尤其便宜：载体随实例创建，create+Release 就坏）。
#   R6  判据已升格进门禁：run_tests.ps1 里 b221_strarr_field_release 命中 = 2（x64 + x86 各一格）。
#
# 用法:  powershell -File scripts\check_class_field_release_order.ps1
# 退出码: 0 = 全绿; 1 = 红
# C3_CHECK_ROOT 只为**负控**存在：把整棵树拷到别处、改一处、再指回来跑，真树一个字节不动。

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
if ($env:C3_CHECK_ROOT) { $root = $env:C3_CHECK_ROOT }

$comCpp  = Join-Path $root 'src\backend\module\cgen_com.cpp'
$fixCls  = Join-Path $root 'tests\b221_strarr\StrArrCls.cls'
$runner  = Join-Path $root 'tests\run_tests.ps1'

foreach ($p in @($comCpp, $fixCls, $runner)) {
    if (-not (Test-Path -LiteralPath $p)) { Write-Host ('FAIL missing ' + $p) -ForegroundColor Red; exit 1 }
}

$textCom = [System.IO.File]::ReadAllText($comCpp)
$textCls = [System.IO.File]::ReadAllText($fixCls)
$textRun = [System.IO.File]::ReadAllText($runner)

$fails = @()
function Expect($tag, $got, $want, $note) {
    if ($got -ne $want) {
        $script:fails += ('{0} = {1} (want {2}) -- {3}' -f $tag, $got, $want, $note)
        Write-Host ('FAIL {0} = {1}, want {2}  ({3})' -f $tag, $got, $want, $note) -ForegroundColor Red
    } else {
        Write-Host ('ok   {0} = {1}' -f $tag, $got) -ForegroundColor DarkGreen
    }
}

function Line-Of($text, $needle) {
    $i = $text.IndexOf($needle)
    if ($i -lt 0) { return -1 }
    return ([regex]::Matches($text.Substring(0, $i), "`n")).Count + 1
}

$emitFree  = 'vb6_BSTR_Free(me->" + cIdent(var.name) + ");'
$askArray  = 'if (var.isDynamicArray || !var.dimensions.empty()) {'
$emitSa    = 'vb6_SA_Destroy(me->" + cIdent(var.name) + ");'
$regBstr   = 'knownBstrVars_.insert(lower);'

# R1 / R3 / R4 —— 命中数
Expect 'R1_scalar_bstr_free_sites' ([regex]::Matches($textCom, [regex]::Escape($emitFree)).Count) 1 `
    '标量 String 字段那一支的发码点只许一枚（0 = 没人放手, 2 = 长出第二份）'
Expect 'R3_array_carrier_destroy_sites' ([regex]::Matches($textCom, [regex]::Escape($emitSa)).Count) 1 `
    '数组载体交回 vb6_SA_Destroy 那一发只许一枚'
Expect 'R4_knownBstrVars_sites' ([regex]::Matches($textCom, [regex]::Escape($regBstr)).Count) 1 `
    '注册腿（服务元素赋值）不许被本刀顺手改'

# R2 —— 问序（两条都在场，且标量那一支排在数组那一问之后）
$lnArray = Line-Of $textCom $askArray
$lnFree  = Line-Of $textCom $emitFree
if ($lnArray -lt 0 -or $lnFree -lt 0) {
    $fails += ('R2_question_order = missing (array line {0}, scalar line {1})' -f $lnArray, $lnFree)
    Write-Host ('FAIL R2_question_order: array line {0}, scalar line {1}' -f $lnArray, $lnFree) -ForegroundColor Red
} elseif ($lnFree -gt $lnArray) {
    Write-Host ('ok   R2_question_order = scalar arm at line {0} after array arm at line {1}' -f $lnFree, $lnArray) -ForegroundColor DarkGreen
} else {
    $fails += ('R2_question_order = inverted (scalar {0} before array {1})' -f $lnFree, $lnArray)
    Write-Host ('FAIL R2_question_order: 标量那一支又排到数组那一问前面了 (scalar {0} < array {1})' -f $lnFree, $lnArray) -ForegroundColor Red
}

# R5 —— 夹具三种字段形状
Expect 'R5_fixture_dyn_field' ([regex]::Matches($textCls, 'Private mK\(\) As String').Count) 1 '动态 String 数组字段'
Expect 'R5_fixture_fixed_field' ([regex]::Matches($textCls, 'Private mF\(2\) As String').Count) 1 '一维定长 String 数组字段'
Expect 'R5_fixture_scalar_field' ([regex]::Matches($textCls, 'Private mS As String').Count) 1 '标量 String 字段（反面证人）'

# R6 —— 判据在门禁里
Expect 'R6_runner_cases' ([regex]::Matches($textRun, 'b221_strarr_field_release').Count) 2 'x64 与 x86 各一格'

if ($fails.Count -gt 0) {
    Write-Host ('CLASS_FIELD_RELEASE_ORDER RED = ' + $fails.Count) -ForegroundColor Red
    $fails | ForEach-Object { Write-Host ('  ' + $_) -ForegroundColor Red }
    exit 1
}
Write-Host 'CLASS_FIELD_RELEASE_ORDER GREEN' -ForegroundColor Green
exit 0
