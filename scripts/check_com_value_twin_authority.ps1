# check_com_value_twin_authority.ps1 - 「值语境吃到交回对象的 COM 出口」只有一位主人
#   (账 #305 第二百四十九刀)
#
# 起因 (读数在第二百四十八刀那一行): 「这发交回的是堆上一枚 VARIANT 还是对象指针」这一问
#   今天有好几份手抄名单，其中**长度对**的那一族(七处)会真的生效，而且生效得错 ——
#   它们把 vb6_ComGetObjectProp( 当成「裸晚绑定调用的结果」包进 vb6_VariantFromComResult。
#   两头都错: 那枚出口交回的是 IDispatch 指针 (VARIANT 结构体已在它内部释放)，而它对
#   **非对象档**压根交回 NULL —— VariantChangeType(VT_I4 -> VT_DISPATCH) 实测
#   0x80020005 (.local/k249_probe/vtprobe.c)。真产物证人 2 处
#   (.local/k247_probe/vbp.emit.txt :80950 / :84099) 都是 `IFont.hFont` 那一类句柄档属性读
#   => 症状 = 句柄恒 0、字体度量整条落空，不是崩溃。
#
# 规则（改坏了会红）：
#   R1  主人 comValueTwinForObjectReadExpr: 定义 1 (cgen_util_com.cpp) + 声明 1
#       (cgen_helpers.inc)；表里恰好**三对**孪生，六个名字逐个在 vb6com.h 里有声明
#       (元数与返回类型同一，才许做这个纯文本换头)。
#   R2  消费者恰好两站 (cgen_expr_call_arg_variant.inc 的两条实参臂: 形参类型已知的那条
#       + 内置/运行时的那条)。别处要接同一问，先改这条并说清为什么。
#   R3  那两站不许再抄对象出口的前缀: 该文件里带引号的 vb6_ComGetObjectProp( 出现即红
#       (含注释 —— 第二百四十六刀立的口径)。
#   R4  两站各有各的「这一位要的是对象指针」闸: 形参那条问 VB 档 (paramBase != Object)，
#       内置那条问 C 档 (rtParamType != void*, 那张表里只有 vb6_UnloadForm 交回对象)。
#       两枚闸各恰好一次 —— 撤掉任何一枚都会把真对象指针换成堆 VARIANT。
#   R5  同问的另一半 (comObjectRefFromCallExpr = 「结果要当对象用」) 还在, 且两半同住
#       cgen_util_com.cpp (一个问题的两个方向拆到两个文件 = 迟早不同答)。
#   R6  判据面: runner 里 b249_value_twin_lesion 登记 1 处 + 两枚反面针各在场；
#       且**没有** b249 的 Test-EmitcShape (整工程产物里带 vb6_BSTR_FromStr 的奇引号行
#       改前改后各 28, 那条守护只对单份文件成立 —— 别把它接回整工程)。
#
# 用法: powershell -File scripts\check_com_value_twin_authority.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
if ($env:C3_CHECK_ROOT) { $root = $env:C3_CHECK_ROOT }

$comC     = Join-Path $root 'src\backend\cgen_util_com.cpp'
$helpers  = Join-Path $root 'src\backend\detail\util\cgen_helpers.inc'
$argVar   = Join-Path $root 'src\backend\detail\expr\cgen_expr_call_arg_variant.inc'
$rtlH     = Join-Path $root 'src\rtl\core\vb6com\vb6com.h'
$runner   = Join-Path $root 'tests\run_tests.ps1'

$script:red = 0
function Read-FileText($p) { return [System.IO.File]::ReadAllText($p) }
function Count-Pattern($text, $pattern) {
    return ([regex]::Matches($text, $pattern)).Count
}
function Expect($what, $got, $want) {
    if ($got -eq $want) { Write-Host ("  ok   " + $what + " = " + $want); return }
    Write-Host ("  RED  " + $what + " = " + $got + "  (期望 " + $want + ")") -ForegroundColor Red
    $script:red++
}
function Expect-Zero($what, $got) { Expect $what $got 0 }

$twinName = 'comValueTwinForObjectReadExpr'
$twinCall = [regex]::Escape($twinName + '(')
$objPrefixQuoted = '"' + [regex]::Escape('vb6_ComGetObjectProp(')
$lesionFont  = [regex]::Escape('vb6_VariantFromComResult(vb6_ComGetObjectProp(Font, L"hFont"))')
$lesionBtn   = [regex]::Escape('vb6_VariantFromComResult(vb6_ComGetObjectProp(ButtonFont, L"hFont"))')
$pairNames = @(
    'vb6_ComGetObjectProp(', 'vb6_ComGetProp(',
    'vb6_ComCallObject(', 'vb6_ComCall(',
    'vb6_ComGetPropertyObject(', 'vb6_ComGetPropertyVariant('
)

$com = Read-FileText $comC
$hel = Read-FileText $helpers
$av  = Read-FileText $argVar
$rtl = Read-FileText $rtlH
$run = Read-FileText $runner

Write-Host 'R1 主人本身与那张三对表'
Expect 'R1 定义处 (cgen_util_com.cpp)' (Count-Pattern $com ('std::string CCodeGen::' + $twinName)) 1
Expect 'R1 声明处 (cgen_helpers.inc)' (Count-Pattern $hel ('std::string ' + $twinName)) 1
$table = [regex]::Match($com, 'kTwins\[\]\[2\] = \{(?s).*?\};')
if (-not $table.Success) {
    Write-Host '  RED  R1 找不到 kTwins 那张表' -ForegroundColor Red
    $script:red++
} else {
    Expect 'R1 表里成对的字符串字面量' ([regex]::Matches($table.Value, '"' + [regex]::Escape('vb6_')).Count) 6
}
Expect 'R1 换头那一腿 (return kTwins[i][1])' (Count-Pattern $com 'return std::string\(kTwins\[i\]\[1\]\)') 1
foreach ($n in $pairNames) {
    Expect ('R1 出口名在 RTL 头里有声明: ' + $n.TrimEnd('(')) (Count-Pattern $rtl ([regex]::Escape($n))) 1
}

Write-Host 'R2 消费者计数 (两臂各一站)'
Expect 'R2 arg_variant 里调用主人的次数' (Count-Pattern $av $twinCall) 2
Expect 'R2 主人四处齐 (定义 + 声明 + 两站调用)' (Count-Pattern ($com + $hel + $av) $twinCall) 4

Write-Host 'R3 那两站不抄对象出口前缀'
Expect-Zero 'R3 arg_variant 里带引号的对象属性读前缀' (Count-Pattern $av $objPrefixQuoted)

Write-Host 'R4 两站那道「要的是对象指针」闸还在'
Expect 'R4 形参臂 paramBase != Vb6Type::Object' (Count-Pattern $av 'paramBase != Vb6Type::Object') 1
Expect 'R4 内置臂 rtParamType != void*' (Count-Pattern $av 'rtParamType != "void\*"') 1

Write-Host 'R5 同一个问题的另一半同住一处'
Expect 'R5 comObjectRefFromCallExpr 定义' (Count-Pattern $com 'std::string CCodeGen::comObjectRefFromCallExpr') 1
Expect 'R5 comObjectRefFromCallExpr 声明' (Count-Pattern $hel 'comObjectRefFromCallExpr') 1

Write-Host 'R6 判据面'
Expect 'R6 反面针登记' (Count-Pattern $run 'Test-EmitcAbsent "b249_value_twin_lesion"') 1
Expect 'R6 病灶针 Font' (Count-Pattern $run $lesionFont) 1
Expect 'R6 病灶针 ButtonFont' (Count-Pattern $run $lesionBtn) 1
Expect 'R6 没把形状针接到整工程' (Count-Pattern $run 'Test-EmitcShape "b249') 0

if ($script:red -gt 0) {
    Write-Host ('FAILED: ' + $script:red + ' 条红') -ForegroundColor Red
    exit 1
}
Write-Host 'ALL PASS'
exit 0
