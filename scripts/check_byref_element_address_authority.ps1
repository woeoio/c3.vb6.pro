# check_byref_element_address_authority.ps1 - 「ByRef 形参吃数组元素」只有一位主人
#   (账 #305 第二百五十三刀, §B164 借来侧那一半)
#
# 起因 (读数在第二百五十二刀那一行): 五份整工程产物里「BSTR 复合字面量再取址」那一形共 375 处,
#   其中 **48 处**内层是 VB6_SA_AT / VB6_SA_ND_AT2 = 数组元素**自己**那枚指针。拷进复合字面量递给
#   被调方的 BSTR 形参后, 被调方发的是 vb6_BSTR_Assign(&(*p), …) ⇒ 那句「退旧」退掉的就是元素
#   自己那一枚串: 写回落进没人看的临时量 (VB6 该填进 arr(i)) + 元素原地悬空 (下次读它
#   use-after-free、下次写它双放) = 堆损坏同形。
#   真跑证人 = tests/test_byref_elem.bas 两台: 修后 BRW:6/6; 把主人短路成恒答否则 BRW:FAIL 2
#   (W1=OLD、W3=10、W5=bad 2000 —— W5 那 2000 次就是「同一格放两次」的反面证人)。
#
# 规则（改坏了会红）：
#   R1  主人 byRefElementAddressExpr: 定义 1 (cgen_util_type.cpp) + 声明 1 (cgen_helpers.inc)；
#       那张元素宏名单恰好 4 枚 (1 维 + 2/3/4 维)；两道闸各一次 —— 右括号必须正好收在整枚文本
#       末尾 (否则取址的是半枚表达式)、元素档必须**逐字等于**形参那一档 (否则交出去是错型指针)；
#       答「是」那一路交出的形状 = 取址 + 原样文本 (不许偷偷改写内层)。
#   R2  消费者恰好两站 (位置实参那一条臂 + 命名实参那一条臂)，全 src 里这个名字只许出现四处
#       (定义 + 声明 + 两站调用) 且只住在四份文件里。第三站要接同一问，先改这条并说清为什么。
#   R3  两站「落副本」那一臂只能挂在主人答否之后：各恰好一次 addr253.empty() 那道闸，
#       且调用主人的那一行必须**早于**拼复合字面量的那一行；两条臂的副本形状各只有一份
#       (位置站那一份带外层括号 = 预处理器不吃宏实参里的逗号，命名站那一份不带)。
#   R4  两站不许把类型名写死：那两份文件里出现 (&(BSTR) 即红 —— 形参那一档只许来自 cType。
#   R5  §B164 还在册的其余站点按**计数**钉住 (常量实参站 1 处 + COM 属性写站 2 处)：给它们
#       接同一位主人时这一条会红，是提醒不是惩罚。
#   R6  判据面: runner 里真跑夹具两台各登记 1 + 三枚正针各在场 + 两枚反面针各在场 (打在真工程上)；
#       夹具文件本身在盘上。
#
# 用法: powershell -File scripts\check_byref_element_address_authority.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
if ($env:C3_CHECK_ROOT) { $root = $env:C3_CHECK_ROOT }

$owner    = Join-Path $root 'src\backend\cgen_util_type.cpp'
$helpers  = Join-Path $root 'src\backend\detail\util\cgen_helpers.inc'
$emit     = Join-Path $root 'src\backend\detail\expr\cgen_expr_call_arg_emit.inc'
$named    = Join-Path $root 'src\backend\detail\expr\cgen_expr_call_named_args.inc'
$naming   = Join-Path $root 'src\backend\cgen_base_naming.cpp'
$comwrite = Join-Path $root 'src\backend\cgen_util_comwrite.cpp'
$runner   = Join-Path $root 'tests\run_tests.ps1'
$fixture  = Join-Path $root 'tests\test_byref_elem.bas'

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

# 先备好原始文本与它的正则形: 数次数用正则形, 比先后顺序用原始文本 (IndexOf 不认转义)。
$ownName      = 'byRefElementAddressExpr'
$callRaw      = $ownName + '(argVal, cType)'
$callTxt      = [regex]::Escape($callRaw)
$emptyGate    = [regex]::Escape('addr253.empty()')
$copyEmitRaw  = 'argVal = "(&(" + cType + "){" + argVal + "})";'
$copyNamedRaw = 'argVal = "&(" + cType + "){" + argVal + "}";'
$bstrHard     = [regex]::Escape('(&(BSTR)')
$rankGate     = [regex]::Escape('if (elem != formalCType) return std::string();')
$tailGate     = [regex]::Escape('if (close + 1 != argText.size()) return std::string();')
$answer       = [regex]::Escape('return "&(" + argText + ")";')
$lesionAt     = [regex]::Escape('vb6_VBFlexGrid_GetCellText((void*)me, iRow, iCol, (&(BSTR){VB6_SA_AT(BSTR, StrArr, OneDimIdx)}));')
$lesionNd     = [regex]::Escape('vb6_VBFlexGrid_GetCellText((void*)me, iRow, iCol, (&(BSTR){VB6_SA_ND_AT2(BSTR, (vb6_SafeArrayND*)StrArr, iColArr, iRowArr)}));')
$shape1d      = [regex]::Escape('vb6_FillStr(&(VB6_SA_AT(BSTR, arrS, 1)));')
$shape2d      = [regex]::Escape('vb6_FillStr(&(VB6_SA_ND_AT2(BSTR, (vb6_SafeArrayND*)arrM, 0, 1)));')
$shapeNum     = [regex]::Escape('vb6_FillNum(&(VB6_SA_AT(int32_t, arrN, 1)));')
$keepNaming   = [regex]::Escape('"(&(BSTR){" + argVal + "})";')
$keepComwrite = [regex]::Escape('"(&(BSTR){("')

$own = Read-FileText $owner
$hel = Read-FileText $helpers
$em  = Read-FileText $emit
$na  = Read-FileText $named
$bn  = Read-FileText $naming
$cw  = Read-FileText $comwrite
$run = Read-FileText $runner

Write-Host 'R1 主人本身 (名单 + 两道闸 + 答是的形状)'
Expect 'R1 定义处 (cgen_util_type.cpp)' (Count-Pattern $own ('std::string CCodeGen::' + $ownName)) 1
Expect 'R1 声明处 (cgen_helpers.inc)' (Count-Pattern $hel ('std::string ' + $ownName)) 1
$table = [regex]::Match($own, 'kElemMacros\[\] = \{(?s).*?\};')
if (-not $table.Success) {
    Write-Host '  RED  R1 找不到那张元素宏名单' -ForegroundColor Red
    $script:red++
} else {
    Expect 'R1 名单里的元素宏' ([regex]::Matches($table.Value, '"VB6_SA_').Count) 4
}
Expect 'R1 那道「右括号收在整枚文本末尾」闸' (Count-Pattern $own $tailGate) 1
Expect 'R1 那道「元素档逐字等于形参档」闸' (Count-Pattern $own $rankGate) 1
Expect 'R1 答是时交出的形状 (取址 + 原样文本)' (Count-Pattern $own $answer) 1

Write-Host 'R2 消费者计数 (两臂各一站, 全 src 四处齐)'
Expect 'R2 位置实参那一站' (Count-Pattern $em $callTxt) 1
Expect 'R2 命名实参那一站' (Count-Pattern $na $callTxt) 1
Expect 'R2 主人四处齐 (定义 + 声明 + 两站调用)' (Count-Pattern ($own + $hel + $em + $na) $ownName) 4
$srcRefs = @(Get-ChildItem -Path (Join-Path $root 'src') -Recurse -File -Include *.cpp, *.hpp, *.inc |
             Where-Object { (Read-FileText $_.FullName).Contains($ownName) })
Expect 'R2 认这个名字的文件只许四份' $srcRefs.Count 4

Write-Host 'R3 副本那一臂挂在主人答否之后'
Expect 'R3 位置站那道答否闸' (Count-Pattern $em $emptyGate) 1
Expect 'R3 命名站那道答否闸' (Count-Pattern $na $emptyGate) 1
Expect 'R3 位置站副本形状只有一份' (Count-Pattern $em ([regex]::Escape($copyEmitRaw))) 1
Expect 'R3 命名站副本形状只有一份' (Count-Pattern $na ([regex]::Escape($copyNamedRaw))) 1
$iCallE = $em.IndexOf($callRaw)
$iCopyE = $em.IndexOf($copyEmitRaw)
if ($iCallE -lt 0 -or $iCopyE -lt 0 -or $iCallE -gt $iCopyE) {
    Write-Host '  RED  R3 位置站问主人排在拼副本之后 (答否闸形同虚设)' -ForegroundColor Red
    $script:red++
} else {
    Write-Host '  ok   R3 位置站先问主人, 答不中才落副本'
}
$iCallN = $na.IndexOf($callRaw)
$iCopyN = $na.IndexOf($copyNamedRaw)
if ($iCallN -lt 0 -or $iCopyN -lt 0 -or $iCallN -gt $iCopyN) {
    Write-Host '  RED  R3 命名站问主人排在拼副本之后 (答否闸形同虚设)' -ForegroundColor Red
    $script:red++
} else {
    Write-Host '  ok   R3 命名站先问主人, 答不中才落副本'
}

Write-Host 'R4 那两站不写死类型名'
Expect-Zero 'R4 位置站里硬写的 BSTR 复合字面量' (Count-Pattern $em $bstrHard)
Expect-Zero 'R4 命名站里硬写的 BSTR 复合字面量' (Count-Pattern $na $bstrHard)

Write-Host 'R5 §B164 还在册的其余站点 (给它们接线那轮这条会红 = 提醒)'
Expect 'R5 常量实参那一站 (cgen_base_naming.cpp)' (Count-Pattern $bn $keepNaming) 1
Expect 'R5 COM 属性写那两站 (cgen_util_comwrite.cpp)' (Count-Pattern $cw $keepComwrite) 2

Write-Host 'R6 判据面'
Expect 'R6 真跑夹具登记 (默认台)' (Count-Pattern $run 'Add-BasTest "test_byref_elem" ') 1
Expect 'R6 真跑夹具登记 (x86 台)' (Count-Pattern $run 'Add-BasTest "test_byref_elem_x86"') 1
Expect 'R6 正针登记 (单文件夹具)' (Count-Pattern $run 'Test-EmitcShape "b253_elem_ref_shape"') 1
Expect 'R6 反面针登记 (真工程)' (Count-Pattern $run 'Test-EmitcAbsent "b253_elem_ref_lesion"') 1
Expect 'R6 正针 一维 String 元素' (Count-Pattern $run $shape1d) 1
Expect 'R6 正针 二维 String 元素' (Count-Pattern $run $shape2d) 1
Expect 'R6 正针 数值档元素' (Count-Pattern $run $shapeNum) 1
Expect 'R6 反面针 一维' (Count-Pattern $run $lesionAt) 1
Expect 'R6 反面针 二维' (Count-Pattern $run $lesionNd) 1
if (-not (Test-Path -LiteralPath $fixture)) {
    Write-Host '  RED  R6 夹具文件不在盘上' -ForegroundColor Red
    $script:red++
} else {
    Write-Host '  ok   R6 夹具文件在盘上'
}

if ($script:red -gt 0) {
    Write-Host ('FAILED: ' + $script:red + ' 条红') -ForegroundColor Red
    exit 1
}
Write-Host 'ALL PASS'
exit 0
