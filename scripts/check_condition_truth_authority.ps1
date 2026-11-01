# check_condition_truth_authority.ps1 - 「把一条 C 表达式折成 VB6 的真值测试」只有一位主人
#   (账 #305 第二百四十四刀; 第二百四十七刀把 `Not <同一族操作数>` 也挂到同一问上)
#
# 起因 (读数在第二百四十三刀那一行): 晚绑定调用 vb6_ComCall(...) 交回的是**堆上一枚 VARIANT 的
#   指针** (C 侧 void*)。条件那一族只有一处包装站 (cgen_stmt.cpp 里补 vb6_VariantToBool 那段)，
#   而它问的是「这串 C 是不是槽」(cExprIsVariant) —— 指针不是槽 ⇒ 答否 ⇒ 一行都不包 ⇒
#   if (ptr) 恒真。真语料按产物数到 82 处 / 27 份文件，最响的一族是 `Do While Not rs.EOF`
#   (恒真 = 死循环)；`Loop While <假条件>` 实测跑满保险丝 (n=4 / VB6 该 1)。
#
# 规则（改坏了会红）：
#   R1  两枚 RTL 出口 vb6_ComResultToBool / vb6_ComResultToNot: 各「定义 1 + 声明 1」，
#       且各自体内三问各 1、**先读值、后收回** (顺序倒了 = 读一枚已清空的槽)。
#       第二百四十七刀起, 全文件里那三问各出现 2 次 = 两枚出口共用同一套读法, 谁也不许多一份。
#   R2  唯一主人 conditionTruthExpr: 定义 1 (cgen_util_type.cpp) + 声明 1 (cgen_helpers.inc)
#       + 站点 7 (If / ElseIf / Do While / Do Until / Loop While / Loop Until / While Wend)。
#   R3  「这发是不是裸晚绑定调用的结果」= cExprIsBareComResult: 定义 1 + 声明 1 +
#       只在两位主人那里问 —— 条件那一站 1 次 (cgen_util_type.cpp)、Not 那一站 1 次
#       (cgen_expr.cpp)，两处都问同一枚函数, 谁也不许自己再抄一遍前缀名单 (名单里出现带
#       引号的它 = 红)。名单恰好三条 (ComCall / ComCallByDispid / ComGetProp) —— 第二百四十六刀
#       撤掉属性读那一枚: 它交回的是**对象指针** (VARIANT 结构体已在出口内释放), 递进
#       ComResultToBool = 拿 IDispatch* 当 VARIANT* 解引用 + free 一个 COM 指针 (账 #255 那位
#       主人早就这么划)。也不许混进带类型的出口 (ComCallInt/BSTR/Double/Object、
#       ComGetStringProp) —— 那些交回的已是具体 C 档, 混进来就是二次解包。
#   R4  站点不许自己再问「是不是槽」: cgen_stmt.cpp 与 cgen_loop.cpp 里 cExprIsVariant(lastExpr_) 各 0。
#   R5  判据面: 夹具存在 + runner 里 7 枚条件正针 (带 ComResultToBool) + 5 枚 Not 正针
#       (带 ComResultToNot) + 5 枚 Not 反面针 (裸 ~(int32_t)(vb6_ComCall( ) + 18 枚键位
#       (条件正反 14 + Not 正反 4) + 真跑期望 COND:16/16 + 四组针各登记一处。
#
# 用法: powershell -File scripts\check_condition_truth_authority.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
if ($env:C3_CHECK_ROOT) { $root = $env:C3_CHECK_ROOT }

$rtlC  = Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_com.c'
$rtlH  = Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_class_com.h'
$typeC = Join-Path $root 'src\backend\cgen_util_type.cpp'
$helpers = Join-Path $root 'src\backend\detail\util\cgen_helpers.inc'
$ifC   = Join-Path $root 'src\backend\stmt\cgen_stmt.cpp'
$loopC = Join-Path $root 'src\backend\stmt\cgen_loop.cpp'
$exprC = Join-Path $root 'src\backend\expr\cgen_expr.cpp'
$harness = Join-Path $root 'tests\run_tests.ps1'
$fix     = Join-Path $root 'tests\test_late_cond.bas'

foreach ($p in @($rtlC, $rtlH, $typeC, $helpers, $ifC, $loopC, $exprC, $harness, $fix)) {
    if (-not (Test-Path -LiteralPath $p)) { Write-Host ('FAIL missing ' + $p) -ForegroundColor Red; exit 1 }
}

function Count-Pattern($path, $pattern) {
    $t = [System.IO.File]::ReadAllText($path)
    return ([regex]::Matches($t, $pattern)).Count
}

function Count-In($text, $pattern) { return ([regex]::Matches($text, $pattern)).Count }

# 取一枚 C 函数的**体内**文本 (从定义行到下一个行首右花括号)。
# 两枚出口 (ToBool / ToNot) 体内问的是同样三件事, 整文件计数分不开谁是谁 —— 只有切片能钉
# 「先读值、后收回」这一条属于哪一枚。行尾按 CRLF/LF 都能切, 故匹配「行首的 }」。
function Get-Body($path, $defPattern) {
    $t = [System.IO.File]::ReadAllText($path)
    $m = [regex]::Match($t, $defPattern)
    if (-not $m.Success) { return '' }
    $rest = $t.Substring($m.Index + $m.Length)
    $c = [regex]::Match($rest, '(?m)^\}')
    if (-not $c.Success) { return $rest }
    return $rest.Substring(0, $c.Index)
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

# R1 —— 两枚 RTL 出口 (真值位 + Not 位)
$defBool = '(?m)^int16_t vb6_ComResultToBool\(void\* variant_ptr\) \{'
$defNot  = '(?m)^int32_t vb6_ComResultToNot\(void\* variant_ptr\) \{'
Expect 'R1 vb6_ComResultToBool defined'   (Count-Pattern $rtlC $defBool) 1 '定义只许一枚, 在 vb6rtl_com.c'
Expect 'R1 vb6_ComResultToBool declared'  (Count-Pattern $rtlH 'int16_t vb6_ComResultToBool\(void\* variant_ptr\);') 1 '声明只许一枚, 在 vb6rtl_class_com.h'
Expect 'R1 vb6_ComResultToNot defined'    (Count-Pattern $rtlC $defNot) 1 '第二百四十七刀那枚出口: 定义只许一枚'
Expect 'R1 vb6_ComResultToNot declared'   (Count-Pattern $rtlH 'int32_t vb6_ComResultToNot\(void\* variant_ptr\);') 1 '声明只许一枚, 在 vb6rtl_class_com.h'
# 体内三问各 1 + 全文件各 2 (两枚出口共用同一套读法, 多一枚就是有人又抄了一遍)
Expect 'R1 取堆块 (全文)'   (Count-Pattern $rtlC 'vb6_VariantFromComResult\(variant_ptr\)') 2 '两枚出口各问一次'
Expect 'R1 收回槽 (全文)'   (Count-Pattern $rtlC 'vb6_VariantClear\(&v\)') 2 '两枚出口各收一次'
foreach ($pair in @(@('Bool', $defBool, 'vb6_VariantToBool'),
                    @('Not',  $defNot,  'vb6_VariantToLong'))) {
    $body = Get-Body $rtlC $pair[1]
    $read = $pair[2] + '(v)'
    Expect ('R1 ' + $pair[0] + ' 体内取堆块') (Count-In $body 'vb6_VariantFromComResult\(variant_ptr\)') 1 ('定义体里必须解引用那一枚堆块 —— 切片为空 = 那枚出口压根不存在')
    Expect ('R1 ' + $pair[0] + ' 体内读数')   (Count-In $body ($pair[2] + '\(v\)')) 1 '值口径问既有的那位主人'
    Expect ('R1 ' + $pair[0] + ' 体内收回')   (Count-In $body 'vb6_VariantClear\(&v\)') 1 '深拷出来的槽要收回 (BSTR/对象/数组那几档有载荷)'
    Expect ('R1 ' + $pair[0] + ' 先读后清') ([int]($body.IndexOf($read) -lt $body.IndexOf('vb6_VariantClear(&v)'))) 1 '顺序倒了 = 读一枚已清空的槽 (IndexOf 给 -1 也算红)'
}

# R2 —— 唯一主人与七站
Expect 'R2 conditionTruthExpr defined'   (Count-Pattern $typeC '(?m)^std::string CCodeGen::conditionTruthExpr\(') 1 '定义只许一枚, 在 cgen_util_type.cpp'
Expect 'R2 conditionTruthExpr declared'  (Count-Pattern $helpers 'std::string conditionTruthExpr\(') 1 '声明在 cgen_helpers.inc'
Expect 'R2 If 那一站'                    (Count-Pattern $ifC   'conditionTruthExpr\(') 2 'If 顶层 + ElseIf (ElseIf 从前连这一问都没问)'
Expect 'R2 循环那五站'                   (Count-Pattern $loopC 'conditionTruthExpr\(') 5 'Do While / Do Until / Loop While / Loop Until / While Wend'

# R3 —— 「裸晚绑定调用」那一问
Expect 'R3 cExprIsBareComResult defined'   (Count-Pattern $typeC '(?m)^bool CCodeGen::cExprIsBareComResult\(') 1 '定义只许一枚'
Expect 'R3 cExprIsBareComResult declared'  (Count-Pattern $helpers 'bool cExprIsBareComResult\(') 1 '声明在 cgen_helpers.inc'
Expect 'R3 asked only by the owner'        (Count-Pattern $typeC 'cExprIsBareComResult\(cExpr\)') 1 '只许主人在条件那一站问它'
Expect 'R3 Not 那一站也问同一位主人'        (Count-Pattern $exprC 'cExprIsBareComResult\(operand\)') 1 '第二百四十七刀: 值位与真值位同问一位主人, 不在表达式站再抄一份名单'
foreach ($q in @('vb6_ComCall(', 'vb6_ComCallByDispid(', 'vb6_ComGetProp(')) {
    Expect ('R3 表达式站不抄名单 ' + $q) (Count-Pattern $exprC ([regex]::Escape('"' + $q + '"'))) 0 '名单只许住在主人那里 —— 抄第二份就是第二百四十六刀撤掉的那类事故之源'
}
foreach ($n in @('vb6_ComCall\(', 'vb6_ComCallByDispid\(', 'vb6_ComGetProp\(')) {
    Expect ('R3 prefix ' + $n) (Count-Pattern $typeC ('"' + $n + '"')) 1 '名单里恰好一条'
}
foreach ($bad in @('vb6_ComCallInt', 'vb6_ComCallBSTR', 'vb6_ComCallDouble', 'vb6_ComCallObject', 'vb6_ComGetStringProp', 'vb6_ComGetObjectProp')) {
    Expect ('R3 forbidden exit ' + $bad) (Count-Pattern $typeC $bad) 0 '交回的已是具体 C 档或对象指针, 混进这一问就是二次解包 / 拿对象指针当 VARIANT*'
}

# R4 —— 站点不许自己再问「是不是槽」
Expect 'R4 If 站不再自问'   (Count-Pattern $ifC   'cExprIsVariant\(lastExpr_\)') 0 '那一问归主人'
Expect 'R4 循环站不再自问' (Count-Pattern $loopC 'cExprIsVariant\(lastExpr_\)') 0 '那一问归主人'

# R5 —— 判据面
$h = [System.IO.File]::ReadAllText($harness)
Expect 'R5 条件正针'      ([regex]::Matches($h, 'vb6_ComResultToBool')).Count 7 '七站各一枚正针'
Expect 'R5 Not 正针'      ([regex]::Matches($h, [regex]::Escape('vb6_ComResultToNot('))).Count 5 'Not 那一族五枚 (条件 2 + 值位 2 + Do While 1); 注释里写名字不带左括号, 不占这五枚'
Expect 'R5 Not 反面针'    ([regex]::Matches($h, [regex]::Escape('(~(int32_t)(vb6_ComCall('))).Count 5 '修复前那五行原样, 一枚不许少 (少 = 反面针没钉住同一发)'
Expect 'R5 keyed needles' ([regex]::Matches($h, 'L"K-')).Count 18 '条件正反 14 + Not 正反 4'
Expect 'R5 bas-test registered' ([regex]::Matches($h, 'Add-BasTest "test_late_cond"')).Count 1 '真跑夹具登记一处'
Expect 'R5 runtime needles' ([regex]::Matches($h, 'COND:16/16')).Count 1 '十六臂全对的总结针'
Expect 'R5 runner 期望同步' ([regex]::Matches($h, 'COND-1:COUNT ok=16/16')).Count 1 '计数针与总结针必须同升, 否则门永远跑不到后四臂'
Expect 'R5 Not 两枚针登记' (([regex]::Matches($h, 'Test-EmitcShape "b247_not_fold_shape"')).Count + [regex]::Matches($h, 'Test-EmitcAbsent "b247_not_fold_lesion"').Count) 2 '形状针与反面针各一处'

if ($fails.Count -gt 0) {
    Write-Host ('RED: ' + $fails.Count + ' guard(s) broken') -ForegroundColor Red
    exit 1
}
Write-Host 'PASS: 晚绑定结果折成 VB6 真值只有一位主人 (条件七站 + Not 一站同问, 两枚出口同读法)' -ForegroundColor Green
exit 0
