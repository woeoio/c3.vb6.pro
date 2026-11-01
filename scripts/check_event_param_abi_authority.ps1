# check_event_param_abi_authority.ps1 - 「这一位事件形参按不按地址递」只有一位主人
#   (账 #305 第二百五十六刀)
#
# 起因: VB6 的 `Public Event GetEmptyMarkup(ByRef Text As String, ByRef Center As Boolean)`
#   里 ByRef 那一档是**出参** —— 处理器写进去的值要回到递出方。C3 从前在六处各自硬答
#   "按值发"，于是:
#     甲 UserControl 路 (ComCtlsDemo 的 ListView 空表提示文本): 四答彼此自洽 ⇒ 编得过、
#        跑不崩, 只是值回不来 ⇒ 那段提示永远不画 (第二百五十五刀「量」的读数)。
#     乙 模块级 Dim WithEvents 路 (tests/test_events 那一族): 处理器**自己**的 VB 签名是
#        ByRef (C 侧 BSTR*), 而包装器发成按值 (BSTR) ⇒ 生成的 C 是「把 wchar_t* 递给
#        BSTR* 槽」—— C 模式下 cl 只给警告, 运行期处理器多解一层引用 ⇒ 实测 0xC0000005
#        (BASE 编译器 2026-10-27 读数: 前五行照常打完, 走到 btn.RenderMarkup() 就 AV)。
#   本哨兵钉的是「六处同问一位主人」这个形状, 不是某一次的读数。
#
# 规则（改坏了会红）：
#   R1 唯一主人 src/common/event_param_abi.hpp: 两枚 inline 出口各恰好一枚。
#   R2 问「按不按地址」的只有三站 (语义层填表 1 + emitEventSink/RaiseEvent 2 + UC prelude 1);
#      模块级那两站包装器**不许自己再问** —— 它们读 Symbol::eventParamByAddr 那张表。
#   R3 写形参声明文本的四处 (typedef / thunk / 模块级包装声明 / 模块级包装实现) 各一枚,
#      且都走 eventParamCDecl —— 谁再手拼一遍 "wType + \" \" + name" 就是形状分家。
#   R4 那张表: 声明 1 处 + 语义层填 1 处 + 跨模块复制 1 处 (复制漏了 = 包装站永远读不到答案)。
#   R5 三处发形参表的站点里 `pi.isByVal = true;` 必须为 0, `pi.isByVal = !byAddr256;` 各 1;
#      模块级包装里 Fix 084k 那枚类形参临时地址只许挂在**按值那一臂**。
#   R6 判据面: 夹具里那一枚出参事件 + 处理器 + runner 的真跑期望 (x64 与 x86 各一处 = 2)
#      与两头发码针各登记一处。
#   R7 第三百八刀: 形参那一档声明的是 Variant 槽、而实参在 C 侧不是槽 —— 「要不要装箱」
#      那一问只许在发码站问一次; 装箱临时量 / 出站回写 / 清槽三件事各只住这一处。
#
# 用法: powershell -File scripts\check_event_param_abi_authority.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
if ($env:C3_CHECK_ROOT) { $root = $env:C3_CHECK_ROOT }

$hdr   = Join-Path $root 'src\common\event_param_abi.hpp'
$sem   = Join-Path $root 'src\semantics\semantic_analyzer.cpp'
$sym   = Join-Path $root 'src\semantics\symbol_table.hpp'
$cev   = Join-Path $root 'src\backend\module\cgen_com_events.cpp'
$pre   = Join-Path $root 'src\backend\detail\module\cgen_form_prelude.inc'
$bod   = Join-Path $root 'src\backend\detail\base\cgen_base_generate_body_pass.inc'
$dcl   = Join-Path $root 'src\backend\detail\base\cgen_base_generate_evt_decl.inc'
$xmd   = Join-Path $root 'src\driver\driver_crossmod.cpp'
$harn  = Join-Path $root 'tests\run_tests.ps1'
$fixC  = Join-Path $root 'tests\test_events\Button.cls'
$fixB  = Join-Path $root 'tests\test_events\test_events.bas'

foreach ($p in @($hdr, $sem, $sym, $cev, $pre, $bod, $dcl, $xmd, $harn, $fixC, $fixB)) {
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

# R1 —— 唯一主人
Expect 'R1 eventParamByAddress' (Count-Pattern $hdr 'inline bool eventParamByAddress\(') 1 '那一问的出口只许一枚'
Expect 'R1 eventParamCDecl'     (Count-Pattern $hdr 'inline std::string eventParamCDecl\(') 1 '形参声明文本只许一枚'

# R2 —— 谁问「按不按地址」
Expect 'R2 语义层填表'   (Count-Pattern $sem 'eventParamByAddress\(') 1 '整工程只算这一次'
Expect 'R2 COM 事件两站' (Count-Pattern $cev 'eventParamByAddress\(') 2 'emitEventSink 的 typedef + visit(RaiseEventStmt) 的实参位'
Expect 'R2 UC prelude'   (Count-Pattern $pre 'eventParamByAddress\(') 1 'thunk 那一站 (声明从 .ctl 文本里切)'
Expect 'R2 模块级实现不自行再问' (Count-Pattern $bod 'eventParamByAddress\(') 0 '读 Symbol::eventParamByAddr 那张表'
Expect 'R2 模块级声明不自行再问' (Count-Pattern $dcl 'eventParamByAddress\(') 0 '读同一张表 —— 声明与定义不同形就是 C2371/C2143'

# R3 —— 谁写形参声明文本
Expect 'R3 typedef'   (Count-Pattern $cev 'eventParamCDecl\(') 1 '回调 typedef 那一处'
Expect 'R3 thunk'     (Count-Pattern $pre 'eventParamCDecl\(') 1 '窗体 thunk 那一处'
Expect 'R3 模块级声明' (Count-Pattern $dcl 'eventParamCDecl\(') 1 '.h 前向声明那一处'
Expect 'R3 模块级实现' (Count-Pattern $bod 'eventParamCDecl\(') 1 '.c 定义那一处'

# R4 —— 那张表
Expect 'R4 表声明'   (Count-Pattern $sym 'std::vector<std::vector<char>> eventParamByAddr;') 1 '与 eventNames 同序'
Expect 'R4 表填写'   (Count-Pattern $sem 'classSym->eventParamByAddr\.push_back') 1 'EventDecl 那一趟顺手填'
Expect 'R4 跨模块复制' (Count-Pattern $xmd 'extSym->eventParamByAddr = srcSym->eventParamByAddr') 1 '漏了 = 别的模块里的包装器读不到答案'

# R5 —— 站点不许再硬答「按值」
foreach ($f in @(@('COM', $cev), @('模块级实现', $bod), @('模块级声明', $dcl))) {
    Expect ('R5 ' + $f[0] + ' 不再硬答按值') (Count-Pattern $f[1] 'pi\.isByVal\s*=\s*true;') 0 '那一问归主人'
    Expect ('R5 ' + $f[0] + ' 问的是主人')   (Count-Pattern $f[1] 'pi\.isByVal\s*=\s*!byAddr256;') 1 '按不按地址翻出来的那个答案'
}
Expect 'R5 类形参临时地址只挂按值臂' (Count-Pattern $bod 'evtParamIsClass\[i\] && !evtParamByAddr\[i\]') 1 '包装器这一位本身按地址时形状已合, 再套一层就是拿 BSTR** 初始化 BSTR*'

# R6 —— 判据面
Expect 'R6 夹具: 出参事件声明' (Count-Pattern $fixC 'Public Event GetEmptyMarkup\(ByRef Text As String, ByRef Center As Boolean\)') 1 '两位都是明写 ByRef'
Expect 'R6 夹具: 递出后读回'   (Count-Pattern $fixC 'RaiseEvent GetEmptyMarkup\(T, C\)') 1 '紧跟着读 T 与 C —— 值回不来就是空串'
Expect 'R6 夹具: 处理器写回'   (Count-Pattern $fixB 'Sub btn_GetEmptyMarkup\(ByRef Text As String, ByRef Center As Boolean\)') 1 '模块级 Dim WithEvents 那一路'
$h = [System.IO.File]::ReadAllText($harn)
# 第三百八刀的判据把同一枚夹具在**两台**各跑一遍 (x64 + x86 —— 20B 槽写穿 4B 存储那一档
# 只在 x86 现形, 而门只跑默认架构), 三条期望串因此各登记两处。钉成 2 而不是放宽成 >=1:
# 少一台红 (曾经就是这样漏的), 多一台也红 (同一枚夹具不许再抄第三遍)。
Expect 'R6 真跑期望(两台)' ([regex]::Matches($h, 'EvtRef test PASSED')).Count 2 'x64 + x86 各一处'
Expect 'R6 x64 那台在场' ([regex]::Matches($h, 'Test-Vbp "test_events" ')).Count 1 '默认架构那一发'
Expect 'R6 x86 那台在场' ([regex]::Matches($h, 'Test-Vbp "test_events_x86" ')).Count 1 '20B-into-4B 那一档只在这台现形'
Expect 'R6 发码正针'   ([regex]::Matches($h, 'Test-EmitcShape "b256_evt_byref_shape"')).Count 1 '三枚同形状针'
Expect 'R6 发码反面针' ([regex]::Matches($h, 'Test-EmitcAbsent "b256_evt_byref_lesion"')).Count 1 '修复前那两行原样, 一枚不许少'

# R7 —— 槽 vs 标量: 装箱那一问与它配套的三件事 (账 #305 第三百八刀)
Expect 'R7 槽那一问只在一处' (Count-Pattern $cev 'prmBase256 == "vb6_VARIANT"') 1 '形参声明档是 vb6_VARIANT 才谈装箱'
Expect 'R7 AST 那条槽判据'  (Count-Pattern $cev 'argTy308 != Vb6Type::Variant') 1 'inferExprType 答 Variant 就留在直递'
Expect 'R7 文本槽判据两处'  (Count-Pattern $cev 'cExprIsVariant\(argVal\)') 2 '新臂一处 + Fix 228 按值臂一处'
Expect 'R7 装箱临时量一枚'  (Count-Pattern $cev '_evt_box_') 1 '那枚槽的名字只在这里拼'
Expect 'R7 回写只住一处'    (Count-Pattern $cev 'evtVariantWriteBack') 2 '文件内定义 1 + 站点调用 1'
Expect 'R7 装与撤成对 A'    (Count-Pattern $cev 'vb6_VariantAssign\(&') 1 '深拷贝进槽 (BSTR 另分配)'
Expect 'R7 装与撤成对 B'    (Count-Pattern $cev 'vb6_VariantClear\(&') 1 '回写搬完就收槽'
if ($fails.Count -gt 0) {
    Write-Host ('RED: ' + $fails.Count + ' guard(s) broken') -ForegroundColor Red
    exit 1
}
Write-Host 'PASS: 事件形参按不按地址只有一位主人 (六站同问/同读那张表)' -ForegroundColor Green
exit 0
