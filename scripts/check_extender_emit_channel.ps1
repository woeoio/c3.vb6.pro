# check_extender_emit_channel.ps1
# 第三片 3a: `Extender.<成员> = v` 必须走**认得实例**的入口, 不能发裸全局赋值。
#
# 背景: 4d5e0280 补齐了 push/pop 存取, eff4de7f/76f91f1e 把通道
# (vb6_UC_ExtenderField) 改指 vb6_UCRec.ext*。于是"用户代码写 vb6_Extender_X = v"
# 与"通道读记录"接不上了 —— 用户的 Extender.Height/Width/Left/Top 会**静默不生效**。
# 3a 的修法是代码生成侧改发 vb6_UC_SetExtenderMember(...), 内部按当前记录落进 ext*。
#
# 本哨兵钉四件事, 每条带负控:
#   1. 代码生成侧闸门用**发射符号前缀**(不是 hostPseudoBareName —— 它不覆盖 Extender 成员)
#   2. 代码生成侧确实改道到实例入口
#   3. RTL 定义了入口, 且头文件声明了它(生成码直接调, 漏声明就 C2065)
#   4. 入口内部走 vb6_UC_ExtenderField(否则又变回只写全局)
#
# 4a 追加**读侧**的两条发射路径 + 共用出口的闸门, 各带负控:
#   5. Property Get 返回槽那一站(value_sem 的 `target = currentReturnVar_` 之后)改道
#   6. withHostPseudoIntrinsic 那一站(With 语境)改道
#   7. 共用出口 extenderMemberReadExpr 的闸门仍是发射符号前缀, 且两个读入口都被挑到

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$cg = Join-Path $root 'src/backend/detail/stmt/cgen_assign_host_pseudo.inc'
$val = Join-Path $root 'src/backend/detail/stmt/cgen_assign_value_sem.inc'
$argP = Join-Path $root 'src/backend/detail/expr/cgen_expr_call_arg_emit.inc'
$selP = Join-Path $root 'src/backend/stmt/cgen_select.cpp'
$evtP = Join-Path $root 'src/backend/module/cgen_com_events.cpp'
$setP = Join-Path $root 'src/backend/detail/stmt/cgen_setlet_set_rhs.inc'
$m22P = Join-Path $root 'src/backend/detail/expr/cgen_expr_member_m22_module.inc'
$binP = Join-Path $root 'src/backend/expr/cgen_expr_binary.cpp'
$withP = Join-Path $root 'src/backend/expr/cgen_expr_with.cpp'
$com = Join-Path $root 'src/backend/cgen_util_com.cpp'
$rtl = Join-Path $root 'src/rtl/core/vb6forms/uc/uc_controls.c'
$hdr = Join-Path $root 'src/rtl/core/vb6rtl/vb6rtl_userctl.h'

foreach ($p in @($cg, $val, $withP, $argP, $selP, $evtP, $setP, $m22P, $binP, $com, $rtl, $hdr)) {
    if (-not (Test-Path -LiteralPath $p)) { Write-Host "missing file: $p" -ForegroundColor Red; exit 1 }
}
$cgT = [System.IO.File]::ReadAllText($cg)
$valT = [System.IO.File]::ReadAllText($val)
$argT = [System.IO.File]::ReadAllText($argP)
$selT = [System.IO.File]::ReadAllText($selP)
$evtT = [System.IO.File]::ReadAllText($evtP)
$setT = [System.IO.File]::ReadAllText($setP)
$m22T = [System.IO.File]::ReadAllText($m22P)
$binT = [System.IO.File]::ReadAllText($binP)
$withT = [System.IO.File]::ReadAllText($withP)
$comT = [System.IO.File]::ReadAllText($com)
$rtlT = [System.IO.File]::ReadAllText($rtl)
$hdrT = [System.IO.File]::ReadAllText($hdr)

$bad = 0
function Check($name, $ok, $why) {
    if ($ok) { Write-Host "  [ok] $name" }
    else { Write-Host "  [FAIL] $name -- $why" -ForegroundColor Red; $script:bad++ }
}

$gateOk = $cgT -match 'rfind\(\s*"vb6_Extender_"\s*,\s*0\s*\)'
Check 'codegen gates on vb6_Extender_ prefix' $gateOk 'gate must match the emitted symbol prefix; hostPseudoBareName does not cover Extender members'

$emitOk = $cgT -match 'vb6_UC_SetExtenderMember'
Check 'codegen emits vb6_UC_SetExtenderMember' $emitOk 'Extender member writes must be redirected to the instance-aware entry point'

$defOk = $rtlT -match 'void\s+vb6_UC_SetExtenderMember\s*\('
Check 'RTL defines vb6_UC_SetExtenderMember' $defOk 'entry point missing from uc_controls.c'

$protOk = $hdrT -match 'void\s+vb6_UC_SetExtenderMember\s*\([^;]*\)\s*;'
Check 'header declares vb6_UC_SetExtenderMember' $protOk 'generated code calls it directly; without the prototype MSVC gives C2065'

$m = [regex]::Match($rtlT, 'void\s+vb6_UC_SetExtenderMember\s*\([^)]*\)\s*\{(?<b>.*?)\r?\n\}', 'Singleline')
Check 'entry point body present' $m.Success 'cannot locate function body'
if ($m.Success) {
    $b = $m.Groups['b'].Value
    # 只认**调用赋值**那一步, 不能拿函数体内的前向声明蒙混过关
    $callRx = '=\s*vb6_UC_ExtenderField\s*\('
    $viaChannel = $b -match $callRx
    Check 'entry point routes through vb6_UC_ExtenderField' $viaChannel 'must resolve via the channel so the write lands in vb6_UCRec.ext*'
    $neg = $rtlT -replace 'void\*\s+fp\s*=\s*vb6_UC_ExtenderField\([^;]*\);', 'void* fp = 0;'
    $mn = [regex]::Match($neg, 'void\s+vb6_UC_SetExtenderMember\s*\([^)]*\)\s*\{(?<b2>.*?)\r?\n\}', 'Singleline')
    $negFails = (-not $mn.Success) -or (-not ($mn.Groups['b2'].Value -match $callRx))
    Check 'negative control (deleting the channel call must go red)' $negFails 'sentinel is vacuous if neutering the call still passes'
}

# 读侧: 两个入口都要有, 且都必须走通道
foreach ($fn in @('vb6_UC_GetExtenderMemberNum', 'vb6_UC_GetExtenderMemberStr')) {
    $defR = $rtlT -match ("int32_t\s+" + $fn + "\s*\(|BSTR\s+" + $fn + "\s*\(")
    Check "RTL defines $fn" $defR 'read-side entry point missing from uc_controls.c'
    $protR = $hdrT -match ([regex]::Escape($fn) + '\s*\([^;]*\)\s*;')
    Check "header declares $fn" $protR 'generated code calls it directly; without the prototype MSVC gives C2065'
}
# 读侧刻意不跟写侧同形: 返回 void* 会逼所有调用点写强制转换
$voidRead = $rtlT -match 'void\*\s+vb6_UC_GetExtenderMember\s*\('
Check 'read side is NOT a single void* getter' (-not $voidRead) 'one void* getter would force a cast at every arithmetic call site'

# 负控: 两个读入口都必须真的调通道
foreach ($fn in @('vb6_UC_GetExtenderMemberNum', 'vb6_UC_GetExtenderMemberStr')) {
    $mr = [regex]::Match($rtlT, ([regex]::Escape($fn) + '\s*\([^)]*\)\s*\{(?<b>.*?)\r?\n\}'), 'Singleline')
    $negR = $rtlT -replace 'void\*\s+fp\s*=\s*vb6_UC_ExtenderField\([^;]*\);', 'void* fp = 0;'
    $mnr = [regex]::Match($negR, ([regex]::Escape($fn) + '\s*\([^)]*\)\s*\{(?<b>.*?)\r?\n\}'), 'Singleline')
    $ok = (-not $mnr.Success) -or (-not ($mnr.Groups['b'].Value -match '=\s*vb6_UC_ExtenderField\s*\('))
    Check "negative control: neutering the channel makes $fn red" $ok 'sentinel is vacuous if neutering the channel call still passes'
}

# --- 4a: 读侧 ---
# 3a 只把**写**拉进了通道。读还发裸全局 (`vb6_ret_Left = vb6_Extender_Left;`),
# 而权威值在 vb6_UCRec.ext* 里 —— getter 读到的是 push/pop 同步的副本, 不是实例记录。
# 4a 把两条读发射路径接到同一个实例感知出口 extenderMemberReadExpr 上。
$getterRx = 'extenderMemberReadExpr\s*\(\s*value\s*\)'
$getterOk = $valT -match $getterRx
Check 'Property Get read path redirects through the instance-aware entry' $getterOk 'vb6_ret_<Prop> = vb6_Extender_<M>; still reads a bare global (4a)'
$negG = $valT -replace '(?m)^[^\n]*extenderMemberReadExpr\s*\(\s*value\s*\)[^\n]*\n', ''
Check 'negative control (restoring the original getter line must go red)' (-not ($negG -match $getterRx)) 'sentinel is vacuous if deleting the redirect still passes'

$withRx = 'extenderMemberReadExpr\s*\(\s*hpRtlW\s*\)'
$withOk = $withT -match $withRx
Check 'withHostPseudoIntrinsic read path is covered' $withOk 'With <pseudo>.<member> whose RTL symbol is vb6_Extender_* must use the same entry point'
$negW = $withT -replace '(?m)^[^\n]*extenderMemberReadExpr\s*\(\s*hpRtlW\s*\)[^\n]*\n', ''
Check 'negative control (restoring the original With line must go red)' (-not ($negW -match $withRx)) 'sentinel is vacuous if deleting the With redirect still passes'

# 共用出口: 闸门必须是**发射符号前缀** (hostPseudoBareName 不覆盖 Extender 成员,
# 拿它当闸门会静默漏掉 —— 3a 踩过一次, 生成码 bare=9 new=0)。
$gateRx = 'extenderMemberReadExpr[\s\S]{0,3000}?rfind\(\s*"vb6_Extender_"\s*,\s*0\s*\)'
Check 'shared read exit gates on the emitted symbol prefix' ($comT -match $gateRx) 'gate must be rfind("vb6_Extender_", 0) on the emitted symbol'
$negC = $comT -replace 'rfind\(\s*"vb6_Extender_"\s*,\s*0\s*\)', 'rfind("vb6_ExtenderX_", 0)'
Check 'negative control (renaming the prefix gate must go red)' (-not ($negC -match $gateRx)) 'sentinel is vacuous if neutering the prefix gate still passes'

# 两个读入口都必须被共用出口挑到 (kind 约定 0=BSTR / 2=int16 / 3=int32 都走 Num)
function Has-Dock($text, $name) { return $text.Contains('"' + $name + '"') }
$numOk = Has-Dock $comT 'vb6_UC_GetExtenderMemberNum'
$strOk = Has-Dock $comT 'vb6_UC_GetExtenderMemberStr'
Check 'shared read exit offers both docks' ($numOk -and $strOk) 'read side needs a numeric and a string entry point (a single void* getter would force a cast per call site)'
$negD = $comT -replace '"vb6_UC_GetExtenderMemberStr"', '"vb6_UC_GetExtenderMemberStrX"'
$negNum = Has-Dock $negD 'vb6_UC_GetExtenderMemberNum'
$negStr = Has-Dock $negD 'vb6_UC_GetExtenderMemberStr'
Check 'negative control (unhooking the string dock must go red)' (-not ($negNum -and $negStr)) 'sentinel is vacuous if either dock can go unreferenced'

# --- 4b: 自由表达式读 ---
# 4a 只接了"目标 == 返回槽"那一问, 于是四种**目标不是返回槽**的形状还发着裸全局。
# ComCtlsDemo 实测 6 处, 分属**三个消费点** + 一个自带发码路径:
#   赋值右值    `Align = vb6_ChkInt(vb6_Extender_Align);`          CoolBar/StatusBar×2/ToolBar
#   Select 临时量 `int32_t _vb6_select_10 = vb6_Extender_Align;`   ToolBar
#   实参        <不适用>  (本语料无)
#   RaiseEvent 实参 `onHeightChanged(h, vb6_Extender_Height);`     CoolBar
# 每道闸用的都是同一根 extenderMemberReadExpr (按发射符号前缀判); 换道一律放在**消费点**,
# 不放在 hostPseudoRtlSymbol 里 —— 与 4a 同形 (符号只作真源, 不作发射点)。

# 判"窄条件没了"的两道讲究:
#  1) 先**剥注释** —— 断言的是代码, 注释里**提到**那串字(说明为什么曾经有)不算数。
#     两层 Replace 都要**各自加作用域**: (?s) 是全局开关, 漏进后半个分支会让 `.` 吃掉换行,
#     `//.*$` 贪到文件尾 (实测把第一个 // 之后的内容全删了)。
#  2) 只认**紧挨着这个调用**的那种 if —— 那个文件下面本来就有两处无关的
#     `target == currentReturnVar_` 判断, 全局搜那串字会误伤。
$valCode = [regex]::Replace([regex]::Replace($valT, '(?s:\/\*.*?\*\/)', ''), '(?m://[^\n]*)', '')
$getterWideRx = 'extenderMemberReadExpr\s*\(\s*value\s*\)'
$guardedRx = 'if\s*\([^;{]*target\s*==\s*currentReturnVar_[^;{]*\)\s*\{\s*value\s*=\s*extenderMemberReadExpr'
Check '4b: assignment RHS gate is no longer narrowed to the return slot' `
    (($valCode -match $getterWideRx) -and (-not ($valCode -match $guardedRx))) `
    'a free-standing `Align = Extender.Align` has a plain local target, so the 4a condition never fires'
$guardLine = '    value = extenderMemberReadExpr(value);'
$reGuarded = $valCode.Replace($guardLine, "    if (!currentReturnVar_.empty() && target == currentReturnVar_) {`r`n        value = extenderMemberReadExpr(value);`r`n    }")
Check 'negative control (re-narrowing the gate to the return slot must go red)' `
    ((-not ($valCode -match $guardedRx)) -and ($reGuarded -match $guardedRx)) `
    'sentinel is vacuous if re-narrowing the gate to the return slot still passes'

# (2) 实参: 换道在 boxIntoLateSlot / VariantToString / wrapToBSTR / boxToVariant **之前**,
#     那一层包装于是套在读入口外面 —— 与 4a "外层 vb6_ChkInt 照旧" 同款。
$argRx = 'extenderMemberReadExpr\s*\(\s*argVal\s*\)'
Check '4b: call-argument gate present' ($argT -match $argRx) `
    'a free-standing call argument reading Extender.<member> still passes a bare global'
$negB2 = $argT -replace '(?m)^[^\n]*extenderMemberReadExpr\s*\(\s*argVal\s*\)[^\n]*\r?\n', ''
Check 'negative control (restoring the original argument line must go red)' (-not ($negB2 -match $argRx)) `
    'sentinel is vacuous if deleting the argument redirect still passes'

# (3) Select Case 临时量: 换道落在**临时量的初始化**上, 不落 testVar 的每个用处 ——
#     下游 cond / switch 读的都是 tempVar, 所以 getter 只调一次, 不会每个 Case 调一遍。
$selRx = 'extenderMemberReadExpr\s*\(\s*lastExpr_\s*\)'
Check '4b: Select Case temp initializer gate present' ($selT -match $selRx) `
    'int32_t _vb6_select_N = vb6_Extender_Align; still reads a bare global'
$negB3 = $selT -replace '(?m)^[^\n]*extenderMemberReadExpr\s*\(\s*lastExpr_\s*\)[^\n]*\r?\n', ''
Check 'negative control (restoring the original Select initializer must go red)' (-not ($negB3 -match $selRx)) `
    'sentinel is vacuous if deleting the Select redirect still passes'

# (4) RaiseEvent 的**事件实参**: 这一路**不走** call_arg_emit —— RaiseEvent 自带一条发码
#     (cgen_com_events.cpp: `emitExpr(arg); argVal = lastExpr_;`), 所以 (2) 那道闸够不着。
#     ComCtlsDemo 实测就栽在这里: CoolBar.ctl `RaiseEvent HeightChanged(Extender.Height)`
#     发成 `onHeightChanged(h, vb6_Extender_Height)`。这是**第四条**消费路径, 少一道就漏。
$evtRx = 'extenderMemberReadExpr\s*\(\s*argVal\s*\)'
Check '4b: RaiseEvent argument gate present' ($evtT -match $evtRx) `
    'RaiseEvent <E>(Extender.<member>) passes a bare global as the event argument'
$negB4 = $evtT -replace '(?m)^[^\n]*extenderMemberReadExpr\s*\(\s*argVal\s*\)[^\n]*\r?\n', ''
Check 'negative control (restoring the original RaiseEvent argument line must go red)' (-not ($negB4 -match $evtRx)) `
    'sentinel is vacuous if deleting the RaiseEvent redirect still passes'
function BodyOf([string]$text, [string]$sig) {
    $i = $text.IndexOf($sig)
    if ($i -lt 0) { return "" }
    $j = $text.IndexOf("`n}", $i)
    if ($j -lt 0) { return $text.Substring($i) }
    return $text.Substring($i, $j - $i)
}
# --- kind 契约: 对象成员不许再报成 BSTR (3b 的前置, 不是清理) ---
# Container/DragIcon 在表里 type=Unknown, 记录字段却是 `void*`
# (vb6forms_uc_internal.h:100-101), 而通道给的是 *kind = 0 = BSTR —— 类型撒谎。
# 今天没炸只因为**读**侧闸门按 type=Unknown 原样返回; 一旦谁把读闸放开,
# `*(BSTR*)&rc->extDragIcon` 会把**活着的对象指针**当 BSTR 交出去, 调用方很可能
# SysFreeString 掉它。**写**侧 3a 的闸门不看类型, 早就在往这条路上走了。
# 约定扩一位: 4 = 对象 (void*)。凡是有 kind switch 的地方都必须跟上 —— 漏掉一处,
# default 分支会 `*(int32_t*)` 把 64 位指针截断, 那才是当场炸。
$KIND_OBJ = 4
$ctl  = [System.IO.File]::ReadAllText($rtl)
$invk = [System.IO.File]::ReadAllText((Join-Path (Split-Path -Parent $PSScriptRoot) 'src/rtl/core/vb6com/vb6com_invoke.c'))

$sites = [ordered]@{
    'vb6_UC_ExtenderField'  = 'void* vb6_UC_ExtenderField(void* p, const wchar_t* name, int32_t* kind) {'
    'vb6_UC_SetExtenderMember' = 'void vb6_UC_SetExtenderMember(const wchar_t* name, void* value) {'
    'vb6_ComGetProp'         = 'void* vb6_ComGetProp(void* disp, const wchar_t* propName) {'
    'vb6_ComSetProp'         = 'void vb6_ComSetProp(void* disp, const wchar_t* propName, void* value_void) {'
}
$bodies = @{}
foreach ($k in $sites.Keys) {
    $b = ''
    if ($k -eq 'vb6_ComGetProp' -or $k -eq 'vb6_ComSetProp') { $b = BodyOf $invk $sites[$k] } else { $b = BodyOf $ctl $sites[$k] }
    $bodies[$k] = $b
    Check "body located: $k" (-not [string]::IsNullOrWhiteSpace($b)) `
        'the signature must name the definition, not an extern declaration; an empty body makes every check below meaningless'
}

$efBody = $bodies['vb6_UC_ExtenderField']
# (a) 两个对象成员的行必须报对象 kind
$rowRx = '_wcsicmp\(name,\s*L"(Container|DragIcon)"\)\s*==\s*0\)\s*\{\s*\*kind\s*=\s*(\d+)'
$rows = [regex]::Matches($efBody, $rowRx)
# 每个成员现在有**两**行 (记录段 + 兜底段), 所以断的是"**没有一行**报成别的 kind",
# 不是"恰好几行" —— 数条数会在补上兜底段之后误报 (自己踩过)。
$objOk = ($rows.Count -ge 2) -and (@($rows | ForEach-Object { $_.Groups[2].Value } | Where-Object { $_ -ne "$KIND_OBJ" }).Count -eq 0)
Check "kind contract: Container/DragIcon are kind $KIND_OBJ (object), not BSTR" $objOk `
    'reporting them as BSTR means a live object pointer can be handed out as a BSTR and freed'
$negK = $efBody -replace '(\{\s*)\*kind\s*=\s*4', '$1*kind = 0'
$negRows = [regex]::Matches($negK, $rowRx)
$negKOk = ($negRows.Count -ge 2) -and (@($negRows | ForEach-Object { $_.Groups[2].Value } | Where-Object { $_ -ne "$KIND_OBJ" }).Count -eq 0)
Check 'negative control (marking the object rows back as BSTR must go red)' (-not $negKOk) `
    'sentinel is vacuous if the object rows can be re-typed to BSTR and still pass'

# (b) 每个 kind switch 都得认这个 kind, 且用**指针宽度**存/取
foreach ($k in @('vb6_UC_SetExtenderMember', 'vb6_ComGetProp', 'vb6_ComSetProp')) {
    $b = $bodies[$k]
    $okSw = $b -match ('case\s+' + $KIND_OBJ + '\s*:.*\*\(void\*\*\)')
    Check "kind $KIND_OBJ handled with a pointer-width access in $k" $okSw `
        'falling through to default uses int32 and truncates a 64-bit pointer'
    $negSw = $b -replace ('case\s+' + $KIND_OBJ + '\s*:'), 'case 99:'
    Check "negative control (removing the kind $KIND_OBJ arm from $k must go red)" (-not ($negSw -match ('case\s+' + $KIND_OBJ + '\s*:.*\*\(void\*\*\)'))) `
        'sentinel is vacuous if the object arm can be dropped and still pass'
}

# (c) 数值入口不许把对象成员的指针截成 int32 读出来
$gnum = BodyOf $ctl 'int32_t vb6_UC_GetExtenderMemberNum(const wchar_t* name) {'
Check 'body located: vb6_UC_GetExtenderMemberNum' (-not [string]::IsNullOrWhiteSpace($gnum)) 'wrong signature'
$gnumOk = $gnum -match ('kind\s*==\s*' + $KIND_OBJ)
Check "numeric read entry refuses kind $KIND_OBJ" $gnumOk `
    'reading 32 bits of an object pointer hands a truncated pointer to arithmetic'
$negG = $gnum -replace ('kind\s*==\s*' + $KIND_OBJ + '\s*\)'), 'kind == 99)'
Check "negative control (removing the kind $KIND_OBJ guard from the numeric entry must go red)" (-not ($negG -match ('kind\s*==\s*' + $KIND_OBJ))) `
    'sentinel is vacuous if the numeric guard can be removed and still pass'
# --- Container/DragIcon 走实例入口 (3b 的最后一格) ---
# 这两个是**对象**成员 (记录字段 void*), 与数值/字符串都不同 ⇒ 要第三个读入口
# `void* vb6_UC_GetExtenderMemberObj`。返回槽实测就是 `void* vb6_ret_Container`
# (ComCtlsDemo CoolBar.c:1651), 指针宽度正好对上, 调用点不需要任何强制转换。
# 表里 type=Unknown 的**方法**行 (Drag/SetFocus/ZOrder) 是**调用**不是读, 必须原样返回 ——
# 用 r.flags & HPF_METHOD 分开; 只看 type 会把这三行也换成一个 getter 调用。
$OBJ = 'vb6_UC_GetExtenderMemberObj'
Check "RTL defines $OBJ" ($rtlT -match ("void\*\s+" + $OBJ + "\s*\(")) 'object read entry point missing from uc_controls.c'
Check "header declares $OBJ" ($hdrT -match ([regex]::Escape($OBJ) + '\s*\([^;]*\)\s*;')) 'generated code calls it directly; without the prototype MSVC gives C2065'
$objBody = BodyOf $ctl ('void* ' + $OBJ + '(const wchar_t* name) {')
Check "body located: $OBJ" (-not [string]::IsNullOrWhiteSpace($objBody)) 'signature must name the definition, not a declaration'
# 对象入口只认 kind 4; 别的 kind 一律 NULL, 别把 BSTR / 数值槽当指针交出去
Check "object read entry accepts only kind 4" ($objBody -match 'kind\s*!=\s*4') `
    'handing a BSTR or int32 slot back as void* is exactly what kind 4 was introduced to stop'
$negO = $objBody -replace 'kind\s*!=\s*4', 'kind != 99'
Check "negative control (widening the kind guard in $OBJ must go red)" (-not ($negO -match 'kind\s*!=\s*4')) `
    'sentinel is vacuous if the guard can be relaxed and still pass'

# 共享出口: 对象行改道到 Obj, 方法行原样返回
Check 'shared read exit offers the object dock' ($comT -match ([regex]::Escape($OBJ))) `
    'Container/DragIcon reads would otherwise keep hitting the bare global'
Check 'shared read exit still returns method rows unchanged' ($comT -match 'flags\s*&\s*HPF_METHOD') `
    'Drag/SetFocus/ZOrder are calls, not reads'
$negM = $comT -replace 'flags\s*&\s*HPF_METHOD', 'flags & 0'
Check 'negative control (removing the method-row guard must go red)' (-not ($negM -match 'flags\s*&\s*HPF_METHOD')) `
    'sentinel is vacuous if the guard can be dropped and still pass'

# 写侧: Property Set 的**目标**那一问在另一条发码路径 (cgen_setlet_set_rhs.inc),
# 3a 那道闸 (cgen_assign_host_pseudo.inc) 够不着 —— 这就是那 66 处
# `vb6_Extender_Container = Value;  /* Set */` 一路活到今天的原因。
$setRx = 'rfind\("vb6_Extender_",\s*0\)\s*==\s*0'
Check 'Set-statement target is redirected to the instance-aware write entry' ($setT -match $setRx) `
    'vb6_Extender_Container = Value; still writes the bare global'
Check 'Set-statement redirect uses vb6_UC_SetExtenderMember' ($setT -match 'vb6_UC_SetExtenderMember') `
    'the target must go through the same write entry 3a uses'
$negS = $setT -replace 'rfind\("vb6_Extender_",\s*0\)\s*==\s*0', 'rfind("vb6_ExtenderX_", 0) == 0'
Check 'negative control (renaming the prefix gate in the Set path must go red)' (-not ($negS -match $setRx)) `
    'sentinel is vacuous if the Set gate can be neutered and still pass'
# Set 语句的**右值**那一问 (`Property Get Container` 的返回槽那句本身也是 Set 语句 ——
# 行尾那个标记就是证据; 目标侧是 vb6_ret_*, 所以目标闸正确地没碰它, 右值还是裸全局)。
$setValRx = 'value\s*=\s*extenderMemberReadExpr\s*\(\s*value\s*\)'
Check 'Set-statement value side is redirected' ($setT -match $setValRx) `
    'vb6_ret_Container = vb6_Extender_Container still reads the bare global (66 sites)'
$negSV = $setT -replace '(?m)^[^\n]*value\s*=\s*extenderMemberReadExpr\s*\(\s*value\s*\)[^\n]*\r?\n', ''
Check 'negative control (restoring the original Set value line must go red)' (-not ($negSV -match $setValRx)) `
    'sentinel is vacuous if deleting the Set value redirect still passes'

# ⚠ 成员表达式**出生地**(m22)**刻意不换道** —— 这条是**反**向的钉, 钉的是一个踩过的坑。
#   出生地 (cgen_expr_member_m22_module.inc) 是 `Pseudo.<成员>` 符号成形的地方, 看着像
#   "读侧总出口", 换道之后比较式那 4 处确实接上了 —— 但它**也服务左值/赋值目标**:
#   合成出来的 `Property Set Container` 存取器体, 它的**目标**是从同一个成员表达式
#   算出来的。换过去以后目标变成 `vb6_UC_GetExtenderMemberObj(L"Container")` —— 一个
#   **函数调用**, 于是 Fix 098 那道"Set 到不可赋值函数引用"闸把它当非左值, 整个赋值
#   **静默丢**成 `(void)(Value);`。实测: ComCtlsDemo 全绿(VBFlexGrid 那条外部属性
#   走的是另一条发码路径, 没暴露), VBFlexGridDemo 的 `Property Set Container/DragIcon`
#   两个写**当场消失**。compile 75/75 + syntax 171/171 **照样全绿** —— 没有任何一条用例
#   断言 Set 存取器真的存进去了。是形状清单把那 1 行哈希变化逐行对出来才发现的。
#   ⇒ 换道只能落在**消费点**, 且要分清这一问是读还是写。比较式那 4 处因此留着没接,
#   留给它们自己那一刀; 宁可留一处有记录的缺口, 也不换一处静默丢写。
$m22Body = BodyOf $m22T 'lastExpr_ = hostPseudoRtlSymbol(modName2, varName2);'
Check 'member-expression birth site is NOT redirected (it also serves lvalue targets)' `
    ($m22T -notmatch 'extenderMemberReadExpr\s*\(\s*hostPseudoRtlSymbol') `
    'gating the symbol birth site turns Set-accessor targets into calls; Fix 098 then silently discards the write'
$negM22 = $m22T -replace 'lastExpr_ = hostPseudoRtlSymbol\(modName2, varName2\);', `
                              'lastExpr_ = extenderMemberReadExpr(hostPseudoRtlSymbol(modName2, varName2));'
Check 'negative control (gating the birth site must go red)' `
    ($negM22 -match 'extenderMemberReadExpr\s*\(\s*hostPseudoRtlSymbol') `
    'sentinel is vacuous if re-introducing the birth-site gate still passes'
# 二元操作数的消费点 (cgen_expr_binary.cpp): 比较式里那一枚宿主伪成员
# (UpDown.ctl 那 4 处 `vb6_ComGetObjectProp(x, L"Container") == vb6_Extender_Container`)。
# 这是 3b 的最后一族, 也是**唯一**能安全换道的地方:
#   * 操作数**永远不是左值** —— 与成员表达式出生地不同, 那里同时服务赋值目标, 换过去会把
#     `Property Set` 存取器的目标变成函数调用, Fix 098 随即把整个赋值静默丢成 (void)(Value);
#   * 所以这一族必须走消费点, 而不是回到出生地 —— 哨兵里那条反向钉就是这个意思。
$binRxL = 'left\s*=\s*extenderMemberReadExpr\s*\(\s*left\s*\)'
$binRxR = 'right\s*=\s*extenderMemberReadExpr\s*\(\s*right\s*\)'
Check 'binary operands are redirected (covers comparison operands)' `
    (($binT -match $binRxL) -and ($binT -match $binRxR)) `
    'a host-pseudo member compared against something still reads the bare global'
$negBin = $binT -replace '(?m)^[^\n]*(left|right)\s*=\s*extenderMemberReadExpr\s*\(\s*(left|right)\s*\);[^\n]*\r?\n', ''
Check 'negative control (restoring the original binary operand lines must go red)' `
    ((-not ($negBin -match $binRxL)) -and (-not ($negBin -match $binRxR))) `
    'sentinel is vacuous if deleting the operand redirect still passes'
if ($bad) { Write-Host "check_extender_emit_channel: FAIL ($bad)"; exit 1 }
Write-Host 'check_extender_emit_channel: PASS'
exit 0