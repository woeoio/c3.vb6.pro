# check_class_refcount_authority.ps1 - 账 #305 §B218 (甲)：工程类实例的引用计数只许有一位主人。
# 三十二条钉（每条都能单独红，红了就说明同一问长出了第二位主人）：
#   ① 出入口那一对的 mint 各恰好 1 处，且住在 module/cgen_com.cpp（每条腿只许转调它）。
#   ② 计数头 `int32_t __refcount;` 的 mint 恰好 1 处，且住在 cgen_iface_vtbl.cpp 的
#      emitIfaceClassFields 里 —— 它必须紧跟 __comObj 那一句（ai/022 D19：偏移 0 恒是 __comObj）。
#   ③ 覆写/置空那一位主人 emitClassSlotRefSet 的消费文件恰好 2 份（Set 的 RHS 臂 + 置空臂）；
#      第 3 份只能由腿④/腿⑤ 那两刀带进来，届时这里跟着改数。
#   ④ 工程类形参登记收成一处出口 registerClassFormal：消费文件恰好 3 份（Sub / Function /
#      Property 三个入口），调用点恰好 9 处（每份三条：跨模块兜底 / Class 符号 / ComClass 同名）。
#   ⑤ procClassFormals_ 那张名册只许住在 10 份文件里（表本体 + 三个入口清表 + 出口本体 +
#      两条腿各问一次 + 第一百四十四刀形A 的三问：出参槽类名 / 直传所有权 / 链形那道闸）。
#   ⑥ 字面量 __refcount 的主人名单 = 在册那 6 份；RTL 那一处必须是 vb6comserver_obj.c 的注释。
#   ⑦ 陈旧注释「没有引用计数」（具体类那三处）不许活着 —— 甲落地之后这句话是错的。
#   ⑧ 过程出口那张带档的登记表 slotsToRelease_ 只许住在 5 份文件里（表本体 + 三个过程入口
#   ⑨ 出参槽那一档 (腿②'): claim-only 的发码形状与站点各恰好一处。
#   ⑩ 腿④ 的第二道夹: 撤登记的主人一处 + 站点一处 (槽地址交给外部 Declare 的那一格出口不放手)。
#   ⑪ 腿⑤ 出厂那一档: `_Release(me->` 全仓恰好一处，且主人文件必须还是 cgen_com.cpp。
#   ⑫ 字段类型那一问 `mapTypeRef(var.asType.get())` 恰好 6 处（第一百四十刀给出厂加了数组那一档；
#      新增一处手抄就得改数）。
#   ⑬ 移交值那一问 `rhsTransfersProjectRef(` 恰好 3 处（声明 + 定义 + 一个调用点），调用点必须在
#      cgen_setlet_set_rhs.inc —— 「右侧带不带对价」与 `New` 共用同一条不 claim 的臂，别处再问就分家。
#      清表 + 主人本体），工程类那一档的注册口 trackClassSlotForRelease 恰好 2 处
#      （Dim As <工程类> 的两条注册路：Class 符号 / projectClassNameOf 兜底）。
#   ⑭ 移交值那一问的三形只住一个函数: 成员返回类那一问 
#      恰好 1 处、 基名那一问  恰好 2 处 (含本刀那一臂,
#      另一处是 myBaseClassOf)、接口视图那道排除  恰好 1 处。
#   ⑮ Static 那道闸 `!var.isStatic && !inStaticProc_` 全仓恰好 3 处、且全住在一份文件 (cgen_localdecl.cpp)：登记「出口放手」的三条腿 (两条工程类 + 一条接口视图) 共用同一道闸 —— 第一百三十三刀补的就是视图那条腿缺的那半。视图注册口 `trackIvrefLocalForRelease(cIdent(var.name));` 恰好 1 处。
#   ⑰ 第一百三十七刀: 实参那一站「交回进站那一枚」只许住在 registerClassArgRefHandback 里 (注释串全仓 1 处、文件 = cgen_com.cpp)，两位站点 (ByVal 那一档 + ByRef 092r 那一档) 都只调它，谁也不许自己补一行 Release。问句 `argHandsOverProjectRef` = 契约问句 + New 那一形的合起来答一次: 声明 + 定义 + 两处调用 = 4 处 / 4 份文件。
#   ⑯ 第一百三十四刀: 视图腿② 的「交回接管来的那一枚」只许住在 emitIvrefQiSet 里 (全仓 1 处、文件 = cgen_iface_vtbl.cpp)，而它的**按名**调用点两位都住在 cgen_setlet_set_prop.inc；第二百零三刀起第三位消费者住在 cgen_setlet_set_rhs.inc (字段槽那一臂，转调同一位主人) —— 移交那一问 (⑬/⑭) 现在是两站四处：工程类站 1 处 (forViewSlot=false) + 视图站 3 处 (true)；再冒出一处自己补 Release 的站点就是第二套判据。
# 退出码：0 = 三十二条都过；1 = 任一条红；2 = 找不到 src/backend。
param(
    [string]$Root = ""
)
$ErrorActionPreference = 'Continue'
if (-not $Root) { $Root = (Split-Path -Parent $PSScriptRoot) }
$backend = Join-Path $Root "src\backend"
if (-not (Test-Path $backend)) {
    Write-Host ("MISSING " + $backend)
    exit 2
}
$srcFiles = Get-ChildItem -Path (Join-Path $Root "src") -Recurse -File |
    Where-Object { $_.Extension -match '^\.(cpp|inc|c|h|hpp)$' }
$backendFiles = Get-ChildItem -Path $backend -Recurse -File |
    Where-Object { $_.Extension -match '^\.(cpp|inc|c|h|hpp)$' }

$fail = @()

function CountOf([string]$pattern) {
    $n = 0
    foreach ($f in $srcFiles) {
        $t = [IO.File]::ReadAllText($f.FullName)
        $n += ([regex]::Matches($t, [regex]::Escape($pattern))).Count
    }
    return $n
}
function FilesOf([string]$pattern) {
    $r = @()
    foreach ($f in $srcFiles) {
        $t = [IO.File]::ReadAllText($f.FullName)
        if (([regex]::Matches($t, [regex]::Escape($pattern))).Count -gt 0) { $r += $f.Name }
    }
    return @($r | Sort-Object -Unique)
}

# ① mint 那一对
foreach ($sig in @('"_AddRef(void* instance) {', '"_Release(void* instance) {')) {
    $bs = @($backendFiles | Select-String -SimpleMatch $sig)
    if ($bs.Count -ne 1) {
        $fail += ("pair mint sites for " + $sig + " = " + $bs.Count + " (want exactly 1)")
    } elseif ($bs[0].Path -notmatch 'cgen_com\.cpp$') {
        $fail += ("pair mint moved into " + $bs[0].Path)
    }
}

# ② 计数头的 mint
$hdr = @($backendFiles | Select-String -SimpleMatch '    int32_t __refcount;')
if ($hdr.Count -ne 1) {
    $fail += ("refcount-header mint = " + $hdr.Count + " (want exactly 1)")
} elseif ($hdr[0].Path -notmatch 'cgen_iface_vtbl\.cpp$') {
    $fail += ("refcount-header mint moved into " + $hdr[0].Path)
}

# ③ 覆写/置空那位主人的消费文件
$ownerFiles = @(FilesOf 'emitClassSlotRefSet' |
    Where-Object { $_ -ne 'cgen_com.cpp' -and $_ -ne 'cgen_helpers.inc' })
if ($ownerFiles.Count -ne 2) {
    $fail += ("emitClassSlotRefSet consumer files = " + $ownerFiles.Count +
              " (want exactly 2): " + ($ownerFiles -join ', '))
}

# ④ 形参登记那枚出口：文件 3 份、调用点 9 处
$regFiles = @(FilesOf 'registerClassFormal' |
    Where-Object { $_ -ne 'cgen_util.cpp' -and $_ -ne 'cgen_helpers.inc' })
if ($regFiles.Count -ne 3) {
    $fail += ("registerClassFormal consumer files = " + $regFiles.Count +
              " (want exactly 3): " + ($regFiles -join ', '))
}
$regSites = CountOf 'registerClassFormal(pLower, '
if ($regSites -ne 9) {
    $fail += ("registerClassFormal call sites = " + $regSites + " (want exactly 9)")
}

# ⑤ procClassFormals_ 的名册面
$cfFiles = FilesOf 'procClassFormals_'
if ($cfFiles.Count -ne 10) {
    $fail += ("procClassFormals_ files = " + $cfFiles.Count +
              " (want exactly 10): " + ($cfFiles -join ', '))
}

# ⑥ __refcount 字面量的主人名单
$rcFiles = FilesOf '__refcount'
$wantRc = @('cgen_com.cpp', 'cgen_expr_call_arg_emit.inc', 'cgen_iface_vtbl.cpp',
            'cgen_setlet_set_prop.inc', 'driver_classchain.cpp', 'vb6comserver_obj.c') |
    Sort-Object -Unique
if (($rcFiles -join ',') -ne ($wantRc -join ',')) {
    $fail += ("__refcount file list changed: got " + ($rcFiles -join ', ') +
              " (want " + ($wantRc -join ', ') + ")")
}
$rtlRc = @(Get-ChildItem -Path (Join-Path $Root "src\rtl") -Recurse -File |
    Where-Object { $_.Extension -match '^\.(c|h)$' } |
    ForEach-Object { ([regex]::Matches([IO.File]::ReadAllText($_.FullName), '__refcount')).Count } |
    Measure-Object -Sum).Sum
if ($rtlRc -ne 1) {
    $fail += ("__refcount hits in src/rtl = " + $rtlRc +
              " (want exactly 1: the comment in vb6comserver_obj.c)")
}

# ⑦ 陈旧注释
$stale = CountOf '没有引用计数'
if ($stale -ne 0) {
    $fail += ("stale '没有引用计数' comment lines = " + $stale + " (want 0)")
}

# ⑧ 出口那一腿的登记面
$srFiles = FilesOf 'slotsToRelease_'
if ($srFiles.Count -ne 5) {
    $fail += ("slotsToRelease_ files = " + $srFiles.Count +
              " (want exactly 5): " + ($srFiles -join ', '))
}
$csSites = CountOf 'trackClassSlotForRelease(cIdent(var.name)'
if ($csSites -ne 2) {
    $fail += ("trackClassSlotForRelease registrant sites = " + $csSites +
              " (want exactly 2)")
}

# ⑨ 出参槽那一档 (腿②'): claim-only 的发码形状只许住在主人本体里一处, 站点也恰好一处。
#    为什么单独钉: 腿④ 会在过程出口放手局部, 写进出参槽的那一枚若不自己取走引用, 调用方
#    还看着它时它就拆了 (门 #763 实测 b187 第五臂 probe exit=-1073740940)。
$primeShapes = CountOf '(甲) 腿②'': Set 出参槽 '
if ($primeShapes -ne 1) {
    $fail += ("out-slot claim shape sites = " + $primeShapes + " (want exactly 1)")
}
$primeCalls = CountOf 'emitClassSlotRefSet(target, clsOfRhs'
if ($primeCalls -ne 1) {
    $fail += ("out-slot claim call sites = " + $primeCalls + " (want exactly 1)")
}

# ⑩ 腿④ 的第二道夹 (地址交给外部 Declare ⇒ 那一格出口不放手): 撤登记的主人只许住在
#    cgen_iface_vtbl.cpp 一处 (与登记那两位同文件 = 同一张表只有一位管), 站点恰好一处
#    (`As Any` ByRef 实参那一站)。为什么单独钉: 那一格的写者在计数系统外面 (C3 的
#    __vbaObjSetAddref 桩对工程内实例刻意不加引用), 出口再退 = 无对价的退 = 当场拆活对象。
$dropDefs = @($backendFiles | Select-String -SimpleMatch 'void CCodeGen::dropClassSlotRelease')
if ($dropDefs.Count -ne 1) {
    $fail += ("dropClassSlotRelease definition sites = " + $dropDefs.Count + " (want exactly 1)")
} elseif ($dropDefs[0].Path -notmatch 'cgen_iface_vtbl\.cpp$') {
    $fail += ("dropClassSlotRelease moved into " + $dropDefs[0].Path)
}
$dropSites = CountOf 'dropClassSlotRelease(argValLv)'
if ($dropSites -ne 1) {
    $fail += ("dropClassSlotRelease call sites = " + $dropSites + " (want exactly 1)")
}

# ⑪ 腽⑤ 出厂那一档（字段那一枚引用放手）: 全仓只许有一处把 _Release 发在 me-> 那一格上，而且必须住在 cgen_com.cpp 的 `_Destroy` 尾巴里。为什么单独钉: 字段进计数系统是从第一刀起（覆写 / 置空两站都问同一位主人），只有出厂这一站缺着 —— 补的时候若再开第二位主人，「谁握字段那一枚」就又是两份答案（本仓那一族病灶: #167/#231/#245/#298/#443）。
$teardownShapes = CountOf '_Release(me->'
if ($teardownShapes -ne 1) {
    $fail += ("field-teardown _Release(me-> sites = " + $teardownShapes + " (want exactly 1)")
}
$teardownFiles = @(FilesOf '_Release(me->')
if ($teardownFiles.Count -ne 1 -or $teardownFiles[0] -ne 'cgen_com.cpp') {
    $fail += ("field-teardown owner file = " + ($teardownFiles -join ', ') + " (want exactly cgen_com.cpp)")
}

# ⑫ 字段类型的问句名单: 出厂这一站的类名是问 `mapTypeRef(var.asType.get())` 要的 —— 与发这枚字段那一行同一个答案，不开第二问。这条钉的是「这一问一共有几处」，新增一处手抄就得改这里的数（有意让它响）。
$fieldQ = CountOf 'mapTypeRef(var.asType.get())'
if ($fieldQ -ne 6) {
    $fail += ("mapTypeRef(var.asType.get()) question sites = " + $fieldQ + " (want exactly 6)")
}


# ⑬ 移交值那一问的名单: 「右侧那一枚是不是已经带对价交过来了」全仓只许问一处 (声明 + 定义 + 一个调用点 = 3 处
#    `rhsTransfersProjectRef(`)，而且调用点必须住在 Set 的 RHS 那一档 (cgen_setlet_set_rhs.inc)。为什么钉: 这一问与
#    `New` 共用同一条「不 claim」的臂 —— 谁在别的站点再问一遍，「实参 / 返回值 / 字段」这三条边的口径就分家了。
#    第一百三十七刀: 实参那一站 (⑰) 也从这里问, 但不**直接**吃 `*node.value` —— 它套在
#    argHandsOverProjectRef 那一层里, 所以下面 `$xferAsk` (只数站点直接问句) 仍是 4, 本条才 7。
#    第二百零三刀: Set 的 RHS 那一档长出**视图**那一臂 (字段槽的声明档是新式接口视图 ⇒ 转调
#    主人 emitIvrefQiSet), 同一条语句按槽档分派时问句的 forViewSlot 参数不同 ⇒ 站点直接问句
#    3→4。这一处依旧不自己补 Release —— 交回那一发只住在主人里 (⑯)。
$xferQ = CountOf 'rhsTransfersProjectRef('
if ($xferQ -ne 7) {
    $fail += ("transfer-question sites = " + $xferQ + " (want exactly 7: decl + def + 4 station calls + 1 ask behind the argument-station wrapper)")
}
$xferFiles = @(FilesOf 'rhsTransfersProjectRef(*node.value')
if ($xferFiles.Count -ne 2 -or -not ($xferFiles -contains 'cgen_setlet_set_rhs.inc') -or -not ($xferFiles -contains 'cgen_setlet_set_prop.inc')) {
    $fail += ("transfer-question call files = " + ($xferFiles -join ', ') + " (want exactly cgen_setlet_set_rhs.inc + cgen_setlet_set_prop.inc)")
}
$xferAsk = CountOf 'rhsTransfersProjectRef(*node.value'
if ($xferAsk -ne 4) {
    $fail += ("transfer-question call sites = " + $xferAsk + " (want exactly 4: 工程类站 1 + 视图站 3)")
}

# ⑭ 移交值那一问的**三形**只许住在这一个函数里 (第一百三十二刀 = (甲) 刀4 加了两形: callee 是成员访问
#    `hh.Make17()` 与无括号的属性读 `hh.Next18`)。三条臂各自沿用一位现成主人，一条都不新立:
#      · 成员是不是过程、返回哪一枚工程类 —— 只问 `getClassMethodReturnType(recvCls, ma.memberName)` 这一处;
#      · 数据字段那一档 (那一枚别人还握着, 该 claim) —— 靠 `classTypedFieldMap_` 命中排除;
#      · `MyBase` 那一形的基名 —— 与 cgen_inherit.cpp 的 myBaseClassOf 同读链表, 全仓只许这两处。
#    谁在别的站点再问一遍这三问，「实参 / 返回值 / 字段 / 视图」的口径就 again 分家 (本仓病灶 #167/#231/#245)。
$xferRetQ = CountOf 'getClassMethodReturnType(recvCls, ma.memberName)'
if ($xferRetQ -ne 1) {
    $fail += ("transfer-question member-return ask sites = " + $xferRetQ + " (want exactly 1, inside the one question)")
}
$xferBaseQ = CountOf 'classViewByName(self->baseKey)'
if ($xferBaseQ -ne 2) {
    $fail += ("base-name question sites = " + $xferBaseQ + " (want exactly 2: myBaseClassOf + 刀4 那一臂)")
}
$xferBaseFiles = @(FilesOf 'classViewByName(self->baseKey)')
if (-not ($xferBaseFiles -contains 'cgen_com.cpp')) {
    $fail += ("base-name question missing from cgen_com.cpp (files = " + ($xferBaseFiles -join ', ') + ")")
}
$xferIvref = CountOf 'ivrefCType(retCls)'
if ($xferIvref -ne 1) {
    $fail += ("interface-view exclusion sites = " + $xferIvref + " (want exactly 1)")
}

# ⑮ 第一百三十三刀 (第一百五十二刀加两条载体臂): Static 那一档「出口不许放手」这条闸, 五条腿都得有 (两条工程类腿本来就有,
#    视图那条腿是这一刀补的)。登记站点与闸成对: 3 处登记 <=> 3 处闸, 少一半就是这道闸被拆了。
$gateN = CountOf '!var.isStatic && !inStaticProc_'
if ($gateN -ne 5) {
    $fail += ("static-gate sites = " + $gateN + " (want exactly 5: two project-class legs + the interface-view leg + 载体档两条臂 = 第一百五十二刀)")
}
$gateFiles = @(FilesOf '!var.isStatic && !inStaticProc_')
if ($gateFiles.Count -ne 1 -or $gateFiles[0] -ne 'cgen_localdecl.cpp') {
    $fail += ("static-gate files = " + ($gateFiles -join ', ') + " (want exactly cgen_localdecl.cpp)")
}
$ivrefReg = CountOf 'trackIvrefLocalForRelease(cIdent(var.name));'
if ($ivrefReg -ne 1) {
    $fail += ("interface-view release-registration sites = " + $ivrefReg + " (want exactly 1)")
}

# ⑯ 第一百三十四刀: 视图腿② QI 完把接管来的那一枚交回 —— 这一发只许住在 emitIvrefQiSet 里
#    (与 ⑬/⑭ 那条问句同一条契约: 工程类那一站「不 claim」、视图这一站「当场交回」, 两处都问同一位主人)。
$handN = CountOf '腿②: 交回接管来的那一枚'
if ($handN -ne 1) {
    $fail += ("view-leg hand-back emit sites = " + $handN + " (want exactly 1, inside emitIvrefQiSet)")
}
$handFiles = @(FilesOf '腿②: 交回接管来的那一枚')
if ($handFiles.Count -ne 1 -or $handFiles[0] -ne 'cgen_iface_vtbl.cpp') {
    $fail += ("view-leg hand-back files = " + ($handFiles -join ', ') + " (want exactly cgen_iface_vtbl.cpp)")
}
$qiCalls = CountOf 'emitIvrefQiSet(target, v->name, value'
if ($qiCalls -ne 2) {
    $fail += ("view-leg QiSet call sites = " + $qiCalls + " (want exactly 2: 借用那一档 + 形① 那一档)")
}
$qiFiles = @(FilesOf 'emitIvrefQiSet(target, v->name, value')
if ($qiFiles.Count -ne 1 -or $qiFiles[0] -ne 'cgen_setlet_set_prop.inc') {
    $fail += ("view-leg QiSet call files = " + ($qiFiles -join ', ') + " (want exactly cgen_setlet_set_prop.inc)")
}
# ⑰ 第一百三十七刀: 实参那一站 (形参是工程类 / 视图) 交完就忘的那一枚, 从此由一位主人登记交回。
#    站点两处 (ByVal 档 / ByRef 092r 档) 都只调这一枚; 注释串全仓 1 处 = 「没有第二站自己补 Release」。
$argTag = CountOf '腿⑤: 实参那一站交回'
if ($argTag -ne 1) {
    $fail += ("argument-station hand-back emit sites = " + $argTag + " (want exactly 1, inside registerClassArgRefHandback)")
}
$argFiles = @(FilesOf '腿⑤: 实参那一站交回')
if ($argFiles.Count -ne 1 -or $argFiles[0] -ne 'cgen_com.cpp') {
    $fail += ("argument-station hand-back files = " + ($argFiles -join ', ') + " (want exactly cgen_com.cpp)")
}
$argHelp = CountOf 'registerClassArgRefHandback('
# 第一百四十四刀: 站点从两处增到四处 (ByVal 档 + ByRef 裸标识符临时量档 + ByRef 复合字面量档 + 命名实参档)
# —— 形A 之后**每一条副本交付都得有人把被调方留在临时量里的那一枚收回来**，漏一站就是一枚永久泄漏。
if ($argHelp -ne 6) {
    $fail += ("argument-station helper sites = " + $argHelp + " (want exactly 6: declaration + definition + 4 call stations)")
}
$argAsk = CountOf 'argHandsOverProjectRef('
if ($argAsk -ne 4) {
    $fail += ("argument-station question sites = " + $argAsk + " (want exactly 4: declaration + definition + 2 call stations)")
}
# ⑱ 第一百四十刀: 下标槽那一档 (数组元素) 的三站只许问一位主人。
#    · classArraySlotElemType = 「这枚目标是不是可寻址的元素槽」: 声明 + 定义 + 两条腿各问一次 = 4 处 / 4 份文件
#      (再多一处 = 有站点绕过问句自己认 VB6_SA_AT)。
#    · projectClsStructOfCType = 「这一档是不是本工程类的指针」: 声明 + 定义 + 出厂两档 = 4 处，两份站点必须在同一份 cgen_com.cpp。
#    · 地址临时量 __clsaddr 的 mint 全仓恰好 1 处 (只有 emitClassSlotRefSet 铸它) ⇒「下标算术算几次」只有一个答案。
#    · 元素出厂那一发 vb6_SA_ReleaseElements = RTL 声明 + 定义 + 字段出厂那一站 = 3 处；数组那一支发的是 fn 变量挑出来的两档之一，所以它不再以 `...Elements(` 的形状出现在后端。
$askElem = CountOf 'classArraySlotElemType('
if ($askElem -ne 4) {
    $fail += ("index-slot question sites = " + $askElem + " (want exactly 4: declaration + definition + 2 legs)")
}
$askElemFiles = @(FilesOf 'classArraySlotElemType(')
if ($askElemFiles.Count -ne 4) {
    $fail += ("index-slot question files = " + ($askElemFiles -join ', ') + " (want exactly 5 files)")
}
$askStruct = CountOf 'projectClsStructOfCType('
if ($askStruct -ne 5) {
    $fail += ("project-class-pointer question sites = " + $askStruct + " (want exactly 5: declaration + definition + 2 teardown legs + 载体销毁那一问)")
}
$askStructFiles = @(FilesOf 'projectClsStructOfCType(')
if (-not ($askStructFiles -contains 'cgen_com.cpp')) {
    $fail += ("project-class-pointer question missing from cgen_com.cpp (files = " + ($askStructFiles -join ', ') + ")")
}
$addrMint = CountOf '"__clsaddr" + std::to_string'
if ($addrMint -ne 1) {
    $fail += ("address-temp mint sites = " + $addrMint + " (want exactly 1, inside emitClassSlotRefSet)")
}
$elemOut = CountOf 'vb6_SA_ReleaseElements('
if ($elemOut -ne 3) {
    $fail += ("element-teardown export sites = " + $elemOut + " (want exactly 3: RTL declaration + definition + 字段出厂那一站 —— 数组那一支由主人用 fn 变量挑两档出口，自己拼那发释放的站点多一处就是第二位主人)")
}
$elemOutFiles = @(FilesOf 'vb6_SA_ReleaseElements')
if (-not ($elemOutFiles -contains 'cgen_com.cpp')) {
    $fail += ("element-teardown emit station missing from cgen_com.cpp (files = " + ($elemOutFiles -join ', ') + ")")
}
# ⑲ 第一百四十四刀: 形A 那座桥的三处问句各只许一位主人。
#    · classOfProjectRefOutSlot = 「(*名) 里的名是不是本过程的工程类形参 (交回类名)」:
#      声明 + 定义 + 覆写/置空两站 = 4 处 / 4 份文件 (再多一处 = 有站点自己拆括号按名查表)。
#    · argHoldsProjectRefInOwnSlot = 「这一枚实参持有它交出去的那一枚引用吗 (才许直传地址)」:
#      声明 + 定义 + 位置实参臂 + 命名实参臂 = 4 处 / 4 份文件。
#    · registerClassArgRefCopyIn = 借来的实参在临时量上先付一枚 (copy-in) 的唯一发码处:
#      声明 + 定义 + 位置实参两臂 (裸标识符 / 复合字面量) + 命名实参臂 = 5 处 / 4 份文件。
$askOutSlot = CountOf 'classOfProjectRefOutSlot('
if ($askOutSlot -ne 4) {
    $fail += ("out-slot class question sites = " + $askOutSlot + " (want exactly 4: declaration + definition + 2 stations)")
}
$askOutSlotFiles = @(FilesOf 'classOfProjectRefOutSlot(')
if ($askOutSlotFiles.Count -ne 4) {
    $fail += ("out-slot class question files = " + ($askOutSlotFiles -join ', ') + " (want exactly 5 files)")
}
$askOwnSlot = CountOf 'argHoldsProjectRefInOwnSlot('
if ($askOwnSlot -ne 4) {
    $fail += ("owner-direct-pass question sites = " + $askOwnSlot + " (want exactly 4: declaration + definition + 2 arg arms)")
}
$askCopyIn = CountOf 'registerClassArgRefCopyIn('
if ($askCopyIn -ne 5) {
    $fail += ("copy-in claim sites = " + $askCopyIn + " (want exactly 5: declaration + definition + 3 temp stations)")
}
$askCopyInFiles = @(FilesOf 'registerClassArgRefCopyIn(')
if ($askCopyInFiles.Count -ne 4) {
    $fail += ("copy-in claim files = " + ($askCopyInFiles -join ', ') + " (want exactly 4 files)")
}
# ⒉ 第一百四十八刀: Preserve 丢尾那一格换的是一枚**带回调的兄弟出口** vb6_SafeArrayReDimPreserve1D_Tr。
#    它全仓恰好 4 处 / 4 份文件: RTL 声明 + 定义 + 发码站 (cgen_redim.cpp) + **载体前缀名单** (cgen_util_type.cpp)。
#    为什么要把名单那一格一起钉: 「这枚表达式**已经是载体**」答错 = 再包一层 = 双重解引用，
#    而它与出口本体住在两份文件里，只改一处就会静默错。新增一枚这样的出口时两格必须一起动。
$trExport = CountOf 'vb6_SafeArrayReDimPreserve1D_Tr('
if ($trExport -ne 4) {
    $fail += ("preserve-with-release export sites = " + $trExport + " (want exactly 4: RTL decl + definition + emit station + carrier-prefix list)")
}
$trExportFiles = @(FilesOf 'vb6_SafeArrayReDimPreserve1D_Tr(')
if ($trExportFiles.Count -ne 4) {
    $fail += ("preserve-with-release export files = " + ($trExportFiles -join ', ') + " (want exactly 4 files: 名单那一格不许漏)")
}
# ⑳ 第一百四十六刀: 载体销毁那一族 (Erase / ReDim 不带 Preserve) 的两问只许一位主人。
#    · classOfArrayCarrier = 「这枚数组变量握着的元素是不是工程类」: 声明 + 定义 + 三位读者 (销毁主人
#      + Preserve 两站) = 5 处 / 4 份文件 (第一百五十六刀起登记口不再问这一句 —— 放宽到每一档局部数组,
#      认不出元素类只是不交回元素, 载体那一发照旧要收尾; 那位主人剩下的读者全是元素出厂那两腿)。
#    · emitClassArrayElemHandback = 销毁载体前那一发的唯一发码处: 声明 + 定义 + 五个站点 (1D Erase / 1D ReDim / ND ReDim 两支 / ND Erase) = 7 处 / 3 份文件，站点必须全住在 cgen_redim.cpp。
#      两位站点必须全住在 cgen_redim.cpp。只要有站自己拼那发释放，“谁握元素那一枚”就又分家。
$askCarrier = CountOf 'classOfArrayCarrier('
if ($askCarrier -ne 5) {
    $fail += ("array-carrier question sites = " + $askCarrier + " (want exactly 5: declaration + definition + 3 readers = 销毁主人 + Preserve 两站)")
}
$askCarrierFiles = @(FilesOf 'classOfArrayCarrier(')
if ($askCarrierFiles.Count -ne 4) {
    $fail += ("array-carrier question files = " + ($askCarrierFiles -join ', ') + " (want exactly 4 files)")
}
$carrierHelp = CountOf 'emitClassArrayElemHandback('
if ($carrierHelp -ne 8) {
    $fail += ("carrier-destroy hand-back sites = " + $carrierHelp + " (want exactly 8: declaration + definition + 5 stations + 过程出口那一档 = 第一百五十二刀)")
}
$carrierHelpFiles = @(FilesOf 'emitClassArrayElemHandback(')
if (-not ($carrierHelpFiles -contains 'cgen_redim.cpp')) {
    $fail += ("carrier-destroy stations missing from cgen_redim.cpp (files = " + ($carrierHelpFiles -join ', ') + ")")
}
# ㉒ 第一百五十刀: 多维载体那一档接进同一位主人之后, 这一问与这枚出口各只许有一位主人。
#    · carrierIsMultidim = 「这枚载体是几维」: 声明 + 定义 + 主人内部那一问 + 第一百五十六刀出口那一问
#      (从 1D / ND 两种销毁出口里挑一种 = 同一问的第二次调用, 不是第二位主人) = 4 处 / 4 份文件。
#      站点若自己按 arrayDimCounts_ 判维度 = 同一问长出第二位主人, 1D 读法套 ND 载体就是步长错位。
#    · vb6_SA_ReleaseElementsND = 多维那一发的出口: RTL 声明 + 定义 + 主人里那枚字符串 = 3 处 / 3 份文件。
#      后端多一处 = 有站点绕过主人自己拼那发释放。
$askDims = CountOf 'carrierIsMultidim('
if ($askDims -ne 4) {
    $fail += ("carrier-shape question sites = " + $askDims + " (want exactly 4: declaration + definition + 销毁主人 + 出口挑销毁那一问)")
}
$askDimsFiles = @(FilesOf 'carrierIsMultidim(')
if ($askDimsFiles.Count -ne 4) {
    $fail += ("carrier-shape question files = " + ($askDimsFiles -join ', ') + " (want exactly 4 files)")
}
$ndExport = CountOf 'vb6_SA_ReleaseElementsND'
if ($ndExport -ne 3) {
    $fail += ("multidim hand-back export sites = " + $ndExport + " (want exactly 3: RTL decl + definition + 主人里那枚字符串)")
}
$ndExportFiles = @(FilesOf 'vb6_SA_ReleaseElementsND')
if ($ndExportFiles.Count -ne 3) {
    $fail += ("multidim hand-back export files = " + ($ndExportFiles -join ', ') + " (want exactly 3 files)")
}

# ㉓ 第一百五十二刀立、第一百五十六刀放宽: 「过程级数组载体出口收尾」这一档的登记口只许有一位,
#    且只有声明那两条臂能调它。名字从 trackClass... 改成 trackArray... 是因为它不再只认工程类元素。
#    · trackArrayCarrierForRelease = 声明 + 定义 + 两条臂 (固定 / 动态) = 4 处 / 3 份文件。
$trackCarrier = CountOf 'trackArrayCarrierForRelease('
if ($trackCarrier -ne 4) {
    $fail += ("carrier-scope-exit registrar sites = " + $trackCarrier + " (want exactly 4: declaration + definition + 2 声明臂)")
}
$trackCarrierFiles = @(FilesOf 'trackArrayCarrierForRelease(')
if ($trackCarrierFiles.Count -ne 3) {
    $fail += ("carrier-scope-exit registrar files = " + ($trackCarrierFiles -join ', ') + " (want exactly 3 files)")
}
# ㉖ 第一百五十六刀: 「出口这一发销毁到底发不发」只有那一枚新字段能答, 且答案只在过程入口算一次。
#    · freeCarrier = 表字段 + 声明默认值 + 定义形参 + 赋值 + 出口两读 + 那条注释 = 7 处 / 3 份文件;
#    · scanLocalArrayEscapes( = 声明 + 定义 + 三条过程入口 (Func / Proc / Prop) = 5 处 / 5 份文件
#      ⇒ 少一条入口 = 那一族的预扫没跑, 登记表会拿上一过程的名单决定这一发的销毁 = 悬空读。
$carrierFree = CountOf 'freeCarrier'
if ($carrierFree -ne 7) {
    $fail += ("carrier-free flag sites = " + $carrierFree + " (want exactly 7: 表字段 + 声明默认 + 定义形参 + 赋值 + 出口两读 + 注释)")
}
$carrierFreeFiles = @(FilesOf 'freeCarrier')
if ($carrierFreeFiles.Count -ne 3) {
    $fail += ("carrier-free flag files = " + ($carrierFreeFiles -join ', ') + " (want exactly 3 files)")
}
$escapeScan = CountOf 'scanLocalArrayEscapes('
if ($escapeScan -ne 5) {
    $fail += ("escape pre-pass sites = " + $escapeScan + " (want exactly 5: declaration + definition + 3 proc entries)")
}
$escapeScanFiles = @(FilesOf 'scanLocalArrayEscapes(')
if ($escapeScanFiles.Count -ne 5) {
    $fail += ("escape pre-pass files = " + ($escapeScanFiles -join ', ') + " (want exactly 5 files)")
}
#    · escapedLocalArrays_ = 「这枚载体在本过程里被整体用掉过没有」那张名单。
#      第一百五十八刀改的是它的含义 (整数组赋值的左值那一格从此豁免 —— 那一发现走复制)，
#      读它的人仍然只有一位 (出口那一发的 freeCarrier) ⇒ 字段 + 注释 + clear + count + insert
#      = 6 处 / 4 份文件；再冒出一处 count = 同一问被重述 (第一百五十九刀多出来的那一处 count
#      就是 dst 那一问的读者 —— 它问的必须是同一张名单, 不能另起一张)。
$escapeRoster = CountOf 'escapedLocalArrays_'
if ($escapeRoster -ne 6) {
    $fail += ("escape roster sites = " + $escapeRoster + " (want exactly 6: 字段 + 注释 + clear + 出口那一发的 count + 第一百五十九刀 dst 那一问的 count + insert)")
}
$escapeRosterFiles = @(FilesOf 'escapedLocalArrays_')
if ($escapeRosterFiles.Count -ne 4) {
    $fail += ("escape roster files = " + ($escapeRosterFiles -join ', ') + " (want exactly 4 files: cgen_state.inc + cgen_helpers.inc + cgen_iface_vtbl.cpp + cgen_util_type.cpp)")
}
#    · 第一百五十八刀: 整数组赋值左值那一格的豁免只许有一位作者，且必须住在 EscapeScan 里。
#      撤掉它 = 每一发整数组赋值净漏一枚载体 (实测把第一百五十六刀那把尺的 B225-take 从 0 抬到 1)；
#      多一处 = 逃逸那条口径被重述。
$lhsExempt = CountOf 'exempt.insert(node.target.get());'
if ($lhsExempt -ne 1) {
    $fail += ("whole-array-assign LHS exemption sites = " + $lhsExempt + " (want exactly 1, inside EscapeScan)")
}
$lhsExemptFiles = @(FilesOf 'exempt.insert(node.target.get());')
if ($lhsExemptFiles.Count -ne 1 -or $lhsExemptFiles[0] -ne 'cgen_iface_vtbl.cpp') {
    $fail += ("whole-array-assign LHS exemption file = " + ($lhsExemptFiles -join ', ') + " (want exactly cgen_iface_vtbl.cpp)")
}
#    · 第一百五十九刀: 同一格的第二半 —— 整数组赋值的**右值**也放行 (复制不吃源的载体)。
#      撤掉它 = 「ReDim 一枚局部数组、只把它复制给另一枚」那一格净漏一枚 (第一百五十六刀那把尺
#      的新针 B227, 拿第一百五十八刀的二进制编同一份夹具就读到 1)；多一处 = 口径被重述。
$rhsExempt = CountOf 'exempt.insert(node.value.get());'
if ($rhsExempt -ne 1) {
    $fail += ("whole-array-assign RHS exemption sites = " + $rhsExempt + " (want exactly 1, inside EscapeScan)")
}
$rhsExemptFiles = @(FilesOf 'exempt.insert(node.value.get());')
if ($rhsExemptFiles.Count -ne 1 -or $rhsExemptFiles[0] -ne 'cgen_iface_vtbl.cpp') {
    $fail += ("whole-array-assign RHS exemption file = " + ($rhsExemptFiles -join ', ') + " (want exactly cgen_iface_vtbl.cpp)")
}
#    · 第一百五十九刀: dst 那一问也只许有一位作者，且必须问的是**那两道现成的闸** (出口收尾同一对)。
#      另起一处 = 「谁持有旧载体」这条口径被重述; 撤掉 = 每一发整数组赋值把旧载体丢在原槽外面。
$dstAsk = CountOf 'exitVisibleArrayCarriers_.count(lower) && escapedLocalArrays_.count(lower) == 0'
if ($dstAsk -ne 1) {
    $fail += ("whole-array-assign dst gate sites = " + $dstAsk + " (want exactly 1, in wrapWholeArrayAssign)")
}
$dstAskFiles = @(FilesOf 'exitVisibleArrayCarriers_.count(lower) && escapedLocalArrays_.count(lower) == 0')
if ($dstAskFiles.Count -ne 1 -or $dstAskFiles[0] -ne 'cgen_util_type.cpp') {
    $fail += ("whole-array-assign dst gate file = " + ($dstAskFiles -join ', ') + " (want exactly cgen_util_type.cpp)")
}
#    · exitVisibleArrayCarriers_ (原名 topLevelLocalArrays_, 第一百六十四刀改口径又改名) = 出口那一发的
#      第二道闸: 「这枚载体的 C 声明在过程出口那一刻还在不在场」。提升开着 ⇒ 每一档局部数组都在场
#      (嵌在块里那几档的声明现在住在序言); C3_NO_HOIST ⇒ 缩回「只有过程体直接子句」那一个答案 (门 #821
#      的 FlexGridX86 十二发 C2065 就是靠这一问挡的)。少这一道闸 = 关掉提升时嵌套块里的名字又被出口放手。
$scopeRoster = CountOf 'exitVisibleArrayCarriers_'
if ($scopeRoster -ne 8) {
    $fail += ("declaration-scope roster sites = " + $scopeRoster + " (want exactly 8: 字段 + clear + 两条 insert (直接子句一趟 + 提升开着再递归一趟) + 出口那道 count + dst 那一问的 count + 递给 EscapeScan 的那一枚指针 + 注释)")
}
$scopeRosterFiles = @(FilesOf 'exitVisibleArrayCarriers_')
if ($scopeRosterFiles.Count -ne 3) {
    $fail += ("declaration-scope roster files = " + ($scopeRosterFiles -join ', ') + " (want exactly 3 files: cgen_state.inc + cgen_iface_vtbl.cpp + cgen_util_type.cpp)")
}
# ㉕ 第一百六十刀: 「这枚局部声明要不要 C 的 static 存储」收成一处 = wantsStaticLocalStorage，
#    数组载体声明那一发的形状收成一处 = emitArrayCarrierDecl。多一处 = 第二位作者又自己拼了一遍
#    (数组那三支从前就是这么掉掉这一档：只有标量那一路问判据，于是 Static 数组每调用换一块)。
$staticAsk = CountOf 'wantsStaticLocalStorage'
if ($staticAsk -ne 9) {
    $fail += ("static-storage question sites = " + $staticAsk + " (want exactly 9: 定义 + 声明 + 数组三支的调用 + 标量那一路的调用 + 第一百六十四刀序言那一发的调用 + 那两条注释)")
}
$staticAskFiles = @(FilesOf 'wantsStaticLocalStorage')
if ($staticAskFiles.Count -ne 2) {
    $fail += ("static-storage question files = " + ($staticAskFiles -join ', ') + " (want exactly 2 files: cgen_localdecl.cpp + cgen_helpers.inc)")
}
$carrierShape = CountOf 'emitArrayCarrierDecl'
if ($carrierShape -ne 5) {
    $fail += ("array-carrier decl shape sites = " + $carrierShape + " (want exactly 5: 定义 + 声明 + 三支站点)")
}
$carrierShapeFiles = @(FilesOf 'emitArrayCarrierDecl')
if ($carrierShapeFiles.Count -ne 2) {
    $fail += ("array-carrier decl shape files = " + ($carrierShapeFiles -join ', ') + " (want exactly 2 files)")
}
# ㉚ 第一百七十一刀 / 新账 #407: 「视图值怎么进 Variant 装箱站」收成一处。
#    病灶形状 = 视图 (vb6_iface_<I>) 是值结构体, 落 vb6_VariantFromValue 那张 _Generic 的 default
#    = vb6_VariantObject(指针) ⇒ cl C2172「实参不是指针」(实物 = 门禁在册的 tests/b189_gate,
#    B189Drv.c(11), 改前两台同红、门一直绿是因为那夹具只走发码针)。现在两位主人各只许一站:
#    · legacyViewBoxExpr = 发码侧那一句的唯一拼法 (声明 + 定义 + wrapVariantValue 与 boxToVariant
#      各一问 + 第一百七十二刀站E 一问 + 那一站自己的注释一行) = 6 处 / 4 份文件。
#    · vb6_VariantFromProjectInstance = RTL 侧那枚出口 (声明 + 定义 + 发码字符串 + 「交回整枚槽」
#      那张表那一行含注释共两处 + 第一百七十二刀两条注释) = 7 处 / 6 份文件。
#    再加一条反面的: .obj 那一截拼装**只许**住在 cgen_util_classtype.cpp —— 站点自己拼 = 第二位作者。
$viewBoxQ = CountOf 'legacyViewBoxExpr'
if ($viewBoxQ -ne 6) {
    $fail += ("legacy view boxing question sites = " + $viewBoxQ + " (want exactly 6: 声明 + 定义 + 三位调用者(含第一百七十二刀站E) + 那一站一条注释)")
}
$viewBoxQFiles = @(FilesOf 'legacyViewBoxExpr')
if ($viewBoxQFiles.Count -ne 4) {
    $fail += ("legacy view boxing question files = " + ($viewBoxQFiles -join ', ') + " (want exactly 3 files)")
}
$viewBoxRtl = CountOf 'vb6_VariantFromProjectInstance'
if ($viewBoxRtl -ne 7) {
    $fail += ("view boxing rtl 出口 sites = " + $viewBoxRtl + " (want exactly 7: RTL 声明 + 定义 + 发码串 + 槽表那一行两处 + 第一百七十二刀两条注释)")
}
$viewBoxRtlFiles = @(FilesOf 'vb6_VariantFromProjectInstance')
if ($viewBoxRtlFiles.Count -ne 6) {
    $fail += ("view boxing rtl 出口 files = " + ($viewBoxRtlFiles -join ', ') + " (want exactly 4 files)")
}
$ctrlTxt171 = [IO.File]::ReadAllText((Join-Path $Root "src\backend\cgen_util_ctrl.cpp"))
if ($ctrlTxt171.Contains('.obj)')) {
    $fail += 'cgen_util_ctrl.cpp 又自己拼 .obj) —— 视图装箱的拼装只许住在 cgen_util_classtype.cpp'
}

# ㉙ 第一百六十六刀 / 新账 #464: legacy VB6 接口 (.cls 宿主、没进 tB 登记表) 那张**槽表的形状**收成一处。
#    病灶形状 = `emitInterfaceVtable` 过去在**每枚实现类里**按自家成员声明顺序 + 自家 `classMeParam()`
#    现铸一份同名 `vb6_vtbl_<I>` / `vb6_iface_<I>` / `vb6_iface_<I>_wrap` —— 一枚接口被 N 枚类 Implements
#    就发 N 份 (真语料 Ray.vbp 实测 54 发 cl C2011、四份表里槽序两套; 新夹具 tests/iface_multi 在改前的
#    编译器上是 8 发 C2011 + 3 发 C2039 "Three 不是 vb6_vtbl_IMul 的成员")。现在四位主人各只许一站:
#    · emitLegacyIfaceContractTypedefs = 表类型的唯一发射口 (声明 + 定义 + 头文件开段那一发调用
#      + cgen_com.cpp 里那句指向主人的注释) = 4 处 / 4 份文件。
#    · legacyIfaceSlots = 槽名单的唯一读法 (声明 + 定义 + 表发射与实例发射各问一次) = 4 处 / 3 份文件。
#    · legacyIfaceWrapFn = wrap 名的唯一出口 (声明 + 定义 + Set 那三臂各一问 + 注释) = 6 处 / 4 份文件。
#    · setLegacyIfaceShapes = 形状表唯一的注入口 (cgen 侧 setter + driver 模块环那一发) = 2 处 / 2 份文件。
#    再加一条反面的: cgen_com.cpp 里**不许**再拼 typedef struct " + vtblName —— 那一发就是病灶本体,
#    谁把它写回去, 多实现类立刻回到 N 份同名表。
$legacyTblEmit = CountOf 'emitLegacyIfaceContractTypedefs'
if ($legacyTblEmit -ne 4) {
    $fail += ("legacy iface table emit sites = " + $legacyTblEmit + " (want exactly 4: 声明 + 定义 + 头文件调用 + 注释)")
}
$legacyTblEmitFiles = @(FilesOf 'emitLegacyIfaceContractTypedefs')
if ($legacyTblEmitFiles.Count -ne 4) {
    $fail += ("legacy iface table emit files = " + ($legacyTblEmitFiles -join ', ') + " (want exactly 4 files)")
}
$legacySlotRead = CountOf 'legacyIfaceSlots'
if ($legacySlotRead -ne 11) {
    $fail += ("legacy iface slot question sites = " + $legacySlotRead + " (want exactly 8: 声明 + 定义 + 两处消费 + 第一百七十四刀 view_from 那道闸 + 第一百七十八刀站F 属性读读者 + 站F 属性写读者 + 第一百八十六刀带参写槽查询 + 第一百八十九刀表达式接收者一问 + 第一百九十刀裸零参读一问)")
}
$legacySlotReadFiles = @(FilesOf 'legacyIfaceSlots')
if ($legacySlotReadFiles.Count -ne 7) {
    $fail += ("legacy iface slot question files = " + ($legacySlotReadFiles -join ', ') + " (want exactly 4 files)")
}
$legacyWrapMint = CountOf 'legacyIfaceWrapFn'
if ($legacyWrapMint -ne 10) {
    $fail += ("legacy iface wrap-name sites = " + $legacyWrapMint + " (want exactly 10: 声明 + 定义 + Set 三臂 + cgen_com 铸表一处 + 第一百七十二刀站B 一问 + 两条注释 + 第一百七十四刀铸 view_from 那一问)")
}
$legacyWrapFiles = @(FilesOf 'legacyIfaceWrapFn')
if ($legacyWrapFiles.Count -ne 6) {
    $fail += ("legacy iface wrap-name files = " + ($legacyWrapFiles -join ', ') + " (want exactly 4 files)")
}
$legacyShapeInject = CountOf 'setLegacyIfaceShapes'
if ($legacyShapeInject -ne 2) {
    $fail += ("legacy iface shape injection sites = " + $legacyShapeInject + " (want exactly 2: setter + driver 注入)")
}
$legacyShapeInjectFiles = @(FilesOf 'setLegacyIfaceShapes')
if ($legacyShapeInjectFiles.Count -ne 2) {
    $fail += ("legacy iface shape injection files = " + ($legacyShapeInjectFiles -join ', ') + " (want exactly 2 files)")
}
$cgenComText = [IO.File]::ReadAllText((Join-Path $Root "src\backend\module\cgen_com.cpp"))
if ($cgenComText.Contains('typedef struct " + vtblName')) {
    $fail += "cgen_com.cpp 又自己铸 vb6_vtbl_ 表类型 (新账 #464 的病灶形状) - 表类型只许 emitLegacyIfaceContractTypedefs 一处发"
}
# ㉘ 第一百六十五刀: 「这一格槽装的是不是 legacy 那一档接口视图」那一问收成一处 = slotIsLegacyIfaceView。
#    病灶形状 = 腿②/腿③ 那两站先问类名 (inferClassTypeOfExpr 把接口名也答出来)，于是铸出一枚声明侧从没造过的
#    `vb6_cls_<接口名>*` = cl C2440 (#407 那十二形全撞在这一发上，且门内没有一份输入走到这一形)。
#    只许一份声明 + 一份定义 + 四处调用 (= 6 处 / 5 份文件)：多一处 = 又有站点自己拼名单，少一处 = 有一站还没问它。
#    (第二百四十八刀: 「Is」 那一站原本问这一位 bool 投影, 但同名接口有两套 (legacy 值视图 / 新式薄指针), 名册分不开
#     ⇒ 那一站改成先问带名字的 slotLegacyIfaceViewName 再用 ivrefCType 问形状, 这一处搬到下面的 ㉛ 去数。)
#    ⚠ 这一问不许被并进 ifaceViewOfExpr: 那一问问的是**新式薄指针** (vb6_ivref_<I>*)，两档的主人与形状都不同，
#    合一起会让 b448 / iface_wrap 那两片真跑夹具当场换形状。
$viewSlotAsk = CountOf 'slotIsLegacyIfaceView'
if ($viewSlotAsk -ne 6) {
    $fail += ("iface-view-slot question sites = " + $viewSlotAsk + " (want exactly 6: 声明 + 定义 + 腿② + 腿③ + 装箱站那一问 + 第一百七十五刀表达式路那一问)")
}
$viewSlotAskFiles = @(FilesOf 'slotIsLegacyIfaceView')
if ($viewSlotAskFiles.Count -ne 5) {
    $fail += ("iface-view-slot question files = " + ($viewSlotAskFiles -join ', ') + " (want exactly 5 files)")
}
# ㉛ 第一百七十四刀 (账 #471 站A): 「集合交回来的那一枚元素是哪一枚实现类」那一问收成一处 ——
#    每枚 legacy 接口在自己那份 TU 里铸一枚 view_from, 按运行期类名选中在册的那一对 wrap。
#    四问各一位主人 + 一把尺只许一处 (多一处 = 又有人在站点自己拼名字或自己比名):
#    · slotLegacyIfaceViewName = 「这一格槽装的是视图, 装的是哪一枚」= 声明 + 定义 + 它自己的 bool 投影
#      (slotIsLegacyIfaceView 从此是这张名册的投影, 不再第二次读 knownIfaceVars_) + For Each 那一站
#      + 第一百七十五刀那位「视图方法调用」主人自己那一问 + 第一百七十七刀 With 那一站
#      + 第一百七十八刀站F 属性读那一位读者 + 新账 #474 写侧那一位读者 = 8 处 / 4 份文件。
#      第一百八十九刀表达式接收者那一问 +1; 第二百一十五刀把 Set 装填那一站的直查名册换成问这一位
#      主人 +1 ⇒ 13 处 / 8 份文件; 第二百四十八刀把 「Is」 那一站的 bool 投影换成这一位主人 +1
#      ⇒ 今天 14 处 / 9 份文件。
#    · legacyIfaceViewFromFn = view_from 那个名字的唯一出口 = 声明 + 定义 + 铸那一发 + 站点 + 一条注释 = 5 处 / 4 份文件。
#    · emitLegacyIfaceViewFrom = 那一发的发码主人 (.h 交原型 / .c 交定义) = 声明 + 定义 + 调用 = 3 处 / 3 份文件,
#      且调用只许挂在接口宿主那一站 (body_pass) —— 别处再挂一处就是第二个 mint。
#      (比名那把尺 _stricmp 只许一处的针在 check_legacy_view_arg_owners.ps1 的 V8, 与 V1 那条同规, 不在这里重复钉)
$viewNameAsk = CountOf 'slotLegacyIfaceViewName'
if ($viewNameAsk -ne 14) {
    $fail += ("iface-view-name question sites = " + $viewNameAsk + " (want exactly 14: 声明 + 定义 + bool 投影 + For Each 那一站 + 视图调用主人自己那一问 + With 那一站 + 站F 读侧读者 + 站F 写侧读者 + 第一百八十六刀带参写槽查询 + 第一百八十九刀表达式接收者一问 + 第二百一十五刀 Set 装填那一站与其定义内两问 + 第二百四十八刀 「Is」 那一站)")
}
$viewNameAskFiles = @(FilesOf 'slotLegacyIfaceViewName')
if ($viewNameAskFiles.Count -ne 9) {
    $fail += ("iface-view-name question files = " + ($viewNameAskFiles -join ', ') + " (want exactly 9 files)")
}
$viewFromName = CountOf 'legacyIfaceViewFromFn'
if ($viewFromName -ne 5) {
    $fail += ("view_from name-owner sites = " + $viewFromName + " (want exactly 5: 声明 + 定义 + 铸那一发 + 站点 + 一条注释)")
}
$viewFromEmit = CountOf 'emitLegacyIfaceViewFrom'
if ($viewFromEmit -ne 3) {
    $fail += ("view_from mint sites = " + $viewFromEmit + " (want exactly 3: 声明 + 定义 + 接口宿主那一站)")
}
$viewFromEmitFiles = @(FilesOf 'emitLegacyIfaceViewFrom')
if ($viewFromEmitFiles -notcontains 'cgen_base_generate_body_pass.inc') {
    $fail += ("view_from 的调用不在接口宿主那一站 (body_pass) - 现在挂在: " + ($viewFromEmitFiles -join ', '))
}
# ㉗ 第一百六十四刀: 「嵌在块里的定长/多维数组载体 —— 声明提到过程序言」这一族五问各只许一位主人:
#    · hoistingLocalDeclsEnabled = 「提升开着没有」(早退那一发与闸① 同一位主人; C3_NO_HOIST 关掉时闸①
#      必须跟着答否, 否则门 #821 那一形立刻回来) = 声明 + 定义 + 两处调用 = 4 处 / 3 份文件。
#    · preludeArrayCarriers_ = 序言发过声明的那几张名 (主人写、原位那一发读) = 5 处 / 4 份文件。
#    · hoistFixedArrayDeclToPrelude = 「把这枚声明提到序言」那一小步 (发一行 + 写名单) = 4 处 / 2 份文件。
#    · arrayCarrierCType = 「哪一档描述符指针」(序言那一发与原位那三支同一个答案) = 6 处 / 2 份文件。
#    · emitArrayCarrierPreludeDecl = 序言那一行的形状 = 3 处 / 3 份文件; emitArrayCarrierGuardedCreate =
#      「拿指针本身当首建守卫」那一块 (Static 档与序言档共用) = 4 处 / 2 份文件。
#    · collectArrayCarrierNames = 「这条声明是不是数组载体」+ 收名 = 4 处 / 1 份文件 (只住在预扫那一趟)。
$hoistAsk = CountOf 'hoistingLocalDeclsEnabled'
if ($hoistAsk -ne 4) {
    $fail += ("hoisting-enabled question sites = " + $hoistAsk + " (want exactly 4: 声明 + 定义 + 早退那一发 + 闸① 那一问)")
}
$hoistAskFiles = @(FilesOf 'hoistingLocalDeclsEnabled')
if ($hoistAskFiles.Count -ne 3) {
    $fail += ("hoisting-enabled question files = " + ($hoistAskFiles -join ', ') + " (want exactly 3 files)")
}
$preludeRoster = CountOf 'preludeArrayCarriers_'
if ($preludeRoster -ne 5) {
    $fail += ("prelude-carrier roster sites = " + $preludeRoster + " (want exactly 5: 字段 + clear + insert + 原位那一发的 count + 那条注释)")
}
$preludeRosterFiles = @(FilesOf 'preludeArrayCarriers_')
if ($preludeRosterFiles.Count -ne 4) {
    $fail += ("prelude-carrier roster files = " + ($preludeRosterFiles -join ', ') + " (want exactly 4 files)")
}
$hoistStep = CountOf 'hoistFixedArrayDeclToPrelude'
if ($hoistStep -ne 4) {
    $fail += ("decl-to-prelude step sites = " + $hoistStep + " (want exactly 4: 声明 + 定义 + VariableDecl 那一支 + MultiDecl 那一支)")
}
$hoistStepFiles = @(FilesOf 'hoistFixedArrayDeclToPrelude')
if ($hoistStepFiles.Count -ne 2) {
    $fail += ("decl-to-prelude step files = " + ($hoistStepFiles -join ', ') + " (want exactly 2 files)")
}
$carrierKind = CountOf 'arrayCarrierCType'
if ($carrierKind -ne 6) {
    $fail += ("carrier descriptor-kind question sites = " + $carrierKind + " (want exactly 6: 声明 + 定义 + 原位三支 + 序言那一发)")
}
$carrierKindFiles = @(FilesOf 'arrayCarrierCType')
if ($carrierKindFiles.Count -ne 2) {
    $fail += ("carrier descriptor-kind question files = " + ($carrierKindFiles -join ', ') + " (want exactly 2 files)")
}
$preludeShape = CountOf 'emitArrayCarrierPreludeDecl'
if ($preludeShape -ne 3) {
    $fail += ("prelude decl shape sites = " + $preludeShape + " (want exactly 3: 声明 + 定义 + 那一小步的调用)")
}
$guardShape = CountOf 'emitArrayCarrierGuardedCreate'
if ($guardShape -ne 4) {
    $fail += ("first-build guard shape sites = " + $guardShape + " (want exactly 4: 声明 + 定义 + Static 那一档 + 序言那一档)")
}
$guardShapeFiles = @(FilesOf 'emitArrayCarrierGuardedCreate')
if ($guardShapeFiles.Count -ne 2) {
    $fail += ("first-build guard shape files = " + ($guardShapeFiles -join ', ') + " (want exactly 2 files)")
}
$carrierNames = CountOf 'collectArrayCarrierNames'
if ($carrierNames -ne 4) {
    $fail += ("carrier-name collector sites = " + $carrierNames + " (want exactly 4: 定义 + 逗号声明递归 + 两趟调用)")
}
$carrierNamesFiles = @(FilesOf 'collectArrayCarrierNames')
if ($carrierNamesFiles.Count -ne 1 -or $carrierNamesFiles[0] -ne 'cgen_iface_vtbl.cpp') {
    $fail += ("carrier-name collector files = " + ($carrierNamesFiles -join ', ') + " (want exactly cgen_iface_vtbl.cpp)")
}
# ㉖ 第一百六十一刀: 「Static 后面跟的是不是过程头」那一问收成一处 = tryParseStaticProcDecl。
#    多一处 = 又有第二位作者自己拼了一遍前瞻形状。这一族从前是**两个分派点各自**把 TokenKind::Static
#    当成变量声明关键字吃（裸写那一形走 :75 那条、带访问修饰符那一形落到 default 那条），于是真语料里
#    10 处 `Private Static Sub` 压根进不了发码阶段（VB2001 expected variable name），后端 inStaticProc_
#    那一腿从头到尾没有读者。只许有一份声明 + 一份定义 + 两个分派点，且两处都得走同一位主人。
$staticProcAsk = CountOf 'tryParseStaticProcDecl'
if ($staticProcAsk -ne 4) {
    $fail += ("static-proc-header question sites = " + $staticProcAsk + " (want exactly 4: 定义 + 声明 + 裸写那一形 + 带访问修饰符那一形)")
}
$staticProcAskFiles = @(FilesOf 'tryParseStaticProcDecl')
if ($staticProcAskFiles.Count -ne 2) {
    $fail += ("static-proc-header question files = " + ($staticProcAskFiles -join ", ") + " (want exactly 2 files: parser.hpp + parser_decl.cpp)")
}

# ㉔ 第一百五十二刀: 那张带档的表 (+第三档) 的形状 —— isCarrier 只许住在表与出口那一趟 (2 份文件 / 4 处)。
$carrierFlag = CountOf 'isCarrier'
if ($carrierFlag -ne 4) {
    $fail += ("carrier-slot flag sites = " + $carrierFlag + " (want exactly 4: 表字段 + 赋值 + 出口判定 + 那条注释)")
}
$carrierFlagFiles = @(FilesOf 'isCarrier')
if ($carrierFlagFiles.Count -ne 2) {
    $fail += ("carrier-slot flag files = " + ($carrierFlagFiles -join ', ') + " (want exactly 2 files: cgen_state.inc + cgen_iface_vtbl.cpp)")
}

# 25 第一百五十四刀: 多维 Preserve 收缩丢尾那位带回调的本体 —— 老名字只剩一句转调,
#    两个名字都要数得着: 本体的声明 + 定义 + 转调里那一发 + 后端发码那一行。
$ndTrExport = CountOf 'vb6_SafeArrayReDimPreserveND_Tr('
if ($ndTrExport -ne 4) {
    $fail += ("ND Preserve-with-release export sites = " + $ndTrExport + " (want exactly 4: RTL decl + def + forwarder call + backend emit)")
}
$ndTrFiles = @(FilesOf 'vb6_SafeArrayReDimPreserveND_Tr(')
if ($ndTrFiles.Count -ne 3) {
    $fail += ("ND Preserve-with-release export files = " + ($ndTrFiles -join ", ") + " (want exactly 3 files: .h + .c + cgen_redim.cpp)")
}
# ㉜ 第三百十四刀 (腿②″ = 泛对象档的**返回槽**从没主人): 这一档的三条问句与那一发形状各只许一位主人
#    (钉法镜像 ⑨ = 出参槽那一档)。病灶形状 = ComCtlsDemo Builds/ImageList/ImageList.ctl:562
#    `Set FListImageExtractIcon = UserControl.Picture` 发成裸拷贝 = 交出去一枚借用, 原主紧接着 563
#    `Set UserControl.Picture = Nothing` 把唯一那一枚退掉 ⇒ 调用方接手的是已回收的 IPictureDisp,
#    运行期一点 Command2 就 0xC0000005 (写进 ole32 的 GetIDsOfNames)。
#    · emitLateSlotRefClaim = 那一发 claim 的唯一发码处: 声明 + 定义 + 站点 = 3 处 / 3 份文件,
#      且站点必须住在 Set 的 RHS 那一档 (cgen_setlet_set_rhs.inc) —— 别处再站一位主人 = 第二套判据。
#    · 注释串「腿②″: Set 泛对象槽」全仓恰好 1 处且住在 cgen_com.cpp ⇒ 没有站点自己补一行 AddRef。
#    · ctypeIsObjectSlotPtr = 「这一档 C 存储是不是对象指针槽」: 声明 + 定义 + 两位消费者 = 4 处 / 4 份文件
#      (第三百十四刀把 cgen_expr_call_com_bind.inc 里那枚站点级 lambda 改成转调, 从此两站同一个答案)。
#    · localSlotHandsOverRef = 「右值那一格是不是本过程的局部(它把唯一那一枚按裸拷贝移交)」= 3 处 / 3 份文件。
#      这一问刻意不并进 ⑬ 的 rhsTransfersProjectRef: 那条答「被调方自带对价交回来」, 这条答「C3 不在过程
#      出口放手局部对象格」(§B164 那族在册) —— 同一个「不 claim」的结论, 两条不同的口径, 不许合成一条。
$lateClaimHelp = CountOf 'emitLateSlotRefClaim('
if ($lateClaimHelp -ne 3) {
    $fail += ("late-bound return-slot claim sites = " + $lateClaimHelp + " (want exactly 3: declaration + definition + 站点)")
}
$lateClaimHelpFiles = @(FilesOf 'emitLateSlotRefClaim(')
if ($lateClaimHelpFiles.Count -ne 3 -or -not ($lateClaimHelpFiles -contains 'cgen_setlet_set_rhs.inc')) {
    $fail += ("late-bound return-slot claim files = " + ($lateClaimHelpFiles -join ', ') + " (want exactly 3 files incl. cgen_setlet_set_rhs.inc)")
}
$lateClaimShape = CountOf '腿②″: Set 泛对象槽'
if ($lateClaimShape -ne 1) {
    $fail += ("late-bound return-slot claim shape sites = " + $lateClaimShape + " (want exactly 1, inside emitLateSlotRefClaim)")
}
$lateClaimShapeFiles = @(FilesOf '腿②″: Set 泛对象槽')
if ($lateClaimShapeFiles.Count -ne 1 -or $lateClaimShapeFiles[0] -ne 'cgen_com.cpp') {
    $fail += ("late-bound return-slot claim shape files = " + ($lateClaimShapeFiles -join ', ') + " (want exactly cgen_com.cpp)")
}
$slotKindAsk = CountOf 'ctypeIsObjectSlotPtr('
if ($slotKindAsk -ne 4) {
    $fail += ("object-pointer-slot question sites = " + $slotKindAsk + " (want exactly 4: declaration + definition + 两位消费者)")
}
$slotKindAskFiles = @(FilesOf 'ctypeIsObjectSlotPtr(')
if ($slotKindAskFiles.Count -ne 4) {
    $fail += ("object-pointer-slot question files = " + ($slotKindAskFiles -join ', ') + " (want exactly 4 files)")
}
$localHandAsk = CountOf 'localSlotHandsOverRef('
if ($localHandAsk -ne 3) {
    $fail += ("local-handover question sites = " + $localHandAsk + " (want exactly 3: declaration + definition + 站点)")
}
$localHandAskFiles = @(FilesOf 'localSlotHandsOverRef(')
if ($localHandAskFiles.Count -ne 3) {
    $fail += ("local-handover question files = " + ($localHandAskFiles -join ', ') + " (want exactly 3 files)")
}
if ($fail.Count -gt 0) {
    Write-Host "CLASS-REFCOUNT-AUTHORITY FAIL"
    foreach ($f in $fail) { Write-Host ("  - " + $f) }
    exit 1
}
Write-Host "CLASS-REFCOUNT-AUTHORITY OK (1 pair mint / 1 header / 2 leg consumers / 3+9 formals / 10 roster files / 6 literal files / 0 stale / 5 registry files + 2 registrants / 1+1 out-slot / 1+1 foreign-write drop / 1 field-teardown shape + file / 6 field-type questions (标量 + 数组两档) / 7 transfer-question sites in 2 files (第二百零三刀给字段槽那一臂加了第 4 处站点问句) / 1 view-leg hand-back + 2 QiSet calls / 1+2+1 transfer-arm questions / 3 static-gate sites in 1 file + 1 ivref registration / 1 arg-station hand-back + 4+4 helper sites in 4 files / 4+4 index-slot + project-class questions / 1 address-temp mint / 3 element-teardown sites + 4+4 carrier-destroy questions + 4 preserve-with-release sites in 4 files / 4+4 形A questions (out-slot class + owner pass) + 5 copy-in claim sites in 4 files / 3+4 carrier-destroy sites in 3 files / 8 hand-back stations (5 in cgen_redim.cpp + 过程出口) + 4 shape questions + 3 multidim export sites + 4+2 carrier-scope registrar + flag sites / 7+3 carrier-free flag + 5+5 escape pre-pass + 6 roster sites in 4 files + 1 LHS + 1 RHS exemption + 1 dst gate + 7+3 declaration-scope roster / 5 static gates / 4+3 ND Preserve-with-release export sites / 6 legacyViewBoxExpr sites in 4 files + 7 rtl-name sites in 6 files + 9 wrap-name sites in 6 files + 6 view-slot sites in 5 files (㉘㉚，第一百七十二刀三站) + 14 view-name sites in 9 files + 5 view_from-name sites + 3 mint sites + 1 name-ruler (㉛，第一百七十四刀站A) / 3+3 late-return-slot claim sites (站点在 cgen_setlet_set_rhs.inc) + 1 shape in cgen_com.cpp + 4+4 object-pointer-slot questions + 3+3 local-handover questions (㉜，第三百十四刀)"
exit 0
