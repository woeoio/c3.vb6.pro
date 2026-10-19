# check_class_refcount_authority.ps1 - 账 #305 §B218 (甲)：工程类实例的引用计数只许有一位主人。
# 十四条钉（每条都能单独红，红了就说明同一问长出了第二位主人）：
#   ① 出入口那一对的 mint 各恰好 1 处，且住在 module/cgen_com.cpp（每条腿只许转调它）。
#   ② 计数头 `int32_t __refcount;` 的 mint 恰好 1 处，且住在 cgen_iface_vtbl.cpp 的
#      emitIfaceClassFields 里 —— 它必须紧跟 __comObj 那一句（ai/022 D19：偏移 0 恒是 __comObj）。
#   ③ 覆写/置空那一位主人 emitClassSlotRefSet 的消费文件恰好 2 份（Set 的 RHS 臂 + 置空臂）；
#      第 3 份只能由腿④/腿⑤ 那两刀带进来，届时这里跟着改数。
#   ④ 工程类形参登记收成一处出口 registerClassFormal：消费文件恰好 3 份（Sub / Function /
#      Property 三个入口），调用点恰好 9 处（每份三条：跨模块兜底 / Class 符号 / ComClass 同名）。
#   ⑤ procClassFormals_ 那张名册只许住在 7 份文件里（表本体 + 三个入口清表 + 出口本体 +
#      两条腿各问一次）。
#   ⑥ 字面量 __refcount 的主人名单 = 在册那 6 份；RTL 那一处必须是 vb6comserver_obj.c 的注释。
#   ⑦ 陈旧注释「没有引用计数」（具体类那三处）不许活着 —— 甲落地之后这句话是错的。
#   ⑧ 过程出口那张带档的登记表 slotsToRelease_ 只许住在 5 份文件里（表本体 + 三个过程入口
#   ⑨ 出参槽那一档 (腿②'): claim-only 的发码形状与站点各恰好一处。
#   ⑩ 腿④ 的第二道夹: 撤登记的主人一处 + 站点一处 (槽地址交给外部 Declare 的那一格出口不放手)。
#   ⑪ 腿⑤ 出厂那一档: `_Release(me->` 全仓恰好一处，且主人文件必须还是 cgen_com.cpp。
#   ⑫ 字段类型那一问 `mapTypeRef(var.asType.get())` 恰好 5 处（新增一处手抄就得改数）。
#   ⑬ 移交值那一问 `rhsTransfersProjectRef(` 恰好 3 处（声明 + 定义 + 一个调用点），调用点必须在
#      cgen_setlet_set_rhs.inc —— 「右侧带不带对价」与 `New` 共用同一条不 claim 的臂，别处再问就分家。
#      清表 + 主人本体），工程类那一档的注册口 trackClassSlotForRelease 恰好 2 处
#      （Dim As <工程类> 的两条注册路：Class 符号 / projectClassNameOf 兜底）。
#   ⑭ 移交值那一问的三形只住一个函数: 成员返回类那一问 
#      恰好 1 处、 基名那一问  恰好 2 处 (含本刀那一臂,
#      另一处是 myBaseClassOf)、接口视图那道排除  恰好 1 处。
# 退出码：0 = 十四条都过；1 = 任一条红；2 = 找不到 src/backend。
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
if ($cfFiles.Count -ne 7) {
    $fail += ("procClassFormals_ files = " + $cfFiles.Count +
              " (want exactly 7): " + ($cfFiles -join ', '))
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
if ($fieldQ -ne 5) {
    $fail += ("mapTypeRef(var.asType.get()) question sites = " + $fieldQ + " (want exactly 5)")
}


# ⑬ 移交值那一问的名单: 「右侧那一枚是不是已经带对价交过来了」全仓只许问一处 (声明 + 定义 + 一个调用点 = 3 处
#    `rhsTransfersProjectRef(`)，而且调用点必须住在 Set 的 RHS 那一档 (cgen_setlet_set_rhs.inc)。为什么钉: 这一问与
#    `New` 共用同一条「不 claim」的臂 —— 谁在别的站点再问一遍，「实参 / 返回值 / 字段」这三条边的口径就分家了。
$xferQ = CountOf 'rhsTransfersProjectRef('
if ($xferQ -ne 3) {
    $fail += ("transfer-question sites = " + $xferQ + " (want exactly 3: decl + def + 1 call)")
}
$xferFiles = @(FilesOf 'rhsTransfersProjectRef(*node.value)')
if ($xferFiles.Count -ne 1 -or $xferFiles[0] -ne 'cgen_setlet_set_rhs.inc') {
    $fail += ("transfer-question call file = " + ($xferFiles -join ', ') + " (want exactly cgen_setlet_set_rhs.inc)")
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
if ($fail.Count -gt 0) {
    Write-Host "CLASS-REFCOUNT-AUTHORITY FAIL"
    foreach ($f in $fail) { Write-Host ("  - " + $f) }
    exit 1
}
Write-Host "CLASS-REFCOUNT-AUTHORITY OK (1 pair mint / 1 header / 2 leg consumers / 3+9 formals / 7 roster / 6 literal files / 0 stale / 5 registry files + 2 registrants / 1+1 out-slot / 1+1 foreign-write drop / 1 field-teardown shape + file / 5 field-type questions / 3 transfer-question sites / 1+2+1 transfer-arm questions)"
exit 0
