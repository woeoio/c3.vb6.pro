# check_variant_slot_store.ps1 - 账 #305 §B164 (第五十五刀): 「把一枚值收进 Variant 槽」只许住一枚出口
#
# 为什么要这一道: 这个决定在 src\backend 里原本有两份答案 —— 普通赋值那一站发的是
# 「先清目标旧的持有, 再把右值搬进去」两行, 而 For Each 的循环变量只发了第二行(裸结构赋值)。
# 后果 = 每遍历一个元素就把上一轮那只 BSTR 忘了(实测 BASE 每元素漏 82/71 字节, 台账 §B164 形④),
# 与 §B163「一站两成因」、§B157「按名字当键必须带全身份」是同一族第三次撞。
#
# 第五十五刀把它收成 CCodeGen::emitVariantSlotStore 一枚, 并让它答**两半句**:
#   ① 放掉槽里原来那份       —— 不放就没人回收(本刀那 82/71 B/元素)
#   ② 槽对新值必须自己持有    —— 照抄借来的指针, 别人一释放这槽就悬垂; 而对借来的值补 ①,
#                              就等于 free 掉别人的内存(实测 arr-sum 820 变 813/815)
# rhsOwnsValue 答 ②: 集合/ParamArray 那三站的右值是刚 mint 的带权值(清+搬), 数组那一站交的是
# 数组自己那只 BSTR(清+深拷贝 = vb6_VariantAssign)。
#
# 七条规则 (R1-R4 = 第五十五刀 §B164; R5 = 第五十九刀 §B168 + 第六十刀 §B169 改问同一张表;
#          R6-R7 = 第六十刀 §B169: 两枚槽位写入出口各一枚, 手抄 BSTR_Assign 的名单钉死)
#   R1  出口存在、本体就是那三形(借用->一次调用 / 自指->临时量+清+搬 / 直发->清+搬):
#       恰好 6 枚 emitLine、含「右值是借来的」那一路、体里不许 switch
#   R2  声明 + 定义各恰好一处, 且调用点 >= $CallFloor (少了 = 有人又开始自己抄)
#   R3  「清+搬」那一对 emitLine 的手抄形状只许出现在**定义出口的这份文件**里, 且恰好两形
#       (两行谁先谁后也由这条钉住: 先清后搬)
#   R4  还剩几处在裸发槽位赋值 —— 名单钉死, 新增即红; 真修掉一处要把名单同步改短(§B164 形⑤起)
param(
    [string]$Root = "",
    [int]$CallFloor = 7
)
$ErrorActionPreference = 'Continue'
if (-not $Root) { $Root = (Split-Path -Parent $PSScriptRoot) }
$backend = Join-Path $Root 'src\backend'
if (-not (Test-Path $backend)) { Write-Host ('FAIL 没有 ' + $backend); exit 2 }

# 本刀刻意没接线的三站(每一站都得先问「右值是不是借来的」, 不能盲改) + 已接线那一站的归零记录
$OPEN = @(
    # 事件实参: pName = vb6_VariantFromStackVARIANT(&args[i])
    @{ File = 'cgen_base_generate_evt_impl.inc'; N = 1 },
    # FnRetUdt 的 Variant 成员那一站本刀已接线(站⑥, 发码仍是「清+搬」) —— 裸发归零;
    # 但「这里的右值是不是借来的」按 Fix 232 那一问还没拍, 记在台账 §B164 形⑥
    @{ File = 'cgen_assign_stmt_special.inc'; N = 0 },
    # 这一份里的三处是 vb6_VariantTo*(从槽里读出, 不是写入) —— R4 的 (?!To) 已排除, 故 N = 0
    @{ File = 'cgen_assign_value_sem.inc'; N = 0 },
    # For 循环起点: varAcc = vb6_VariantFromValue(start)
    @{ File = 'cgen_control.cpp'; N = 1 },
    # Property Let 进槽: target = vb6_VariantFromValue(value)
    @{ File = 'cgen_setlet.cpp'; N = 1 }
)

$files = @(Get-ChildItem -Path $backend -Recurse -File -Include *.cpp, *.hpp, *.inc)
$bad = @()
$defFile = ''
$defBody = ''
$declCount = 0
$totalRefs = 0
$handPairs = @()
$bareByFile = @{}

$DEF_RE   = '(?m)^\s*void\s+CCodeGen::emitVariantSlotStore\s*\('
$DECL_RE  = '^\s*void\s+emitVariantSlotStore\s*\('
$PAIR_RE  = 'c_\.emitLine\("\s*vb6_VariantClear\(&" \+ (\w+) \+ "\);"\)[\s;]+c_\.emitLine\((?:"\s*" \+ )?\1 \+ "'
$BARE_RE  = 'c_\.emitLine\(\w+ \+ " = vb6_Variant(?!To)'

foreach ($f in $files) {
    $t = [System.IO.File]::ReadAllText($f.FullName)
    $refs = [regex]::Matches($t, 'emitVariantSlotStore')
    foreach ($p in [regex]::Matches($t, $PAIR_RE)) { $handPairs += $f.Name }
    foreach ($b in [regex]::Matches($t, $BARE_RE)) {
        $k = Split-Path $f.FullName -Leaf
        if (-not $bareByFile.ContainsKey($k)) { $bareByFile[$k] = 0 }
        $bareByFile[$k] = $bareByFile[$k] + 1
    }
    if ($refs.Count -eq 0) { continue }
    $totalRefs += $refs.Count
    if ([regex]::IsMatch($t, $DECL_RE, 'Multiline')) { $declCount++ }
    $dm = [regex]::Matches($t, $DEF_RE)
    if ($dm.Count -gt 1) { $bad += ('R1 ' + $f.Name + ': 出口定义了 ' + $dm.Count + ' 遍') }
    if ($dm.Count -eq 1) {
        if ($defFile -ne '') { $bad += ('R1 出口定义落在两份文件: ' + (Split-Path $defFile -Leaf) + ' + ' + $f.Name) }
        $defFile = $f.FullName
        $i = $t.IndexOf($dm[0].Value)
        $j = $t.IndexOf([char]10 + '}', $i)
        $defBody = $t.Substring($i, [Math]::Max(0, $j - $i))
    }
}

if ($defFile -eq '') { $bad += 'R1 src\backend 里找不到 emitVariantSlotStore 的定义 —— 那一问又没人管了' }
if ($declCount -ne 1) { $bad += ('R2 声明 found ' + $declCount + ' 处 (要恰好 1, 在 detail\util\cgen_helpers.inc)') }

$callSites = $totalRefs - $declCount - 1          # 减去声明与定义那两处
if ($callSites -lt $CallFloor) {
    $bad += ('R2 调用点 ' + $callSites + ' < ' + $CallFloor + ' —— 站点该问这一枚出口, 不该自己抄')
}
if ($defBody -ne '') {
    $n = ([regex]::Matches($defBody, 'c_\.emitLine')).Count
    if ($n -ne 6) { $bad += ('R1 出口本体有 ' + $n + ' 枚 emitLine (要恰好 6: 借用一路 1 + 自指一路 3 + 直发一路 2)') }
    if ($defBody.IndexOf('vb6_VariantAssign(&') -lt 0) { $bad += 'R1 出口缺了「右值是借来的」那一路 (vb6_VariantAssign)' }
    if ($defBody -match 'switch\s*\(') { $bad += 'R1 出口体里出现 switch —— 槽位写入不该再分档' }
}
$wantFile = Split-Path $defFile -Leaf
foreach ($g in ($handPairs | Group-Object)) {
    if ($g.Name -ne $wantFile -or $g.Count -ne 2) {
        $bad += ('R3 手抄「清+搬」出现在 ' + $g.Name + ' x' + $g.Count + ' (只许住 ' + $wantFile + ' x2: 直发 + 自指)')
    }
}
$bareTotal = 0
foreach ($v in $bareByFile.Values) { $bareTotal += $v }
foreach ($o in $OPEN) {
    $got = 0
    if ($bareByFile.ContainsKey($o.File)) { $got = $bareByFile[$o.File] }
    if ($got -ne $o.N) {
        $bad += ('R4 裸发槽位赋值 ' + $o.File + ' = ' + $got + ' 处, 名单钉的是 ' + $o.N + ' (新增即红; 修掉了就把名单改短)')
    }
}
foreach ($k in $bareByFile.Keys) {
    if (@($OPEN | Where-Object { $_.File -eq $k }).Count -eq 0) {
        $bad += ('R4 名单外的新站点 ' + $k + ' x' + $bareByFile[$k] + ' —— 那一问有人又开始自己抄了')
    }
}

# ---- R5 (账 #305 §B168/§B169): 「`<对象>.<字段>` 这一枚左值是哪一档」—— 只许一张表一处问 ----
# 第五十九刀把「是不是 Variant 槽」收进 driver 注入的工程级表 + 一处查询出口;
# 第六十刀 (§B169) 把同一趟预扫描的那张表**扩成答声明档** (Variant / String / 其它标量),
# 于是三站都问它: 赋值站③ (BSTR 档, 修形② 的裸指针别名)、赋值站④ (Variant 档)、
# With 块那一站 (修形①: 查不到 Property Let/Set 就凭空拼 vb6_<Cls>_prop_let_<F> ⇒ LNK2019)。
#   R5a 查询出口 classFieldKind 恰好定义 1 遍且住 cgen_api.inc; 三个布尔投影
#       (VariantSlot / BstrSlot / ScalarData) 各 1 遍, 且投影里只许**转问** classFieldKind
#       (在投影里再抄一遍 find = 两份答案)
#   R5b 直接问 classFieldKind 的站点 >= 1 (With 块那一站) 且三个投影的站点 >= 2
#       (赋值站③ String 档 + 赋值站④ Variant 档) —— 归零 = 那一问又没人管了
#   R5c 那张表只在 driver 的一趟预扫描里建 (恰好 1 份文件
#       driver_codegen_fieldkind_scan.inc), 注入站点恰好 1 处
#   R5d 「认文本尾注」这条歧路不许复活: `class var .` 这份拼写逐文件钉死处数, 新增即红
#       (今天 9 处 / 6 份文件 —— 第六十三刀 (§B172) 之前 11 处: 那两处**生产**尾注的字面量收进了
#        唯一出口 classFieldReadExpr, 它按 label 实参拼这句, 于是发码文件里再无该字面量; 「哪几站
#        调它」改由 R9c 钉。剩下的都是讲这件事的注释或消费点 —— 注释也一起钉, 因为「谁在提这句
#        文本」正是这一族重新长出来的地方)
#        注释也一起钉, 因为「谁在提这句文本」正是这一族重新长出来的地方)
#   R5e 本模块那张老表 classVariantMembers_ 的写者仍只住一份文件, 恰好 4 枚 insert
$QDEF_RE_168    = 'Vb6Type\s+classFieldKind\s*\('
$QPROJ_DEF_169  = 'bool\s+classFieldIs(VariantSlot|BstrSlot|ScalarData)\s*\('
$QPROJ_USE_169  = 'classFieldKind\(className, memberName\)'
$QCALL_RE_168   = 'classFieldKind\s*\('
$MAPDECL_RE_168 = 'std::unordered_map<std::string,\s*std::unordered_map<std::string,\s*Vb6Type>>\s+classFieldKindMap\s*;'
$INJ_RE_168     = 'setClassFieldKindMap\(&classFieldKindMap\)'
$INS_RE_168     = 'classVariantMembers_\.insert'
$MARK_RE_168    = 'class var \.'
$srcFiles168 = @(Get-ChildItem -Path (Join-Path $Root 'src') -Recurse -File -Include *.cpp, *.hpp, *.inc)
$qDefs168 = 0; $qDefFiles168 = @(); $qAll168 = 0; $mapFiles168 = @(); $inj168 = 0
$projDefs169 = 0; $projForwards169 = 0; $projSites169 = 0
$markByFile168 = @{}; $insByFile168 = @{}
foreach ($f in $srcFiles168) {
    $t = [System.IO.File]::ReadAllText($f.FullName)
    $dm168 = [regex]::Matches($t, $QDEF_RE_168).Count
    if ($dm168 -gt 0) { $qDefs168 += $dm168; $qDefFiles168 += $f.Name }
    $projDefs169 += [regex]::Matches($t, $QPROJ_DEF_169).Count
    $projForwards169 += [regex]::Matches($t, $QPROJ_USE_169).Count
    # 注释行里提一句名字不算站点 —— 逐行扫, 只数**代码行**上的调用
    $pd169 = [regex]::Matches($t, $QPROJ_DEF_169).Count
    foreach ($ln168 in ($t -split "`r?`n")) {
        $tr168 = $ln168.Trim()
        if ($tr168.StartsWith('//') -or $tr168.StartsWith('*')) { continue }
        $qAll168 += [regex]::Matches($ln168, $QCALL_RE_168).Count
        $projSites169 += [regex]::Matches($ln168, 'classFieldIs(VariantSlot|BstrSlot|ScalarData)\s*\(').Count
    }
    $projSites169 -= $pd169          # 投影自身那三行定义不是站点
    if ([regex]::IsMatch($t, $MAPDECL_RE_168)) { $mapFiles168 += $f.Name }
    $inj168 += [regex]::Matches($t, $INJ_RE_168).Count
    $leaf168 = Split-Path $f.FullName -Leaf
    $mc168 = [regex]::Matches($t, $MARK_RE_168).Count
    if ($mc168 -gt 0) { $markByFile168[$leaf168] = $mc168 }
    $ic168 = [regex]::Matches($t, $INS_RE_168).Count
    if ($ic168 -gt 0) { $insByFile168[$leaf168] = $ic168 }
}
if ($qDefs168 -ne 1) { $bad += ('R5a classFieldKind 定义了 ' + $qDefs168 + ' 遍 (要恰好 1)') }
if ($qDefFiles168.Count -ne 1 -or $qDefFiles168[0] -ne 'cgen_api.inc') {
    $bad += ('R5a 查询出口住在 [' + ($qDefFiles168 -join ',') + '] (要住 cgen_api.inc)')
}
if ($projDefs169 -ne 3) {
    $bad += ('R5a 三个投影定义了 ' + $projDefs169 + ' 遍 (要恰好 3: VariantSlot/BstrSlot/ScalarData)')
}
if ($projForwards169 -lt $projDefs169) {
    $bad += ('R5a 投影里有 ' + ($projDefs169 - $projForwards169) + ' 枚没转问 classFieldKind —— 投影里再抄 find 就是两份答案')
}
$qSites168 = $qAll168 - $qDefs168 - $projForwards169
if ($qSites168 -lt 1) { $bad += ('R5b 直接问 classFieldKind 的站点 ' + $qSites168 + ' 处 (>=1: With 块那一站) —— 归零 = 那一问又没人管了') }
if ($projSites169 -lt 2) { $bad += ('R5b 三个投影的站点 ' + $projSites169 + ' 处 (>=2: 赋值站③ + 站④)') }
if ($mapFiles168.Count -ne 1 -or $mapFiles168[0] -ne 'driver_codegen_fieldkind_scan.inc') {
    $bad += ('R5c 工程级表建在 [' + ($mapFiles168 -join ',') + '] (只许 driver_codegen_fieldkind_scan.inc)')
}
if ($inj168 -ne 1) { $bad += ('R5c 表的注入站点 ' + $inj168 + ' 处 (要恰好 1)') }
$WantMark168 = [ordered]@{
    'cgen_expr_call_callee_member.inc'    = 1
    'cgen_expr_member_class_fallback.inc' = 1
    'cgen_expr_member_voidptr_com.inc'    = 4
    'cgen_state.inc'                      = 1
    'cgen_call.cpp'                       = 1
    'cgen_assign_value_sem.inc'           = 1
}
foreach ($k in $WantMark168.Keys) {
    $g168 = 0
    if ($markByFile168.ContainsKey($k)) { $g168 = $markByFile168[$k] }
    if ($g168 -ne $WantMark168[$k]) {
        $bad += ('R5d ' + $k + ' 里 class var . 出现 ' + $g168 + ' 处, 名单钉的是 ' + $WantMark168[$k])
    }
}
foreach ($k in $markByFile168.Keys) {
    if (-not $WantMark168.Contains($k)) {
        $bad += ('R5d 名单外的新文件 ' + $k + ' x' + $markByFile168[$k] + ' —— 又有人把那句尾注当判据了')
    }
}
$insFiles168 = @($insByFile168.Keys)
if ($insFiles168.Count -ne 1 -or $insByFile168[$insFiles168[0]] -ne 4 -or $insFiles168[0] -ne 'cgen_base_generate_state_scan.inc') {
    $bad += ('R5e classVariantMembers_ 的写者散在 [' + ($insFiles168 -join ',') + '] (只许 state_scan x4)')
}

# ---- R6 (账 #305 §B169): 收进 BSTR 槽 / 收进 Variant 槽这两问各只许住一枚出口 ----
# emitBstrSlotStore = 原先住在赋值站③ 那一支 (右值折算 + 自有还是借来 + 定长收口 + 清旧值),
# emitVariantSlotAssign = 原先住在赋值站④ 那一支 (装箱 + 所有权 + 自指要先收临时量)。
# With 块那一站 (§B169 形①) 需要的正是同一串决定, 所以它**共用**这两枚而不是再抄一遍。
#   R6a 各恰好定义 1 遍 (cgen_util_ctrl.cpp) + 声明 1 遍 (cgen_helpers.inc)
#   R6b 各调用点 >= 2 (赋值那一站 + With 那一站) —— 少了就是有人又开始自己抄
$AUTH169 = @('emitBstrSlotStore', 'emitVariantSlotAssign')
$authDef169 = @{}; $authDecl169 = @{}; $authCall169 = @{}
foreach ($a169 in $AUTH169) {
    $d169 = 0; $k169 = 0; $c169 = 0
    foreach ($f in $srcFiles168) {
        $t = [System.IO.File]::ReadAllText($f.FullName)
        $d169 += [regex]::Matches($t, ('void\s+CCodeGen::' + $a169 + '\s*\(')).Count
        $k169 += [regex]::Matches($t, ('^\s*void\s+' + $a169 + '\s*\('), 'Multiline').Count
        foreach ($ln169 in ($t -split "`r?`n")) {
            $tr169 = $ln169.Trim()
            if ($tr169.StartsWith('//') -or $tr169.StartsWith('*')) { continue }
            $c169 += [regex]::Matches($ln169, ($a169 + '\s*\(')).Count
        }
    }
    $authDef169[$a169] = $d169; $authDecl169[$a169] = $k169; $authCall169[$a169] = $c169 - $d169 - $k169
    if ($d169 -ne 1) { $bad += ('R6a ' + $a169 + ' 的定义 ' + $d169 + ' 遍 (要恰好 1, 在 cgen_util_ctrl.cpp)') }
    if ($k169 -ne 1) { $bad += ('R6a ' + $a169 + ' 的声明 ' + $k169 + ' 处 (要恰好 1, 在 cgen_helpers.inc)') }
    if (($c169 - $d169 - $k169) -lt 2) {
        $bad += ('R6b ' + $a169 + ' 的调用点 ' + ($c169 - $d169 - $k169) + ' 处 (<2 = With 那一站又不问它了)')
    }
}

# ---- R7 (账 #305 §B169): 手抄的 vb6_BSTR_Assign 发码点名单钉死, 新增即红 ----
# 出口本体那一文件交三形 (定长收口 / 自有接管 / 借来深拷贝); 另外两份是**还没并进来**的站:
#   cgen_assign_stmt_special.inc 4 处 = 定长串那两形 + 成员字段那两形 (每一站都得先问右值是不是自有)
#   cgen_stmt.cpp 1 处 = 控件属性写那一形
# 把哪一站改走 emitBstrSlotStore 了, 就把名单同步改短。
$WantBstr169 = [ordered]@{
    'cgen_util_ctrl.cpp'            = 3
    'cgen_assign_stmt_special.inc'  = 4
    'cgen_stmt.cpp'                 = 1
}
$bstrByFile169 = @{}
foreach ($f in $srcFiles168) {
    $t = [System.IO.File]::ReadAllText($f.FullName)
    $n169 = [regex]::Matches($t, 'c_\.emitLine\("vb6_BSTR_Assign').Count
    if ($n169 -gt 0) { $bstrByFile169[(Split-Path $f.FullName -Leaf)] = $n169 }
}
foreach ($k in $WantBstr169.Keys) {
    $g169 = 0
    if ($bstrByFile169.ContainsKey($k)) { $g169 = $bstrByFile169[$k] }
    if ($g169 -ne $WantBstr169[$k]) {
        $bad += ('R7 ' + $k + ' 里手写 vb6_BSTR_Assign 发码点 ' + $g169 + ' 处, 名单钉的是 ' + $WantBstr169[$k])
    }
}
foreach ($k in $bstrByFile169.Keys) {
    if (-not $WantBstr169.Contains($k)) {
        $bad += ('R7 名单外的新文件 ' + $k + ' x' + $bstrByFile169[$k] + ' —— BSTR 槽位写入又长了一份答案')
    }
}
$bstrTotal169 = 0
foreach ($v in $bstrByFile169.Values) { $bstrTotal169 += $v }

# ---- R8 (账 #305 §B170 第六十一刀): 「这枚成员读是哪一档」那两站都问那张表 ----
# 钉的是**读侧**三站 (inferExprType 的点号 case + With 体内 case + isDefinitelyVariantExpr 的认槽 case) 各问一次 classFieldKind,
# 并且问点只许长在名单里那几个文件。少一站 = 那一站回到按裸成员名查全局符号 —— 实测: 字段名
# 撞上内置函数名就读成内置的返回档 (`CStr(h.PV)` 撞 PV() 答 Double ⇒ 串值 0|1; 一枚
# `Public Command As Long` 撞 Command() 答 String ⇒ 把 7 当 BSTR 指针解, 两台原生 AV)。
# 第六十二刀 (§B171) 的第三站是**同一问的第三个读点** (isDefinitelyVariantExpr 的成员读分支), 不是新抄一份答案; 它同时把「尾注 variant 词」收进 cExprIsVariantSlotFieldRead 一处 (见 R9).
$WantReadAsk170 = [ordered]@{ 'cgen_util_type.cpp' = 3 }
$KnownAskFiles170 = @('cgen_api.inc', 'cgen_state.inc', 'driver_codegen_fieldkind_scan.inc',
                      'cgen_assign_prop_write.inc', 'cgen_util_type.cpp')
$askByFile170 = @{}
foreach ($f in $srcFiles168) {
    $t = [System.IO.File]::ReadAllText($f.FullName)
    $n170 = [regex]::Matches($t, 'classFieldKind\(').Count
    if ($n170 -gt 0) { $askByFile170[(Split-Path $f.FullName -Leaf)] = $n170 }
}
foreach ($k in $WantReadAsk170.Keys) {
    $g170 = 0
    if ($askByFile170.ContainsKey($k)) { $g170 = $askByFile170[$k] }
    if ($g170 -ne $WantReadAsk170[$k]) {
        $bad += ('R8 ' + $k + ' 问 classFieldKind ' + $g170 + ' 处, 名单钉的是 ' + $WantReadAsk170[$k] + ' (点号 + With + 认槽三站)')
    }
}
foreach ($k in $askByFile170.Keys) {
    if ($KnownAskFiles170 -notcontains $k) {
        $bad += ('R8 名单外的新问点 ' + $k + ' x' + $askByFile170[$k] + ' —— 读侧那一问又长了第二份地方')
    }
}
$readAsk170 = 0
if ($askByFile170.ContainsKey('cgen_util_type.cpp')) { $readAsk170 = $askByFile170['cgen_util_type.cpp'] }
# ---- R9 (账 #305 §B171 第六十二刀 → §B172 第六十三刀): 那条尾注只许一处写、三站调、一处读 ----
# 第六十三刀把「拼这条成员读尾注」本身收进唯一出口 classFieldReadExpr (定义在 cgen_util_type.cpp,
# 档位那一问就在它身上问 classFieldIsVariantSlot)。三站 (点号 / Fix 088c 的类形参 AST 推 / With)
# 只许调它; 少一站 = 那一站的槽读从此不带档位, 「按 C 文本认槽」那一问一律答否 —— 实测: 内在函数
# 吃整枚类形参的 Variant 槽字段 (`InStr(1, h.PV, ...)') 把 vb6_VARIANT 递进 BSTR 槽 ⇒ cl C2440.
# 「这枚 C 文本是不是一枚**可寻址的槽字段读」**收成一处 (cExprIsVariantSlotFieldRead):
#   定义与 cExprIsVariant 的那枚委托都在 cgen_util_type.cpp, 装箱站要地址就问这一处。
#   写词那一头只认得两条成员读的发文路 (点号 / With) —— 第三处自己拼尾注 = 下一次漏档的现场。
$WantWord171 = [ordered]@{ 'cgen_util_type.cpp' = 1 }
$WantAsk171 = [ordered]@{ 'cgen_util_type.cpp' = 2; 'cgen_helpers.inc' = 1;
                          'cgen_util_comwrite.cpp' = 1; 'cgen_expr_call_arg_emit.inc' = 2 }
$wordByFile171 = @{}; $askByFile171 = @{}
foreach ($f in $srcFiles168) {
    $t171 = [System.IO.File]::ReadAllText($f.FullName)
    $leaf171 = Split-Path $f.FullName -Leaf
    $w171 = [regex]::Matches($t171, '\?\s*" variant" : ""').Count
    $a171 = [regex]::Matches($t171, 'cExprIsVariantSlotFieldRead\(').Count
    if ($w171 -gt 0) { $wordByFile171[$leaf171] = $w171 }
    if ($a171 -gt 0) { $askByFile171[$leaf171] = $a171 }
}
foreach ($k in $WantWord171.Keys) {
    $gw171 = 0
    if ($wordByFile171.ContainsKey($k)) { $gw171 = $wordByFile171[$k] }
    if ($gw171 -ne $WantWord171[$k]) {
        $bad += ('R9 ' + $k + ' 写 variant 尾注 ' + $gw171 + ' 处, 名单钉的是 ' + $WantWord171[$k] + ' (三站只许调唯一出口 classFieldReadExpr, 词由它写)')
    }
}
foreach ($k in $wordByFile171.Keys) {
    if (-not $WantWord171.Contains($k)) {
        $bad += ('R9 第三处手写 variant 尾注 ' + $k + ' x' + $wordByFile171[$k] + ' —— 写词那一问又长了第三份地方')
    }
}
foreach ($k in $WantAsk171.Keys) {
    $ga171 = 0
    if ($askByFile171.ContainsKey($k)) { $ga171 = $askByFile171[$k] }
    if ($ga171 -ne $WantAsk171[$k]) {
        $bad += ('R9 ' + $k + ' 问 cExprIsVariantSlotFieldRead ' + $ga171 + ' 处, 名单钉的是 ' + $WantAsk171[$k])
    }
}
foreach ($k in $askByFile171.Keys) {
    if (-not $WantAsk171.Contains($k)) {
        $bad += ('R9 名单外的新问点 ' + $k + ' x' + $askByFile171[$k] + ' —— 这一问又有人自己抄一份')
    }
}
# ---- R9c (第六十三刀 §B172): 成员读发文站三处, 每一处都只许调那条唯一出口 ----
$WantStation172 = [ordered]@{ 'cgen_expr_member_class_fallback.inc' = 2; 'cgen_expr_with.cpp' = 1 }
$stationByFile172 = @{}
foreach ($f in $srcFiles168) {
    $t172 = [System.IO.File]::ReadAllText($f.FullName)
    $n172 = [regex]::Matches($t172, '= classFieldReadExpr\(').Count
    if ($n172 -gt 0) { $stationByFile172[(Split-Path $f.FullName -Leaf)] = $n172 }
}
foreach ($k in $WantStation172.Keys) {
    $g172 = 0
    if ($stationByFile172.ContainsKey($k)) { $g172 = $stationByFile172[$k] }
    if ($g172 -ne $WantStation172[$k]) {
        $bad += ('R9 ' + $k + ' 调 classFieldReadExpr ' + $g172 + ' 处, 名单钉的是 ' + $WantStation172[$k] + ' (点号 + 类形参 AST + With 三站)')
    }
}
foreach ($k in $stationByFile172.Keys) {
    if (-not $WantStation172.Contains($k)) {
        $bad += ('R9 名单外的新发文站 ' + $k + ' x' + $stationByFile172[$k] + ' —— 成员读那一站又有人自己拼尾注')
    }
}
$stationTotal172 = 0
foreach ($v in $stationByFile172.Values) { $stationTotal172 += $v }

# ---- R10 (账 #305 §B173 第六十四刀): 「整枚槽交进具体档形参」那两条臂都必须先问 AST ----
# cgen_expr_call_arg_variant.inc 里同一句决定抄了两遍: 上面那条臂 (Fix 029, 形参类型已知) 问
# AST ∪ 文本, 下面那条臂 (Fix 038b-2, 内置/运行时形参表为空或档为 Variant) 此前只问文本 ——
# 于是同一枚 UDT 的 Variant 成员 `u.UV` 在 Left/Len/Replace/Val (带类型化形参表) 那里被解包,
# 在 InStr (只登记返回档) 那里裸交 vb6_VARIANT 进 BSTR 形参 = cl C2440 (两台都没 exe)。
# 口径 = §B147「改道吃 AST ∪ 文本」: 两条臂各问一次, 总数钉在 2。掉回 1 = 那条臂又只认文本了。
$WantAstAsk173 = [ordered]@{ 'cgen_expr_call_arg_variant.inc' = 2; 'cgen_expr_call_arg_emit.inc' = 2 }
$astByFile173 = @{}
foreach ($f in $srcFiles168) {
    $t173 = [System.IO.File]::ReadAllText($f.FullName)
    $n173 = [regex]::Matches($t173, 'isDefinitelyVariantExpr\(\s*\*node\.positional\[i\]').Count
    if ($n173 -gt 0) { $astByFile173[(Split-Path $f.FullName -Leaf)] = $n173 }
}
foreach ($k in $WantAstAsk173.Keys) {
    $g173 = 0
    if ($astByFile173.ContainsKey($k)) { $g173 = $astByFile173[$k] }
    if ($g173 -ne $WantAstAsk173[$k]) {
        $bad += ('R10 ' + $k + ' 问 isDefinitelyVariantExpr ' + $g173 + ' 处, 名单钉的是 ' + $WantAstAsk173[$k] + ' (形参已知/形参未知两条臂都要吃 AST)')
    }
}
foreach ($k in $astByFile173.Keys) {
    if (-not $WantAstAsk173.Contains($k)) {
        $bad += ('R10 名单外的新问点 ' + $k + ' x' + $astByFile173[$k] + ' —— 「实参是不是整枚槽」那一问又长了第五处地方')
    }
}
$astTotal173 = 0
foreach ($v in $astByFile173.Values) { $astTotal173 += $v }

$wordTotal171 = 0
foreach ($v in $wordByFile171.Values) { $wordTotal171 += $v }
$askTotal171 = 0
foreach ($v in $askByFile171.Values) { $askTotal171 += $v }
$markTotal168 = 0
foreach ($v in $markByFile168.Values) { $markTotal168 += $v }

$nHand = $handPairs.Count
if ($bad.Count -eq 0) {
    Write-Host ('PASS variant_slot_store: def 1 | calls ' + $callSites + ' | shapes ' + $nHand + ' | open stations ' + $bareTotal + ' | slotask168 ' + $qSites168 + ' | projask169 ' + $projSites169 + ' | mark168 ' + $markTotal168 + ' | bstr169 ' + $bstrTotal169 + ' | readask170 ' + $readAsk170 + ' | slotword171 ' + $wordTotal171 + ' | slotst172 ' + $stationTotal172 + ' | slotask171 ' + $askTotal171 + ' | argarm173 ' + $astTotal173 + ' | auth169 ' + $authCall169['emitBstrSlotStore'] + '/' + $authCall169['emitVariantSlotAssign'] + ' | violations 0') -ForegroundColor Green
    exit 0
}
foreach ($b in $bad) { Write-Host ('FAIL ' + $b) -ForegroundColor Red }
exit 1
