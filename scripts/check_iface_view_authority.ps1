# check_iface_view_authority.ps1 - 账 #305 §B221/#449（§B222）：「这一枚表达式交出来的是视图还是实例」
# 这一问只许有一位主人。十一条钉：
#   ① 视图的成员调用发射形状 `->vt-><slot>(x, args)` 在全 src/backend 恰好 **1** 处，
#      且必须住在 module/cgen_iface_vtbl.cpp 的 ivrefVtableCall 里（com_bind 与「接收者是
#      视图值的调用」那一站都调它）。长出第 2 处 = 同一形状两位作者。
#   ② 问句 ifaceViewOfExpr 的**消费文件**恰好 **6** 份（今天在册的名单，逐文件钉住）：Set 目标侧 / 实参侧 /
#      成员访问接收者（字段兜底臂）/ 成员访问接收者（类模块臂，第二百一十刀 S2）/
#      带实参调用的接收者 / 成员访问接收者（零参裸读那一臂，2026-11-10 R4 收口）。少一份 = 有一站回到各写一遍；多一份 = 判据被复制。
#   ③ 直接拿字面量 "vb6_ivref_" 拼类型/问档的文件恰好 **7** 份（今天在册的名单，逐文件钉住）。
#      第 8 份长出来就意味着「是不是接口」又有了第二位主人。
#   ④ 「Set 的槽位交给哪一位主人」那一问 (第二百零三刀) 两站同问:
#      cgen_setlet_set_prop.inc (置空腿③) 与 cgen_setlet_set_rhs.inc (覆写腿②) **各自**至少
#      问一次 ivrefCType(<声明档>)。少一站 = 那一站照旧铸 `vb6_cls_<接口名>_AddRef/_Release`
#      = 链接期 LNK2019 (第二百零三刀实测: 只补覆写那一站时, 夹具红在 `Set me->m_iv = Nothing`)。
#   ⑤ 视图计数主人 emitIvrefRefSet 的**调用点**恰好 **5** 处, 且只在那两份 Set 站点文件里
#      (cgen_setlet_set_prop.inc 3 处: 按名 B05 / 置空 / 字段覆写; cgen_setlet_set_rhs.inc 2 处: 覆写腿②的赋值/置空)。长出第 6 处或换文件 = 同一发计数两位作者。
#   ⑥ 「一枚工程类实例怎么交进视图槽」那一发形状 (`&(<实例>)->__iv_<I>`, 第二百零四刀) 有**一位**
#      作者 = ivrefInstanceViewExpr；它的**消费点恰好 3 处**、分居三份文件：
#      Set 覆写腿② / Set 按名 B04 / 实参交付侧 deliverToClassSlot。少一处 = 那一站回到自己拼；
#      多一处 = 同一形状第二位作者。第二百零五刀给交付侧再接一档（借来的实例走 nullable 那一档）
#      ⇒ 消费点 3 → **4**，仍分居那三份文件。
#   ⑦ 字面量 `"__iv_" + cIdent(` 那份拼法只许住在两份文件里 = 主人自己 + ByVal 实参适配器
#      (那一发带 `(vb6_ivref_I*)` 强转、形状不同，刻意留在册)。第 3 份长出来 = 又一处现拼视图槽字段。
#   ⑧ 视图槽「谁付钱」那一发 AddRef (第二百零五刀给 copy-in 补上视图档) 的**发射点**在全 src/backend
#      恰好 **5** 处、只居这 3 份文件 = cgen_com.cpp (registerClassArgRefCopyIn 的视图臂，恰 1 处) +
#      cgen_iface_vtbl.cpp (emitIvrefRefSet) + cgen_setlet_set_prop.inc (仍在册的两处内联形状 + 置空前退旧)。
#      站点自己再拼一发 = 「谁付钱」长出第二位主人；这一条掉到 4 处 = 视图档又没人付（实测修前全模块 0 条
#      copy-in ⇒ 交回腿净 −1 ⇒ 过程出口拆槽时 0xC0000374）。
#   ⑨ 「接收者是视图 ⇒ 属性读走 vtable 槽」那一臂有一位主人 ivrefViewGetterReadExpr：
#      定义恰 1 (cgen_iface_vtbl.cpp) / 声明恰 1 (cgen_helpers.inc) / 调用点恰 2 且住在
#      cgen_util_com.cpp (站S1) + cgen_call.cpp (Debug.Print 的 COM 属性读那一站, 2026-11-10 R4 收口接上)。
#      它问名册那一头 slotIvrefViewName: 名册 (knownIvrefVars_) 是**按名**那张表, 而 ifaceViewOfExpr 是**按表达式**
#      那一问 —— 第二百一十刀 S1 起, 名册的**文本投影**就是给「手里只有发射文本、拿不到 AST」的那一族站点用的
#      (R4 收口接进来的六处都属这一族: With 接收者 / 语句位接收者 / For Each 接收者 / 跨模块成员调用那一站)。
#      它们问的是**同一个函数**, 不是各抄一份 ⇒ 不变量「名册只由一个函数读」仍然成立 (定义 1 + 声明 1);
#      被读多少次是**读者名单**, 名单是闭的 ⇒ 逐文件钉死: 带 `(` 的出现恰 10 处 = 定义 1 + 主人内 2 +
#      声明 1 + 六位站点读者各 1。长出第 11 处 = 有站点绕过主人自己查名册 = 「是不是视图」有了第二位判据的读者。
#   ⑩ 「字段档」那一层 `me->` 前缀只许**两族名册主人各自剥一次** (第二百一十五刀把旧式那一族补齐):
#      发射形状 `static const std::string mePre = "me->";` 在全 src/backend 恰好 **2** 处，且只住
#      cgen_iface_vtbl.cpp (新式 slotIvrefViewName) + cgen_util_classtype.cpp (旧式 slotLegacyIfaceViewName)。
#      同一趟还钉「Set 那两站不许再自己查那张旧式名册」: cgen_setlet_set_prop.inc 里
#      `knownIfaceVars_.find(targetLower)` 必须 **0** 处 (第二百一十五刀之前它是 1 处，字段槽
#      `me->m_lv` 因此在装填那一站落空 = 裸拷贝实例指针进值结构体槽 = cl C2440)。
#   ⑪ 「一枚**值结构体槽**怎么清成空」那一发有一位主人 emitValueStructSlotClear (第二百一十七刀):
#      带 `(` 的出现恰 **4** 处 = 定义 1 (cgen_util_classtype.cpp) + 声明 1 (cgen_helpers.inc) +
#      调用 2 (类工厂字段初始化 / Set 置空那一站，只许这两份文件)；且发射形状那一串 `"memset(&"` 在
#      全 src/backend 恰 **1** 处、住在主人自己那份文件。长出第 2 处 = 站点又自己拼清零那一发
#      (病灶实物 = 置空那一站从前把整枚结构体的地址交给 vb6_ReleaseObject(void**)，读的首字段是
#      那张函数指针表；字段初始化那一档从前发 `= NULL` = cl C2440)。
# 退出码：0 = 十一条都过；1 = 任一条红。
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
$files = Get-ChildItem -Path $backend -Recurse -File |
    Where-Object { $_.Extension -match '^\.(cpp|inc|h|hpp)$' }

$fail = @()

# ① vtable 发射形状唯一作者
$vt = @($files | Select-String -Pattern '"->vt->"')
if ($vt.Count -ne 1) {
    $fail += ("vt-call sites = " + $vt.Count + " (want exactly 1)")
} elseif ($vt[0].Path -notmatch 'cgen_iface_vtbl\.cpp$') {
    $fail += ("vt-call formula moved into " + $vt[0].Path)
}

# ② ifaceViewOfExpr 的消费文件（去掉它自己的定义与声明）
#   ⚠ 2026-11-12 (第五十六刀) 改钉: 原来钉 5 份。2026-11-10 (650b7bac, R4 收口) 给
#   **成员访问接收者的零参裸读**那一臂 (cgen_expr_member.cpp) 也接上了同一位主人 ⇒ 6 份。
#   按 §56.38 A7 先例: 计数照实改, 再**逐文件钉住** (名单是闭的 —— 只钉总数会漏掉
#   「少一处、多一处但总数不变」)。下面那一行 roster 就是权威。
$ivAll = @($files | Select-String -Pattern 'ifaceViewOfExpr')
$ivFiles = @($ivAll | ForEach-Object { Split-Path -Leaf $_.Path } | Sort-Object -Unique |
    Where-Object { $_ -ne 'cgen_util_classtype.cpp' -and $_ -ne 'cgen_helpers.inc' })
if ($ivFiles.Count -ne 6) {
    $fail += ("ifaceViewOfExpr consumer files = " + $ivFiles.Count +
              " (want exactly 6): " + ($ivFiles -join ', '))
}
$ivWant = 'cgen_expr_call_arg_emit.inc,cgen_expr_call_callee_member.inc,cgen_expr_member.cpp,' +
          'cgen_expr_member_class_fallback.inc,cgen_expr_member_class_module.inc,cgen_setlet_set_prop.inc'
if (($ivFiles -join ',') -ne $ivWant) {
    $fail += ("ifaceViewOfExpr consumer roster = " + ($ivFiles -join ', ') +
              " (want exactly: " + $ivWant + ")")
}

# ③ 字面量 "vb6_ivref_" 的主人名单
#   ⚠ 2026-11-12 改钉: 6 → 7。R4 收口 (43528bb8) 给 With 那一站 (cgen_with.cpp) 接上了
#   「临时量的 C 类型 = vb6_ivref_<I>*」那一发 ⇒ 第 7 份。同样逐文件钉住。
$litFiles = @($files | Select-String -Pattern '"vb6_ivref_"' |
    ForEach-Object { Split-Path -Leaf $_.Path } | Sort-Object -Unique)
if ($litFiles.Count -ne 7) {
    $fail += ("vb6_ivref_ literal files = " + $litFiles.Count +
              " (want exactly 7): " + ($litFiles -join ', '))
}
$litWant = 'cgen_base_type.cpp,cgen_decl_func.cpp,cgen_expr_call_arg_emit.inc,cgen_iface_vtbl.cpp,' +
           'cgen_setlet_set_prop.inc,cgen_util_comwrite.cpp,cgen_with.cpp'
if (($litFiles -join ',') -ne $litWant) {
    $fail += ("vb6_ivref_ literal roster = " + ($litFiles -join ', ') +
              " (want exactly: " + $litWant + ")")
}

# ④ Set 两站同问: 置空腿③ 与 覆写腿② 各问一次「这格槽的声明档是不是新式接口视图」
$nullStation = Join-Path $backend "detail\stmt\cgen_setlet_set_prop.inc"
$overStation = Join-Path $backend "detail\stmt\cgen_setlet_set_rhs.inc"
foreach ($st in @(@{ F = $nullStation; P = 'ivrefCType\(targetCls101\)' },
                  @{ F = $overStation; P = 'ivrefCType\(clsOfSlot\)' })) {
    if (-not (Test-Path $st.F)) { $fail += ("station missing: " + $st.F); continue }
    $asked = @(Select-String -Path $st.F -Pattern $st.P)
    if ($asked.Count -lt 1) {
        $fail += ("station " + (Split-Path -Leaf $st.F) + " no longer asks " + $st.P +
                  " (= 那一格会再铸一枚 vb6_cls_<接口名>_AddRef/_Release, LNK2019)")
    }
}

# ⑤ 视图计数主人的调用点名单
#   ⚠ 2026-11-12 改钉: 4 → 5。2026-11-10 (1075d990, Set <接口变量> = <Variant/Object>
#   走运行期类名拆箱) 在 cgen_setlet_set_prop.inc 的字段覆写那一支加了第 5 处 ⇒
#   逐文件钉住 (覆写腿那一站 = 3 处: 置空 / 按名 / 字段覆写; 覆写腿②那一站 = 2 处)。
$refSetCalls = @($files | Select-String -Pattern 'emitIvrefRefSet\(target')
if ($refSetCalls.Count -ne 5) {
    $fail += ("emitIvrefRefSet call sites = " + $refSetCalls.Count + " (want exactly 5)")
}
$refSetFiles = @($refSetCalls | ForEach-Object { Split-Path -Leaf $_.Path } | Sort-Object -Unique)
if (($refSetFiles -join ',') -ne 'cgen_setlet_set_prop.inc,cgen_setlet_set_rhs.inc') {
    $fail += ("emitIvrefRefSet consumers = " + ($refSetFiles -join ', ') +
              " (want only the two Set station files)")
}
$refSetProp = @($refSetCalls | Where-Object { (Split-Path -Leaf $_.Path) -eq 'cgen_setlet_set_prop.inc' })
$refSetRhs = @($refSetCalls | Where-Object { (Split-Path -Leaf $_.Path) -eq 'cgen_setlet_set_rhs.inc' })
if ($refSetProp.Count -ne 3 -or $refSetRhs.Count -ne 2) {
    $fail += ("emitIvrefRefSet per-file = cgen_setlet_set_prop.inc " + $refSetProp.Count +
              " + cgen_setlet_set_rhs.inc " + $refSetRhs.Count +
              " (want 3 + 2 = 置空 / 按名 / 字段覆写 + 覆写腿②两处)")
}

# ⑥ 视图形状主人 ivrefInstanceViewExpr: 一位作者 + 三处消费 (Set 覆写 / Set 按名 / 实参交付侧)
$viewAll = @($files | Select-String -Pattern 'ivrefInstanceViewExpr\(')
$viewDef = @($viewAll | Where-Object { (Split-Path -Leaf $_.Path) -eq 'cgen_iface_vtbl.cpp' })
$viewDecl = @($viewAll | Where-Object { (Split-Path -Leaf $_.Path) -eq 'cgen_helpers.inc' })
$viewCalls = @($viewAll | Where-Object { $_ -notin $viewDef -and $_ -notin $viewDecl })
if ($viewDef.Count -ne 1) {
    $fail += ("ivrefInstanceViewExpr definition sites = " + $viewDef.Count + " (want exactly 1)")
}
if ($viewDecl.Count -ne 1) {
    $fail += ("ivrefInstanceViewExpr declarations = " + $viewDecl.Count + " (want exactly 1)")
}
if ($viewCalls.Count -ne 4) {
    $fail += ("ivrefInstanceViewExpr call sites = " + $viewCalls.Count + " (want exactly 4)")
}
$viewCallFiles = @($viewCalls | ForEach-Object { Split-Path -Leaf $_.Path } | Sort-Object -Unique)
if (($viewCallFiles -join ',') -ne 'cgen_base_type.cpp,cgen_setlet_set_prop.inc,cgen_setlet_set_rhs.inc') {
    $fail += ("view-shape consumers = " + ($viewCallFiles -join ', ') +
              " (want the two Set stations + the arg delivery side)")
}
# 交付侧那一站必须**经主人**拿这一发 (自己内联拼 = 这一条掉下来)
$deliverVia = @(Select-String -Path (Join-Path $backend "cgen_base_type.cpp") `
                              -Pattern 'ivrefInstanceViewExpr\(')
if ($deliverVia.Count -ne 2) {
    $fail += ("deliverToClassSlot routes the view shape through ivrefInstanceViewExpr " +
              ("" + $deliverVia.Count) + " times (want exactly 2 = New 那一档 + 借来的实例那一档)" +
              " (少一档 = 实参站回到把实例基址当视图交出去, 实测 rc=0xC0000005)")
}

# ⑦ `"__iv_" + cIdent(` 那份拼法的住家名单
$shapeFiles = @($files | Select-String -Pattern '"__iv_" \+ cIdent' |
    ForEach-Object { Split-Path -Leaf $_.Path } | Sort-Object -Unique)
if (($shapeFiles -join ',') -ne 'cgen_expr_call_arg_emit.inc,cgen_iface_vtbl.cpp') {
    $fail += ("inline __iv_ spellings live in " + ($shapeFiles -join ', ') +
              " (want only the owner + the ByVal arg adapter)")
}

# ⑧ 视图槽「谁付钱」那一发 AddRef 的发射点名单 (第二百零五刀)
$addRefs = @($files | Select-String -Pattern '"->vt->AddRef\("' -Encoding UTF8)
if ($addRefs.Count -ne 5) {
    $fail += ("view AddRef emission sites = " + $addRefs.Count + " (want exactly 5)")
}
$addRefFiles = @($addRefs | ForEach-Object { Split-Path -Leaf $_.Path } | Sort-Object -Unique)
if (($addRefFiles -join ',') -ne 'cgen_com.cpp,cgen_iface_vtbl.cpp,cgen_setlet_set_prop.inc') {
    $fail += ("view AddRef owners = " + ($addRefFiles -join ', ') +
              " (want copy-in + emitIvrefRefSet + the two recorded inline Set shapes)")
}
$copyInAddRef = @(Select-String -Path (Join-Path $backend "module\cgen_com.cpp") `
                              -Pattern '"->vt->AddRef\("' -Encoding UTF8)
if ($copyInAddRef.Count -ne 1) {
    $fail += ("registerClassArgRefCopyIn view arm = " + $copyInAddRef.Count +
              " (want exactly 1; 0 = 视图槽又没人付钱, 实测净 −1 → 出口 0xC0000374)")
}


# ⑨ 新式视图属性读那一臂的主人 (第二百一十刀 S1)
#   ⚠ 2026-11-12 (第五十六刀) 改钉两处:
#   (a) 调用点 1 → 2: R4 收口 (43528bb8) 给 Debug.Print 的 COM 属性读那一站
#       (cgen_call.cpp) 也接上了同一位主人 ⇒ 那一站原来漏问 (视图属性读恒 0)。
#       逐文件钉住: cgen_util_com.cpp (站S1) + cgen_call.cpp (Debug.Print 那一站)。
#   (b) 名册那一问 slotIvrefViewName 的读者 3 → 10: 原来钉的是「只许主人自己用」
#       (定义 + 声明 + 主人内那一次)。**那一钉把「主人」与「读者」混成了一件事** ——
#       名册 (knownIvrefVars_) 是**按名**那张表, 而 `ifaceViewOfExpr` 是**按表达式**那一问;
#       第二百一十刀 S1 起, 名册的**文本投影**就是给「手里只有发射文本、拿不到 AST」的
#       那一族站点用的 (R4 收口接进来的六处都属这一族: With 接收者 / 语句位接收者 /
#       For Each 接收者 / 跨模块成员调用那一站)。它们问的是**同一个函数**,
#       不是各抄一份 ⇒ 不变量「名册只由一个函数读」仍然成立 (定义 1 + 声明 1);
#       被读多少次是**读者名单**, 名单是闭的 ⇒ 下面逐文件钉死, 多一位读者立刻红。
$rdAll = @($files | Select-String -Pattern 'ivrefViewGetterReadExpr\(')
$rdDef = @($rdAll | Where-Object { (Split-Path -Leaf $_.Path) -eq 'cgen_iface_vtbl.cpp' })
$rdDecl = @($rdAll | Where-Object { (Split-Path -Leaf $_.Path) -eq 'cgen_helpers.inc' })
$rdCalls = @($rdAll | Where-Object { $_ -notin $rdDef -and $_ -notin $rdDecl })
if ($rdDef.Count -ne 1) {
    $fail += ("ivrefViewGetterReadExpr definition sites = " + $rdDef.Count + " (want exactly 1)")
}
if ($rdDecl.Count -ne 1) {
    $fail += ("ivrefViewGetterReadExpr declarations = " + $rdDecl.Count + " (want exactly 1)")
}
if ($rdCalls.Count -ne 2) {
    $fail += ("ivrefViewGetterReadExpr call sites = " + $rdCalls.Count + " (want exactly 2)")
}
$rdCallFiles = @($rdCalls | ForEach-Object { Split-Path -Leaf $_.Path } | Sort-Object -Unique)
if (($rdCallFiles -join ',') -ne 'cgen_call.cpp,cgen_util_com.cpp') {
    $fail += ("view property-read consumers = " + ($rdCallFiles -join ', ') +
              " (want exactly: cgen_call.cpp, cgen_util_com.cpp = Debug.Print 那一站 + 站S1)")
}
$nameAll = @($files | Select-String -Pattern 'slotIvrefViewName\(')
if ($nameAll.Count -ne 10) {
    $fail += ("slotIvrefViewName sites = " + $nameAll.Count +
              " (want exactly 10 = 定义 1 + 声明 1 + 主人内 2 + 站点读者 6)")
}
$nameFiles = @($nameAll | ForEach-Object { Split-Path -Leaf $_.Path } | Sort-Object -Unique)
if (($nameFiles -join ',') -ne 'cgen_assign_stmt_special.inc,cgen_call.cpp,cgen_control.cpp,' +
    'cgen_expr_call_com_bind.inc,cgen_expr_with.cpp,cgen_helpers.inc,cgen_iface_vtbl.cpp,cgen_with.cpp') {
    $fail += ("roster ask lives in " + ($nameFiles -join ', ') +
              " (want the owner + its declaration + the six text-level stations)")
}
# 名册的读者名单**逐文件**钉死: 主人那份文件里 3 处 (定义 + 主人内两问), 其余六个站点各 1 处。
$nameOwner = @($nameAll | Where-Object { (Split-Path -Leaf $_.Path) -eq 'cgen_iface_vtbl.cpp' })
$nameDecl = @($nameAll | Where-Object { (Split-Path -Leaf $_.Path) -eq 'cgen_helpers.inc' })
$nameStations = @($nameAll | Where-Object { $_ -notin $nameOwner -and $_ -notin $nameDecl })
if ($nameOwner.Count -ne 3 -or $nameDecl.Count -ne 1 -or $nameStations.Count -ne 6) {
    $fail += ("slotIvrefViewName per-file = owner " + $nameOwner.Count + " + declaration " +
              $nameDecl.Count + " + stations " + $nameStations.Count +
              " (want 3 + 1 + 6: 定义 + 主人内两问 / 声明 / 六位站点读者)")
}
$nameStationFiles = @($nameStations | ForEach-Object { Split-Path -Leaf $_.Path } | Sort-Object -Unique)
if (($nameStationFiles -join ',') -ne 'cgen_assign_stmt_special.inc,cgen_call.cpp,cgen_control.cpp,' +
    'cgen_expr_call_com_bind.inc,cgen_expr_with.cpp,cgen_with.cpp') {
    $fail += ("slotIvrefViewName station roster = " + ($nameStationFiles -join ', ') +
              " (want exactly six text-level stations)")
}


#   ⑩ 字段档那一层 `me->` 前缀只许**两族名册主人各自剥一次** (第二百一十五刀补齐旧式那一族):
#      发射形状 `static const std::string mePre = "me->";` 在全 src/backend 恰好 **2** 处, 且只住
#      cgen_iface_vtbl.cpp (新式 slotIvrefViewName) + cgen_util_classtype.cpp (旧式 slotLegacyIfaceViewName)。
#      同一趟钉「Set 那一站不许再自己查那张旧式名册」: prop 站里那条直查必须 **0** 处 —— 第二百一十五刀
#      之前它是 1 处, 字段槽 `me->m_lv` 因此在装填那一站落空 = 裸拷贝实例指针进值结构体槽 = cl C2440。
$stripAll = @($files | Select-String -Pattern 'static const std::string mePre = "me->";')
if ($stripAll.Count -ne 2) {
    $fail += ("me-> strip sites in the two roster owners = " + $stripAll.Count +
              " (want exactly 2 = one per family)")
}
$stripFiles = @($stripAll | ForEach-Object { Split-Path -Leaf $_.Path } | Sort-Object -Unique)
if (($stripFiles -join ',') -ne 'cgen_iface_vtbl.cpp,cgen_util_classtype.cpp') {
    $fail += ("me-> strip lives in " + ($stripFiles -join ', ') +
              " (want the new-style owner + the legacy owner; a 3rd file = a station strips it itself)")
}
$directAsk = @(Select-String -Path (Join-Path $backend "detail\stmt\cgen_setlet_set_prop.inc") `
                          -Pattern 'knownIfaceVars_\s*\.\s*find\s*\(\s*targetLower')
if ($directAsk.Count -ne 0) {
    $fail += ("the Set station asks the legacy roster directly again (" + $directAsk.Count +
              " sites) = 字段槽 (me->x) 问不中名册, 装填那一站落回裸拷贝 (实测 cl C2440)")
}


# ⑪ 账 #305 第二百一十七刀: 「一枚值结构体槽怎么清成空」那一发的唯一主人。
$vsAll = @($files | Select-String -Pattern 'emitValueStructSlotClear\(')
$vsDef = @($vsAll | Where-Object { (Split-Path -Leaf $_.Path) -eq 'cgen_util_classtype.cpp' })
$vsDecl = @($vsAll | Where-Object { (Split-Path -Leaf $_.Path) -eq 'cgen_helpers.inc' })
$vsCalls = @($vsAll | Where-Object { $_ -notin $vsDef -and $_ -notin $vsDecl })
if ($vsDef.Count -ne 1) {
    $fail += ("value-struct empty-form owner definitions = " + $vsDef.Count + " (want exactly 1)")
}
if ($vsDecl.Count -ne 1) {
    $fail += ("value-struct empty-form owner declarations = " + $vsDecl.Count + " (want exactly 1)")
}
if ($vsCalls.Count -ne 2) {
    $fail += ("value-struct empty-form call sites = " + $vsCalls.Count +
              " (want exactly 2 = 类工厂字段初始化 + Set 置空那一站)")
}
$vsCallFiles = @($vsCalls | ForEach-Object { Split-Path -Leaf $_.Path } | Sort-Object -Unique)
if (($vsCallFiles -join ',') -ne 'cgen_com.cpp,cgen_setlet_set_prop.inc') {
    $fail += ("value-struct empty-form consumers = " + ($vsCallFiles -join ', ') +
              " (want the factory field-init station + the Set clear station)")
}
$memsetSites = @($files | Select-String -Pattern '"memset\(&"' -Encoding UTF8)
if ($memsetSites.Count -ne 1) {
    $fail += ("value-struct zeroing shape sites = " + $memsetSites.Count +
              " (want exactly 1 = 主人自己拼那一条; 长出第 2 处 = 站点又自己拼清零那一发)")
} elseif ((Split-Path -Leaf $memsetSites[0].Path) -ne 'cgen_util_classtype.cpp') {
    $fail += ("value-struct zeroing shape moved into " + (Split-Path -Leaf $memsetSites[0].Path))
}

if ($fail.Count -gt 0) {
    Write-Host "IFACE-VIEW-AUTHORITY FAIL"
    foreach ($f in $fail) { Write-Host ("  - " + $f) }
    exit 1
}
Write-Host "IFACE-VIEW-AUTHORITY OK (1 formula site / 6 ifaceViewOfExpr consumer files in an exact roster / 7 vb6_ivref_ literal files in an exact roster / 2 Set stations ask / 5 refSet calls (3+2) / 4 view-shape consumers / 2 inline __iv_ files / 5 view AddRef sites in 3 files / 1 view property-read owner with 2 call sites in 2 stations + 10 roster ask sites (owner 3 + declaration 1 + 6 stations) / 2 me- prefix strips in 2 owner files + 0 direct roster asks in the Set station / 4 empty-form owner sites in 4 files + 1 zeroing shape site)"
exit 0
