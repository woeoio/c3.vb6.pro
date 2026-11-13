# check_legacy_view_arg_owners.ps1 - legacy 接口**视图**在被接收那一头只有这几位主人
#
# 起因 (2026-10-22, 账 #305 第一百七十二刀): 真语料 Ray.vbp 里 `vb6_iface_<I> { vtbl; obj }`
# 这枚值结构体站在被赋值/被传参那一头共 23 站。修的时候三条都不许长第二种答案 ——
# 谁拼 wrap 名 (legacyIfaceWrapFn, 第一百六十六刀)、谁认视图槽 (slotIsLegacyIfaceView,
# 第一百六十五刀)、谁答视图进装箱 (legacyViewBoxExpr, 第一百七十一刀)、谁把类型名认成
# 视图槽 (cParamLegacyIfaceSlotName, 本刀新增, 与 cParamClassPtrType 同住一份文件)。
# 这条线一旦有人自己拼 `vb6_iface_` 前缀、自己查 knownIfaceVars_、或者在别处再发一句
# `.obj ==`, 同一件事就又有了两位主人 ⇒ 本哨兵按"位数"钉死, 加读者要回台账改数。
#
# 规则 (改坏了会红, 不是装饰):
#   V1  `vb6_iface_` 那一串前缀比较 compare(0, 10, "vb6_iface_") 在代码里恰好 1 处,
#       且住在 cgen_base_type.cpp (那张形参槽名单出生的地方)。
#   V2  cParamLegacyIfaceSlotName 恰好 3 处 = 定义 1 (cgen_base_type.cpp) + 声明 1
#       (cgen_helpers.inc) + 站点 1 (cgen_expr_call_arg_emit.inc)。
#   V3  legacyIfaceWrapFn( 在代码里恰好 8 处 = 定义 1 + 声明 1 + 调用 6
#       (Set 那三臂 3 / cgen_com 发定义那 1 / 实参位 1)。
#   V4  legacyViewBoxExpr( 在代码里恰好 5 处 = 定义 1 + 声明 1 + 调用 3
#       (wrapVariantValue / boxToVariant / ByRef Variant 实参位)。
#   V5  slotIsLegacyIfaceView( 在代码里恰好 6 处 = 定义 1 + 声明 1 + 调用 4
#       (legacyViewBoxExpr 自己 / 腿② 覆写 / 腿③ 置空 / 第一百七十五刀表达式路那一问) —— 第二百四十八刀把 「Is」 比较那一站换成**带名字**那一问 (要拿接口名再问一次形状), 所以这一档从 5 处调用降到 4 处, 那一处搬到 V7 去数。
#   V6  视图参与比较那两句发码: `).obj == (` 恰好 1 处、`).obj == NULL` 恰好 2 处,
#       都住在 cgen_expr_binary.cpp。
#
# 用法: powershell -File scripts\check_legacy_view_arg_owners.ps1 [-Root <副本目录>]  (PASS = exit 0)

param([string]$Root = "")

$ErrorActionPreference = "Stop"
if ($Root -eq "") { $Root = Split-Path -Parent $PSScriptRoot }

$srcDir = Join-Path $Root "src"
if (-not (Test-Path -LiteralPath $srcDir)) { Write-Host ("FAIL missing src dir: " + $srcDir); exit 1 }

# 只认**代码行**: 以 // 、/* 、* 开头的整行不算, 其余行裁掉行内 // 之后才算。
$code = @{}
foreach ($f in Get-ChildItem -LiteralPath $srcDir -Recurse -Include *.cpp,*.h,*.inc -File) {
    $keep = New-Object System.Collections.Generic.List[string]
    foreach ($line in [System.IO.File]::ReadAllLines($f.FullName)) {
        $t = $line.TrimStart()
        if ($t.StartsWith("//") -or $t.StartsWith("/*") -or $t.StartsWith("*")) { continue }
        $idx = $line.IndexOf("//")
        if ($idx -ge 0) { $line = $line.Substring(0, $idx) }
        $keep.Add($line)
    }
    $code[$f.FullName] = $keep
}

function CountWhere([string]$needle, [string]$fileName) {
    $n = 0
    foreach ($kv in $code.GetEnumerator()) {
        if ($fileName -and ($kv.Key -notlike ("*" + $fileName))) { continue }
        foreach ($line in $kv.Value) {
            if ($line.Contains($needle)) { $n++ }
        }
    }
    return $n
}

# 取某个函数的花括号平衡块体 (认字符串字面量 —— 那两发 view_from 里各有一行
#   c_.emitLine("... (void* instance, const char* className) {"); ，串里那个 '{' 不算深度)。
function Get-FnBlock([string]$text, [string]$sig) {
    if ($text -eq '') { return $null }
    $i = $text.IndexOf($sig)
    if ($i -lt 0) { return $null }
    $open = $text.IndexOf('{', $i)
    if ($open -lt 0) { return $null }
    $depth = 0
    $inStr = $false
    for ($k = $open; $k -lt $text.Length; $k++) {
        $ch = $text[$k]
        if ($inStr) {
            if ($ch -eq '\') { $k++; continue }
            if ($ch -eq '"') { $inStr = $false }
            continue
        }
        if ($ch -eq '"') { $inStr = $true; continue }
        if ($ch -eq '{') { $depth++ }
        elseif ($ch -eq '}') {
            $depth--
            if ($depth -eq 0) { return $text.Substring($open + 1, $k - $open - 1) }
        }
    }
    return $null
}

$fail = @()

$v1 = CountWhere 'compare(0, 10, "vb6_iface_")' "cgen_base_type.cpp"
$v1all = CountWhere 'compare(0, 10, "vb6_iface_")' ""
if ($v1 -ne 1 -or $v1all -ne 1) { $fail += ("V1 前缀名单: 期望 1 处且住在 cgen_base_type.cpp, 实到 本文件=" + $v1 + " 全仓=" + $v1all) }

$v2 = CountWhere "cParamLegacyIfaceSlotName" ""
if ($v2 -ne 3) { $fail += ("V2 cParamLegacyIfaceSlotName: 期望 3 处(定义+声明+站点), 实到 " + $v2) }
if ((CountWhere "cParamLegacyIfaceSlotName" "cgen_expr_call_arg_emit.inc") -ne 1) {
    $fail += "V2 站点: 实参位那一问恰好 1 处才许包视图"
}

$v3 = CountWhere "legacyIfaceWrapFn(" ""
if ($v3 -ne 8) { $fail += ("V3 wrap 主人(定义+声明+调用)位数: 期望 8 = 6 处调用 + 定义 + 声明, 实到 " + $v3) }

$v4 = CountWhere "legacyViewBoxExpr(" ""
if ($v4 -ne 5) { $fail += ("V4 装箱主人(定义+声明+调用)位数: 期望 5 = 3 处调用 + 定义 + 声明, 实到 " + $v4) }

$v5 = CountWhere "slotIsLegacyIfaceView(" ""
if ($v5 -ne 6) { $fail += ("V5 认视图槽(定义+声明+调用)位数: 期望 6 = 4 处调用 + 定义 + 声明, 实到 " + $v5) }

$v6a = CountWhere ").obj == (" "cgen_expr_binary.cpp"
$v6b = CountWhere ").obj == NULL" "cgen_expr_binary.cpp"
if ($v6a -ne 1 -or $v6b -ne 2) {
    $fail += ("V6 视图比较发码: 期望 ).obj == ( =1 / ).obj == NULL =2, 实到 " + $v6a + " / " + $v6b)
}
if ((CountWhere ").obj == " "") -gt 3) { $fail += "V6 视图比较发码在别处又长了读者" }

# V7/V8 开线 (2026-10-23, 账 #305 第一百七十四刀 = 新账 #471 站A): For Each 那一站把集合交回来的元素
#   换成视图, 靠的是「认视图槽」那一问 (slotLegacyIfaceViewName = 名册只读一次, slotIsLegacyIfaceView
#   从此是它的投影) 与「按运行期类名选中那一对 wrap」的那一发。两条都不许在站点自己长出来。
#   V7  slotLegacyIfaceViewName( 在代码行里恰好 8 处 = 定义 1 + 声明 1 + bool 投影 1 + For Each 那一站 1 + 视图调用主人自己那一问 1
#        + With 那一站 1 (第一百七十七刀站D′) + 站F 属性读那一位读者 1 (第一百七十八刀),
    #        (新账 #474 写侧又添一位读者 ⇒ 位数见下面 V7 那条针)
#        且 cgen_control.cpp / cgen_with.cpp 各恰好 1 处、cgen_util_classtype.cpp 恰好 5 处 (定义 + bool 投影 + 方法主人那一问 + 站F 读侧读者 + 站F 写侧读者)。
#   V8  比名那把尺 `_stricmp(className` 全仓恰好 2 处, 且**按发射器分别钉住**:
#        · emitLegacyIfaceViewFrom (cgen_iface_vtbl.cpp:597) —— legacy 值视图那一发 (接口自己铸的 view_from);
#        · emitIvrefIfaceViewFrom  (cgen_iface_vtbl.cpp:638) —— 新式薄指针视图的**兄弟发射器**
#          (R4 收口, 提交 43528bb8)。
#        两处问的是**同一把尺** (运行期 className 是不是这枚实现类) = 同问同一权威, 不是放宽
#        (台账 §56.38 A7 先例); 站点自己比名 = 第二条答案; mint 外再比一次名 = 又一位主人。
$v7 = CountWhere "slotLegacyIfaceViewName(" ""
if ($v7 -ne 14) { $fail += ("V7 认视图槽那一问位址: 期望 14 = 定义 + 声明 + bool 投影 + For Each 那一站 + 调用主人自己那一问 + With 那一站 + 站F 读侧读者 + 站F 写侧读者 + 带参写槽只读查询 (第186刀) + With 体内零参读 (第193刀) + Set 装填那一站 (第215刀) + 「Is」 那一站 (第248刀: 认完名册再用 ivrefCType 问形状), 实际到 " + $v7) }
if ((CountWhere "slotLegacyIfaceViewName(" "cgen_control.cpp") -ne 1) {
    $fail += "V7 站点: For Each 那一站只许问一次认视图槽那一问"
}
if ((CountWhere "slotLegacyIfaceViewName(" "cgen_with.cpp") -ne 1) {
    $fail += "V7 站点: With 那一站 (站D') 只许问一次认视图槽那一问"
}
if ((CountWhere "slotLegacyIfaceViewName(" "cgen_util_classtype.cpp") -ne 6) {
    $fail += "V7 站点: 视图那一族自己这一份文件里, 认视图槽那一问只许六处 (定义 + bool 投影 + 方法主人那一问 + 站F 读侧读者 + 站F 写侧读者 + 带参写槽查询)"
}
$v8 = CountWhere "_stricmp(className" ""
if ($v8 -ne 2) { $fail += ("V8 视图那一族的比名尺: 期望 2 = legacy 值视图那一发 + 新式薄指针那一发 (同问同一权威, 不是放宽), 实际到 " + $v8) }
# 按**发射器**分别钉住 (两条都在 cgen_iface_vtbl.cpp 里 —— 旧措辞「比名尺住错了文件」是误报:
#   两条同住一份文件, 真正要钉的是「哪个发射器有几处」)。
$vtblKey = ($code.Keys | Where-Object { $_ -like "*cgen_iface_vtbl.cpp" } | Select-Object -First 1)
$vtblText = if ($vtblKey) { ($code[$vtblKey] -join "`n") } else { "" }
$blkLegacy = Get-FnBlock $vtblText "void CCodeGen::emitLegacyIfaceViewFrom("
$blkIvref  = Get-FnBlock $vtblText "void CCodeGen::emitIvrefIfaceViewFrom("
if ($null -eq $blkLegacy -or $null -eq $blkIvref) {
    $fail += "V8 读不到那两个发射器的函数体 (cgen_iface_vtbl.cpp 的 emitLegacyIfaceViewFrom / emitIvrefIfaceViewFrom)"
} else {
    $v8l = @([regex]::Matches($blkLegacy, '_stricmp\(className')).Count
    $v8i = @([regex]::Matches($blkIvref,  '_stricmp\(className')).Count
    if ($v8l -ne 1) { $fail += ("V8 legacy 值视图那一发 (emitLegacyIfaceViewFrom) 的比名尺: 期望 1, 实到 " + $v8l) }
    if ($v8i -ne 1) { $fail += ("V8 新式薄指针那一发 (emitIvrefIfaceViewFrom, 提交 43528bb8) 的比名尺: 期望 1, 实到 " + $v8i) }
}
# V9 开线 (2026-10-23, 账 #305 第一百七十五刀 = 新账 #471 站D): 「接收者是一枚视图时那一发方法调用怎么拼」
#   收成一处 = legacyViewCallExpr。病灶形状 = 表达式路现场拼了一份 (P6.4)，语句路压根没问、直接落
#   vb6_ComCall(视图, …) ⇒ 整枚结构体进 void* 槽 = cl C2172 (真语料 24 行)。只许 定义 + 声明 + 两路各一问。
#   语句路那一问住在 cgen_call.cpp 的 COM 兜底之前 (与 C29-SL-l 那条「语句路也要问一次」同规)；
#   With 那一头 (站D′) 由第一百七十七刀接上，接的**不是形状**: 它把 With 临时量登记进名册，
#   于是体内成员由上面那两路主人接手 —— 所以 V9 的位数一枚都不许因为站D′ 而涨。
#   第一百七十八刀站F 的读者是**第三位消费者** (属性读的形状也交回这位主人, 站点自己一枚都不拼)。
$v9 = CountWhere "legacyViewCallExpr(" ""
if ($v9 -ne 7) { $fail += ("V9 视图调用主人位数: 期望 7 = 定义 + 声明 + 表达式路 + 语句路 + 站F 读侧读者 (值档 + 对象档两发) + 站F 写侧读者, 实际到 " + $v9) }
if ((CountWhere "legacyViewCallExpr(" "cgen_call.cpp") -ne 1) {
    $fail += "V9 语句路: cgen_call.cpp 只许问一次视图调用主人"
}
if ((CountWhere "legacyViewCallExpr(" "cgen_expr_call_com_bind.inc") -ne 1) {
    $fail += "V9 表达式路: com_bind 只许问一次视图调用主人 (别把形状又现场拼一遍)"
}
if ((CountWhere "legacyViewCallExpr(" "cgen_util_classtype.cpp") -ne 4) {
    $fail += "V9 主人自己这一份文件: 只许定义 + 站F 读侧 (值档 + 对象档) / 写侧读者各一发, 别处再拼一份形状就是第二位主人"
}
# V10 开线 (2026-10-23, 账 #305 新账 #474 写侧): 属性写那一站的读者 = 视图调用主人的第四位消费者;
# V11 开线 (2026-10-23, 第一百八十一刀): 「VB 型 ↔ 晚绑定打包器」这一张配对表只许住一处。
#   正向 comPackFnForVbType = 定义 + 声明 + 两处消费 (comPackExpr 的类型分支 / 裸控件名折默认属性) = 4 处;
#   反向 comVbTypeForPackFn = 定义 + 声明 + With 体内那一问 = 3 处;
#   表本体 kComPackTiers = 3 处 (表定义 + 两个循环)、且只住 cgen_util_com.cpp 一份文件。
#   谁在别处再抄一份 switch / 名单，这两问就会在不同站点给出不同答案。
#   ⚠ V11 (正向 comPackFnForVbType 期望 4 / 实到 5) 曾经是**真回归**, 不是判据漂移 ——
#     多出来的那处 (src/backend/cgen_util_ctrl.cpp:2569, 提交 356f21ae = F1 收口引入) 要改
#     **发码** (新增消费点绕过既有权威), 另起一刀收口。**第五十六刀 (2026-11-12) 已收口**:
#     发码侧新增唯一一问 comPackFnForDefaultPropType (cgen_util_com.cpp, 住权威那份文件),
#     把「折叠值只认那六档」这道闸与那张配对表合成一问 —— 两处消费点
#     (comPackExpr 的「裸控件名折默认属性」那一支 + emitUcDefaultPropLet 写那一腿) 都问它,
#     于是 comPackFnForVbType 的读者回到 2 位 (正向问位数仍是 4 = 定义 + 声明 + 两处消费),
#     V11 的**判据一个字没动**。形状零改动 (逐字节证, 见 cgen_util_com.cpp 那处注释)。
#   V11b (第五十六刀新钉): 新问句 comPackFnForDefaultPropType = 定义 1 + 声明 1 + 消费 2 = 4,
#     且住处只许 cgen_util_com.cpp / cgen_helpers.inc / cgen_util_ctrl.cpp 三份 ——
#     新增第三位读者 = 又一处绕过权威, 立刻红。
$v11r = CountWhere "comVbTypeForPackFn(" ""
if ($v11r -ne 3) { $fail += ("V11 打包器反向问位数: 期望 3 = 定义 + 声明 + With 站点一问, 实际到 " + $v11r) }
$v11f = CountWhere "comPackFnForVbType(" ""
if ($v11f -ne 4) { $fail += ("V11 打包器正向问位数: 期望 4 = 定义 + 声明 + 两处消费, 实际到 " + $v11f) }
$v11t = CountWhere "kComPackTiers" ""
if ($v11t -ne 3) { $fail += ("V11 配对表本体位数: 期望 3 = 表定义 + 正向循环 + 反向循环, 实际到 " + $v11t) }
if ((CountWhere "kComPackTiers" "cgen_util_com.cpp") -ne 3) {
    $fail += "V11 配对表住处: 三处全部只许在 cgen_util_com.cpp (别在别处抄第二份)"
}
# V11b: 「折叠值该按哪一档装箱」这一问 (第五十六刀新增的那一问) —— 读者名单是闭的。
$v11b = CountWhere "comPackFnForDefaultPropType(" ""
if ($v11b -ne 4) { $fail += ("V11b 折叠值装箱问位数: 期望 4 = 定义 + 声明 + 两处消费, 实际到 " + $v11b) }
$v11bAll = @()
foreach ($f in $code.Keys) {
    foreach ($ln in $code[$f]) { if ($ln.Contains("comPackFnForDefaultPropType(")) { $v11bAll += (Split-Path -Leaf $f) } }
}
$v11bFileSet = @($v11bAll | Sort-Object -Unique)
if (($v11bFileSet -join ',') -ne 'cgen_helpers.inc,cgen_util_com.cpp,cgen_util_ctrl.cpp') {
    $fail += ("V11b 折叠值装箱问的住处 = " + ($v11bFileSet -join ', ') + " (want cgen_helpers.inc, cgen_util_com.cpp, cgen_util_ctrl.cpp = 声明 + 权威 + UC 默认属性写那一腿)")
}
#   站点 (cgen_assign_prop_write.inc) 只许问一次, 形状仍在主人手里 (站点一枚都不拼)。
# V12 (第一百八十六刀): 「带参写槽的末形参是什么 C 类型」只许住一处查询 —— 定义 + 声明 + 一个站点
$v12 = CountWhere "legacyViewSetterLastParamCType(" ""
if ($v12 -ne 3) { $fail += ("V12 带参写槽查询位数: 期望 3 = 定义 + 声明 + 多入口参数站点一问, 实际到 " + $v12) }
# V13 (第一百八十八刀): 「这一档返回类型是不是值聚合」只许住一处判据 (`vb6_type_` ∪ `vb6_iface_`) ——
#   定义 + 声明 + Function / Property Get 两处返回槽 = 4; `vb6_iface_` 前缀比较仍只住 cgen_base_type.cpp (V1 守)。
#   第二百一十五刀再添一位读者 = 类工厂的**字段初始化**那一档 (视图字段槽没有 NULL 这一形) ⇒ 5。
$v13 = CountWhere "cTypeIsValueAggregate(" ""
if ($v13 -ne 5) { $fail += ("V13 值聚合判据: 期望 5 = 定义 + 声明 + 两处返回槽 + 类工厂字段初始化 (第215刀), 实际到 " + $v13) }
# V14 (第一百八十八刀): 「返回类型若为视图, 交回接口名」只许住一处 —— 定义 + 声明 + 两处登记站点 = 4。
$v14 = CountWhere "legacyIfaceNameOfCType(" ""
if ($v14 -ne 4) { $fail += ("V14 视图返回名: 期望 4 = 定义 + 声明 + 两处登记站点, 实际到 " + $v14) }
# V15 (第一百八十九刀): 表达式接收者的按值适配器名只许住一处 —— 定义 + 声明 + 铸 + 问 = 4。
$v15 = CountWhere "legacyIfaceSlotCallFn(" ""
if ($v15 -ne 6) { $fail += ("V15 视图按值适配器名: 期望 6 = 定义 + 声明 + 铸 + 调用站一问 + 裸零参读一问 + With 体内零参读一问, 实际到 " + $v15) }
if ((CountWhere "legacyViewSetterCallExpr(" "cgen_assign_com_prop.inc") -ne 1) {
    $fail += "V10 站点: 多入口参数那一站 (cgen_assign_com_prop.inc) 只许问一次视图写读者"
}
$v10 = CountWhere "legacyViewSetterCallExpr(" ""
if ($v10 -ne 8) { $fail += ("V10 视图属性写读者位数: 期望 8 = 定义 + 声明 + 块外站点一问 + With 体内站点一问 + Set 语句三处 + 多入口参数站点一处 (第186刀), 实际到 " + $v10) }
if ((CountWhere "legacyViewSetterCallExpr(" "cgen_assign_prop_write.inc") -ne 2) {
    $fail += "V10 站点: 属性写只许问两次视图写读者 = 块外一站 + With 体内一站 (别在站点拼 .vtbl-><槽>(….obj, 值))"
}
if ((CountWhere "legacyViewSetterCallExpr(" "cgen_setlet_set_prop.inc") -ne 3) {
    $fail += "V10 站点: Set 语句只许问三次视图写读者 = Nothing 那支 + 命名属性那支 + objRef 那支 (putref_ 档同一主人)"
}
if ($fail.Count -gt 0) {
    foreach ($m in $fail) { Write-Host ("FAIL " + $m) }
    exit 1
}
Write-Host ("PASS 前缀名单=" + $v1all + "(本文件 " + $v1 + ") 形参问句=" + $v2 + " wrap主人=" + $v3 + " 装箱主人=" + $v4 + " 认槽主人=" + $v5 + " 比较发码=" + $v6a + "+" + $v6b + " 认视图槽那一问=" + $v7 + " 比名尺=" + $v8 + " 视图调用主人=" + $v9 + " 写读者=" + $v10 + " 打包器配对=" + $v11f + "/" + $v11r + "/" + $v11t + " 值聚合=" + $v13 + " 视图返回名=" + $v14 + " 适配器名=" + $v15)
exit 0
