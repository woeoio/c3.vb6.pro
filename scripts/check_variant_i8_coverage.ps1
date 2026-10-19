# check_variant_i8_coverage.ps1 - 「指针宽度那一档属于 Long 同族」这条决定的覆盖面哨兵 (账 #305 / §B138 尾巴)
#
# 为什么单开这一道：同一个问句（「这枚 Variant 装的算不算数值 / 该报什么类型名 / 打包成什么档」）在
# src\rtl 里以 `switch (v.vt)` 的形状实现了**十一遍**。§B138 把 VarPtr/StrPtr/ObjPtr 的类型答案换成
# LongPtr 之后，x64 第一次真的往这些表里喂 VT_I8 —— 于是四处当场答错：vb6_TypeName 落 default 报
# "Variant"、vb6_IsNumeric 报假、vb6_ComPackVariant 落 default 把地址**静默打包成 VT_EMPTY**、
# UC 宿主模型两枚取数读回 0（后两处按不同读数只算一处）。补 case 只是止血；这一道钉的是**闭合 census**：
# 以后任何人新写一张 vt 表，要么进 must 名单并答这一档，要么进 exempt 并写明理由，否则门红。
#
# 规则（改坏了会红，不是装饰）：
#   V1  src\rtl 里发现的 vt switch 站点数 >= 13 —— 突然变少或数出 0 = glob/正则坏了，那比红更糟
#   V2  闭合性：每一处站点的宿主函数必须**要么在 must、要么在 exempt**；两边都没有 ⇒ 红
#   V3  must 名单里每个函数体必须真有 VT_I8 档（挂了名字没写档 = 静默漏答，正是本账的实物）
#   V4  名单（must + exempt + decMust + decExempt）里每个名字都必须在 src\rtl 里真的存在（拼错的名字会让规则断电）
#   V5  must 名单不许空着或缩水到 < 11；decMust 不许缩水到 < 8 (§B145 把 Decimal 的名单加到 8 枚)
#   D1  decMust 里每个函数体必须真有 VT_DECIMAL 档 —— 与 V3 同形，只是问的是另一档
#   D2  闭合性按**档**算：每一处站点的宿主函数在 Decimal 这一档也必须二选一（decMust 或 decExempt）
#   D3  decExempt 的每条理由不许是空话（少于 12 字 ⇒ 红），豁免必须说清「为什么不答这一档」
#   C1  「往里读」那一族的核只许一枚 switch，两枚公开薄壳的体里不许再出现 switch
#   C2  值档(BSTR + x86 DECIMAL)的「谁分配谁释放」只许住 vb6_VariantValueFree / ValueCopy 这一对;
#       宿主环那五站(取回/清值/存集合/取一份/删)必须转问它们，且体里不许再出现自己的 vt 表
#
# 账 #305 §B145 之后「槽 -> 真 VARIANT」那一族只有一张表(vb6_VariantPackBorrow)：出站打包(V3 名单里
# 原来是 vb6_ComPackVariant)与 Variant 算术都读它，所以名单换名不换数量——两处各写一遍的那些年，
# 就是这张表连漏 Decimal/Currency/Date 三档的来历。
# 站点匹配前会剥掉注释(§B145: 引用了 `switch (v.vt)` 的那句注释自己变成一个站点，
# 还被算给上一个函数 —— phantom 站点比漏站点更坏，因为它凭空造出一格「要豁免」。
# 口径从「一档 × 每表」换成「每档 × 每表」的来历（账 #305 §B143）：VT_I8 那一撞之后 Decimal 是第二次，
# 而两次都只钉了一档 —— 补了 I8 的表照样可以漏 Decimal。现在同一张表要答两档，名单各列各的。
# 账 #305 §B142 (第五十三刀): 上面那句 ⚠ 收了 —— 站点现在认**任何**接收者拼法(`p->vt` / `in->vt` /
    #   `pv->vt` / `varElem.vt` / `retVal.vt` / `expr.vt` / 裸 `comVt` 都在内)，且 `.inc` 一起扫(Format 那张
    #   分诊表就住在 detail/*.inc 里，改前压根不在 census 眼皮底下)。放宽之后站点 13 -> 19，
    #   每一张都重新逐张归过因(must 里 13 名 / exempt 里 9 名)，V1 的下限跟着钉成 19。
    #   ⚠ 放宽正则之后**再写一张新表就一定红**：它必须在某一档上二选一(答它 / 写明为什么不答)。
    # 账 #305 §B171: vb6_bag_deserialize 那张表(UC 属性包反序列化, uc_propbag.c)就是这条 ⚠ 的实物 ——
    #   它按 iVarType 分档解 (name,value), 漏 I8/Decimal 时那两项**静默不进表**, .ctl 侧只读回默认值
    #   (与「抹成 0」同形: 错的不是值, 是那一项根本不存在)。两档都答了 —— 顺带把 VB6 真会写出的
    #   VT_CY(6) / VT_DATE(7) 也补上, 那两档不在这道哨兵的钉子上, 但同一个 default 会吃掉它们。
# 账 #305 §B165 (第五十四刀): 上面那 19 张里有一张是宿主存值腿(vb6_Host_FromWinVariant)的九档表,
#   它当年进 exempt 是因为 x86 的 Decimal 槽里装的是一枚指针、借用回来存进集合就悬空(实测读回
#   2066035336255469780997 那一类数)。本刀撤了那张表(改问 §B144 的布局权威 + 本刀新立的所有权出口),
#   站点 19 -> **18**，两条豁免随表一起撤;「这一格从此怎么办」由新规则 C2 钉住 ——
#   环上五站转问那一对出口, 谁再手写一档(或再养一张自己的 vt 表)就红。
#
# 用法:  powershell -File scripts\check_variant_i8_coverage.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$rtl  = Join-Path $root "src\rtl"
if (-not (Test-Path -LiteralPath $rtl)) {
    Write-Host ("FAIL rtl dir missing: " + $rtl) -ForegroundColor Red
    exit 1
}
$bad = @()

$must = @(
    "vb6_VariantToLong", "vb6_VariantToLongPtr", "vb6_VariantToBool", "vb6_VariantToDouble",
    # 账 #305 §B155 第四十六刀: 名单随表走 —— TypeName 那张表整个住在 vb6_TypeNameScalar 里了
    # (数组那一支现在也拼同一串元素名)，vb6_TypeName 自己不再有 switch，挂着它的名字等于挂一条死规则。
    "vb6_TypeNameScalar", "vb6_IsNumeric", "vb6_VariantPackBorrow", "memVariantToI4",
    "vb6_ho_variantToLong", "vb6_ho_variantToDouble", "vb6_VarToDouble_internal",
    # 账 #305 §B142 (第五十三刀): 往里读那一族的两枚公开名以前各抄一遍 23 档表(134 行 + 121 行)，
    # 现在表只住 vb6_ComVariantIntoSlot 一枚(下面 C1 钉这件事)；Format 那张分诊表住在 .inc 里，
    # 放宽扫描后它进了 census，两档都答(§B142 第三十四刀补的 Decimal + Task #44 补的 I8)。
    "vb6_ComVariantIntoSlot", "vb6_Format",
    # 账 #305 §B171: UC 属性包反序列化那张表(见上面同名的注释)
    "vb6_bag_deserialize"
)
$exempt = @{
    "vb6_CDec" = "两架构读数本来就相同（CStr(CDec(整数)) 今天就是空串），缺的不是这一档 -> 另立 §B141"
    "vb6_ResIdOf" = "资源 id 的实参解包：只认 BSTR 与数值档，而 default 交 lVal(不是 0)，与本账「漏档静默交 0」不同形；Decimal 档全仓 0 个调用点"
    # 账 #305 §B142 (第五十三刀): 放宽正则之后新见光的三张，逐张归因如下。
    "vb6_EnumNext" = "IEnumVARIANT 取元素那一族，全仓(src + backend + 哨兵)0 个调用者的死码；真 For Each 走 vb6_ForEach_Next。撤不撤由用户拍，先别给它写档"
    "vb6_SaElemTypeFromVt" = "这不是槽表，是 SAFEARRAY 的**元素档**表(§B159 与尺寸表/标签表同族三张)；I8/Decimal 那两档枚举里本来没有，回落 variant 是 §B155 写死的口径"
# 账 #305 §B165 (第五十四刀): vb6_Host_FromWinVariant 那两条豁免随它那张九档表一起撤了。
    "pb_Read" = "ActiveX 属性包(IPropertyBag::Read)只交回**设计期存过的属性**，写侧那几档由 C3 自己的下发决定；语料里 I8/Decimal 走到属性包的读数 0 处，要补先证明可达"
}

# Decimal 这一档（§B143）：同一张表要答两档，所以名单也列两份。
$decMust = @(
    "vb6_VariantToLong", "vb6_VariantToLongPtr", "vb6_VariantToBool", "vb6_VariantToDouble",
    "vb6_TypeNameScalar", "vb6_IsNumeric", "vb6_VarToDouble_internal",
    # 账 #305 §B145: 「槽 -> 真 VARIANT」那张表唯一的住处(出站打包与 Variant 算术都读它)。
    "vb6_VariantPackBorrow",
    # 账 #305 §B142 (第五十三刀): 往里读那一枚核 + Format 那张分诊表(见上面 I8 名单同两条的理由)。
    "vb6_ComVariantIntoSlot", "vb6_Format",
    # 账 #305 §B171: 同上 —— 属性包那张表两档都答
    "vb6_bag_deserialize"
)
$decExempt = @{
    # 账 #305 §B145: 这一格撤掉——那张 vt 表整个搬进 vb6_VariantPackBorrow(它现在两档都答)，
    # vb6_ComPackVariant 只剩「借用 -> 带权」那一行升级，自己不再有 switch。
    "memVariantToI4" = "UC 宿主模型的取数口，全仓 0 个 Decimal 调用点（§B143 量过），要补先证明可达"
    "vb6_ho_variantToLong" = "UC 宿主模型的取数口，全仓 0 个 Decimal 调用点（§B143 量过），要补先证明可达"
    "vb6_ho_variantToDouble" = "UC 宿主模型的取数口，全仓 0 个 Decimal 调用点（§B143 量过），要补先证明可达"
    "vb6_CDec" = "它是造 Decimal 的那一档，不是读表；补它等于把同一个问题抄第二遍"
    "vb6_ResIdOf" = "资源 id 的实参解包：LoadRes*(CDec(...)) 全仓 0 处，而 default 交的是 lVal 不是 0"
    # 账 #305 §B142 (第五十三刀): 同一批新见光的三张(理由与 I8 名单那三条同形，逐条写不合并)。
    "vb6_EnumNext" = "IEnumVARIANT 取元素那一族，全仓 0 个调用者的死码(§B142 ⑤ 量的)，死码上不写档"
    "vb6_SaElemTypeFromVt" = "SAFEARRAY 元素档表而不是槽表：枚举里没有 Decimal 那一档，回落 variant 是 §B155 写死的口径"
# 账 #305 §B165 (第五十四刀): 同上 —— 宿主存值腿不再有表, 它的收支由 C2 那五条锚点钉住。
    "pb_Read" = "ActiveX 属性包(IPropertyBag::Read)只把**设计期存过的属性**交回容器，写侧那几档由 C3 自己的下发决定；语料里 Decimal/I8 走到属性包的读数 0 处，要补先证明可达"
}

# ai/032 §6.9: 小整型两档（VT_I1 / VT_UI2）—— 同一张表要答的第四、五档，名单也各列一份。
# 与 I8/Decimal 的区别：这两档在 C3 里**只在显式 SByte / UInteger 表达式装箱时**才产生，
# 而 SByte=int8_t(=char)、UInteger=uint16_t(=wchar_t)，以前都被 _Generic 接走装成 VT_I4 ——
# 所以「漏答」在既往语料里读不出来（值对、TypeName 答 "Long"），是本账第一次把这两档摆上台面。
$smMust = @(
    "vb6_VariantToLong", "vb6_VariantToLongPtr", "vb6_VariantToBool", "vb6_VariantToDouble",
    "vb6_TypeNameScalar", "vb6_IsNumeric", "vb6_VarToDouble_internal", "vb6_VariantPackBorrow",
    "vb6_ComVariantIntoSlot", "vb6_Format", "memVariantToI4", "vb6_ho_variantToLong", "vb6_ho_variantToDouble"
)
$smExempt = @{
    "vb6_CDec"            = "造 Decimal 的那一档不是读表；补它等于把同一个问题抄第二遍（与 decExempt 同理由）"
    "vb6_EnumNext"        = "IEnumVARIANT 取元素，全仓 0 个调用者的死码；死码上不写档（与 I8 豁免同理由）"
    "vb6_bag_deserialize" = "UC 属性包按 **iVarType 标签**分档解(name,value)，不是槽表；这两档在包格式里没有编码，补它等于改线格式"
    "vb6_ResIdOf"         = "资源 id 的实参解包只认 BSTR 与数值档，且 default 交 lVal(不是 0)；语料里这两档走到它 0 处"
    "vb6_SaElemTypeFromVt"= "SAFEARRAY 的元素档表不是槽表；枚举里没有 SByte/UInteger 两档，回落 variant 是既有口径"
    "pb_Read"             = "ActiveX 属性包只交回设计期存过的属性；UI2 那档它本来就答，I1 要补先证明语料可达"
}

$fnRe     = [regex]'^[A-Za-z_][A-Za-z0-9_ \*]*?([A-Za-z_][A-Za-z0-9_]*)\s*\([^;]*\)\s*\{\s*$'
# 接收者叫 v 的 vt switch，**允许带强转**：`switch ((vb6_vartype)v->vt)` 与 `switch (v.vt)` 是同一族表。
# 以前这条正则不认强转写法，于是「比较走哪个数值」那一口(vb6_VarToDouble_internal)整个不在 census 里。
# 账 #305 §B155 第四十六刀: 再认一枚**裸形参** `switch (vt)` —— vb6_TypeName 那张名单抽成了 static
# 出口 vb6_TypeNameScalar(数组那一支要与标量拼同一个答案)，表还在、只是不再带 `v.` 前缀。不放宽的话
# census 静默少一张表(V1 的站点数 13 -> 12)，而「站点少了」正是这道哨兵唯一看得见名单断电的途径。
# 接收者拼法一律认: `<名>.vt` / `<名>->vt`，以及**变量名本身就以 vt 结尾**的裸形参(`vt` / `comVt`)。
# 强转写法照旧支持(`switch ((vb6_vartype)v->vt)`)。
$switchRe = [regex]'switch\s*\(\s*(?:\([^)]*\)\s*)?(?:[A-Za-z_][A-Za-z0-9_]*(?:\.vt|->vt)|(?:[A-Za-z_][A-Za-z0-9_]*[Vv][Tt]|vt))\s*\)'
# .inc 里的表没有函数签名可依附，宿主函数按这张名数表认领(新加一个 .inc 里的表而不登记，
# 站点会以「无人认领」报红 —— 宁可红在名单上，也不要静默少一张表)。
$incOwner = @{ 'vb6rtl_format_extract.inc' = 'vb6_Format' }
# 档的写法有两种拼法：SDK 名与 vb6_ 前缀名，且可能带强转 ⇒ 按同一个模式认。
$armRe = @{
    I8  = [regex]'case\s+(?:\([^)]*\)\s*)?(?:VT_I8|vb6_vtI8)\s*:'
    DEC = [regex]'case\s+(?:\([^)]*\)\s*)?(?:VT_DECIMAL|vb6_vtDecimal)\s*:'
    # ai/032 §6.9: C3 扩展的小整型两档。两档都**值不缺**（8/16 位塞得进 VT_I4），缺的是**类型身份** ——
    #   以前被 _Generic 的 `char:` / `wchar_t:` 档接走装成 VT_I4，TypeName 答 "Long"。与 I8/Decimal 同口径
    #   逐张 vt 表都要么答要么写明豁免；正转 (VT_I1) 两档本体都要这枚 arm。
    I1  = [regex]'case\s+(?:\([^)]*\)\s*)?(?:VT_I1|vb6_vtSByte|vb6_vtByte1)\s*:'
    UI2 = [regex]'case\s+(?:\([^)]*\)\s*)?(?:VT_UI2|vb6_vtUInteger)\s*:'
}

# ---------- 扫一遍：站点 + 每个函数的体 ----------
$files = @(Get-ChildItem -LiteralPath $rtl -Recurse -File -Include *.c, *.inc)
$sites = @()
$bodies = @{}
$allNames = @($must + $decMust + $smMust + @($exempt.Keys) + @($decExempt.Keys) + @($smExempt.Keys) + 
               @('vb6_VariantFromComResult', 'vb6_VariantFromStackVARIANT',
                 # 账 #305 §B165 (第五十四刀): C2 要读这一对出口与宿主环五站的体 —— 名字不在这张表里
                 # 就进不了 $bodies, C2 会红在「函数不存在」上(那是假红, 不是发现)。
                 'vb6_VariantValueFree', 'vb6_VariantValueCopy',
                 'vb6_Host_FromWinVariant', 'vb6_Host_ClearVariant',
                 'vb6_coll_push', 'vb6_coll_itemCopy',
                 'vb6_Collection_Remove')) | Select-Object -Unique
foreach ($f in $files) {
    $text  = [System.IO.File]::ReadAllText($f.FullName)
    $lines = $text -split "`r?`n"
    $cur   = ""
    for ($i = 0; $i -lt $lines.Count; $i++) {
        # 账 #305 §B145: 注释里**引用**的形状不是站点。第三十六刀往 vb6rtl_com.c 写了一句
        # 「接收者叫 v 的 switch (v.vt) 是这道哨兵认的形状」，于是那句注释自己变成了一个
        # 站点，还被算给上一个函数(vb6_VariantToObject)—— phantom 站点会把闭合 census 变成
        # 一份凭空的豁免名单，所以先把注释剥掉再匹配。
        $ln = ($lines[$i] -replace '/\*.*?\*/', '')
        $t = $ln.TrimStart()
        if ($t.StartsWith('*') -or $t.StartsWith('//') -or $t.StartsWith('/*')) { continue }
        $m = $fnRe.Match($ln)
        if ($m.Success) { $cur = $m.Groups[1].Value }
        if ($switchRe.IsMatch($ln)) {
            # 账 #305 §B142: .inc 里的表没有签名可依附 => 按 $incOwner 认领；没登记就留空，
            # 由下面的 V2/D2「could not be attributed」报红(宁可红，不要静默少一张表)。
            $who = $cur
            if ($f.Extension -eq '.inc' -and $incOwner.ContainsKey($f.Name)) { $who = $incOwner[$f.Name] }
            $sites += [pscustomobject]@{ File = $f.FullName.Substring($root.Length + 1)
                                         Line = $i + 1; Fn = $who }
        }
    }
    foreach ($name in ($allNames)) {
        if ($bodies.ContainsKey($name)) { continue }
        $sig = [regex]::Match($text, '(?m)^[\w \*]*' + [regex]::Escape($name) + '\s*\([^\r\n;]*\)\s*\{')
        if (-not $sig.Success) { continue }
        $depth = 0; $from = $sig.Index + $sig.Value.Length - 1; $to = -1
        for ($k = $from; $k -lt $text.Length; $k++) {
            $c = $text[$k]
            if ($c -eq '{') { $depth++ }
            elseif ($c -eq '}') { $depth--; if ($depth -eq 0) { $to = $k; break } }
        }
        if ($to -le $from) { continue }
        $body = $text.Substring($from, $to - $from + 1)
        # 账 #305 §B142: 体里的 `#include "detail/xxx.inc"` 要**就地展开** —— Format 那张分诊表整个
        # 住在 .inc 里(六个 .inc 在函数体内被 include，非独立编译单元)，不展开的话 must 名单会红在
        # 「这一档没写」上，而那其实是规则看不见它住的地方。
        foreach ($im in [regex]::Matches($body, '#include\s*"([^"]+\.inc)"')) {
            # 片段名**不带** detail/ 前缀(编译器的 include 路径里有那一层)，所以按文件名找。
            $ip = Join-Path $f.DirectoryName $im.Groups[1].Value
            if (-not (Test-Path -LiteralPath $ip)) {
                $incName = Split-Path -Leaf $im.Groups[1].Value
                $hit = @(Get-ChildItem -LiteralPath $f.DirectoryName -Recurse -File -Filter $incName)
                if ($hit.Count -ge 1) { $ip = $hit[0].FullName }
}
            if (Test-Path -LiteralPath $ip) { $body += [System.IO.File]::ReadAllText($ip) }
        }
        $bodies[$name] = $body
    }
}

# ---------- V1: 覆盖面 ----------
# 账 #305 §B165 (第五十四刀): 19 -> 18。少的那一张是宿主存值腿的九档表, 它整个撤了 —— 不是漏数:
# V2/D2 若看见「无人认领」的表会另报红, 而那条环的五个站点由 C2 钉着。
if ($sites.Count -lt 18) {
    $bad += ("V1 only " + $sites.Count + " vt switches found under src\rtl (floor 18) -> the glob or the " +
             "regex broke; the closure checks below would be checking almost nothing")
}

# ---------- 每一档各问一遍：闭合 census + 名单自己 ----------
$arms = @(
    [pscustomobject]@{ Tag = "V2/V3"; Arm = "VT_I8";      Must = $must;    Exempt = $exempt;    Floor = 13; Re = $armRe.I8 },
    [pscustomobject]@{ Tag = "D1/D2"; Arm = "VT_DECIMAL"; Must = $decMust; Exempt = $decExempt; Floor = 10;  Re = $armRe.DEC },
    # ai/032 §6.9: 两档共用一组名单（每张表要么答正转+无符号，要么两者都豁免），故两组用同一份 must/exempt。
    [pscustomobject]@{ Tag = "S1/S2"; Arm = "VT_I1";      Must = $smMust;  Exempt = $smExempt;  Floor = 12; Re = $armRe.I1 },
    [pscustomobject]@{ Tag = "S1/S2"; Arm = "VT_UI2";     Must = $smMust;  Exempt = $smExempt;  Floor = 12; Re = $armRe.UI2 }
)
foreach ($a in $arms) {
    # 闭合性：每张 vt 表都要在这一档二选一(答它 / 写明为什么不答)
    foreach ($s in $sites) {
        if ($s.Fn -eq "") {
            $bad += ($a.Tag + " " + $s.File + ":" + $s.Line + " switch could not be attributed to a function " +
                     "-> the census can't tell whether it answers the " + $a.Arm + " 档; check the signature shape")
            continue
        }
        if (($a.Must -notcontains $s.Fn) -and (-not $a.Exempt.ContainsKey($s.Fn))) {
            $bad += ($a.Tag + " " + $s.File + ":" + $s.Line + " in " + $s.Fn + ": a vt switch that is in " +
                     "neither list for the " + $a.Arm + " 档 -> it drops that 档 into default. Either add it " +
                     "to the must list AND write the case, or to the exempt list with the reason it doesn't need one")
        }
    }
    # must 名单：名字要在，档也要在
    foreach ($name in $a.Must) {
        if (-not $bodies.ContainsKey($name)) {
            $bad += ("V4 must-list entry " + $name + " has no function body in src\rtl -> renamed or deleted; " +
                     "this rule is now dead power")
            continue
        }
        if (-not $a.Re.IsMatch($bodies[$name])) {
            $bad += ($a.Tag + " " + $name + ": in the must-list for the " + $a.Arm + " 档 but its body has no " +
                     "case for it -> this table silently answers that 档 through default")
        }
    }
    # exempt 名单：名字还要在(删了函数就撤豁免)，理由还要写得住
    foreach ($name in @($a.Exempt.Keys)) {
        if (-not $bodies.ContainsKey($name)) {
            $bad += ("V4 exempt entry " + $name + " is not a function in src\rtl anymore -> drop the exemption " +
                     "together with the function")
        }
        if (([string]$a.Exempt[$name]).Length -lt 12) {
            $bad += ("D3 exempt entry " + $name + " (" + $a.Arm + " 档) has a reason shorter than 12 chars -> an " +
                     "empty reason turns the exemption into a silent hole in the census")
        }
    }
    # 针面非空
    if ($a.Must.Count -lt $a.Floor) {
        $bad += ("V5 " + $a.Arm + " must-list holds " + $a.Must.Count + " names (floor " + $a.Floor +
                 ") -> the needle surface collapsed; if these tables really merged into one authority, " +
                 "delete this sentinel together with them")
    }
}

# ---------- C1: 「往里读」那一族只许一枚写者 (账 #305 §B142 第五十三刀) ----------
# 改前两枚公开名各带一张 23 档的表(134 行 + 121 行)，剥掉注释逐行 diff 只剩四处不同、全是所有权。
# 现在表只许住在 vb6_ComVariantIntoSlot 一枚里；两枚公开名是**薄壳**，体里再出现 switch 就是第二次实现。
$coreSwitches = 0
if ($bodies.ContainsKey('vb6_ComVariantIntoSlot')) {
    $coreSwitches = ([regex]::Matches($bodies['vb6_ComVariantIntoSlot'], 'switch\s*\(')).Count
} else {
    $bad += 'C1 the one inbound core vb6_ComVariantIntoSlot has no body in src\rtl -> the census is blind'
}
if ($coreSwitches -ne 1) {
    $bad += ('C1 inbound core carries ' + $coreSwitches + ' switches (want exactly 1) -> the vt table is ' +
             'being written a second time inside the core')
}
foreach ($shell in @('vb6_VariantFromComResult', 'vb6_VariantFromStackVARIANT')) {
    if (-not $bodies.ContainsKey($shell)) {
        $bad += ('C1 public wrapper ' + $shell + ' vanished -> codegen emits it by name; check before deleting')
        continue
    }
    $n = ([regex]::Matches($bodies[$shell], 'switch\s*\(')).Count
    if ($n -ne 0) {
        $bad += ('C1 ' + $shell + ' has its own switch again (' + $n + ') -> two writers for the same question ' +
                 '(this is exactly the shape §B142 counted: 23 arms written twice, diverging on ownership)')
    }
}

# ---------- C2: 值档的「谁分配谁释放」只许住那一对出口 (账 #305 §B165 第五十四刀) ----------
# 改前这一问住两遍(vb6_VariantClear 里的 rev11 三行 + vb6_VariantAssign 里的 rev16 一段)，而夹在中间的
# vb6_VariantCopy 压根不认 Decimal ⇒ 任何「只 Copy 不 Assign」的路径在 x86 上都是把 src 那 16 字节的
# 指针搬进 dst，两边各 Clear 一次 = 二次 free。这一条按形状还能推出两处调点(数组整赋 / ArraySetVariant)
# 也坏，但 A/B 量不出来(§B165 那一格的六行读数两台相同)；量到真坏的是宿主环那五站 —— 摘掉这一对里的
# DECIMAL 分配，x86 当场 0xC0000374。五站今天全部转问这一对出口。
$pairMust = [ordered]@{
    'vb6_VariantValueFree' = @('vb6_vtBSTR', 'CoTaskMemFree')
    'vb6_VariantValueCopy' = @('vb6_vtBSTR', 'CoTaskMemAlloc')
}
foreach ($fn in @($pairMust.Keys)) {
    if (-not $bodies.ContainsKey($fn)) {
        $bad += ('C2 value-ownership authority ' + $fn + ' has no body in src\rtl -> the pair collapsed')
        continue
    }
    foreach ($tok in $pairMust[$fn]) {
        if ($bodies[$fn] -notmatch [regex]::Escape($tok)) {
            $bad += ('C2 ' + $fn + ': body lacks ' + $tok + ' -> the BSTR/DECIMAL ownership question is ' +
                     'being hand-copied somewhere again (that is the shape this knife closed)')
        }
    }
}
$ringStations = [ordered]@{
    'vb6_Host_FromWinVariant' = 'vb6_VariantValueCopy'
    'vb6_Host_ClearVariant'   = 'vb6_VariantValueFree'
    'vb6_coll_push'           = 'vb6_VariantValueCopy'
    'vb6_coll_itemCopy'       = 'vb6_VariantValueCopy'
    'vb6_Collection_Remove'   = 'vb6_VariantValueFree'
}
foreach ($fn in @($ringStations.Keys)) {
    $ask = $ringStations[$fn]
    if (-not $bodies.ContainsKey($fn)) {
        $bad += ('C2 host-ring station ' + $fn + ' has no body in src\rtl -> this ring calls it by name ' +
                 'from five places; check before deleting')
        continue
    }
    if ($bodies[$fn] -notmatch [regex]::Escape($ask)) {
        $bad += ('C2 ' + $fn + ' no longer asks ' + $ask + ' -> a value 档 is being copied by hand a ' +
                 'second time; the ring has exactly one authority for alloc/free')
    }
    if ($bodies[$fn] -match 'switch\s*\(') {
        $bad += ('C2 ' + $fn + ' owns a switch again -> the host ring must not keep its own vt table ' +
                 '(that table is what dropped Date/Currency/I8/Decimal into default)')
    }
}

if ($bad.Count -eq 0) {
    Write-Host ("PASS variant_i8_coverage: switches " + $sites.Count + " I8 must " + $must.Count +
                " exempt " + $exempt.Count + " | DEC must " + $decMust.Count + " exempt " +
                $decExempt.Count + " | I1 must " + $smMust.Count + " | UI2 must " + $smMust.Count +
                " exempt " + $smExempt.Count +
                " | inbound core 1 | value pair 2 ring 5 | violations 0") -ForegroundColor Green
    exit 0
}
foreach ($b in $bad) { Write-Host ("FAIL " + $b) -ForegroundColor Red }
exit 1
