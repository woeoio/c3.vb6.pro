# check_builtin_slot_returns.ps1 —— 「内置出口交出来的 C 是不是整只槽」这张表与 RTL 对账（账 #305 §B146，第三十九刀）
#
# 为什么要有这一道：variantPrefixes（名单本体在 checked-in 生成物
# src/backend/detail/expr/cgen_variant_slot_prefixes.inc，由 scripts/gen_variant_slot_prefixes.ps1
# 从 RTL 派生）回答的那个问句（「vb6_X(...) 交回来的是 vb6_VARIANT 槽还是裸标量」）的答案
# **住在 RTL 的声明里**。这份表与 RTL 两个方向都会走散，且都贵：
#   少一名 = 消费点按裸标量发码。实物（第三十九刀量到，BASE 真编译 6 条错）：
#          `CDec(1234567) + 1` 发成 (int32_t)(vb6_CDec(...) + 1) => C2088「+ 对于 struct 非法」，
#          `d = CDec("2.5")` => C2440。整条夹具编不过。
#   多一名 = 死名躺在表里骗人。本刀清掉 11 条（RTL 没有这枚符号、400 份语料发码里一次没出现）。
#   错位一格 = 指针返回当槽。`vb6_VARIANT *vb6_VariantArrayElemPtr(...)` 交回的是**槽的地址**，
#          按名字登记它就是把「地址」答成「槽」（本刀的 census 第二版栽过一次，故 R3 单独钉）。
#
# 判据面（任何一条不齐就红）：
#   R1 表里每一枚都必须真是「RTL 声明返回 vb6_VARIANT」的那一族 —— 例外只有两张白名单，
#      每条都要写理由：RTL 的 _Generic 宏（只有每一档变目都交槽才许登记），以及
#      **不是 RTL 出口**的名字（语料工程自己的过程在单模块模式下拼成裸 vb6_<proc>）。
#   R2 RTL 里「声明返回 vb6_VARIANT 且后端发码字符串里出现过」的每一枚都必须在表里。
#   R3 指针返回（vb6_VARIANT *f(...)）一律不许进表。
#   R4 census must not read empty: table >=25, rtlSlot >=25, backend literal names >=50,
#      少一个就报「扫描没匹配到东西」（红比 vacuous PASS 好）。
#   R5 「这串已发出的 C 是不是晚绑定槽的元素读」只许在 cgen_util_type.cpp 的那张
#      kSlotElemForms 里答: src/backend 任何别的文件在**字符串字面量**里拼
#      vb6_VariantArray(Get|GetND|Set|SetND|ElemPtr|ElemPtrND)( 就红。实物（第四十刀开格前
#      量到）: 那一问原先由六处手写字符串比较各答一遍, 抄的还是两种前缀 (带尾括号 / 不带),
#      于是新增多维那枚名进去之后带括号那四处对它一律答 false —— VarPtr(v(1,2)) 退成
#      复合字面量 (C2440)、Set v(1,2) = x 退成整条赋值被丢掉 (静默)。钉的是「不许有第二份
#      抄法」, 不是「抄了几处」; 剥注释 + 只看字面量, 与 §B145 那条课同口径。
# 用法: powershell -File scripts\check_builtin_slot_returns.ps1   (PASS = exit 0)

$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot
$typeCpp = Join-Path $Root 'src\backend\cgen_util_type.cpp'
if (-not (Test-Path -LiteralPath $typeCpp)) { Write-Host ('FAIL: 读不到 ' + $typeCpp); exit 1 }
# 名单本体已是 checked-in 生成物 (scripts/gen_variant_slot_prefixes.ps1 从 RTL 派生)。
$tblInc = Join-Path $Root 'src\backend\detail\expr\cgen_variant_slot_prefixes.inc'
if (-not (Test-Path -LiteralPath $tblInc)) { Write-Host ('FAIL: 读不到 ' + $tblInc); exit 1 }

function Strip-CComments([string]$t) {
    $t = [regex]::Replace($t, '/\*[\s\S]*?\*/', ' ')
    $t = [regex]::Replace($t, '(?m)//.*$', ' ')
    return $t
}

# ---- 1) 那张表（生成物 .inc） ----
$srcRaw = [IO.File]::ReadAllText($tblInc)
$m = [regex]::Match($srcRaw, '(?s)static const std::vector<std::string> variantPrefixes = \{(?<body>.*?)\};')
if (-not $m.Success) { Write-Host 'FAIL: 在 cgen_variant_slot_prefixes.inc 里找不到 variantPrefixes 的初始化块（表被改名/挪走了？）'; exit 1 }
$tbl = @{}
foreach ($q in [regex]::Matches($m.Groups['body'].Value, '"(vb6_[A-Za-z0-9_]+)\("')) { $tbl[$q.Groups[1].Value] = $true }
$tableNames = @($tbl.Keys | Sort-Object)
if ($tableNames.Count -lt 25) { Write-Host ('FAIL R4: 表只读出 ' + $tableNames.Count + ' 条 (<25) —— 扫描读空, 不是一条绿着的判据'); exit 1 }

# ---- 2) RTL: 谁声明返回 vb6_VARIANT ----
$rtlSlot = @{}; $rtlPtr = @{}; $rtlMacro = @{}
$rtlFiles = @(Get-ChildItem -Path (Join-Path $Root 'src\rtl') -Recurse -File -Include *.c,*.h)
foreach ($f in $rtlFiles) {
    $t = Strip-CComments ([IO.File]::ReadAllText($f.FullName))
    foreach ($k in [regex]::Matches($t, '(?m)^\s*#\s*define\s+(vb6_[A-Za-z0-9_]+)')) { $rtlMacro[$k.Groups[1].Value] = $true }
    foreach ($k in [regex]::Matches($t, 'vb6_VARIANT\s*\*\s*(vb6_[A-Za-z0-9_]+)\s*\(')) { $rtlPtr[$k.Groups[1].Value] = $true }
    foreach ($k in [regex]::Matches($t, 'vb6_VARIANT\s+(vb6_[A-Za-z0-9_]+)\s*\(')) {
        $n = $k.Groups[1].Value
        if (-not $rtlPtr.ContainsKey($n)) { $rtlSlot[$n] = $true }
    }
}
if ($rtlSlot.Keys.Count -lt 25) { Write-Host ('FAIL R4: RTL 槽声明只数出 ' + $rtlSlot.Keys.Count + ' 枚 (<25) —— 扫描读空'); exit 1 }

# ---- 3) 后端: 哪些名字真出现在发码字面量里 ----
$emitted = @{}
$beFiles = @(Get-ChildItem -Path (Join-Path $Root 'src\backend') -Recurse -File -Include *.cpp,*.inc)
foreach ($f in $beFiles) {
    $t = Strip-CComments ([IO.File]::ReadAllText($f.FullName))
    foreach ($lit in [regex]::Matches($t, '"(?:[^"\\]|\\.)*"')) {
        foreach ($k in [regex]::Matches($lit.Value, '(vb6_[A-Za-z0-9_]+)')) { $emitted[$k.Groups[1].Value] = $true }
    }
}
if ($emitted.Keys.Count -lt 50) { Write-Host ('FAIL R4: backend literals gave only ' + $emitted.Keys.Count + ' names (min 50) -- scan read nothing'); exit 1 }

# ---- 白名单（每一格都要有理由，不是「先让它绿」） ----
$allowMacro = @{ 'vb6_VariantFromValue' = '_Generic 宏: 每一档变目都交槽 (含 vb6_VARIANT: vb6_VariantIdentity 档)' }
$allowNonRtl = @{ 'vb6_DispCallByVtbl' = '不是 RTL 出口 —— 语料 mdQRCodegen.bas 自己的 Private Function, 单模块模式下 C 名是裸 vb6_<proc>, variantReturnFuncs_ 够不着 (撤掉它会丢五处 vb6_VariantToLong 提取, 见 §B146)' }

$viol = @()

# R1: 表 -> RTL
foreach ($n in $tableNames) {
    if ($rtlSlot.ContainsKey($n)) { continue }
    if ($allowMacro.ContainsKey($n) -and $rtlMacro.ContainsKey($n)) { continue }
    if ($allowNonRtl.ContainsKey($n)) { continue }
    $why = 'RTL 里没有这枚符号返回 vb6_VARIANT'
    if ($rtlPtr.ContainsKey($n)) { $why = 'RTL 里它是**指针返回** (交回槽的地址, 不是槽) —— 见 R3' }
    elseif ($rtlMacro.ContainsKey($n)) { $why = 'RTL 里它是宏, 且不在 allowMacro 的逐档核过名单里 —— 宏按名字登记会一律答错' }
    $viol += ('R1 死名/错档: ' + $n + '(' + ' —— ' + $why)
}

# R2: (RTL 槽 ∩ 后端 emit) -> 表
foreach ($n in @($rtlSlot.Keys | Sort-Object)) {
    if (-not $emitted.ContainsKey($n)) { continue }
    if ($tbl.ContainsKey($n)) { continue }
    $viol += ('R2 漏登记: ' + $n + '( —— RTL 声明返回 vb6_VARIANT, 后端发码字面量里出现过, 但表里没有 (消费点会按裸标量发码)')
}

# R3: 指针返回一律不许在表里
foreach ($n in @($rtlPtr.Keys | Sort-Object)) {
    if ($tbl.ContainsKey($n)) { $viol += ('R3 指针返回被登记成槽: ' + $n + '(') }
}

# ---- 5) R5: 槽元素读的「认文本」那一问只许住在权威那张文件里 ----
$slotLit = 'vb6_VariantArray(?:Get|GetND|Set|SetND|ElemPtr|ElemPtrND)\('
$typeFull = (Resolve-Path -LiteralPath $typeCpp).Path
# 名单生成物 (.inc) 不是「第二份抄法」: 它只是槽前缀数据, 不回答「是不是元素读」那一问。
$tblFull = (Resolve-Path -LiteralPath $tblInc).Path
$authHits = 0
$litRe = '"(?:[^"\\]|\\.)*"'
foreach ($f in $beFiles) {
    if ($f.FullName -eq $tblFull) { continue }
    $t = Strip-CComments ([IO.File]::ReadAllText($f.FullName))
    foreach ($lit in [regex]::Matches($t, $litRe)) {
        if (-not [regex]::IsMatch($lit.Value, $slotLit)) { continue }
        if ($f.FullName -eq $typeFull) { $authHits++; continue }
        $viol += ('R5 第二份抄法: ' + (Split-Path -Leaf $f.FullName) + ' 的字面量里拼了 ' + $lit.Value +
                  ' —— 认文本那一问只在 cgen_util_type.cpp 的 kSlotElemForms 里答 (把这段改写换成问那两个出口)')
    }
}
if ($authHits -lt 3) { Write-Host ('FAIL R5: 权威文件里只数出 ' + $authHits + ' 处槽元素读字面量 (<3) —— 扫描读空, R5 是一条绿着的假判据'); exit 1 }

# ---- 6) R6/R7: Print#/Write# 的字符串分流**不许再养名字名单**（第二步的删除不可回潮） ----
# 旧实现是 34 名 C 出口名表 (bstrFuncs + isBstrExpr)，两份 visitor 各抄一遍。删掉之后
# 判据只剩两问：inferExprType==String，以及发射形状（BSTR 槽/IIf 串支/数组元素/静态变量）。
# 名字表一旦回来，类型推断再退化也不会红 —— 那正是本刀要防的回潮。
$fileIoCpp = Join-Path $Root 'src\backend\stmt\cgen_file_io.cpp'
if (-not (Test-Path -LiteralPath $fileIoCpp)) { Write-Host ('FAIL R6: 读不到 ' + $fileIoCpp); exit 1 }
$fileIo = Strip-CComments ([IO.File]::ReadAllText($fileIoCpp))
if ($fileIo -match 'bstrFuncs|isBstrExpr') { $viol += 'R6 名字名单回潮: cgen_file_io.cpp 里又出现了 bstrFuncs / isBstrExpr（第二步已删；分流只许走两问）' }
# 第三刀 (账 #608/第331刀) 把两份 visitor 的 lambda 收成一位主人 fileOutArgIsBstr:
#   ① 语义类型 (inferExprType==String) 是主判据，形状只作补充，knownBstrVars_ 只兜底裸标识符;
#   ② 两条 visitor (Print#/Write#) 与窗体面 Print 必须都转调这一位主人 —— 各自再抄一份就是
#      第二个判据，名字/形状会重新变成主判据。
$ownerHits = [regex]::Matches($fileIo, 'bool\s+CCodeGen::fileOutArgIsBstr\(').Count
if ($ownerHits -ne 1) { $viol += ('R7 「这枚实参是不是 BSTR」的主人应为 1 枚，实测 ' + $ownerHits) }
$isStringItemHits = [regex]::Matches($fileIo, 'auto\s+isStringItem\s*=').Count
if ($isStringItemHits -ne 0) { $viol += ('R7 isStringItem 抄本回潮: ' + $isStringItemHits + ' 处 (分流口只许住 fileOutArgIsBstr)') }
if ($fileIo -notmatch 'inferExprType\(argNode\)\s*==\s*Vb6Type::String') { $viol += 'R7 主人不再问 inferExprType（类型那一问被拆掉，只剩形状猜测）' }
foreach ($needle in @('vb6_BSTR', 'vb6_IIfBSTR', 'VB6_SA_AT(BSTR,', 'knownBstrVars_')) {
    if ($fileIo -notmatch [regex]::Escape($needle)) { $viol += ('R7 主人缺发射形状判据: ' + $needle) }
}
$siteHits = [regex]::Matches($fileIo, 'fileOutArgIsBstr\(val,\s*\*expr\)').Count
if ($siteHits -ne 2) { $viol += ('R7 Print#/Write# 转调主人的站点应为 2 处，实测 ' + $siteHits) }
$formHits = [regex]::Matches($fileIo, 'fileOutArgBstrExpr\(val,\s*\*expr\)').Count
if ($formHits -ne 1) { $viol += ('R7 窗体面 Print 转调主人的站点应为 1 处，实测 ' + $formHits) }

Write-Host ('census: table=' + $tableNames.Count + ' rtlSlot=' + $rtlSlot.Keys.Count +
            ' rtlPtr=' + $rtlPtr.Keys.Count + ' rtlMacro=' + $rtlMacro.Keys.Count +
            ' emittedNames=' + $emitted.Keys.Count +
            ' (RTL 槽 ∩ emit) = ' + (@($rtlSlot.Keys | Where-Object { $emitted.ContainsKey($_) }).Count) +
            ' slotReadLit=' + $authHits)

if ($viol.Count -gt 0) {
    foreach ($v in $viol) { Write-Host ('  ' + $v) }
    Write-Host ('FAIL builtin_slot_returns: ' + $viol.Count + ' 条违规 —— 修法是把表与 RTL 对账 (补名/撤名), 不是往白名单里加名字; 真要走例外请连带写清理由并重跑 §B146 那把 400 份 A/B')
    exit 1
}
Write-Host 'PASS builtin_slot_returns (R1..R7)'
exit 0
