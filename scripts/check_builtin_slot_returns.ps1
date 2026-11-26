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

# ---- 7) R8..R11: P1.4 三问各一组正反用例 ----
# 这三条刻意**各问各的事**, 不许三条都只测同一件:
#   R8  = 生成物与 RTL 派生同源 (手改 .inc 或改了 RTL 没重跑 gen 脚本 ⇒ 红)
#   R9  = 问句①「C 文本是不是整只槽」: 表里要有真槽出口(正), 地址出口不许进表(反),
#         AST 那一半必须排除 Variant 数组本体(反)
#   R10 = 问句③「VB 类型是不是 Variant」: 只吃 AST(正), **不许**排除数组本体(反) ——
#         这一条与 R9 相反, 两条同时绿才证明「类型」与「整只槽」真是两个问句
#   R11 = 问句②「是不是 SafeArray 载体」: 名单命中 vb6_Split( (正), 标量出口不许命中(反),
#         _arr_ 特例仍在(正)
#   R12 = 数组本体 (Array/Split/Filter, VarType=8204) 的分歧: isDefinitelyVariantExpr 必须
#         答 Variant 且置 isArrOut=true ⇒ 类型问句 true / 槽问句 false (P2 的 C1)
$genScript = Join-Path $PSScriptRoot 'gen_variant_slot_prefixes.ps1'
if (-not (Test-Path -LiteralPath $genScript)) {
    $viol += ('R8 找不到派生脚本 ' + $genScript)
} else {
    $genOut = (& powershell -NoProfile -ExecutionPolicy Bypass -File $genScript -Check 2>&1 | Out-String)
    if ($LASTEXITCODE -ne 0) {
        $viol += ('R8 派生不一致: checked-in .inc 与从 RTL 重新派生的结果不同 (手改 .inc, 或改了 RTL 没重跑 gen 脚本) —— ' + $genOut.Trim())
    }
}

$typeSrc = Strip-CComments ([IO.File]::ReadAllText($typeCpp))
$astM = [regex]::Match($typeSrc, '(?s)bool CCodeGen::variantOperandAST\(const Expr\* e\) const \{(?<b>.*?)\r?\n\}')
$typeM = [regex]::Match($typeSrc, '(?s)bool CCodeGen::isVariantType\(const Expr\* e\) const \{(?<b>.*?)\r?\n\}')
$carM = [regex]::Match($typeSrc, '(?s)bool CCodeGen::isSafeArrayCarrier\(const std::string& cExpr\) const \{(?<b>.*?)\r?\n\}')
if (-not $astM.Success -or -not $typeM.Success -or -not $carM.Success) {
    $viol += 'R9-R11 三问的函数体没找全 (改名/换签名了? 这一组正反用例会静默失效)'
} else {
    $astBody = $astM.Groups['b'].Value
    $typeBody = $typeM.Groups['b'].Value
    $carBody = $carM.Groups['b'].Value
    if (-not $tbl.ContainsKey('vb6_VariantFromValue')) { $viol += 'R9 槽问句正例缺失: 表里没有 vb6_VariantFromValue' }
    if ($tbl.ContainsKey('vb6_VariantArrayElemPtr')) { $viol += 'R9 槽问句反例失败: vb6_VariantArrayElemPtr 交回的是**地址**不是槽, 却进了槽表' }
    if ($astBody -notmatch 'isDefinitelyVariantExpr\([\s\S]*?&isArr') { $viol += 'R9 槽问句反例失败: variantOperandAST 没向 isDefinitelyVariantExpr 要 isArr' }
    if ($astBody -notmatch '!\s*isArr') { $viol += 'R9 槽问句反例失败: variantOperandAST 没排除 Variant 数组本体 (!isArr)' }
    if ($typeBody -notmatch 'isDefinitelyVariantExpr\(') { $viol += 'R10 类型问句正例缺失: isVariantType 不调 isDefinitelyVariantExpr' }
    if ($typeBody -match '!\s*isArr') { $viol += 'R10 类型问句反例失败: isVariantType 也排除数组本体 —— 两个问句又合成一个' }
    if ($carBody -notmatch '"vb6_Split\("') { $viol += 'R11 载体问句正例缺失: 名单里没有 vb6_Split(' }
    if ($carBody -match '"vb6_ChkLong\("') { $viol += 'R11 载体问句反例失败: 标量出口 vb6_ChkLong( 进了载体名单' }
    if ($carBody -notmatch '"_arr_"') { $viol += 'R11 载体问句正例缺失: _arr_ 特例不见了 (Fix 092g)' }
}

# ---- R12: Array/Split/Filter 是「类型 Variant、不是整只槽」的实物分歧 ----
# 判据（P2，C1）：VB6 里 VarType(Array(1,2)) = 8204 (vbVariant + vbArray)，Split/Filter 同理。
# 所以 isDefinitelyVariantExpr 必须把这三枚**当 Variant**（类型 true），同时置 isArrOut=true，
# 让 variantOperandAST 的 !isArr 把它们从「槽」里排除。这句分歧是两张问句分开的最强证据:
# 同一枚 `Array(1,2,3)` 上 isVariantType 必须答 true、isVariantSlot 必须答 false。
# 只查源结构（不起 cl）: 三枚名字 + 那条 isArrOut 赋值 + 不许落回 kBuiltinFuncNames 的拒绝。
$defM = [regex]::Match($typeSrc, '(?s)bool CCodeGen::isDefinitelyVariantExpr\((?<sig>[^\n]*)\{(?<b>.*?)\r?\n\}')
if (-not $defM.Success) {
    $viol += 'R12 找不到 isDefinitelyVariantExpr 的函数体 (改名/换签名了? 这条断言会静默失效)'
} else {
    $defBody = $defM.Groups['b'].Value
    foreach ($bn in @('"array"', '"split"', '"filter"')) {
        if ($defBody -notmatch [regex]::Escape($bn)) {
            $viol += ('R12 类型问句正例缺失: isDefinitelyVariantExpr 里没有数组本体内置 ' + $bn +
                      ' —— VarType=8204 的那一族会被当成「非 Variant」')
        }
    }
    if ($defBody -notmatch '\*isArrOut\s*=\s*true') {
        $viol += 'R12 槽问句反例缺失: 数组本体内置没置 isArrOut=true —— 它们会同时被当成整只槽'
    }
}

# ---- R13: Format 数值段的字面量/零段/前导 0 与 Percent 分组 (§11.167 第一百八十四刀) ----
# 判据 (真机 P197): H1/H3/H4/H5/H6/H7 的字面量必须出现; H9 零段只打字面量;
# H15 `#.##` 的 0.5 = `.5`; H10/H11 Percent 不带千分位。
$tailSrc = $null
$tailPath = Join-Path $Root 'src/rtl/core/vb6rtl/detail/vb6rtl_format_numeric_tail.inc'
if (Test-Path -LiteralPath $tailPath) { $tailSrc = Get-Content -LiteralPath $tailPath -Raw -Encoding UTF8 }
if ($null -eq $tailSrc) { $viol += 'R13 找不到 Format 数值尾部片段' }
else {
    if ($tailSrc -notmatch '§11\.167') { $viol += 'R13 字面量拼装没有登记 (§11.167 段标缺)' }
    foreach ($t in @('prefix', 'suffix', 'dropLeadZero', 'lastPh')) {
        if ($tailSrc -notmatch [regex]::Escape($t)) { $viol += ('R13 拼装缺 ' + $t) }
    }
    if ($tailSrc -match 'int dotEmitted') { $viol += 'R13 旧的「不定位字面量」那段还在 (dotEmitted 是它的标记)' }
}
$parsePath = Join-Path $Root 'src/rtl/core/vb6rtl/detail/vb6rtl_format_parse.inc'
$parseSrc = $null
if (Test-Path -LiteralPath $parsePath) { $parseSrc = Get-Content -LiteralPath $parsePath -Raw -Encoding UTF8 }
if ($null -eq $parseSrc) { $viol += 'R13 找不到 Format 解析片段' }
else {
    if ($parseSrc -notmatch 'vb6_FormatPercent\(numVal, 2, -1, 0, 0\)') {
        $viol += 'R13 Percent 命名格式必须不带千分位 (groupDigits=0)'
    }
}
$runSrcFmt = Join-Path $Root 'tests/run_tests.ps1'
if (Test-Path -LiteralPath $runSrcFmt) {
    $runTxt = Get-Content -LiteralPath $runSrcFmt -Raw -Encoding UTF8
    $fmtCells = ([regex]::Matches($runTxt, "Test-Vbp 'b192_format_literal")).Count
    if ($fmtCells -ne 2) { $viol += ('R13 b192 夹具格数 = ' + $fmtCells + ', 要 2 (x64 + x86)') }
    $fmtMain = Join-Path $Root 'tests/b192_format_literal/B192Main.bas'
    if (-not (Test-Path -LiteralPath $fmtMain)) { $viol += 'R13 夹具 B192Main.bas 不在' }
    else {
        $fmtTxt = Get-Content -LiteralPath $fmtMain -Raw -Encoding UTF8
        foreach ($w in @('H9', 'H15')) {
            if ($fmtTxt -notmatch ('"' + $w + '"')) { $viol += ('R13 夹具缺证人 ' + $w) }
        }
    }
}

# ---- R14: 「谁先进格式串, 谁赢」与区末 e 记号 (§11.168 第一百八十五刀) ----
# 判据 (真机 P200~P203): 0/# 先出现 ⇒ 数值档 (字母全是字面量); 字母先出现或串里有冒号
# ⇒ 日期档 (反向证人 U16..U20)。另: 区末那枚 e/E 算指数记号 (U1 = "1234 Byt")。
$fmtCPath = Join-Path $Root 'src/rtl/core/vb6rtl/vb6rtl_format.c'
$fmtC = $null
if (Test-Path -LiteralPath $fmtCPath) { $fmtC = Get-Content -LiteralPath $fmtCPath -Raw -Encoding UTF8 }
if ($null -eq $fmtC) { $viol += 'R14 找不到 vb6rtl_format.c' }
else {
    if ($fmtC -notmatch '§11\.168') { $viol += 'R14 「谁先出现」那道没有登记 (§11.168 段标缺)' }
    foreach ($t in @('firstNum', 'firstStrong', 'hasColon')) {
        if ($fmtC -notmatch $t) { $viol += ('R14 判别缺 ' + $t) }
    }
    if ($fmtC -notmatch 'if \(firstNum >= 0 && \(firstStrong < 0 \|\| firstNum < firstStrong\)\) return 0;') {
        $viol += 'R14 数值优先的判据丢了 (0/# 先出现时未回 0)'
    }
}
if ($null -ne $tailSrc) {
    if ($tailSrc -notmatch "§11\.168") { $viol += 'R14 区末 e 记号没有登记 (§11.168 段标缺)' }
    if ($tailSrc -notmatch "i == secLen - 1") { $viol += 'R14 区末 e/E 记号的判据丢了' }
}
if (Test-Path -LiteralPath $runSrcFmt) {
    $unitCells = ([regex]::Matches($runTxt, "Test-Vbp 'b193_numeric_unit_text")).Count
    if ($unitCells -ne 2) { $viol += ('R14 b193 夹具格数 = ' + $unitCells + ', 要 2 (x64 + x86)') }
    $unitMain = Join-Path $Root 'tests/b193_numeric_unit_text/B193Main.bas'
    if (-not (Test-Path -LiteralPath $unitMain)) { $viol += 'R14 夹具 B193Main.bas 不在' }
    else {
        $unitTxt = Get-Content -LiteralPath $unitMain -Raw -Encoding UTF8
        foreach ($w in @('U1', 'U16', 'U19')) {
            if ($unitTxt -notmatch ('"' + $w + '"')) { $viol += ('R14 夹具缺证人 ' + $w) }
        }
    }
}

# ---- R15: 整数区多组占位符的数字分配 (§11.169 第一百八十六刀) ----
# 判据 (真机 P204): 整串先插千分位、右往左按容量切组、最左组吸收余数、空组按 0/# 补位;
# 只在没有小数点时启用。G18/G20 是反向证人。
if ($null -ne $tailSrc) {
    if ($tailSrc -notmatch '§11\.169') { $viol += 'R15 多组分配没有登记 (§11.169 段标缺)' }
    foreach ($t in @('gCap', 'gZero', 'gsrc', 'slStart', 'slEnd')) {
        if ($tailSrc -notmatch $t) { $viol += ('R15 多组分配缺 ' + $t) }
    }
    if ($tailSrc -notmatch 'if \(dotPos < 0\)') { $viol += 'R15 多组分配必须限定在无小数点那一族' }
    if ($tailSrc -notmatch 'if \(ng >= 2') { $viol += 'R15 多组分配的入口判据丢了' }
}
if (Test-Path -LiteralPath $runSrcFmt) {
    $grpCells = ([regex]::Matches($runTxt, "Test-Vbp 'b194_multi_int_group")).Count
    if ($grpCells -ne 2) { $viol += ('R15 b194 夹具格数 = ' + $grpCells + ', 要 2 (x64 + x86)') }
    $grpMain = Join-Path $Root 'tests/b194_multi_int_group/B194Main.bas'
    if (-not (Test-Path -LiteralPath $grpMain)) { $viol += 'R15 夹具 B194Main.bas 不在' }
    else {
        $grpTxt = Get-Content -LiteralPath $grpMain -Raw -Encoding UTF8
        foreach ($w in @('G7', 'G8', 'G21')) {
            if ($grpTxt -notmatch ('"' + $w + '"')) { $viol += ('R15 夹具缺证人 ' + $w) }
        }
    }
}

# ---- R16: FormatNumber / FormatCurrency / FormatPercent 的十进制半值进位 (§11.170) ----
# 判据 (真机 P207/P208): 不能回到 C 运行库的 %.*f (二进制五成双), 必须走十进制文本进位。
$cmpPath = Join-Path $Root 'src/rtl/core/vb6rtl/vb6rtl_compat.c'
$cmpSrc = $null
if (Test-Path -LiteralPath $cmpPath) { $cmpSrc = Get-Content -LiteralPath $cmpPath -Raw -Encoding UTF8 }
if ($null -eq $cmpSrc) { $viol += 'R16 找不到 vb6rtl_compat.c' }
else {
    if ($cmpSrc -notmatch '§11\.170') { $viol += 'R16 十进制定型没有登记 (§11.170 段标缺)' }
    foreach ($t in @('vb6_fmtRoundDecimalText', 'vb6_fmtFixedText', '%.15g')) {
        if ($cmpSrc -notmatch [regex]::Escape($t)) { $viol += ('R16 定型链缺 ' + $t) }
    }
    foreach ($fn in @('vb6_FormatCurrency', 'vb6_FormatNumber', 'vb6_FormatPercent')) {
        $body = [regex]::Match($cmpSrc, ('BSTR ' + $fn + '\([\s\S]{0,1500}')).Value
        if ($body -eq '') { $viol += ('R16 找不到 ' + $fn) }
        elseif ($body -notmatch 'vb6_fmtFixedText') { $viol += ($fn + ' 没走十进制定型') }
    }
}
if (Test-Path -LiteralPath $runSrcFmt) {
    $rndCells = ([regex]::Matches($runTxt, "Test-Vbp 'b195_format_round_decimal")).Count
    if ($rndCells -ne 2) { $viol += ('R16 b195 夹具格数 = ' + $rndCells + ', 要 2 (x64 + x86)') }
    $rndMain = Join-Path $Root 'tests/b195_format_round_decimal/B195Main.bas'
    if (-not (Test-Path -LiteralPath $rndMain)) { $viol += 'R16 夹具 B195Main.bas 不在' }
    else {
        $rndTxt = Get-Content -LiteralPath $rndMain -Raw -Encoding UTF8
        foreach ($w in @('R1', 'R4', 'R10', 'B9')) {
            if ($rndTxt -notmatch ('"' + $w + '"')) { $viol += ('R16 夹具缺证人 ' + $w) }
        }
    }
}

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
Write-Host 'PASS builtin_slot_returns (R1..R16)'
exit 0

