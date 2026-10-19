# check_implicit_cstr_station.ps1 —— 「数值实参交进 VB 的字符串形参 ⇒ 按 VB6 的隐式 CStr 折」这一条
# 规则的站点 / 名单 / 形状对账（账 #305 §B174，第六十五刀）
#
# 为什么要有这一道：这条规则在发码侧原本住**两家**——
#   Fix 175 (src/backend/detail/expr/cgen_expr_call_arg_emit.inc) 问**形参表** (ByVal String)，
#          形状走 wrapToBSTR；
#   Fix 108c (src/backend/detail/expr/cgen_expr_call_builtin_argtype.inc) 靠一张 13 枚名单，
#          自己写了一遍折算 if/else，而且做在 argList **已拼成一整串之后** —— 找回第 N 个实参
#          只能靠"从上次命中往后数子串"这种对齐。
# 两家的档位清单从不汇合（一家有 Byte / LongPtr，另一家有 SByte / UInteger / ULong / ULongLong /
# Decimal），Boolean 那一档还各自给出两种答案（wrapToBSTR 的 "True" 与名单自己写的 "-1"）。
# 第六十五刀收成一位，并补上这一位以前管不着的那一族（**故意没有形参表**的 InStr：位号答在
# builtinNoParamTableStringSlot 一处）。这一道钉的就是"别再长出第二位、别再把形状就地重写"。
#
# 判据（任何一条不齐就红）：
#   R1 站点唯一：typeTakesImplicitCStr 的**调用点**恰好 1 处，且住在 cgen_expr_call_arg_emit.inc
#      （声明 cgen_helpers.inc + 定义 cgen_util_type.cpp 各 1 处，是同一枚问句的两半，不算调用）。
#   R2 位号表唯一：builtinNoParamTableStringSlot 同一条，也必须恰好一处调用、同一文件。
#   R3 旧名单不许回来：strArgPos108c 在 src/ 里 0 处。
#   R4 那一站不再发形状：cgen_expr_call_builtin_argtype.inc 的**代码**里不许出现 vb6_CStr。
#   R5 名单自己的 cast 形状不许在别处重现：vb6_CStrLong((int32_t)( 在 src/backend 里 0 处。
#   R6 census 不许读空：档位闸的 case >= 14、**位号表认到的内在函数名 >= 11**（第六十六刀
#      从 1 枚扩到 11 枚 = InStr + 注册表四枚 + MsgBox / InputBox / SendKeys / SetAttr /
#      SavePicture / Format），arg_emit 里 wrapToBSTR 的调用 >= 1（形状出口还在被问）。
#   R7 第二位主人不许回来：cgen_expr_call_conv_output.inc 的**代码**里 wrapToBSTR 0 处 ——
#      那一站曾经按 args[0] 的文本前缀连判六条自己折 MsgBox 的 prompt（§B175 量过：它在自己
#      管的位上折得对，但**只管 prompt 一位**，标题位今天仍裸交 double）。同时钉
#      'callee = "vb6_MsgBox1"' 恰好 1 处 —— 那一形是 arity 改道，问的是另一件事，必须还在。
#
#   R11 字节数组的两位主人登记路数相等（各 7 条：模块级 2 + 局部 2 + 形参 3）。
#   R12 字节数组那两站的剥壳闸 = 2 条 `>= 4`，且 builtin_fixup 里不许再留 `> 4`。
#   R13 形参那一腿：三份声明文件里「按 VB 类型的裸 knownLongVars_.insert(pLower)」与
#      「knownIntVars_.insert(pLower)」必须同数（guarded 兜底臂只接 enum/未知档，不计入）。
#   R14 类字段那一腿（账 #305 §B183）：窄档表 classIntMembers_ 只许一份声明；扫描里按模块 clear、
#      两拼法各登一次，且这一登必须**住在放宽臂内**（把 Integer 从那张表里摘走会饿着放宽档的其他
#      消费者）；三处过程入口放宽档与窄档同数。
#   R15 「这枚左值有多窄 / 这一档要套哪只范围闸」只许有一个答案（账 #305 §B184，第七十五刀）：
#      三只闸的名字串只住那张表（3 处 / 1 文件）；表是被**问**的不是被抄的（3 个调用点，改前那枚
#      手写 cast 在注释剥净的代码里绝迹）；narrowCheckAssign 被 3 站问到（赋值两腿 + With 那一腿）；
#      成员臂的类字段一问恰好 1 次且住在 §B170 否决之后；With 那一站的裸形状绝迹而站点本身还在。
# 用法: powershell -File scripts\check_implicit_cstr_station.ps1 [-RootOverride <目录>]  (PASS = exit 0)
#   -RootOverride 是给负控用的: 把 src\backend 抄到别处再判, 真树一字不动。

param([string]$RootOverride = '')
$ErrorActionPreference = 'Stop'
$Root = if ($RootOverride) { $RootOverride } else { Split-Path -Parent $PSScriptRoot }
$backend = Join-Path $Root 'src\backend'
$utilCpp = Join-Path $backend 'cgen_util_type.cpp'
$helpers = Join-Path $backend 'detail\util\cgen_helpers.inc'
$emitInc = Join-Path $backend 'detail\expr\cgen_expr_call_arg_emit.inc'
$argtypeInc = Join-Path $backend 'detail\expr\cgen_expr_call_builtin_argtype.inc'
$convInc    = Join-Path $backend 'detail\expr\cgen_expr_call_conv_output.inc'
$argVariantInc = Join-Path $backend 'detail\expr\cgen_expr_call_arg_variant.inc'
foreach ($p in @($utilCpp, $helpers, $emitInc, $argtypeInc, $convInc, $argVariantInc)) {
    if (-not (Test-Path -LiteralPath $p)) { Write-Host ('FAIL: 读不到 ' + $p); exit 1 }
}

function Strip-CComments([string]$t) {
    $t = [regex]::Replace($t, '/\*[\s\S]*?\*/', ' ')
    $t = [regex]::Replace($t, '(?m)//.*$', ' ')
    return $t
}

# 取一段花括号平衡的块体 (start 指向块首 '{' 的位置)
function Get-BlockBody([string]$t, [int]$start) {
    $depth = 0
    for ($j = $start; $j -lt $t.Length; $j++) {
        $ch = $t[$j]
        if ($ch -eq '{') { $depth++ }
        elseif ($ch -eq '}') {
            $depth--
            if ($depth -eq 0) { return $t.Substring($start + 1, $j - $start - 1) }
        }
    }
    return $null
}

$allFiles = @(Get-ChildItem -Path $backend -Recurse -File |
              Where-Object { $_.Extension -eq '.cpp' -or $_.Extension -eq '.inc' })
$codeByFile = @{}
foreach ($f in $allFiles) { $codeByFile[$f.FullName] = Strip-CComments ([IO.File]::ReadAllText($f.FullName)) }
$util = $codeByFile[$utilCpp]
$emit = $codeByFile[$emitInc]
$argtype = $codeByFile[$argtypeInc]
$argVariant = $codeByFile[$argVariantInc]
$allCode = ($codeByFile.Values -join [Environment]::NewLine)

$srcRaw = ''
foreach ($f in @(Get-ChildItem -Path (Join-Path $Root 'src') -Recurse -File |
                 Where-Object { $_.Extension -eq '.cpp' -or $_.Extension -eq '.inc' -or $_.Extension -eq '.h' })) {
    $srcRaw += [IO.File]::ReadAllText($f.FullName) + [Environment]::NewLine
}

# 调用点 = 名字后面紧跟 '(' ，且前面不是"定义 (CCodeGen::)"也不是"声明 (bool / std::string)"
function Get-CallSites([string]$name) {
    $rx = [regex]([regex]::Escape($name) + '\(')
    $hits = @()
    foreach ($path in $codeByFile.Keys) {
        $t = $codeByFile[$path]
        foreach ($m in $rx.Matches($t)) {
            # $m.Index 指的是**名字**那一处, 所以 $pre 以名字前的那个词收尾:
            # 定义 = "…bool CCodeGen::" / 声明 = "…bool " —— 两者都不算调用点。
            $pre = $t.Substring(0, $m.Index)
            if ($pre -match 'CCodeGen::\s*$') { continue }
            if ($pre -match '(bool|std::string|size_t)\s+$') { continue }
            $hits += (Split-Path -Leaf $path)
        }
    }
    return $hits
}

$viol = @()

# ---- R1 / R2 ----
foreach ($pair in @(@('typeTakesImplicitCStr', 'cgen_expr_call_arg_emit.inc'),
                    @('builtinNoParamTableStringSlot', 'cgen_expr_call_arg_emit.inc'))) {
    $nm = $pair[0]
    $want = $pair[1]
    $sites = @(Get-CallSites $nm)
    if ($sites.Count -ne 1) {
        $viol += ('R1/R2 ' + $nm + ' 的调用点数应为 1, 实为 ' + $sites.Count +
                  ' (' + (($sites | Sort-Object -Unique) -join ', ') +
                  ') —— 这条规则又长出第二位主人了')
    } elseif ($sites[0] -ne $want) {
        $viol += ('R1/R2 ' + $nm + ' 的唯一调用点应在 ' + $want + ', 现在在 ' + $sites[0])
    }
}

# ---- R3 ----
if ($srcRaw.Contains('strArgPos108c')) {
    $viol += 'R3 strArgPos108c 又出现了 (那张 13 枚名单已并入 Fix 175 + 位号表, 见 §B174)'
}

# ---- R4 ----
if ([regex]::IsMatch($argtype, 'vb6_CStr')) {
    $viol += ('R4 cgen_expr_call_builtin_argtype.inc 的代码里出现 vb6_CStr* —— 数值折串的形状' +
              '只许出自 wrapToBSTR 一处 (这一站现在只该管裸 COM 结果解包 / ObjPtr 取值)')
}

# ---- R5 ----
$castN = @([regex]::Matches($allCode, [regex]::Escape('vb6_CStrLong((int32_t)('))).Count
if ($castN -ne 0) {
    $viol += ('R5 vb6_CStrLong((int32_t)( 在 src/backend 出现 ' + $castN +
              ' 处, 应为 0 —— 那是旧名单自己的形状, 与 wrapToBSTR 不同口径')
}

# ---- R6 ----
$gateCases = -1
$gateDef = [regex]::Match($util, 'bool CCodeGen::typeTakesImplicitCStr[^\{]*\{')
if ($gateDef.Success) {
    $gateBody = Get-BlockBody $util ($gateDef.Index + $gateDef.Value.Length - 1)
    if ($gateBody) { $gateCases = @([regex]::Matches($gateBody, 'case Vb6Type::')).Count }
}
if ($gateCases -lt 14) {
    $viol += ('R6 档位闸读出 ' + $gateCases + ' 档 (< 14) —— 两位旧主人的清单并集是 14 档, ' +
              '少一档就是 Left/Replace 那类实参重新裸交给 BSTR')
}
# 位号表的覆盖面 = 两条形状各数一次：InStr 那一枚带 arity 判断 (callee == "vb6_InStr")，
# 其余十枚住在 kNoTableFormalKinds176 那张档表里 ({"vb6_GetSetting", {"BSTR",…}}, ...)。
# 账 #305 §B176: 这张表从「哪几位是字符串」升成「第 pos 位要什么 C 档」，是这一问唯一的主人；
# `builtinNoParamTableStringSlot` 只剩一行投影 (== "BSTR")，所以名单必须按**档表**那一段数。
$slotDef = [regex]::Match($util, 'std::string CCodeGen::builtinNoParamTableFormalCType[^\{]*\{')
$tableNames = 0
if ($slotDef.Success) {
    $slotBody = Get-BlockBody $util ($slotDef.Index + $slotDef.Value.Length - 1)
    if ($slotBody) {
        $tableNames = @([regex]::Matches($slotBody, 'callee == "vb6_[A-Za-z]+"')).Count
        $tableNames += @([regex]::Matches($slotBody, '\{"vb6_[A-Za-z]+",\s*\{')).Count
    }
}
if ($tableNames -lt 11) {
    $viol += ('R6 位号表只认到 ' + $tableNames + ' 枚内在函数名 (< 11) —— 少一枚就是那一族' +
              '的字符串形参位又没人答 (注册表那四位今天语料命中 461 处数字字面量, §B175)')
}
$wrapN = @([regex]::Matches($emit, 'wrapToBSTR\(')).Count
if ($wrapN -lt 1) {
    $viol += 'R6 cgen_expr_call_arg_emit.inc 里不再问 wrapToBSTR —— 折算形状那一站断了'
}

# ---- R7 ----
$conv = $codeByFile[$convInc]
$convWrap = @([regex]::Matches($conv, 'wrapToBSTR\(')).Count
$convRename = @([regex]::Matches($conv, 'callee = "vb6_MsgBox1"')).Count
if ($convWrap -ne 0) {
    $viol += ('R7 cgen_expr_call_conv_output.inc 的代码里又有 ' + $convWrap +
              ' 处 wrapToBSTR —— MsgBox 那一形又长出第二位主人 (它只吃文本前缀, 且只管' +
              ' prompt 一位; 位号与形状都问 Fix 175 那一站, 见 §B175)')
}
if ($convRename -ne 1) {
    $viol += ('R7 conv_output 里的 arity 改道 (callee = "vb6_MsgBox1") 应为 1 处, 实为 ' +
              $convRename + ' 处 —— 那一问与形参档位无关, 撤了就没人补 1/2/3 参三形')
}

# ---- R8 (账 #305 §B176, 第六十七刀) ----
# 「没有形参表的那一族, 第 pos 位要什么 C 档」= 一枚出口, 读者只许两位:
#   cgen_util_type.cpp 里那行投影 (`== "BSTR"`) + arg_variant 那条臂 (Fix 038b-2)。
# 同时钉住: 那张按 callee 名登的 C 签名表 (`getRuntimeParamCType`) 只剩**一枚回落读者**,
# 它的行是按 VB 写出来的 arity 登的 —— 谁再拿它答位号, 就会把 InStr 的 int32_t start 位答成 BSTR。
$kindSites = @(Get-CallSites 'builtinNoParamTableFormalCType')
$kindWant = @('cgen_expr_call_arg_variant.inc', 'cgen_util_type.cpp')
if ($kindSites.Count -ne 2) {
    $viol += ('R8 builtinNoParamTableFormalCType 的调用点数应为 2 (投影 + arg_variant), 实为 ' +
              $kindSites.Count + ' (' + (($kindSites | Sort-Object -Unique) -join ', ') +
              ') —— 这一问又长出第三位读者/主人了')
} elseif ((($kindSites | Sort-Object -Unique) -join ',') -ne ($kindWant -join ',')) {
    $viol += ('R8 档表的读者应是 ' + ($kindWant -join ' + ') + ', 现在是 ' +
              (($kindSites | Sort-Object -Unique) -join ', '))
}
$rtSites = @(Get-CallSites 'getRuntimeParamCType')
if ($rtSites.Count -ne 1 -or $rtSites[0] -ne 'cgen_expr_call_arg_variant.inc') {
    $viol += ('R8 getRuntimeParamCType 的调用点应只剩 1 处 (arg_variant 的回落), 实为 ' +
              $rtSites.Count + ' (' + (($rtSites | Sort-Object -Unique) -join ', ') +
              ') —— 那张 C 签名表不许再答位号 (§B176 形② 就是这么来的)')
}
if (-not [regex]::IsMatch($argVariant, 'builtinNoParamTableFormalCType\(')) {
    $viol += 'R8 arg_variant 那条臂不再问档表 —— 位号又回到只认 C 签名表的状态 (§B176)'
}

# ---- R9 (账 #305 §B177, 第六十八刀) ----
# needle 位 = RTL 那枚 _Generic 宏的**分派位**, 它一位两答: 数值实参要折成串 (Fix 175 那一站),
# Variant 实参要**裸交** (宏自己走 vb6_InStrVar 解包并 free)。两答必须同源 —— 位号只许在
# builtinNoParamTableDispatchSlot 里数一遍, 否则下一刀又会像六十七刀那样「改一答、掉一答」。
$rawSites = @(Get-CallSites 'builtinNoParamTableVariantMayStayRaw')
if ($rawSites.Count -ne 1 -or $rawSites[0] -ne 'cgen_expr_call_arg_variant.inc') {
    $viol += ('R9 「这一位的 Variant 可否裸交」的调用点应恰好 1 处 (arg_variant), 实为 ' +
              $rawSites.Count + ' (' + (($rawSites | Sort-Object -Unique) -join ', ') +
              ') —— 这一问又长出第二位主人')
}
$slotSites = @(Get-CallSites 'builtinNoParamTableDispatchSlot')
if ($slotSites.Count -ne 2 -or (($slotSites | Sort-Object -Unique) -join ',') -ne 'cgen_util_type.cpp') {
    $viol += ('R9 分派位那一问的调用点应恰好 2 处且都在 cgen_util_type.cpp (档 + 可否裸交两枚投影), ' +
              '实为 ' + $slotSites.Count + ' (' + (($slotSites | Sort-Object -Unique) -join ', ') +
              ') —— 位号被抄了第二遍, §B176 形② 就是这么长回来的')
}
$arityN = @([regex]::Matches($util, 'firstArgType\s*==\s*Vb6Type::String')).Count
if ($arityN -ne 1) {
    $viol += ('R9 cgen_util_type.cpp 里「首参是不是串」那一问出现 ' + $arityN +
              ' 次, 应为 1 次 —— InStr 的位号只能有一位主人')
}
$slotWantsWant = @('cgen_expr_call_arg_variant.inc', 'cgen_expr_call_builtin_fixup.inc') -join ','

# ---- R10 (账 #305 §B180, 第六十九刀) ----
# 「这一位要不要收**整枚槽**而不是收串」= 一枚问句，读者**恰好两枚**（arg_variant 撤回 Variant→BSTR
# 的解包 / builtin_fixup 照着同一问句把 callee 换成收槽的出口 vb6_StrPtrVar）。两站判据必须同源 ——
# 只改一头就会造出「发码把整枚槽交给 void* vb6_StrPtr(BSTR)」(C2440) 或「callee 没改名、槽被解包
# 又现造一枚临时串」(本刀要撤的那条) 两种半边病。
# 它刻意**不进**上面那张按位号答档的表：`vb6_StrPtr(` 在全语料 1720 处里确定不是现造串的有 953 处，
# 另有 535 处这次逐枚读过（数组元素左值 / 解引用 / 返回 BSTR 的调用），只有 41 处是槽 ——
# 逐枚分解与订正见台账 §B180；无条件按位号答槽会把那 953 处一并打成 C2440。
# 变量（`.build/k69_strptr_census.txt`），无条件按位号答槽会把它们一并打成 C2440。
$slotWants = @(Get-CallSites 'builtinWantsRawSlotFromVariant')
if ($slotWants.Count -ne 2) {
    $viol += ('R10 「这一位要不要收整枚槽」的调用点应为 2 (arg_variant 撤回解包 + builtin_fixup 改名), ' +
              '实为 ' + $slotWants.Count + ' (' + (($slotWants | Sort-Object -Unique) -join ', ') +
              ') —— 两站又开始各判各的')
} elseif ((($slotWants | Sort-Object -Unique) -join ',') -ne $slotWantsWant) {
    $viol += ('R10 这一问的读者应是 arg_variant + builtin_fixup, 现在是 ' +
              (($slotWants | Sort-Object -Unique) -join ', '))
}
$strPtrVarN = @([regex]::Matches($allCode, [regex]::Escape('vb6_StrPtrVar'))).Count
if ($strPtrVarN -ne 1) {
    $viol += ('R10 vb6_StrPtrVar 在 src/backend 出现 ' + $strPtrVarN + ' 处, 应为 1 处 —— ' +
              '改道只许在 builtin_fixup 那一处发生, 别在第二站再拼一次出口名')
}

# ---- R11 (账 #305 §B181, 第七十一刀) ----
# 「这一枚实参是不是字节数组」有两位主人：`arrayElemTypes_`（按名字答元素档，六条声明路都填、**含形参**）
# 与 `knownByteArrayVars_` ∪ `classByteArrayMembers_`（Fix 140 那两站查的文本名单）。第七十一刀之前
# 名单缺「形参」整条路 ⇒ 同一枚数组在 `StrPtr(局部)` 与 `StrPtr(形参)` 下交出两种指针（后者把 SafeArray
# **描述符**地址交给要 BSTR 的 Win32 形参，cl 在 C 里只警告）。现在补齐，于是两位主人的**登记路数**
# 必须相等：以后谁再添一条声明路而只填一张表，这一条就红。
$aeWriters = @([regex]::Matches($allCode, 'arrayElemTypes_\[\s*\w+\s*\]\s*=\s*elemType')).Count
$baWriters = @([regex]::Matches($allCode, 'knownByteArrayVars_\.insert')).Count
if ($aeWriters -ne 7 -or $baWriters -ne 7) {
    $viol += ('R11 数组元素档的登记路数 = ' + $aeWriters + '，字节数组名单的登记路数 = ' + $baWriters +
              '，两者应各为 7（模块级 2 + 局部 2 + 形参 3）—— 一张表多一条腿而另一张没有，' +
              '就是「StrPtr(形参) 交出描述符指针」那一族静默病的成因')
}

# ---- R12 (账 #305 §B181, 第七十一刀) ----
# 剥 `(*x)` 这层壳的规则在 src/backend 里有 **24 位主人**，长度闸却写着三种（`> 2` / `> 3` / `> 4`）
# ⇒ 同一个问题在不同站点答案不一样，而 `> 4` 那批对**单字母名**整片失灵（`(*b)` 正好 4 个字符）。
# 收成一处是下一格；本刀先把字节数组那两站（`vb6_StrPtr` / `vb6_StrConv` 的改名）的闸写成 `>= 4`，
# 这里钉两件事：那两站各一条 `>= 4`，且 builtin_fixup 里**不许**再留 `> 4` 的剥壳闸。
$fixupPath = Join-Path $backend 'detail\expr\cgen_expr_call_builtin_fixup.inc'
$fixupCode = if ($codeByFile.ContainsKey($fixupPath)) { $codeByFile[$fixupPath] } else { '' }
$derefGe4 = @([regex]::Matches($fixupCode, '\.size\(\)\s*>=\s*4\s*&&\s*\w+\[0\]\s*==\s*''\(''')).Count
if ($derefGe4 -ne 2) {
    $viol += ('R12 字节数组那两站的 (*x) 剥壳闸应有 2 条 `>= 4`，实为 ' + $derefGe4 +
              ' 条 —— 闸退回 `> 4` 会让单字母形参连查表那一步都没进（病形 = 描述符指针交给要 BSTR 的形参）')
}
$derefGt4Fixup = @([regex]::Matches($fixupCode, '\.size\(\)\s*>\s*4\s*&&\s*\w+\[0\]\s*==\s*''\(''')).Count
if ($derefGt4Fixup -ne 0) {
    $viol += ('R12 builtin_fixup 里还剩 ' + $derefGt4Fixup + ' 条 `> 4` 的 (*x) 剥壳闸：同一条规则的两半' +
              '不许一个修一个不修')
}

Write-Host ('census: backendFiles=' + $allFiles.Count +
            ' gateCases=' + $gateCases +
            ' slotTableNames=' + $tableNames +
            ' wrapToBSTR@emit=' + $wrapN +
            ' wrapToBSTR@conv=' + $convWrap +
            ' kindSites=' + $kindSites.Count +
            ' rtSites=' + $rtSites.Count +
            ' rawSites=' + $rawSites.Count +
            ' dispatchSites=' + $slotSites.Count +
            ' arityQuestions=' + $arityN +
            ' rawSlotReaders=' + $slotWants.Count +
            ' strPtrVarNames=' + $strPtrVarN +
            ' castShapes=' + $castN)

if ($viol.Count -gt 0) {
    foreach ($v in $viol) { Write-Host ('  ' + $v) }
    Write-Host ('FAIL implicit_cstr_station: ' + $viol.Count +
                ' 条违规 —— 这条规则的站点在 cgen_expr_call_arg_emit.inc (Fix 175): ' +
                '位号问形参表 (缺表时问 builtinNoParamTableFormalCType 那一枚档表), ' +
                '档位问 typeTakesImplicitCStr, ' +
                '形状走 wrapToBSTR; 要加第二位主人请先说明这四个问为什么分不开')
    exit 1
}

# --- R13 (knife 73, ledger #305 B182): a formal registered in the WIDENED (32-bit) table by its VB
# type must also reach the NARROW table when that type is Integer. Otherwise a store into an Integer
# formal skips vb6_ChkInt: VB6 raises run-time error 6 and leaves the caller's variable alone, while
# C truncates silently and ByRef writes -25536 back through the reference.
# Two widened insert shapes live in these files: the bare one (exact VB type) and a guarded fallback
# for enum/unknown parameters whose C type is int32_t/int16_t/VBABOOL. Only the bare arm can carry an
# Integer, so R13 compares the narrow count against the bare count and requires exactly one of each
# per file - a fix that lands in only one of the three declaration paths goes red here.
$intFormalFiles = @('src\backend\decl\cgen_decl_func.cpp', 'src\backend\decl\cgen_decl_prop.cpp',
                   'src\backend\decl\cgen_decl_proc.cpp')
$r13Bad = @()
foreach ($f in $intFormalFiles) {
    $txt = [IO.File]::ReadAllText((Join-Path $Root $f))
    $allWide = ([regex]::Matches($txt, 'knownLongVars_\.insert\(pLower\)')).Count
    $fallback = ([regex]::Matches($txt, 'if \(!knownLongVars_\.count\(pLower\)\) knownLongVars_\.insert\(pLower\)')).Count
    $bareWide = $allWide - $fallback
    $narrow = ([regex]::Matches($txt, 'knownIntVars_\.insert\(pLower\)')).Count
    if ($bareWide -lt 1) { $r13Bad += ('{0}: no bare widened formal arm found' -f (Split-Path $f -Leaf)) }
    if ($narrow -ne $bareWide) {
        $r13Bad += ('{0}: Integer reaches the narrow table in {1} of {2} bare widened formal arms' -f
            (Split-Path $f -Leaf), $narrow, $bareWide)
    }
}
if ($r13Bad.Count -gt 0) {
    Write-Host ('FAIL implicit_cstr_station R13 (formal Integer must reach the narrow table): ' + ($r13Bad -join ' | '))
    exit 1
}
Write-Host ('  R13 formal-Integer narrow registration OK ({0} files, bare-widened + narrow in lockstep)' -f $intFormalFiles.Count)
# --- R14 (knife 74, ledger #305 B183): the class-FIELD leg of the same narrow/widened question.
# The class-member scan folds Long, Integer and Boolean into ONE widened table (classLongMembers_)
# and that table is the only thing a procedure entry restores, so inside its own class module a bare
# member name answered tgtBits=32: `ci = 40000` skipped vb6_ChkInt - VB6 raises error 6 and keeps the
# old value, C truncated it to -25536. Integer is therefore ALSO registered in a narrow table
# (classIntMembers_) WITHOUT being removed from the widened one; double registration is this file
# family's existing caliber (Byte = ledger 123, Boolean = ai/022 W1, Date = Fix 175). Every pin below
# reads comment-stripped code text, so a WHY comment can never be what keeps this gate green.
#   R14a the narrow table has exactly one declaration (one owner);
#   R14b the scan clears it per module and fills BOTH spellings (m_ + bare), and that fill sits
#        INSIDE the widened arm: move Integer out of the arm and the widened table's other consumers
#        lose the answer, so the position is pinned, not only the count;
#   R14c exactly 3 sites restore the widened table into knownLongVars_ and each of them restores the
#        narrow table into knownIntVars_ in the same breath (function / sub / property entry).
$scanInc = Join-Path $backend 'detail\base\cgen_base_generate_state_scan.inc'
$stateInc = Join-Path $backend 'detail\util\cgen_state.inc'
if (-not (Test-Path -LiteralPath $scanInc) -or -not (Test-Path -LiteralPath $stateInc)) {
    Write-Host 'FAIL implicit_cstr_station R14: read neither the class-member scan nor the state file'
    exit 1
}
$scanCode = $codeByFile[$scanInc]
$stateCode = $codeByFile[$stateInc]
$r14Bad = @()
$declN = ([regex]::Matches($stateCode, 'std::unordered_set<std::string>\s+classIntMembers_')).Count
if ($declN -ne 1) { $r14Bad += ('R14a classIntMembers_ declared in ' + $declN + ' places, want 1') }
$clearN = ([regex]::Matches($scanCode, 'classIntMembers_\.clear\(\)')).Count
$insM = ([regex]::Matches($scanCode, 'classIntMembers_\.insert\(mLower\)')).Count
$insO = ([regex]::Matches($scanCode, 'classIntMembers_\.insert\(oLower\)')).Count
if ($clearN -ne 1) { $r14Bad += ('R14b classIntMembers_ cleared ' + $clearN + ' times in the scan, want 1') }
if ($insM -ne 1 -or $insO -ne 1) {
    $r14Bad += ('R14b narrow class-field spellings registered m_=' + $insM + ' bare=' + $insO + ', want 1/1')
}
# The widened arm must STILL name Integer (that is the double registration: the widened table keeps
# its answer for its other consumers) and the narrow fill must sit inside that same arm body, which
# is what this ordering pins - the four inserts of that arm appear as wide m_ / wide bare / narrow m_
# / narrow bare, and the next arm (Double) only starts after them.
$wideArm = 'vtype == Vb6Type::Long || vtype == Vb6Type::Integer || vtype == Vb6Type::Boolean'
$wideArmN = ([regex]::Matches($scanCode, [regex]::Escape($wideArm))).Count
if ($wideArmN -ne 1) {
    $r14Bad += ('R14b widened arm naming Long+Integer+Boolean found ' + $wideArmN + ' times, want 1 ' +
                '(Integer must stay in the widened table too)')
}
$iWideM = $scanCode.IndexOf('classLongMembers_.insert(mLower)')
$iWideO = $scanCode.IndexOf('classLongMembers_.insert(oLower)')
$iNarrowM = $scanCode.IndexOf('classIntMembers_.insert(mLower)')
$iNarrowO = $scanCode.IndexOf('classIntMembers_.insert(oLower)')
$iNextArm = $scanCode.IndexOf('classDoubleMembers_.insert(mLower)')
if ($iWideM -lt 0 -or $iWideO -lt 0 -or $iNarrowM -lt 0 -or $iNarrowO -lt 0 -or $iNextArm -lt 0 -or
    -not ($iWideM -lt $iWideO -and $iWideO -lt $iNarrowM -and $iNarrowM -lt $iNarrowO -and $iNarrowO -lt $iNextArm)) {
    $r14Bad += 'R14b the Integer narrow registration must sit inside the widened Long/Integer/Boolean arm'
}
$wideRx = 'knownLongVars_\.insert\(classLongMembers_\.begin\(\), classLongMembers_\.end\(\)\)'
$narrowRx = 'knownIntVars_\.insert\(classIntMembers_\.begin\(\), classIntMembers_\.end\(\)\)'
$wideByFile = @{}
foreach ($path in $codeByFile.Keys) {
    $n = ([regex]::Matches($codeByFile[$path], $wideRx)).Count
    if ($n -gt 0) { $wideByFile[$path] = $n }
}
$wideTotal = 0
foreach ($v in $wideByFile.Values) { $wideTotal += $v }
if ($wideTotal -ne 3) { $r14Bad += ('R14c widened class-member restore sites = ' + $wideTotal + ', want 3') }
foreach ($path in $wideByFile.Keys) {
    $m = ([regex]::Matches($codeByFile[$path], $narrowRx)).Count
    if ($m -ne $wideByFile[$path]) {
        $r14Bad += ('R14c ' + (Split-Path -Leaf $path) + ': narrow restore ' + $m + ' vs widened restore ' + $wideByFile[$path])
    }
}
if ($r14Bad.Count -gt 0) {
    Write-Host ('FAIL implicit_cstr_station R14 (Integer class field must reach the narrow table at every proc entry): ' + ($r14Bad -join ' | '))
    exit 1
}
Write-Host ('  R14 class-field-Integer narrow registration OK (1 table, widened+narrow in lockstep at ' + $wideTotal + ' entries)')
# --- R15 (knife 75, ledger #305 B184): 「这枚左值有多窄 / 这一档窄存要套哪只范围闸」只许有一个答案,
# 四个站点只许去问它, 不许各自写形状. 改前这一族有四位主人: narrowCheckAssign 里的那个 switch、COM
# 字段写回路那一条按 VB 档自己 cast 的臂、narrowTargetTypeOf 的成员臂（它对**类实例**一律答 Unknown,
# 所以typed 直接字段存值压根没被问过范围）、以及 With 那一站（直接从 lastExpr_ 发 temp->f = <expr>）。
#   R15a 三只闸的**函数名串**是那张表独有的: 注释剥净后的整个 src 里恰好 3 处, 且全住
#        cgen_util_type.cpp。第二位主人手写名字 = 重新长出表要消掉的那种漂移。
#   R15b 表是被**问**的, 不是被抄的: narrowCheckFnForVbType 恰好 3 个调用点
#        (1 = narrowCheckAssign 自己那一问, 2 = COM 的 Integer/Byte 两臂), 且改前那枚手写
#        (int16_t)value.lVal 已从**注释剥净**的代码里绝迹 —— 它还活在一条 WHY 注释里, 而这一族的
#        尺子量的正是"注释绝不允许是唯一撑着判据的东西"。
#   R15c narrowCheckAssign 恰好被 3 站问到, 且分配已知: 赋值两腿 2 + With 那一腿 1（第七十五刀新增）。
#   R15d narrowTargetTypeOf 的成员臂里, 类字段那一问恰好 1 次, 且必须住在 §B170 那枚否决
#        (resolveClassMemberCall) **之后** —— 方法/属性名永不被字段表的答案抢走。
#   R15e With 那一站的裸形状 (" = " + lastExpr_ + ";  /* With class field write */") 已绝迹, 而那一站
#        本身还在 —— 读的是**原文**（/* */ 会被 $codeByFile 剥掉）, 两头钉, 所以删掉整站也绿不了 R15e。
$propWriteInc = Join-Path $backend 'detail\stmt\cgen_assign_prop_write.inc'
if (-not (Test-Path -LiteralPath $propWriteInc)) {
    Write-Host 'FAIL implicit_cstr_station R15: neither the With station file nor its raw text is readable'
    exit 1
}
# 「Strip-CComments」把 /**/ 按全文成对匹配, 而这一族文件里有 **C 字符串字面量**
# 里面就写着「/* ... */」(实物: cgen_util_type.cpp:853/855 那两条 marker 拼接) ——
# 它会从那个「/*」一直吞到下一个真正的「*/」, 吞掉整段函数 (在 cgen_util_type.cpp 上实测少 55292 字符,
# narrowTargetTypeOf 的定义被整段吃掉)。后果是双向的: 拿它定义锚点会报不到,
# 拿它数禁忌（「这个形状已绝迹」）会因为违规代码恰巧落在被吞的那段里而假绿。
# R15 自己的那把尺子只剥同一行内的块注释 + 整行/行尾的 //, 从不跨行配对。
function Strip-CComments15([string]$t) {
    $t = [regex]::Replace($t, '/\*[^\r\n]*?\*/', ' ')
    $t = [regex]::Replace($t, '(?m)//.*$', ' ')
    return $t
}
$srcCodeByFile = @{}
foreach ($f in @(Get-ChildItem -Path (Join-Path $Root 'src') -Recurse -File |
                 Where-Object { $_.Extension -in @('.cpp', '.inc', '.h', '.hpp') })) {
    $srcCodeByFile[$f.FullName] = Strip-CComments15 ([IO.File]::ReadAllText($f.FullName))
}
# R15 自己的调用点尺子: 排除「定义 (CCodeGen::)」与「声明」, 而声明档要含 `const char*` ——
# 表返回的正是 const char*, 上面那把 Get-CallSites 的名单 (bool|std::string|size_t) 认不到它。
function Get-CallSites15([string]$name) {
    $rx = [regex]([regex]::Escape($name) + '\(')
    $hits = @()
    foreach ($path in $srcCodeByFile.Keys) {
        $t = $srcCodeByFile[$path]
        foreach ($m in $rx.Matches($t)) {
            $pre = $t.Substring(0, $m.Index)
            if ($pre -match 'CCodeGen::\s*$') { continue }
            if ($pre -match '(bool|std::string|size_t|const char\*)\s*$') { continue }
            $hits += $path
        }
    }
    return $hits
}
function Count-ByLeaf($sites) {
    $h = @{}
    foreach ($p in $sites) {
        $k = Split-Path -Leaf $p
        if ($h.ContainsKey($k)) { $h[$k] = $h[$k] + 1 } else { $h[$k] = 1 }
    }
    return $h
}
$r15Bad = @()

# ---- R15a ----
$chkByFile = @{}
foreach ($path in $srcCodeByFile.Keys) {
    $n = ([regex]::Matches($srcCodeByFile[$path], '"vb6_Chk[A-Za-z]*"')).Count
    if ($n -gt 0) { $chkByFile[$path] = $n }
}
$chkTotal = 0
foreach ($v in $chkByFile.Values) { $chkTotal += $v }
$chkFiles = @($chkByFile.Keys | ForEach-Object { Split-Path -Leaf $_ } | Sort-Object)
if ($chkTotal -ne 3 -or $chkFiles.Count -ne 1 -or $chkFiles[0] -ne 'cgen_util_type.cpp') {
    $r15Bad += ('R15a quoted vb6_Chk* names = ' + $chkTotal + ' in [' + ($chkFiles -join ', ') +
                '], want 3 in [cgen_util_type.cpp] (the table is the only owner of the names)')
}

# ---- R15b ----
$tblByFile = Count-ByLeaf @(Get-CallSites15 'narrowCheckFnForVbType')
$tblTotal = 0
foreach ($v in $tblByFile.Values) { $tblTotal += $v }
if ($tblTotal -ne 3 -or $tblByFile['cgen_util_type.cpp'] -ne 1 -or $tblByFile['cgen_com.cpp'] -ne 2) {
    $r15Bad += ('R15b narrowCheckFnForVbType call sites = ' + $tblTotal +
                ' (util=' + $tblByFile['cgen_util_type.cpp'] + ' com=' + $tblByFile['cgen_com.cpp'] +
                '), want 3 (1 ask + the 2 COM field arms) - the table is asked, never copied')
}
$castTotal = 0
foreach ($path in $srcCodeByFile.Keys) {
    $castTotal += ([regex]::Matches($srcCodeByFile[$path],
                  [regex]::Escape('(int16_t)value.lVal'))).Count
}
if ($castTotal -ne 0) {
    $r15Bad += ('R15b the hand-written (int16_t)value.lVal cast is back in comment-stripped code: ' +
                $castTotal + ' site(s) - ask the table instead')
}

# ---- R15c ----
$ncaByFile = Count-ByLeaf @(Get-CallSites15 'narrowCheckAssign')
$ncaTotal = 0
foreach ($v in $ncaByFile.Values) { $ncaTotal += $v }
if ($ncaTotal -ne 3 -or $ncaByFile['cgen_assign_value_sem.inc'] -ne 2 -or
    $ncaByFile['cgen_assign_prop_write.inc'] -ne 1) {
    $r15Bad += ('R15c narrowCheckAssign call sites = ' + $ncaTotal + ' (assign-legs=' +
                $ncaByFile['cgen_assign_value_sem.inc'] + ' With=' +
                $ncaByFile['cgen_assign_prop_write.inc'] +
                '), want 3 = 2 + 1 - a station that writes its own shape instead of asking')
}

# ---- R15d ----
$util15 = Strip-CComments15 ([IO.File]::ReadAllText($utilCpp))
$defNTO = $util15.IndexOf('Vb6Type CCodeGen::narrowTargetTypeOf(')
if ($defNTO -lt 0) {
    $r15Bad += 'R15d narrowTargetTypeOf definition not found in cgen_util_type.cpp'
} else {
    $obNTO = $util15.IndexOf('{', $defNTO)
    $bodyNTO = $null
    if ($obNTO -ge 0) { $bodyNTO = Get-BlockBody $util15 $obNTO }
    if ($null -eq $bodyNTO) {
        $r15Bad += 'R15d the body of narrowTargetTypeOf is not brace-balanced - cannot read the member arm'
    } else {
        $cfkN = ([regex]::Matches($bodyNTO, 'classFieldKind\(')).Count
        $vetN = ([regex]::Matches($bodyNTO, 'resolveClassMemberCall\(')).Count
        if ($cfkN -ne 1 -or $vetN -lt 1 -or
            $bodyNTO.IndexOf('resolveClassMemberCall(') -gt $bodyNTO.IndexOf('classFieldKind(')) {
            $r15Bad += ('R15d member arm asks classFieldKind ' + $cfkN + ' time(s) with ' + $vetN +
                        ' veto ask(s) BEFORE it - want exactly 1 ask, after the B170 veto')
        }
    }
}

# ---- R15e ----
$propRaw = [IO.File]::ReadAllText($propWriteInc)
if ($propRaw.Contains('" = " + lastExpr_ + ";  /* With class field write */"')) {
    $r15Bad += 'R15e the With station emits lastExpr_ bare again - wrap it with narrowCheckAssign'
}
$shapeN = ([regex]::Matches($propRaw, [regex]::Escape('/* With class field write */'))).Count
if ($shapeN -ne 1) {
    $r15Bad += ('R15e the With class-field station itself is gone: shape marker count = ' + $shapeN +
                ', want 1 (R15e may not be greened by deleting the station)')
}

if ($r15Bad.Count -gt 0) {
    Write-Host ('FAIL implicit_cstr_station R15 (the narrow-store question has one answer and every station asks it): ' + ($r15Bad -join ' | '))
    exit 1
}
Write-Host ('  R15 narrow-store single authority OK (names 3/1 file, table asked at ' + $tblTotal +
            ' sites, assigned at ' + $ncaTotal + ', veto-ordered class-field ask, With leg wrapped)')
Write-Host 'PASS implicit_cstr_station (R1..R15)'
exit 0
