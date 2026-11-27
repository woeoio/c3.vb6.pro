# check_float_to_int_round.ps1 - 账 #248 的结构性哨兵 (只扫源码, 不起 cl)
#
# 为什么单开这一道: 浮点交给整数目标这一件事, 以前**有两条路两个答案** ——
# 发码把 double 直接递进 vb6_ChkLong(int64_t)，C 在调用边界上截断 (`l = 7 / 2` ⇒ 3)；
# 同一句写成 `l = CLng(7 / 2)` 走 vb6_CLng 内部的 round() ⇒ 4。这类"同一个决定抄了两遍"
# 是本仓反复量出来的缺陷族 (#234 拿 DC / #235 画笔色 / #239 笔位 / #247 宿主层几何)，
# 而它的形状在编译期与链接期**一声不响** —— 只有拿 exact .75 / .5 这类数去跑才看得见。
#
# 规则 (改坏了会红, 不是装饰):
#   S1  取整出口只有一枚: vb6_FltToLng 在 src/rtl 里恰好定义一次，在 vb6rtl_builtin.h 里声明一次
#   S2  半值取偶那一位也只许有一处: `((int64_t)f & 1) == 0` 在 src/rtl 里恰好 1 次，且住在
#       「半到偶」的唯一主人 vb6_RoundHalfEven 体内（第二百二十三刀把身子从 vb6_FltToLng
#       提到这里，Round 也问它）; vb6_FltToLng 必须只委托一次；旧的 `(int64_t)round(` 形
#       必须为 0 (它现在是半个答案: 半值向外)。
#       (Fix 431 把这里从 round() 改成 floor + 奇偶判定 —— 银行家舍入, 见 vb6rtl_conv.c 那条)
#       (改前 vb6_CInt / vb6_CLng 各抄了一遍 —— 那正是"两个答案"的另一半)
#   S3  显式转换与隐式赋值同源: vb6_CInt / vb6_CLng 的体内必须调 vb6_FltToLng (两处)
#   S4  发码侧浮点档齐: narrowCheckAssign 里那一道取整分支必须认 Single / Double / Currency / Date 四档
#       (第一百零二刀把 Date 接进来: Date 在 C 侧就是 double 的 OLE 序列号，交给窄整型目标时
#        与浮点同一条命 —— 少它一档就是 `r = DateValue(...)` 在调用边界上截断, 账 #431 §B205)
#   S5  旧形不回潮: 套检查的两条发射路径写死 —— 浮点那条必须裹 helper，整数那条必须不裹；
#       出现第三种 `std::string(fn) + ...` 就是又开了一条不经 helper 的路
#   S6  helper 只住一处: vb6_FltToLng( 在 src/backend 里只许出现在 narrowCheckAssign 那一处，
#       且 vb6_ChkByte/Int/Long 这三个名字在 src/backend 里恰好 3 次 (就是那张 switch) ——
#       别处再拼一份 Chk 调用，等于把这条链挪出判据面
#   ---- 以下两条是账 #261 (§B77②) 加的目标侧 ----
#   S7  目标侧也只有一个答案: 「这枚左值是哪档窄整型」只住在 narrowTargetTypeOf (定义一次、
#       声明一次、只被 narrowCheckAssign 问一次)，且它的体内必须问那两条既有出口
#       (UDT 字段 = inferUdtFieldVb6Type / 数组元素 = arrayElemTypes_) —— 抄一份新的类型表
#       就是把这条链再劈成两半；数组那一档还必须**拒绝空下标** (`dst() = src()` 是整枚数组
#       描述符指针，不是元素槽 —— 套上检查就是 rev36 那枚 error 6 换路重来)
#   S8  「只认裸标识符」那道闸门不许回潮: 改前 narrowCheckAssign 第一行就是
#       `target->kind != IdentifierExpr -> 原样发`，成员/数组元素两形整片不经检查 (实测六形
#       全截断)。这条判据专门拦"为了稳妥又把闸门加回来"
#
# 用法:  pwsh -File scripts\check_float_to_int_round.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$bad = @()

$rtlDir = Join-Path $root "src\rtl"
$beDir = Join-Path $root "src\backend"
$convRel = "src\rtl\core\vb6rtl\vb6rtl_conv.c"
$hdrRel = "src\rtl\core\vb6rtl\vb6rtl_builtin.h"
$typeRel = "src\backend\cgen_util_type.cpp"
$helpRel = "src\backend\detail\util\cgen_helpers.inc"
$conv = Join-Path $root $convRel
$hdr = Join-Path $root $hdrRel
$typeFile = Join-Path $root $typeRel
$helpFile = Join-Path $root $helpRel
foreach ($need in @($conv, $hdr, $typeFile, $helpFile)) {
    if (-not (Test-Path -LiteralPath $need)) {
        Write-Host ("FAIL S0 missing " + $need) -ForegroundColor Red
        exit 1
    }
}

$convText = [System.IO.File]::ReadAllText($conv)
$hdrText = [System.IO.File]::ReadAllText($hdr)
$typeText = [System.IO.File]::ReadAllText($typeFile)
$helpText = [System.IO.File]::ReadAllText($helpFile)

function Get-SrcFiles($dir) {
    Get-ChildItem -LiteralPath $dir -Recurse -File |
        Where-Object { $_.Extension -in ".c", ".h", ".cpp", ".inc", ".hpp" }
}

# 读一份源码的**代码部分**: 把 // 注释去掉再数。两道普查 (S6 的"别人不许出现 Chk 名字"、
# S8 的"别人不许问 narrowTargetTypeOf") 数的是**消费点**, 而注释里提一句权威的名字是给人看的
# 说明, 不是调用 —— 门 #634 就红在这一点上 (第七十三刀给三条形参臂写的 WHY 注释各提了一次
# narrowTargetTypeOf, 被 S8 数成三处消费)。只裁"行首或空白后的 //"两种形状, 免得吃掉字符串里的
# 路径; 块注释不裁 (真要藏一个调用进 /* */ 里, 那是另一回事, 这里不放宽)。
function Get-CodeText($path) {
    $raw = [System.IO.File]::ReadAllText($path)
    $kept = New-Object System.Collections.Generic.List[string]
    foreach ($line in ($raw -split "`n")) {
        $line = $line.TrimEnd("`r")
        $m = [regex]::Match($line, '(^|\s)//')
        if ($m.Success) { $kept.Add($line.Substring(0, $m.Index)) } else { $kept.Add($line) }
    }
    return ($kept -join "`n")
}

# ---- S1: 取整出口只有一枚 ----
$defSites = @()
$defCount = 0
foreach ($f in Get-SrcFiles $rtlDir) {
    $n = @([regex]::Matches([System.IO.File]::ReadAllText($f.FullName), 'int64_t\s+vb6_FltToLng\s*\(\s*double[^)]*\)\s*\{')).Count
    if ($n -gt 0) { $defSites += ($f.Name + "=" + $n); $defCount += $n }
}
if ($defCount -ne 1) {
    $bad += ("S1 vb6_FltToLng definitions = " + $defCount + " (" + ($defSites -join ",") + ", expected 1)")
}
if (@([regex]::Matches($hdrText, 'int64_t vb6_FltToLng\(double x\);')).Count -ne 1) {
    $bad += "S1 vb6_FltToLng is not declared exactly once in vb6rtl_builtin.h"
}

# ---- S2: 半值取偶那一位只住在 helper 体内；旧的 round() 形不许回潮 ----
$evenTag = '\(\(int64_t\)\s*f\s*&\s*1\)\s*==\s*0'
$evenSites = @()
$evenCount = 0
$roundSites = @()
$roundCount = 0
foreach ($f in Get-SrcFiles $rtlDir) {
    $t = [System.IO.File]::ReadAllText($f.FullName)
    $n = @([regex]::Matches($t, $evenTag)).Count
    if ($n -gt 0) { $evenSites += ($f.Name + "=" + $n); $evenCount += $n }
    $r = @([regex]::Matches($t, '\(int64_t\)\s*round\s*\(')).Count
    if ($r -gt 0) { $roundSites += ($f.Name + "=" + $r); $roundCount += $r }
}
if ($evenCount -ne 1) {
    $bad += ("S2 half-even tie-break sites in src/rtl = " + $evenCount + " (" + ($evenSites -join ",") +
             ", expected 1 -- the helper body; a second copy is the two-answers bug)")
}
if ($roundCount -ne 0) {
    $bad += ("S2 (int64_t)round( sites in src/rtl = " + $roundCount + " (" + ($roundSites -join ",") +
             ", expected 0 -- half-away-from-zero is the other answer Fix 431 removed)")
}
$mHelp = [regex]::Match($convText, 'int64_t\s+vb6_FltToLng\s*\([^)]*\)\s*\{[\s\S]*?\}')
if (-not $mHelp.Success) {
    $bad += "S2 helper body not found in vb6rtl_conv.c"
} elseif (@([regex]::Matches($mHelp.Value, 'vb6_RoundHalfEven\(')).Count -ne 1) {
    $bad += "S2 vb6_FltToLng no longer delegates to vb6_RoundHalfEven exactly once"
}
$mEven = [regex]::Match($convText, 'double\s+vb6_RoundHalfEven\s*\([^)]*\)\s*\{[\s\S]*?\n\}')
if (-not $mEven.Success) {
    $bad += "S2 vb6_RoundHalfEven body not found in vb6rtl_conv.c"
} elseif (@([regex]::Matches($mEven.Value, $evenTag)).Count -ne 1) {
    $bad += "S2 the one half-even tie-break is not inside vb6_RoundHalfEven's body"
}

# ---- S3: CLng / CInt 同源 ----
foreach ($fn in @("vb6_CInt", "vb6_CLng")) {
    $mFn = [regex]::Match($convText, '\b' + $fn + '\s*\(\s*double[^)]*\)\s*\{[^}]*\}')
    if (-not $mFn.Success) { $bad += ("S3 " + $fn + " body not found"); continue }
    if (@([regex]::Matches($mFn.Value, 'vb6_FltToLng\(')).Count -ne 1) {
        $bad += ($fn + " does not route through vb6_FltToLng exactly once -- its body is: " +
                 $mFn.Value.Replace("`r", " ").Replace("`n", " ").Substring(0, [Math]::Min(90, $mFn.Value.Length)))
    }
}

# ---- S4: 发码侧浮点档齐 ----
# 三道档必须一起在册: cgenIntBits(Single) 答 32，Long 目标也 32 —— 少了 Single 那一档，
# `l = 某Single` 会被下面那句"装得下, 不套"放过，连截断都省了 (实测 F2L05 那一形)。
$mGate = [regex]::Match($typeText, 'if\s*\(\s*vt == Vb6Type::Single\s*\|\|\s*vt == Vb6Type::Double\s*\|\|\s*vt == Vb6Type::Currency\s*\|\|\s*vt == Vb6Type::Date\s*\)')
if (-not $mGate.Success) {
    $bad += "S4 narrowCheckAssign's rounding gate no longer names all four of Single / Double / Currency / Date"
}

# ---- S5: 两条发射路径写死 ----
$nHelper = @([regex]::Matches($typeText, 'return std::string\(fn\) \+ "\(vb6_FltToLng\(" \+ cValue \+ "\)\)";')).Count
$nPlain = @([regex]::Matches($typeText, 'return std::string\(fn\) \+ "\(" \+ cValue \+ "\)";')).Count
if ($nHelper -ne 1) { $bad += ("S5 float-source wrap through the helper = " + $nHelper + " (expected 1)") }
if ($nPlain -ne 1) { $bad += ("S5 integer-source wrap = " + $nPlain + " (expected 1)") }
$allWrap = @([regex]::Matches($typeText, 'std::string\(fn\)\s*\+')).Count
if ($allWrap -ne 2) { $bad += ("S5 emission paths for the narrow check = " + $allWrap + " (expected exactly 2)") }

# ---- S6: helper / Chk* 名字的站点普查 ----
$bHelp = @([regex]::Matches($typeText, 'vb6_FltToLng\(')).Count
if ($bHelp -ne 1) { $bad += ("S6 vb6_FltToLng( occurrences in cgen_util_type.cpp = " + $bHelp + " (expected 1)") }
$others = @()
foreach ($f in Get-SrcFiles $beDir) {
    if ($f.FullName -eq $typeFile) { continue }
    $n = @([regex]::Matches((Get-CodeText $f.FullName), '"vb6_Chk(Byte|Int|Long)"')).Count
    if ($n -gt 0) { $others += ($f.Name + "=" + $n) }
}
if ($others.Count -ne 0) {
    $bad += ("S6 the narrow-check names leaked into another emitter: " + ($others -join ","))
}
$swNames = @([regex]::Matches($typeText, '"vb6_Chk(Byte|Int|Long)"')).Count
if ($swNames -ne 3) { $bad += ("S6 narrow-type table names = " + $swNames + " (expected 3: Byte / Integer / Long)") }

# ---- S7: 目标侧只有一处回答 (账 #261) ----
# "这枚左值是哪档窄整型" 必须只住在 narrowTargetTypeOf 里，narrowCheckAssign 只许问它一次。
# 三条判据各自会红的假改动: 把答案抄回调用方 (nHelper2=2 / asked=2)、认不出时猜一档
# (accept=3 之外的档位在 switch 里自己红)、数组元素不按声明类型答 (S7 的第二条)。
$nResolve = @([regex]::Matches($typeText, 'Vb6Type CCodeGen::narrowTargetTypeOf\s*\(\s*Expr\s*\*\s*target\s*\)\s*const\s*\{')).Count
if ($nResolve -ne 1) { $bad += ("S7 narrowTargetTypeOf definitions = " + $nResolve + " (expected 1)") }
$nDecl = @([regex]::Matches($helpText, 'Vb6Type narrowTargetTypeOf\s*\(\s*Expr\s*\*\s*target\s*\)\s*const;')).Count
if ($nDecl -ne 1) { $bad += ("S7 narrowTargetTypeOf declarations in cgen_helpers.inc = " + $nDecl + " (expected 1)") }
$nAsk = @([regex]::Matches($typeText, 'narrowTargetTypeOf\(\s*target\s*\)')).Count
if ($nAsk -ne 1) { $bad += ("S7 narrowCheckAssign asks the target type " + $nAsk + " times (expected 1)") }
$mRes = [regex]::Match($typeText, 'Vb6Type CCodeGen::narrowTargetTypeOf\s*\([^)]*\)\s*const\s*\{[\s\S]*?\n\}')
if (-not $mRes.Success) {
    $bad += "S7 narrowTargetTypeOf body not found in cgen_util_type.cpp"
} else {
    $resBody = $mRes.Value
    # 只许答三档 (Byte / Integer / Long)，且必须是从那两条既有出口 (字段声明类型 /
    # 数组元素声明类型) 拿来的值 —— 不许在这里现拼类型表。
    $accept = @([regex]::Matches($resBody, '== Vb6Type::(Byte|Integer|Long)')).Count
    if ($accept -lt 3) { $bad += ("S7 the resolver's three accepted narrow tiers = " + $accept + " (expected >= 3)") }
    if ($resBody -notmatch 'inferUdtFieldVb6Type\s*\(\s*target\s*,\s*&') {
        $bad += "S7 the member branch no longer asks inferUdtFieldVb6Type(target, &isArray) -- array members hold a descriptor, not a scalar slot (VbQRCodegen error 6)"
    }
    if ($resBody -notmatch 'if\s*\(\s*fldIsArray\s*\)\s*return Vb6Type::Unknown') {
        $bad += "S7 the member branch no longer refuses array members (a member array is a descriptor, not a scalar slot)"
    }
    if ($resBody -notmatch 'arrayElemTypes_\.find\s*\(') {
        $bad += "S7 the array branch no longer reads arrayElemTypes_.find( (the declared element type)"
    }
    # 空下标那一形 (`dst() = src()`, Fix 170) 发的是**整枚数组描述符指针**，不是元素槽。
    # 实测漏这道闸的产物: test_array.bas 改后 EXIT=0x00000006、四条 wa-* 判据整片不打印
    # (`vb6_ChkByte(vb6_StringToByteArray(s))` = rev36 那枚雷换一条入口重来)。
    if ($resBody -notmatch 'positional\.empty\s*\(\)') {
        $bad += "S7 the array branch no longer refuses the empty-index whole-array form (Fix 170)"
    }
    if ($resBody -match 'Vb6Type::Boolean') {
        $bad += "S7 the resolver claims Boolean -- its value range is -1/0, wrapping it is noise (pre-#261 behaviour)"
    }
}

# ---- S8: 旧的"只认裸标识符"那道闸门不许回潮 (账 #261) ----
# 那一道闸门就是这一格的根因本身: 目标的窄整型问题被"目标长什么样"提前判死，
# 成员/数组元素两形整片走截断。留着这条判据, 有人"为了稳妥"把它加回来时就会红。
$nOldGate = @([regex]::Matches($typeText, 'target->kind\s*!=\s*ASTNodeKind::IdentifierExpr\s*\)\s*return cValue')).Count
if ($nOldGate -ne 0) {
    $bad += ("S8 the identifier-only target gate is back (" + $nOldGate + " site(s)) -- that gate IS ledger 261")
}
# 出口普查: 整个 backend 里除 narrowTargetTypeOf 的定义与声明之外, 别人不许问它
# (声明那份在 cgen_helpers.inc, S7 已经按"恰好一次"钉住了, 这里只查真调用)
$nOtherAsk = 0
foreach ($f in Get-SrcFiles $beDir) {
    if ($f.FullName -eq $typeFile) { continue }
    if ($f.FullName -eq $helpFile) { continue }
    $nOtherAsk += @([regex]::Matches((Get-CodeText $f.FullName), 'narrowTargetTypeOf')).Count
}
if ($nOtherAsk -ne 0) {
    $bad += ("S8 narrowTargetTypeOf consumed outside cgen_util_type.cpp: " + $nOtherAsk + " site(s)")
}

# ---- S9: Round 的口径 = 半到偶 + 位域 [0,22] (账 §11.206 第二百二十三刀) ----
# 真机 .temp/k223p/P290..P296: 平局往偶、位域外 5、且 2.675→2.68 与 1.005→1 两枚读数
# 只有在「先乘再半到偶」下才能同时成立（十进制串取整会把 1.005 打成 1.01）。
$roundDef = @([regex]::Matches((Get-CodeText $conv), 'double\s+vb6_RoundHalfEven\s*\(')).Count
if ($roundDef -ne 1) { $bad += ("S9 vb6_RoundHalfEven definitions in vb6rtl_conv.c = " + $roundDef + " (want 1)") }
$roundDecl = @([regex]::Matches((Get-CodeText $hdr), 'double\s+vb6_RoundHalfEven\s*\(\s*double\s+x\s*\)\s*;')).Count
if ($roundDecl -ne 1) { $bad += ("S9 vb6_RoundHalfEven declaration = " + $roundDecl + " (want 1)") }
$roundUse = 0
foreach ($f in Get-SrcFiles $rtlDir) {
    if ($f.Extension -ne ".c") { continue }   # 头里的声明单算，不混进站点普查
    $roundUse += @([regex]::Matches((Get-CodeText $f.FullName), 'vb6_RoundHalfEven\(')).Count
}
if ($roundUse -ne 3) { $bad += ("S9 vb6_RoundHalfEven sites in .c = " + $roundUse + " (want 3 = def + FltToLng + Round)") }
$mRound = [regex]::Match((Get-CodeText $conv), 'double\s+vb6_Round\s*\(\s*double\s+x[^)]*\)\s*\{[\s\S]*?\n\}')
if (-not $mRound.Success) {
    $bad += "S9 vb6_Round body not found"
} else {
    if (@([regex]::Matches($mRound.Value, 'decimals\s*<\s*0\s*\|\|\s*decimals\s*>\s*22')).Count -ne 1) {
        $bad += "S9 vb6_Round no longer gates numdecimalplaces to 0..22 with error 5"
    }
    if (@([regex]::Matches($mRound.Value, 'vb6_ErrRaiseNumber\s*\(\s*5\s*\)')).Count -ne 1) {
        $bad += "S9 vb6_Round does not raise error 5 outside 0..22"
    }
    if (@([regex]::Matches($mRound.Value, '(?<!_)\bround\s*\(')).Count -ne 0) {
        $bad += "S9 vb6_Round calls round() again (half-away-from-zero is the old answer)"
    }
}
$runnerText = [System.IO.File]::ReadAllText((Join-Path $root "tests\run_tests.ps1"))
$b230Cells = @([regex]::Matches($runnerText, 'b230_round_banker\\B230Res\.vbp')).Count
if ($b230Cells -ne 2) { $bad += ("S9 b230 round-banker run cells = " + $b230Cells + " (want 2 = x64 + x86)") }

# ---- S10: Hex/Oct 的实参收窄 = CLng 那一档 (账 §11.209 第二百二十六刀) ----
# 真机 .temp/k225p/P305+306: Hex(44259.62)=ACE4 / Hex(1.5)=2 / Hex("2.5")=2 /
# Hex(2147483647.6)=Error 6。改前 C 在 int32_t 边界上隐式截断 ⇒ ACE3 / 1 / 指针值 / 饱和值。
# 判据两条：
#  (a) 收窄住**调用站**恰一处，且问的是静态档位 (inferExprType) 而不是文本前缀 ——
#      文本判据在 §11.208 刚量过一次失效（包一层 Borrow 就永远对不上）；
#      四个浮点形档 (Single/Double/Date/Currency) + String 五档都得在。
#  (b) RTL 那边不许另开一条路：vb6_Hex / vb6_Oct 的形参仍是 int32_t（收窄是调用方的事），
#      而取整那一份答案仍只有 vb6_FltToLng → vb6_RoundHalfEven 一条链（S1/S2/S3 已钉）。
$argtypeRel = "src\backend\detail\expr\cgen_expr_call_builtin_argtype.inc"
$argtype = Join-Path $root $argtypeRel
if (-not (Test-Path $argtype)) {
    $bad += "S10 cgen_expr_call_builtin_argtype.inc not found"
} else {
    $atText = Get-CodeText $argtype
    $nNarrow = @([regex]::Matches($atText, 'const Vb6Type t209 = inferExprType')).Count
    if ($nNarrow -ne 1) { $bad += ("S10 the Hex/Oct narrowing station = " + $nNarrow + " (want 1)") }
    foreach ($t in @('Vb6Type::Single', 'Vb6Type::Double', 'Vb6Type::Date', 'Vb6Type::Currency', 'Vb6Type::String')) {
        if ($atText -notmatch [regex]::Escape($t)) { $bad += ("S10 the narrowing station no longer gates " + $t) }
    }
    if ($atText -notmatch 'inferExprType\(\*node\.positional\[0\]\)') {
        $bad += "S10 the narrowing station no longer asks inferExprType (text matching is the §11.208 failure mode)"
    }
    if (@([regex]::Matches($atText, 'vb6_CLng\(')).Count -ne 2) {
        $bad += "S10 the narrowing station must ship exactly two vb6_CLng outlets (String via vb6_NumVal, float via direct)"
    }
}
$conv10 = Get-CodeText $conv
if ($conv10 -notmatch 'BSTR vb6_Hex\(int32_t n\)') {
    $bad += "S10 vb6_Hex no longer takes int32_t -- the narrow must stay at the call site"
}
if ($conv10 -notmatch 'BSTR vb6_Oct\(int32_t n\)') {
    $bad += "S10 vb6_Oct no longer takes int32_t -- the narrow must stay at the call site"
}
$b233Cells = @([regex]::Matches($runnerText, 'b233_hex_oct_narrow\\B233Res\.vbp')).Count
if ($b233Cells -ne 2) { $bad += ("S10 b233 hex/oct run cells = " + $b233Cells + " (want 2 = x64 + x86)") }

# ---- S11: Hex/Oct 的答案宽度跟载体档走 (账 §11.210 第二百二十七刀) ----
# 真机 .temp/k225p/P308/P310：Hex(CByte(255))="FF" / Hex(CInt(-1))="FFFF" /
# Hex(True)="FFFF" / Hex(CLng(-1))="FFFFFFFF"；正值不补零 ⇒ 宽度只体现在负值的二补码上。
# 三条规则：
#  (a) 宽度表只许一张：vb6_HexMaskOf 定义 1 处，且它认 Byte/Integer/Boolean 三档；
#      进制格式化（%X / %o）各只许 1 处 —— 多一处就是另一份答案。
#  (b) 静态档由发码选名字（*Int/*Byte），变体档由 vt 选（vb6_HexW/Slot），两条共一张表。
#  (c) 字符串上下文与变体上下文**同一个发射形状**（HexSlot），所以两个装箱站
#      （mathNullFamilyV / boxToVariant 的 k64 表）都必须认得它 —— 否则 Null 会丢
#      （真机 P310 Z2：IsNull(Hex(Null))=True，而 HexSlot 在串上下文必须给空串 Z1）。
$conv11 = Get-CodeText $conv
$maskDef = @([regex]::Matches($conv11, 'static uint32_t vb6_HexMaskOf\s*\(\s*int32_t vartype\s*\)')).Count
if ($maskDef -ne 1) { $bad += ("S11 vb6_HexMaskOf definitions = " + $maskDef + " (want 1)") }
foreach ($vt in @('vb6_vtByte', 'vb6_vtInteger', 'vb6_vtBoolean')) {
    $m = [regex]::Match($conv11, 'static uint32_t vb6_HexMaskOf[\s\S]*?\n\}')
    if (-not $m.Success -or $m.Value -notmatch [regex]::Escape($vt)) {
        $bad += ("S11 the width table no longer gates " + $vt)
    }
}
$nX = @([regex]::Matches($conv11, 'swprintf\(buf, 16, L"%X"')).Count
$nO = @([regex]::Matches($conv11, 'swprintf\(buf, 16, L"%o"')).Count
if ($nX -ne 1) { $bad += ("S11 %X formatters = " + $nX + " (want 1 = the only hex outlet)") }
if ($nO -ne 1) { $bad += ("S11 %o formatters = " + $nO + " (want 1 = the only octal outlet)") }
$slotNull = @([regex]::Matches($conv11, 'BSTR vb6_HexSlot\(vb6_VARIANT v\) \{[\s\S]{0,400}?vb6_vtNull')).Count
if ($slotNull -lt 1) { $bad += "S11 vb6_HexSlot no longer maps Null to the empty string (oracle P310 Z1)" }
if ($conv11 -notmatch 'vb6_VariantString\(vb6_HexSlot\(v\)\)') {
    $bad += "S11 vb6_HexV no longer delegates to vb6_HexSlot (one width table, two consumers)"
}
$ctrl11 = Get-CodeText (Join-Path $root "src\backend\cgen_util_ctrl.cpp")
$nSlotBox = @([regex]::Matches($ctrl11, 'splitOneArgSlotCall\(cExpr, "vb6_HexSlot\(')).Count
if ($nSlotBox -ne 2) { $bad += ("S11 HexSlot recognised by boxing stations = " + $nSlotBox + " (want 2 = boxToVariant + mathNullFamilyV)") }
$argtype11 = Get-CodeText $argtype
if (@([regex]::Matches($argtype11, 'callee \+= "Byte"')).Count -ne 1) {
    $bad += "S11 the Byte-width outlet selection is gone (Hex(CByte(255)) must answer FF)"
}
if (@([regex]::Matches($argtype11, 'callee \+= "Int"')).Count -ne 1) {
    $bad += "S11 the Integer-width outlet selection is gone (Hex(CInt(-1)) must answer FFFF)"
}
if ($argtype11 -notmatch 'isNarrowIntExpr\(node\.positional\[0\]\.get\(\)\)') {
    $bad += "S11 the Integer-width branch no longer asks isNarrowIntExpr (Integer vars answer Long in inferExprType)"
}
$b234Cells = @([regex]::Matches($runnerText, 'b234_hex_oct_width\\B234Res\.vbp')).Count
if ($b234Cells -ne 2) { $bad += ("S11 b234 hex/oct width run cells = " + $b234Cells + " (want 2 = x64 + x86)") }

if ($bad.Count -ne 0) {
    foreach ($b in $bad) { Write-Host ("FAIL " + $b) -ForegroundColor Red }
    exit 1
}
Write-Host "PASS float_to_int_round (S1 one helper, S2 one round() in vb6_RoundHalfEven, S3 CLng/CInt same source, S4 four float-ish types gated, S5 two emission paths, S6 census clean, S7 one target-type answer, S8 no identifier-only gate, S9 Round banker + 0..22, S10 Hex/Oct narrow through CLng at the call site, S11 Hex/Oct width follows the carrier)"
