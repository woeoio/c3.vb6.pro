# 账 #262 / #265 的结构性哨兵 (只扫源码, 不起 cl)
#
# 为什么单开这一道: 「一枚 UDT 成员数组有几维、每维几格」在发码里原本只有**一个**答案 ——
# parser 把第 2..N 维"解析并丢弃"，所以 cgen_decl 只折得出第一维，结构体发出 `float M[5]`
# 而不是 `float M[5][5]`，而 `m.M(3, 3)` 的下标也只折成 `M[3]`。三个不同下标写进同一格：
# 不响、不崩、静默错值 (真流量 = Charts 2020 LabelPlus 的 GDI+ 色彩矩阵，GDI+ 从一枚 20
# 字节的结构体里读 100 字节)。**同一条走查里还量出第二格** (账 #265): 含所有权元素
# (String / Variant / 子 UDT) 的定长成员数组，自动深拷贝助手发的是
# `sizeof(d->S[_i])` —— `_i` 用在它自己那条声明**之前** ⇒ C2065，工程编不过；语料里 0 处
# 这种形状，所以从没响过。
# 这两件事在源码层面看都跟"正常"一模一样，所以判据必须是结构性的。
#
# 规则 (改坏了会红, 不是装饰):
#   U1  下标折算只住一处: udtFixedMemberIndex 定义=1 / 声明=1 / 调用点=2 (obj.M(i,j) 与
#       With 里的 .M(i,j) 两条路必须问同一个出口)，别的文件不许再拼一层括号
#   U2  那一处里: 层数只从秩来 (layers = rank)，缺的那层补 0，括号只拼一次
#   U3  维数的唯一来源是语法: parser 每条后续维 push 一次 moreDims，且"解析并丢弃"那句话
#       不许回来 (那就是这一格的根因本身)
#   U4  语义层只许传**秩**、不许传各维格数 (mi.arrayRank 写 1 次，发码侧读 2 次) ——
#       格数住 in cgen_decl 的 tryEvalConstInt，抄第二份就会和它分家
#   U5  结构体那侧: 第一维与"多出来的那些维"共用同一条折叠出口，brackets 只拼这一处
#   U6  深拷贝助手那侧: 逐格必须走扁平元素指针 (`*_dp`)，`_i` 必须先声明后用，
#       `sizeof(d->X[_i])` 这一形一个不许留
#   U7  元素档那条闸必须问数组性: narrowTargetTypeOf 的成员元素分支里 `!fldIsArray` 那道
#       拒绝要在 (账 #261 的自伤形状 = 把数组描述符指针当标量槽套 Chk* ⇒ run-time error 6)
#   U8  ByRef 数组实参那处的元素尺寸/格数必须按秩剥到底 (秩 >=2 时 `sizeof((arg)[0])` 数的是一行)
#
# 用法:  pwsh -File scripts\check_udt_member_dims.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$bad = @()

function Get-SrcText($rel) {
    return [System.IO.File]::ReadAllText((Join-Path $root $rel))
}

$clsRel    = "src\backend\cgen_util_classtype.cpp"
$typeRel   = "src\backend\cgen_util_type.cpp"
$declRel   = "src\backend\decl\cgen_decl.cpp"
$helpRel   = "src\backend\detail\util\cgen_helpers.inc"
$parseRel  = "src\parser\parser_decl.cpp"
$astRel    = "src\ast\detail\ast_decl.hpp"
$cloneRel  = "src\ast\ast_clone.cpp"
$semRel    = "src\semantics\semantic_analyzer_decl_type.cpp"
$memberRel = "src\backend\detail\expr\cgen_expr_call_callee_member.inc"
$withmRel  = "src\backend\detail\expr\cgen_expr_call_callee_withm.inc"

$clsText    = Get-SrcText $clsRel
$typeText   = Get-SrcText $typeRel
$declText   = Get-SrcText $declRel
$helpText   = Get-SrcText $helpRel
$parseText  = Get-SrcText $parseRel
$astText    = Get-SrcText $astRel
$cloneText  = Get-SrcText $cloneRel
$semText    = Get-SrcText $semRel
$memberText = Get-SrcText $memberRel
$withmText  = Get-SrcText $withmRel

# 注释里可以写"以前这里丢弃了后续维度"，那句话不该把哨兵弄红 —— 判"有没有某段代码"之前
# 先把注释剥掉 (账 #261 之后的一条通用口径: 注释是待测主张，不是被测代码)。
function Remove-LineComments($s) {
    $out = New-Object System.Text.StringBuilder
    foreach ($l in ($s -split "`r?`n")) {
        $cut = $l.IndexOf("//")
        if ($cut -ge 0) { [void]$out.AppendLine($l.Substring(0, $cut)) } else { [void]$out.AppendLine($l) }
    }
    return $out.ToString()
}
$parseCode = Remove-LineComments $parseText

# ---- U1: 一个出口, 两条路 ----
$def = @([regex]::Matches($clsText, 'std::string CCodeGen::udtFixedMemberIndex\s*\(')).Count
if ($def -ne 1) { $bad += ("U1 udtFixedMemberIndex definitions = " + $def + " (expected 1)") }
$dec = @([regex]::Matches($helpText, 'std::string udtFixedMemberIndex\s*\(')).Count
if ($dec -ne 1) { $bad += ("U1 udtFixedMemberIndex declarations in cgen_helpers.inc = " + $dec + " (expected 1)") }
$callMember = @([regex]::Matches($memberText, 'udtFixedMemberIndex\(')).Count
$callWith   = @([regex]::Matches($withmText,   'udtFixedMemberIndex\(')).Count
if ($callMember -ne 1) { $bad += ("U1 the obj.member(i, j) road calls it " + $callMember + " times (expected 1)") }
if ($callWith -ne 1) { $bad += ("U1 the With .member(i, j) road calls it " + $callWith + " times (expected 1)") }
foreach ($rel in @($declRel, $typeRel)) {
    $n = @([regex]::Matches((Get-SrcText $rel), 'udtFixedMemberIndex')).Count
    if ($n -ne 0) { $bad += ("U1 " + $rel + "拼了自己的下标层 (" + $n + " 处) —— 层数只许问那一个出口") }
}

# ---- U2: 层数只从秩来 ----
$mFold = [regex]::Match($clsText, 'std::string CCodeGen::udtFixedMemberIndex\s*\([^)]*\)\s*\{[\s\S]*?\n\}')
if (-not $mFold.Success) {
    $bad += "U2 udtFixedMemberIndex body not found in cgen_util_classtype.cpp"
} else {
    $fb = $mFold.Value
    if ($fb -notmatch 'layers\s*=\s*rank\s*<\s*1\s*\?\s*1\s*:\s*rank') {
        $bad += "U2 the fold no longer takes its layer count from the rank alone"
    }
    if (@([regex]::Matches($fb, 'out\s*\+=\s*"\["')).Count -ne 1) {
        $bad += "U2 the bracket is assembled in more than one place inside the fold"
    }
    if ($fb -notmatch 'idx\s*=\s*"0"') {
        $bad += "U2 the missing layer is no longer padded with 0 (an under-supplied subscript would emit field[] = invalid C)"
    }
}

# ---- U3: 维数的唯一来源是语法 ----
$push = @([regex]::Matches($parseText, 'moreDims\.push_back')).Count
if ($push -ne 1) { $bad += ("U3 parser pushes extra dimensions " + $push + " times (expected 1)") }
$store = @([regex]::Matches($parseText, 'memberNode->moreDims\s*=\s*std::move\(moreDims\)')).Count
if ($store -ne 1) { $bad += ("U3 parser hands the dims to the node " + $store + " times (expected 1)") }
if ($parseText -notmatch 'std::vector<TypeMember::MoreDim> moreDims') {
    $bad += "U3 the parser no longer collects the extra dims into a typed list"
}
$drop = 0
foreach ($l in ($parseCode -split "`r?`n")) {
    if ($l -match '解析并丢弃') { $drop++ }
}
if ($drop -ne 0) { $bad += ("U3 'parse and throw away' is back in the code (" + $drop + " line(s)) -- that discard IS ledger 262") }
# 反向牙口: 那条丢弃本来是"注释 + 一句不要返回值的 parseExpression()"。把赋值改回裸调用,
# 这一条就先红 (U3 的 push_back 那条同时也会红)。
$mdLower = @([regex]::Matches($parseCode, 'md\.lower\s*=\s*parseExpression\(\)')).Count
if ($mdLower -ne 1) { $bad += ("U3 the extra dimension's first expression is assigned " + $mdLower + " times (expected 1 -- it must be kept, not discarded)") }
if ($astText -notmatch 'int arrayRank\(\) const \{ return arraySize \? \(1 \+ \(int\)moreDims\.size\(\)\) : 0; \}') {
    $bad += "U3 TypeMember::arrayRank() no longer derives the rank from arraySize + moreDims (one source per dimension)"
}

# ---- U4: 语义层只传秩 ----
$wr = @([regex]::Matches($semText, 'mi\.arrayRank\s*=')).Count
if ($wr -ne 1) { $bad += ("U4 semantics writes arrayRank " + $wr + " times (expected exactly 1)") }
if ($semText -notmatch 'mi\.arrayRank\s*=\s*1\s*\+\s*\(int32_t\)memberPtr->moreDims\.size\(\)') {
    $bad += "U4 arrayRank is no longer counted from the syntax (rank only -- per-dim counts stay in cgen)"
}
foreach ($rel in @($declRel, $clsRel, $memberRel, $withmRel)) {
    $n = @([regex]::Matches((Get-SrcText $rel), 'arrayRank')).Count
    if ($n -gt 2) { $bad += ("U4 " + $rel + " reads arrayRank " + $n + " times (expected <= 2: the two subscript roads)") }
}

# ---- U5: 每维折一次, 只在这一处 ----
$bFirst = @([regex]::Matches($declText, 'brackets\s*=\s*"\["')).Count
$bMore  = @([regex]::Matches($declText, 'brackets\s*\+=\s*"\["')).Count
if ($bFirst -ne 1) { $bad += ("U5 the first dimension builds its bracket " + $bFirst + " times (expected 1)") }
if ($bMore -ne 1) { $bad += ("U5 the extra dimensions build their brackets " + $bMore + " times (expected exactly 1)") }
if ($declText -notmatch 'tryEvalConstInt\(md\.upper\.get\(\)') {
    $bad += "U5 the extra dimensions no longer fold through tryEvalConstInt (the same exit the first one uses)"
}
if (@([regex]::Matches($declText, 'for \(auto& md : member->moreDims\)')).Count -ne 1) {
    $bad += "U5 the walk over member->moreDims is missing or duplicated"
}

# ---- U6: 深拷贝助手: 扁平指针 + 先声明后用 ----
if ($clsText -notmatch 'int32_t _i; int32_t _n = \(int32_t\)\(sizeof\(d->" \+ fn\s*\+ "\) / sizeof\(\*_dp\)\)') {
    $bad += "U6 the owned-member copy no longer declares _i before using it / counts cells by element type"
}
$badIdx = @([regex]::Matches($clsText, 'sizeof\(d->" \+ idx')).Count
if ($badIdx -ne 0) { $bad += ("U6 the pre-265 shape sizeof(d->X[_i]) is back (" + $badIdx + " site(s)) -- that is C2065") }
if (-not $clsText.Contains('elemC + "* _dp = ("')) {
    $bad += "U6 the flat destination element pointer is missing"
}

# ---- U7: 元素档那道闸还问数组性 ----
$mNarrow = [regex]::Match($typeText, 'Vb6Type CCodeGen::narrowTargetTypeOf\s*\([^)]*\)\s*const\s*\{[\s\S]*?\n\}')
if (-not $mNarrow.Success) {
    $bad += "U7 narrowTargetTypeOf body not found in cgen_util_type.cpp"
} else {
    $nb = $mNarrow.Value
    if (@([regex]::Matches($nb, 'if \(!fldIsArray\) return Vb6Type::Unknown;')).Count -ne 1) {
        $bad += "U7 the member-element branch no longer refuses a non-array field (that is how a descriptor pointer gets wrapped into a scalar check = run-time error 6)"
    }
    if ($nb -notmatch 'inferUdtFieldVb6Type\(ioc\.callee\.get\(\)') {
        $bad += "U7 the member-element branch no longer asks the one field-declared-type exit"
    }
}

# ---- U8: ByRef 数组实参那一侧也按秩剥到底 ----
# 那处的旧写法是 `sizeof((arg)[0])` —— 秩 1 时正好是元素，秩 >=2 时是**一行**。
$argRel = "src\backend\detail\expr\cgen_expr_call_arg_emit.inc"
$argText = Get-SrcText $argRel
if (@([regex]::Matches($argText, 'mi44\.arrayRank')).Count -ne 1) {
    $bad += ("U8 the ByRef-array wrap asks the rank " + @([regex]::Matches($argText, 'mi44\.arrayRank')).Count + " times (expected 1)")
}
if ($argText -notmatch 'deep44 \+= "\[0\]"') {
    $bad += "U8 the ByRef-array wrap no longer peels the element chain down to the rank"
}
$blind = @([regex]::Matches($argText, 'sizeof\(\(" \+ argVal \+ "\)\[0\]\)\)')).Count
if ($blind -ne 0) { $bad += ("U8 the rank-blind sizeof((arg)[0]) is back (" + $blind + " site(s)) -- that counts ROWS on a multi-D member") }

if ($bad.Count -ne 0) {
    foreach ($b in $bad) { Write-Host ("FAIL " + $b) -ForegroundColor Red }
    exit 1
}
Write-Host "PASS udt_member_dims (U1 one subscript exit x2 roads, U2 layers from rank, U3 dims counted at parse, U4 rank-only metadata, U5 one fold per dimension, U6 flat element copy, U7 element gate asks array-ness)"
