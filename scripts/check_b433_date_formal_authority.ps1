# check_b433_date_formal_authority.ps1 - 「日期形参」与「数值形参」这两问不许拿 C 签名反推
#   (账 #433 §B207 第一百零五刀 + 账 #434 §B208 第一百零六刀)
#
# 为什么单开这一道: 「这一位实参要不要按**日期**解析」/「要不要按**数值**折算」这两个问句,
#   C 侧看不见答案 —— 日期形参与数值形参**都是 double**。而引擎对文本实参的答案**按 builtin
#   分两种**:
#     FormatDateTime("46027", 0) = 2026/1/5   (日期解析: 纯数字文本按序列号那天)
#     Sqr("4")                  = 2          (数值折算)
#     Sqr("2026/1/5")           = err 13     (判别格: 日期串进数值形参要**报错**)
#   拿 `rtParamType == "double"` / `paramBase == Double` 当判据, 会把两族一起互相误伤
#   (§B174 教训① 的同一形状: 那张「哪几位是字符串」的表按 C 签名反推就答错过)。所以这两问
#   **各自**带一份名单, 与 C 签名各走各的, 谁也不许替谁答。
#
# 关键口径 —— 判别格**逐位**成立, 不是逐 builtin: 同一枚 `DateAdd` 的两个位答案不同,
#   DateAdd("d","3","2026/1/5") = 2026/1/8 ⇒ 位 1(number) 走**数值**; 位 2(date) 走**日期**。
#   所以两张名单按**发码位号**登记, 且同一 builtin 的同一位不许同时出现在两张名单里(R7)。
#
# 规则（改坏了会红，不是装饰）:
#   R1  builtinNoParamTableFormalIsDate 必须存在，且**自带**一张表（不许退化成"读 C 签名"）
#   R2  名单里每一枚都必须在 cgen_util_type.cpp 的 kNoTableFormalKinds176 里也登了同一个名字
#       —— 因为发码那一臂要先 rtParamType 答得出 "double" 才会走到本问（否则整条臂跳过）
#   R3  名单里的位次必须 <= 同一行在 kNoTableFormalKinds176 登记的**长度**（不许指到不存在的位）
#   R4  **反名单**：kNoTableFormalKinds176 里所有登记了 double 形参、却**不在**日期名单里的
#       名字，必须逐个写明"为什么不按日期解析"（Sqr/Abs/CLng/CDbl/Cur/Format$ 那一族）
#   R5  发码那一臂必须真的按日期走 vb6_CDate，且**装箱**（vb6_CDate 的形参是 vb6_VARIANT，
#       裸发 BSTR 是反向 C2440 —— 本刀先踩过这一脚）
#   R6  数值那一族的名字**不许**出现在日期名单里（防止 R4 被自己抄破）
#
# 用法:  powershell -File scripts\check_b433_date_formal_authority.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$typeCpp = Join-Path $root 'src\backend\cgen_util_type.cpp'
$argInc = Join-Path $root 'src\backend\detail\expr\cgen_expr_call_arg_variant.inc'
foreach ($p in @($typeCpp, $argInc)) {
    if (-not (Test-Path -LiteralPath $p)) { Write-Host ("FAIL missing " + $p) -ForegroundColor Red; exit 1 }
}
$typeTxt = [IO.File]::ReadAllText($typeCpp)
$argTxt  = [IO.File]::ReadAllText($argInc)
$bad = @()

# ---- R2/R3/R6: 把两张表的对应关系读出来 ----
# 日期名单 = builtinNoParamTableFormalIsDate 里那张 kDateFormalPos433
$dm = [regex]::Match($typeTxt, '(?s)kDateFormalPos433\s*=\s*\{(?<body>.*?)\n\s*\};')
if (-not $dm.Success) {
    $bad += 'R1 builtinNoParamTableFormalIsDate has no kDateFormalPos433 table -> the question lost its own authority'
    Write-Host 'FAIL b433_date_formal_authority: builtinNoParamTableFormalIsDate has no table' -ForegroundColor Red
    exit 1
}
$datePos = @{}
foreach ($m in [regex]::Matches($dm.Groups['body'].Value, '\{"(?<n>vb6_[A-Za-z0-9_]+)",\s*\{(?<p>[0-9,\s]*)\}\}')) {
    $ps = @([regex]::Matches($m.Groups['p'].Value, '[0-9]+') | ForEach-Object { [int]$_.Value })
    $datePos[$m.Groups['n'].Value] = $ps
}
if ($datePos.Count -lt 8) {
    $bad += ('R1 date list holds only ' + $datePos.Count + ' names (floor 8) -> the scan or the table shrank')
}

# C 档表 = kNoTableFormalKinds176
$cm = [regex]::Match($typeTxt, '(?s)kNoTableFormalKinds176\s*=\s*\{(?<body>.*?)\n\s*\};')
if (-not $cm.Success) { $bad += 'R2 kNoTableFormalKinds176 table not found'; $cm = [pscustomobject]@{ Groups = @{ 'body' = @{ Value = '' } } } }
$cTable = @{}
foreach ($m in [regex]::Matches($cm.Groups['body'].Value, '\{"(?<n>vb6_[A-Za-z0-9_]+)",\s*\{(?<p>[^}]*)\}\}')) {
    $cTable[$m.Groups['n'].Value] = @([regex]::Matches($m.Groups['p'].Value, '"[^"]*"')).Count
}

# 账 #434 §B208: 数值名单先读出来 —— R4/R7/R8 都要用它, 所以**在 R4 之前**解析。
$numPos = @{}
$numPosNames = @{}
$nm = [regex]::Match($typeTxt, '(?s)kNumericFoldPos434\s*=\s*\{(?<body>.*?)\n\s*\};')
$nmOk = $nm.Success
if ($nmOk) {
    foreach ($m in [regex]::Matches($nm.Groups['body'].Value, '\{"(?<n>vb6_[A-Za-z0-9_]+)",\s*\{(?<p>[0-9,\s]*)\}\}')) {
        $numPosNames[$m.Groups['n'].Value] = $true
        $numPos[$m.Groups['n'].Value] = @([regex]::Matches($m.Groups['p'].Value, '[0-9]+') | ForEach-Object { [int]$_.Value })
    }
}

foreach ($n in @($datePos.Keys | Sort-Object)) {
    if (-not $cTable.ContainsKey($n)) {
        $bad += ('R2 ' + $n + ' is in the date list but NOT in kNoTableFormalKinds176 -> rtParamType stays empty, ' +
                 'so the emit arm never reaches the date question (this is why it silently did nothing)')
        continue
    }
    foreach ($p in $datePos[$n]) {
        if ($p -ge $cTable[$n]) {
            $bad += ('R3 ' + $n + ' date slot ' + $p + ' is beyond its registered arity ' + $cTable[$n] +
                     ' -> the slot does not exist in the C-type table')
        }
    }
}

# ---- R4: double 形参却不在日期名单的，必须逐个有理由 ----
$NOT_DATE = @{
    'vb6_Abs' = '绝对值: 数值形参, 文本实参走数值折算(引擎 Abs("-4")=4)'
    'vb6_Sgn' = '符号: 同上, 数值形参'
    'vb6_Sqr' = '开方: 引擎 Sqr("4")=2 —— 本族反例, 按日期解析会答 1899 年'
    'vb6_CLng' = 'CLng 形参是数值, 文本按 Val 折算'
    'vb6_CDbl' = 'CDbl 形参是数值, 文本按 Val 折算'
    'vb6_CCur' = 'CCur 形参是数值(货币), 文本按 Val 折算'
}
foreach ($m in [regex]::Matches($cm.Groups['body'].Value, '\{"(?<n>vb6_[A-Za-z0-9_]+)",\s*\{(?<p>[^}]*)\}\}')) {
    $n = $m.Groups['n'].Value
    $p = $m.Groups['p'].Value
    if ($p -notmatch '"double"') { continue }          # 只查有 double 形参的行
    if ($datePos.ContainsKey($n)) { continue }        # 日期名单里的不在本检查范围
    # 账 #434 §B208: 数值名单里的(那一族走数值折算，见 kNumericFoldPos434)也不算漏网 ——
    #   Sqr/Abs/Int/Fix/Sgn/Sin/Cos/Tan/Atn/Log/Exp/Round 现在**有**答案了。
    if ($numPosNames.ContainsKey($n)) { continue }
    if (-not $NOT_DATE.ContainsKey($n)) {
        $bad += ('R4 ' + $n + ' registers a double formal but is in NEITHER the date list NOR a written reason ' +
                 '-> either answer it as a date or write down why it is not one (silent hole)')
    }
}
foreach ($n in @($NOT_DATE.Keys)) {
    if (([string]($NOT_DATE[$n])).Length -lt 12) {
        $bad += ('R4 not-date entry ' + $n + ' has a reason shorter than 12 chars -> an empty reason is a silent hole')
    }
}

# R5 追加: 数值那一臂必须走 vb6_CDbl，且不许落到 vb6_CDate / vb6_Val
if ($argTxt -notmatch 'vb6_CDbl\(vb6_VariantFromValue\(') {
    $bad += 'R5 the emit arm does not emit vb6_CDbl(vb6_VariantFromValue(...)) for numeric formals -> the numeric ' +
             'half of the question is unanswered (bare BSTR = C2440)'
}
if ($argTxt -match 'vb6_Val\(') {
    $bad += 'R5 the emit arm reaches for vb6_Val -> Val is lenient (stops at the first non-numeric char, never ' +
             'raises); the engine reports 13 on "abc", so Val answers the wrong question'
}

# ---- 账 #434 §B208: 数值名单的逐位互斥 (名单本身已在上面解析) ----
if (-not $nmOk) {
    $bad += 'R7 builtinNoParamTableFormalWantsNumericFold has no kNumericFoldPos434 table -> the numeric half lost its authority'
} else {
    if ($numPos.Count -lt 10) {
        $bad += ('R7 numeric list holds only ' + $numPos.Count + ' names (floor 10) -> the table shrank')
    }
    # 逐位互斥: 同一枚 builtin 的同一个位不许同时在日期名单与数值名单里。
    # 这一格是判别式的代码化 —— 引擎 DateAdd("d","3",d)=+3 天(数值) 而 DateAdd("d",3,"2026/1/5")
    # = 那天(日期); 两张名单抄错一位, 读数就从「+3 天」变成「1900 年那天」。
    foreach ($n in @($numPos.Keys)) {
        foreach ($p in $numPos[$n]) {
            if ($datePos.ContainsKey($n) -and ($datePos[$n] -contains $p)) {
                $bad += ('R7 ' + $n + ' slot ' + $p + ' is in BOTH the date list and the numeric list -> one slot cannot ' +
                         'want both a date parse and a numeric fold (engine DateAdd("d","3",d) = +3 days)')
            }
        }
        foreach ($p in $numPos[$n]) {
            if ($cTable.ContainsKey($n) -and $p -ge $cTable[$n]) {
                $bad += ('R7 ' + $n + ' numeric slot ' + $p + ' is beyond its registered arity ' + $cTable[$n])
            }
        }
    }
    # R8: 数值名单的位次必须存在 —— 但**不要求**在 kNoTableFormalKinds176 里登过。
    # 为什么这一族与日期那一族的条件不同: 日期那一臂历史上要 rtParamType 答出 "double"
    #   才走得到 (那正是 §B207 第一版只治好 1/9 的成因, R2/R3 钉的就是它)。而数值那一臂
    #   **落在两条臂之外**(与 §B207 那条同位), 不看 rtParamType ⇒ 一元数值族
    #   (Sqr/Abs/Int/…) 两张表都没登也照样走到 (实测 Sqr(s)=2 编译+读数都对)。
    #   把 R2 那条复制过来会把 12 枚全判成「走不到」= 假红, 而真红反而被淹掉。
    # 真正要钉的是: 名单里的位不许越界, 且不许与日期名单同位互撞(上面已查)。
    foreach ($n in @($numPos.Keys | Sort-Object)) {
        foreach ($p in $numPos[$n]) {
            if ($cTable.ContainsKey($n) -and $p -ge $cTable[$n]) {
                $bad += ('R8 ' + $n + ' numeric slot ' + $p + ' is beyond its registered arity ' + $cTable[$n] +
                         ' in kNoTableFormalKinds176')
            }
        }
    }
}

# ---- R5 追加: 数值那一臂必须走 vb6_CDbl，且不许落到 vb6_CDate / vb6_Val ----
if ($argTxt -notmatch 'vb6_CDbl\(vb6_VariantFromValue\(') {
    $bad += 'R5 the emit arm does not emit vb6_CDbl(vb6_VariantFromValue(...)) for numeric formals -> the numeric ' +
             'half of the question is unanswered (bare BSTR = C2440)'
}
if ($argTxt -match 'vb6_Val\(') {
    $bad += 'R5 the emit arm reaches for vb6_Val -> Val is lenient (stops at the first non-numeric char, never ' +
             'raises); the engine reports 13 on "abc", so Val answers the wrong question'
}

# ---- R6: 数值那一族不许被抄进日期名单 ----
foreach ($n in @($NOT_DATE.Keys)) {
    if ($datePos.ContainsKey($n)) {
        $bad += ('R6 ' + $n + ' is a numeric-formal builtin but appears in the DATE list -> numeric text ' +
                 'would be parsed as a date (engine Sqr("4")=2, not a 1900 date)')
    }
}

# ---- R5: 发码那一臂必须真的走 vb6_CDate 且装箱 ----
if ($argTxt -notmatch 'vb6_CDate\(vb6_VariantFromValue\(') {
    $bad += 'R5 the emit arm does not box the text before vb6_CDate -> vb6_CDate takes vb6_VARIANT, ' +
             'so a bare BSTR is the reverse C2440 (BSTR -> vb6_VARIANT); this knife hit that first'
}
if ($argTxt -notmatch 'builtinNoParamTableFormalIsDate') {
    $bad += 'R5 the emit arm never calls builtinNoParamTableFormalIsDate -> the date question has two answers again'
}
if ($argTxt -match 'argVal\s*=\s*"vb6_DateValue\(') {
    $bad += 'R5 the emit arm uses vb6_DateValue -> engine DateValue("46027") reports 13 (strict), ' +
             'but FormatDateTime("46027",0) answers 2026/1/5 (permissive); the strict outlet kills a legal arg'
}
# 判别式不许退化成 C 签名
if ($argTxt -match 'rtParamType\s*==\s*"double"[^\n]*\n[^\n]*builtinNoParamTableFormalIsDate') {
    $bad += 'R5 the emit arm keys the date conversion off rtParamType == "double" -> that is exactly the ' +
             'C-signature inference this knife exists to forbid (it would capture Sqr/Abs/CLng)'
}

if ($bad.Count -eq 0) {
    Write-Host ('PASS b433_date_formal_authority: date list ' + $datePos.Count + ' names / ' +
                (($datePos.Values | ForEach-Object { $_.Count }) | Measure-Object -Sum).Sum + ' slots | ' +
                'numeric list ' + $numPos.Count + ' names / ' +
                (($numPos.Values | ForEach-Object { $_.Count }) | Measure-Object -Sum).Sum + ' slots | ' +
                'not-date ' + $NOT_DATE.Count + ' | C-table rows ' + $cTable.Count + ' | violations 0') -ForegroundColor Green
    exit 0
}
foreach ($b in $bad) { Write-Host ('FAIL ' + $b) -ForegroundColor Red }
exit 1