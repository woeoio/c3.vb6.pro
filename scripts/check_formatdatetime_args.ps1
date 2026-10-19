# ============================================================
# check_formatdatetime_args.ps1 -- 账 #305 / §B206 / 第一百零三刀 (2026-10-17)
#
# 「FormatDateTime(<一枚槽>, n) 那两位各要什么 C 档」这一问之前**两级答案全空** —— 没有类型化
# 形参表的那族内在函数先问 kNoTableFormalKinds176 (src/backend/cgen_util_type.cpp)，答不出才回落
# getRuntimeParamCType (同文件)，两边都没有 vb6_FormatDateTime 那一行 ⇒ 发码里 rtParamType 为空 ⇒
# 不包 ⇒ 整枚 vb6_VARIANT 裸交进 `double dateSerial` / `int32_t namedFormat` = cl C2440，**整份
# 工程编不过** (第一百零二刀 BASE 那台 14 条 C2440，两台同一行集，尺 .build/k691_base_read.txt)。
# 真语料 D:/vb_yqt4qPac 里 FormatDateTime 44 处 / 15 份，第一实参明确不是数值那一档的 13 处全吃
# 这一格 (尺 .build/k611_fdt_args.txt)；逐形真编译 = 十二形里八形硬错误 / 四形编得过
# (.build/k672_k103_readings.md)。
#
# 本哨兵钉的是**这条问句的形状**，不是「原型 == 表行」: §B174 那条教训还在 —— CStr / CInt /
# WeekdayName / Partition / NPer / NPV 那几枚是 _Generic 宏或 inline 包装，RTL 原型压根不答这一问，
# 所以 R4 的 census 只问**两张表答不答得出**，答不出的名字全列在下面那份名单里当**刻意排除**。
# 名单不许悄悄变长: 新登记一枚有参内在函数而不给它档位答案，R4 就红。
#
# 规则: R1 表里 vb6_FormatDateTime 那一行恰好一行、两档 = double,int32_t (三种写坏的形状各钉 0);
#       R2 那张表本身良构 (行数 21、键唯一、键一律 vb6_ 前缀、FormatDateTime 那一行 arity 2);
#       R3 站点次序仍是「先问那张表、答不出才回落」，且回落只此一处;
#       R4 census 两头各钉一份名单 (识别表里没有 callee 的 7 枚 + 两级都没答案的 34 枚);
#       R5 判据两头都在 (b433 两片 cell、期望表 12 行、夹具 12 针、本哨兵已登记**且被调用**)。
# 本文件里不写反斜杠字面量: 路径走 [IO.Path]::DirectorySeparatorChar，计数走 [regex]::Escape。
# ============================================================

param(
    [string]$RepoRoot = (Split-Path -Parent $PSScriptRoot)
)

$ErrorActionPreference = 'Stop'
$fail = @()
$SQ = [char]39
$DQ = [char]34
$NL = [char]10

function C([string]$lit, [string]$txt) { return ([regex]::Matches($txt, [regex]::Escape($lit))).Count }
function Want([string]$rule, [int]$got, [int]$expect, [string]$why) {
    if ($got -ne $expect) { $script:fail += ($rule + ' = ' + $got + ' (want ' + $expect + '): ' + $why) }
}
# 名字别起 RD / CD / MD: PowerShell 的别名**排在函数前面**, rd 就是 Remove-Item。
function ReadSrc([string]$root, [string]$rel) {
    $p = Join-Path $root ($rel.replace('/', $script:sep))
    if (-not (Test-Path $p)) { Write-Output ('FAIL: missing ' + $p); exit 1 }
    return [System.IO.File]::ReadAllText($p).Replace("`r`n", "`n")
}
function TableBody([string]$txt, [string]$anchor, [string]$rule) {
    $i = $txt.IndexOf($anchor)
    if ($i -lt 0) { $script:fail += ($rule + ' anchor missing: ' + $anchor); return '' }
    $j = $txt.IndexOf($NL + '    };', $i)
    if ($j -lt 0) { $script:fail += ($rule + ' end of table not found after the anchor'); return '' }
    return $txt.Substring($i, $j - $i)
}

$sep = [IO.Path]::DirectorySeparatorChar
$TAB = ReadSrc $RepoRoot 'src/semantics/builtin/builtin_funcs.inc'
$CT = ReadSrc $RepoRoot 'src/backend/cgen_util_type.cpp'
$AR = ReadSrc $RepoRoot 'src/backend/detail/expr/cgen_expr_call_arg_variant.inc'
$NM = ReadSrc $RepoRoot 'src/backend/detail/expr/cgen_expr_ident_builtin_table.inc'
$RT = ReadSrc $RepoRoot 'tests/run_tests.ps1'
$FX = ReadSrc $RepoRoot 'tests/b433/B433Main.bas'

$ROWPAT = '\{\s*"([A-Za-z0-9_]+)",\s*\{([^}]*)\}\s*\}'
$B176 = TableBody $CT 'kNoTableFormalKinds176 = {' 'R2'
$BRT = TableBody $CT 'std::string CCodeGen::getRuntimeParamCType(const std::string& funcName, size_t paramIdx) {' 'R4'

# ---- R1 那一行只有形状对了才有人发 vb6_VariantToDouble / vb6_VariantToLong ----
$want1 = '{"vb6_FormatDateTime", {"double", "int32_t"}}'
Want 'R1 the row as shipped' (C $want1 $B176) 1 '表里那一行必须是 double,int32_t 这两档 (顺序也算答案)'
Want 'R1 tiers swapped' (C '{"vb6_FormatDateTime", {"int32_t", "double"}}' $B176) 0 '两位反了 = 把槽解成 int 交进日期形参'
Want 'R1 one tier only' (C '{"vb6_FormatDateTime", {"double"}}' $B176) 0 '只答第一位 = NamedFormat 那位仍是裸槽'
Want 'R1 three tiers' (C '{"vb6_FormatDateTime", {"double", "int32_t", "int32_t"}}' $B176) 0 '这一族只有两条形参，第三档是凭空加的'
$fdtRows = [regex]::Matches($B176, '"vb6_FormatDateTime"').Count
Want 'R1 one row in the table' $fdtRows 1 '这一枚 callee 在那张表里只许登记一次'

# ---- R2 那张表良构 ----
$rows = @([regex]::Matches($B176, $ROWPAT) | ForEach-Object { $_.Groups[1].Value })
$keys = @{}
$badPrefix = 0
foreach ($r in $rows) {
    if (-not $keys.ContainsKey($r)) { $keys[$r] = 0 }
    $keys[$r] = $keys[$r] + 1
    if (-not $r.StartsWith('vb6_')) { $badPrefix++ }
}
$dups = 0
foreach ($kk in $keys.Keys) { if ($keys[$kk] -gt 1) { $dups++ } }
Want 'R2 row count' $rows.Count 21 '那张表现有 21 行 (11 + 账 #433 §B207 补的 10 枚日期形参: Year/Month/Day/Hour/Minute/Second/Weekday/DateAdd/DateDiff/DatePart); 加了人就来这里把计数与台账一起改 (减人更要改)'
Want 'R2 duplicate keys' $dups 0 '同一枚 callee 登记两行 = 后一行没人读，前一行骗人'
Want 'R2 vb6_ prefix' $badPrefix 0 '键是发码里的 callee 名，裸名当键永远答不出'
$fdtTiers = ''
foreach ($m in [regex]::Matches($B176, $ROWPAT)) {
    if ($m.Groups[1].Value -eq 'vb6_FormatDateTime') { $fdtTiers = $m.Groups[2].Value }
}
$arity = 0
if ($fdtTiers.Length -gt 0) { $arity = @($fdtTiers.split(',')).Count }
Want 'R2 FormatDateTime arity' $arity 2 '两条形参: dateSerial 与 namedFormat (省略形参那一档由 #428 的默认值表管)'

# ---- R3 站点次序: 先问那张表，答不出才回落 ----
Want 'R3 the ternary' (C 'rtParamType = kind176.empty() ? getRuntimeParamCType(callee, i) : kind176;' $AR) 1 '那一站的次序 = 唯一权威先答，回落只补它答不出的'
Want 'R3 fallback used once' (C 'getRuntimeParamCType(callee, i)' $AR) 1 '回落那张表在本文件里只许被问一次 (两处问 = 两个主人)'
Want 'R3 table asked once' (C 'builtinNoParamTableFormalCType(' $AR) 1 '问唯一权威那一处在这一臂里只此一处'

# ---- R4 census: 没有类型化形参表的那族内在函数，档位要么有人答、要么在名单里 ----
$withP = @{}
foreach ($m in [regex]::Matches($TAB, 'addBuiltinFuncWithParams\("([A-Za-z0-9_]+)"')) { $withP[$m.Groups[1].Value] = 1 }
$noTable = @{}
foreach ($m in [regex]::Matches($TAB, 'addBuiltinFunc\("([A-Za-z0-9_]+)"')) {
    if (-not $withP.ContainsKey($m.Groups[1].Value)) { $noTable[$m.Groups[1].Value] = 1 }
}
$entryOf = @{}
foreach ($m in [regex]::Matches($NM, '\{"([a-z0-9_]+)",\s*"(vb6_[A-Za-z0-9_]+)"\}')) {
    $k = $m.Groups[1].Value
    if (-not $entryOf.ContainsKey($k)) { $entryOf[$k] = @() }
    $entryOf[$k] += $m.Groups[2].Value
}
$k176 = @{}
foreach ($m in [regex]::Matches($B176, $ROWPAT)) { if ($m.Groups[1].Value.StartsWith('vb6_')) { $k176[$m.Groups[1].Value] = 1 } }
$krt = @{}
foreach ($m in [regex]::Matches($BRT, $ROWPAT)) { if ($m.Groups[1].Value.StartsWith('vb6_')) { $krt[$m.Groups[1].Value] = 1 } }

$noCallee = @()
$noAnswer = @()
foreach ($n in $noTable.Keys) {
    $e = $entryOf[$n.ToLower()]
    if (-not $e) { $noCallee += $n; continue }
    $answered = $false
    foreach ($x in $e) { if ($k176.ContainsKey($x) -or $krt.ContainsKey($x)) { $answered = $true } }
    if (-not $answered) { $noAnswer += $n }
}
$cmp = [System.StringComparer]::Ordinal
$noCallee = [string[]]@($noCallee)
$noAnswer = [string[]]@($noAnswer)
[Array]::Sort($noCallee, $cmp)
[Array]::Sort($noAnswer, $cmp)
# 名单一 = 内在函数名在识别表里压根没有 callee 那一行 (变长实参 / 别处特判)，它们不落这张表:
$WANT_NO_CALLER = 'Array,Beep,Choose,IIf,If,Switch,VarPtr'
# 名单二 = 有 callee、但两级答案全空的那些 (§B176 那族里剩下的整族 census，本刀只补 FormatDateTime
# 一站，其余刻意不动 —— 零实参那一档 (Date / Now / Time / Timer / Erl / FreeFile / Command) 与
# _Generic 宏 / inline 包装 (CVar / DoEvents) 也在名单里，等各自的量收完再逐格答):
$WANT_NO_ANSWER = 'AppActivate,CVar,Command,DDB,Date,DoEvents,Erl,FV,FormatCurrency,FormatNumber,FormatPercent,FreeFile,IPmt,IRR,IsNothing,LSet,Load,MIRR,MonthName,NPV,NPer,Now,PPmt,PV,Partition,Pmt,RATE,RSet,Rnd,SLN,SYD,Time,Timer,WeekdayName'
$gotNoCaller = ($noCallee -join ',')
$gotNoAnswer = ($noAnswer -join ',')
if ($gotNoCaller -cne $WANT_NO_CALLER) {
    $fail += ('R4 no-callee set differs (got ' + $noCallee.Count + '): ' + $gotNoCaller + ' | want ' + $WANT_NO_CALLER)
}
if ($gotNoAnswer -cne $WANT_NO_ANSWER) {
    $fail += ('R4 unanswered set differs (got ' + $noAnswer.Count + '): ' + $gotNoAnswer + ' | want ' + $WANT_NO_ANSWER)
}
if ($k176.ContainsKey('vb6_FormatDateTime') -eq $false) {
    $fail += 'R4 vb6_FormatDateTime left the answered side of the census (那一行被删了 = 账 #432 回潮)'
}

# ---- R5 判据两头都在 ----
Want 'R5 cells' (C ('Test-Vbp ' + $SQ + 'b433_formatdatetime_args') $RT) 2 '两台各一片 (x64 + x86)'
$tstart = $RT.IndexOf('$b433Expected = @(')
if ($tstart -lt 0) { $fail += 'R5 expectation table missing ($b433Expected)'; $trows = 0 }
else {
    $tend = $RT.IndexOf(')', $tstart)
    $trows = [regex]::Matches($RT.Substring($tstart, $tend - $tstart), '"B433-').Count
}
Want 'R5 table rows' $trows 12 '期望表 11 针 + 一枚 DONE'
Want 'R5 fixture prints' (C ('Debug.Print ' + $DQ + 'B433-') $FX) 12 '夹具里 11 针 + DONE 证人 (两头都要，只钉一头会漏改)'
Want 'R5 stray escaped arch' (C '-Arch "x86\' $RT) 0 '那一行结尾多一个反斜杠会把 C3 的参数打成 "x86\" (第九十八刀真栽过)'
Want 'R5 sentinel registered' (C 'check_formatdatetime_args.ps1' $RT) 1 '本哨兵在 run_tests.ps1 里被调用一次'
Want 'R5 sentinel function' (C 'function Test-FormatDateTimeArgs' $RT) 1 '那道静态检查自己有一处定义'
Want 'R5 sentinel called' ([regex]::Matches($RT, '(?m)^    Test-FormatDateTimeArgs[ \t]*\r?$').Count) 1 '那条裸调用自己有一行 (只钉定义 = 拆掉调用照样绿，第一百零二刀 P12 收的纪律)'

if ($fail.Count -gt 0) {
    foreach ($m in $fail) { Write-Output ('FAIL ' + $m) }
    Write-Output ('FAIL check_formatdatetime_args: ' + $fail.Count + ' rule(s) broken')
    exit 1
}
Write-Output ('PASS check_formatdatetime_args: rows=' + $rows.Count +
    ' arity=' + $arity + ' noTable=' + $noTable.Count +
    ' noCaller=' + $noCallee.Count + ' noAnswer=' + $noAnswer.Count +
    ' cells=' + (C ('Test-Vbp ' + $SQ + 'b433_formatdatetime_args') $RT) +
    ' needles=' + (C ('Debug.Print ' + $DQ + 'B433-') $FX))
exit 0
