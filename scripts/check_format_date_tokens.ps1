# ============================================================
# check_format_date_tokens.ps1 —— 账 #305 / §B201 / 第一百刀 (2026-10-17)
#
# Format 的日期面新认了五枚令牌: w / ww / q / ttttt / ddddd。三条结构性规则钉住它:
#   ① 「这枚格式串算不算日期格式」只许住一处, 两条臂各自显式表态 (allowWQT 传 1 / 0)。数值臂那条还带着
#      「串里有 0 或 # 就不算日期」这道闸: 语料把单位文本裸写在数值格式串里 (#,##0 Twip 3 / T0 2 /
#      #### Byte 1 / Settings 4 = 15 处), 裸字母一视同仁就把那 15 处从「那个数」打成日期 (尺
#      .build/k496_census.txt + k497_wqt_where.txt; 现状读数 .build/k504 / k505 / k511)。
#   ② 「一年里第几周」与「星期序号」两问各只有一位主人 (vb6_WeekOfYear / vb6_Weekday)。抄第二份公式
#      正是这一族以前坏掉的原因, 而本刀给周数那一问添了第二个读者 —— 它不再是 date.c 的内部静态函数。
#   ③ 「locale 的短日期 / 长时间」在 Format 家族里收成一条出口, 六处读者共吃它;
#      vb6_Format 那两条形参不许再被 (void) 丢掉。
#
# 第一百零一刀订正了 R5 那一组: 「locale 短日期 / 长时间」的读者从 4+4 变成 2+2 再各加一枚新出口的读者 ——
#   general date 那一档与无格式串那条兜底臂现在都吃 vb6_fmtLocaleGeneralDate, 由它去问这两条出口 (§B203)。
# 判据面: tests/b430 二十七针 (两台逐字节相同)。B430-24/25 是「数值臂没被牵连」的反面证人,
# B430-19/20/21 是「两条形参真被读」的证人, B430-23 是「单枚 t 仍是字面量」的证人。
# 本文件里不写反斜杠字面量: 字面计数走 [regex]::Escape, 真正则的转义符全部由 [char]92 现拼
# (工具层会先折一级反斜杠, 那道坑本线已经踩过)。
# ============================================================

param(
    [string]$RepoRoot = (Split-Path -Parent $PSScriptRoot)
)

$ErrorActionPreference = 'Stop'
$fail = @()
$SQ = [char]39   # 单引号
$DQ = [char]34   # 双引号
$B  = [char]92   # 反斜杠 (只给真正则用)

# 字面量计数: 不碰正则语法, 免得数错
function C([string]$lit, [string]$txt) { return ([regex]::Matches($txt, [regex]::Escape($lit))).Count }
# 真正则计数 (Multiline: 行首/行尾锚点用得上)
function N([string]$pat, [string]$txt) { return ([regex]::Matches($txt, $pat, 'Multiline')).Count }
function Want([string]$rule, [int]$got, [int]$expect, [string]$why) {
    if ($got -ne $expect) { $script:fail += ($rule + ' = ' + $got + ' (want ' + $expect + '): ' + $why) }
}

$sep = [IO.Path]::DirectorySeparatorChar
$files = [ordered]@{
    fmt   = 'src/rtl/core/vb6rtl/vb6rtl_format.c'
    date  = 'src/rtl/core/vb6rtl/vb6rtl_date.c'
    hdr   = 'src/rtl/core/vb6rtl/vb6rtl_builtin.h'
    extr  = 'src/rtl/core/vb6rtl/detail/vb6rtl_format_extract.inc'
    parse = 'src/rtl/core/vb6rtl/detail/vb6rtl_format_parse.inc'
    run   = 'tests/run_tests.ps1'
    fx    = 'tests/b430/B430Main.bas'
}
$T = @{}
foreach ($k in $files.Keys) {
    $p = Join-Path $RepoRoot ($files[$k].replace('/', $sep))
    if (-not (Test-Path $p)) { Write-Output ('FAIL: missing ' + $p); exit 1 }
    $T[$k] = [System.IO.File]::ReadAllText($p)
}
$F = $T['fmt']; $DC = $T['date']; $HD = $T['hdr']; $EX = $T['extr']; $PA = $T['parse']; $RT = $T['run']; $FX = $T['fx']
$BOTH = $F + $EX
$GATE = 'static int vb6_fmtIsDateFormat(BSTR fmt, int allowWQT) {'

# ---- R1 门票: 一处定义、两条臂各自表态、没有第三种问法 ----
Want 'R1 gate definition' (C $GATE $F) 1 '只许一处定义'
Want 'R1 date-valued station' (C 'vb6_fmtIsDateFormat(fmt, 1)' $EX) 1 'extract.inc 那条 VT_DATE 臂必须传 1'
Want 'R1 numeric station' (C 'vb6_fmtIsDateFormat(fmt, 0)' $PA) 1 'parse.inc 那条数值兜底臂必须传 0'
Want 'R1 bare call' (C 'vb6_fmtIsDateFormat(fmt)' $BOTH) 0 '不许有第三种不带档的调用 (两条臂必须各自表态)'

# ---- R2 数值臂那道闸 + 三枚令牌的游程规则 ----
Want 'R2 placeholder guard' (C ('if (fmt[i] == L' + $SQ + '0' + $SQ + ' || fmt[i] == L' + $SQ + '#' + $SQ + ') return 0;') $F) 1 '数值占位符闸只许一处'
Want 'R2 allowWQT branch' (C 'if (!allowWQT) {' $F) 1 '那道闸挂在 allowWQT 这一问后面'
Want 'R2 w-or-q' (C ('if (lc == L' + $SQ + 'w' + $SQ + ' || lc == L' + $SQ + 'q' + $SQ + ') return 1;') $F) 1 'w / q 认单枚: 单枚 w 一视同仁会吃掉语料里 Twip 那 3 处, 那道伤靠下面 R2 的数值占位符闸挡住'
Want 'R2 t' (C ('if (lc == L' + $SQ + 't' + $SQ + ' && j - i >= 5) return 1;') $F) 1 't 只认五连 ttttt; 单枚 t 必须是字面量 (ISO 串里那枚 T)'

# ---- R3 三枚新臂各一处, 且都问现成的主人 ----
foreach ($c in @('q', 'w', 't')) {
    Want ('R3 case ' + $c) (C ('case L' + $SQ + $c + $SQ + ':') $F) 1 ('日期臂里 case L' + $SQ + $c + $SQ + ' 恰好一处')
}
Want 'R3 week outlet asked' (C 'vb6_WeekOfYear((int32_t)serial, firstDayOfWeek, firstWeekOfYear)' $F) 1 'ww 必须问那条唯一出口'
Want 'R3 weekday outlet asked' (C 'vb6_Weekday(serial, 1)' $F) 1 'w 必须问 vb6_Weekday, 不许再抄 mod-7'
Want 'R3 second week formula' (C '(doy - 1) / 7' $F) 0 'Format 里不许出现第二份周数公式'

# ---- R4 两条形参真被读 + 五连那一档只两处 ----
Want 'R4 voided formals' (C '(void)firstDayOfWeek' $F) 0 'vb6_Format 不许再把两条形参丢掉'
Want 'R4 formatter arity' (C 'static BSTR vb6_fmtDateSerial(double serial, BSTR fmt, int32_t firstDayOfWeek, int32_t firstWeekOfYear)' $F) 1 '日期臂自己得带着这两条形参'
Want 'R4 date-value call' (C 'vb6_fmtDateSerial(expr.dblVal, fmt, firstDayOfWeek, firstWeekOfYear)' $EX) 1 'extract.inc 把两条形参递下去'
Want 'R4 numeric call' (C 'vb6_fmtDateSerial(numVal, fmt, firstDayOfWeek, firstWeekOfYear)' $PA) 1 'parse.inc 把两条形参递下去'
Want 'R4 five-run arms' (C 'if (cnt >= 5) {' $F) 2 '五连只有两处: 日档 ddddd 与 t 档 ttttt'
Want 'R4 dddd intact' (C 'else if (cnt == 4) vb6_fmtAppendW(out, 256, &o, kDayLong[dow]);' $F) 1 'dddd 那一档不许被牵连'

# ---- R5 locale 两问各一条出口, 六处读者共吃 ----
Want 'R5 short-date outlet' (C 'void vb6_fmtLocaleShortDate(' $F) 1 '一条出口 (第一百零一刀起不再是 static)'
Want 'R5 long-time outlet' (C 'void vb6_fmtLocaleLongTime(' $F) 1 '一条出口 (第一百零一刀起不再是 static)'
Want 'R5 short-date readers' (C 'vb6_fmtLocaleShortDate(&st' $BOTH) 2 '读者 = short date 那一档 + ddddd (第一百零一刀: general date 与无格式串兜底改吃那条新出口)'
Want 'R5 short-date via new outlet' (C 'vb6_fmtLocaleShortDate(st,' $BOTH) 1 '新出口自己也是这一问的读者, 且只此一处'
Want 'R5 long-time readers' (C 'vb6_fmtLocaleLongTime(&st' $BOTH) 2 '读者 = long time 那一档 + ttttt (同上, general date 那两站搬到新出口里去了)'
Want 'R5 long-time via new outlet' (C 'vb6_fmtLocaleLongTime(st, out + n + 1' $BOTH) 1 '新出口拼时间那一段只许这一处'
Want 'R5 raw short-date' (C 'GetDateFormatW(LOCALE_USER_DEFAULT, DATE_SHORTDATE' $BOTH) 1 'DATE_SHORTDATE 只许住在那条出口里'
Want 'R5 raw long-time' (C 'GetTimeFormatW(LOCALE_USER_DEFAULT, 0, st, NULL' $BOTH) 1 'locale 默认时间格式只许住在那条出口里'

# ---- R6 星期序号那一问也只一位主人 (第一百刀并掉的第二份抄写) ----
Want 'R6 DatePart asks outlet' (C 'return vb6_Weekday((double)d, 1);' $DC) 1 'DatePart 的 w 臂改问出口'
Want 'R6 hand remap gone' (C 'if (wd == 0) return 7;' $DC) 0 '手抄的周六 remap 不许回来'
Want 'R6 outlet declared' (C 'int32_t vb6_WeekOfYear(int32_t serial, int32_t fdow, int32_t fwoy);' $HD) 1 '周数出口得有声明'
Want 'R6 outlet not static' (C 'static int32_t vb6_WeekOfYear(' $DC) 0 '定义不再是 file-static (它有两位读者)'

# ---- R7 判据两头都在: 定义了没被调用 = 静默不跑的那一格 ----
Want 'R7 b430 cells' (C ('Test-Vbp ' + $SQ + 'b430_format_date') $RT) 2 'run_tests.ps1 里两台各一行'
Want 'R7 expectation table' (C '$b430Expected = @(' $RT) 1 '期望表恰好一张'
Want 'R7 x86 cell shape' (N ('Test-Vbp ' + $SQ + 'b430_format_date_x86' + $SQ + ' [^' + $B + 'r' + $B + 'n]*-Arch ' + $DQ + 'x86' + $DQ + $B + 'r?$') $RT) 1 'x86 那一整行的形状 (结尾多一个反斜杠就会死在编译期)'
Want 'R7 stray arch backslash' (C ('-Arch ' + $DQ + 'x86' + $DQ + $B + $B) $RT) 0 '不许有带反斜杠的 -Arch'
Want 'R7 sentinel function' (N '^function Test-FormatDateTokens ' $RT) 1 '静态哨兵的函数恰好定义一处'
Want 'R7 sentinel called' (N '^    Test-FormatDateTokens' $RT) 1 '定义了必须被调用'
Want 'R7 self registered' (C 'check_format_date_tokens' $RT) 1 '本哨兵得写进 runner'
Want 'R7 fixture needles' (C ('Debug.Print ' + $DQ + 'B430-') $FX) 28 '夹具二十七针加一枚 DONE 证人都在'
Want 'R7 fixture witness' (C ('Debug.Print ' + $DQ + 'B430-DONE' + $DQ) $FX) 1 '跑到底的那枚证人'

if ($fail.Count -gt 0) {
    foreach ($m in $fail) { Write-Output ('FAIL ' + $m) }
    Write-Output ('FAIL check_format_date_tokens: ' + $fail.Count + ' rule(s) broken')
    exit 1
}
Write-Output ('PASS check_format_date_tokens: gate=' + (C $GATE $F) +
    ' stations=' + (C 'vb6_fmtIsDateFormat(fmt, 1)' $EX) + '/' + (C 'vb6_fmtIsDateFormat(fmt, 0)' $PA) +
    ' arms=' + (C ('case L' + $SQ + 'q' + $SQ + ':') $F) + (C ('case L' + $SQ + 'w' + $SQ + ':') $F) + (C ('case L' + $SQ + 't' + $SQ + ':') $F) +
    ' outlets=' + (C 'vb6_WeekOfYear((int32_t)serial, firstDayOfWeek, firstWeekOfYear)' $F) + '/' + (C 'vb6_Weekday(serial, 1)' $F) +
    ' locale=' + (C 'vb6_fmtLocaleShortDate(&st' $BOTH) + '/' + (C 'vb6_fmtLocaleLongTime(&st' $BOTH) +
    ' cells=' + (C ('Test-Vbp ' + $SQ + 'b430_format_date') $RT) + ' needles=' + (C ('Debug.Print ' + $DQ + 'B430-') $FX))
exit 0
