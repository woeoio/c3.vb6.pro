# ============================================================
# check_format_ampm.ps1 - 账 #305 #426 (第九十八刀) 的结构性哨兵
#
# 为什么单开这一道: vb6_fmtDateSerial 的日期面有三格答错, 而三格是同一个问题的三个侧面 ——
# 「这枚格式串里有没有上午/下午记号」这一问以前没有人答。① h/hh 恒按 12 小时报
# (那条算式是 st.wHour % 12 无条件走), 而真语料把「hh:mm:ss」「yyyy-mm-dd hh:mm:ss」当**时间戳**
# 用 (1693 处日期类格式串里各占 70 / 101, 尺 .build/k392_fmt_census.py) => 一整个下午被写成半个
# 下午; VB 的口径是 h/hh 走 24 小时, 只有串里出现 AM/PM 那一族才切 12 小时。②「AM/PM」令牌被
# 逐字符循环当成月份吃掉 => Format(d,「AM/PM」) 交 A1/P1 (那两枚 1 / 7 就是两个月号: 两枚探针
# 时间分别在 1 月与 7 月, 尺 .build/k391_run.txt)。③ 只写记号、不带任何日期字母的串 (「A/P」
# 「AMPM」) 过不了 vb6_fmtIsDateFormat => 掉回数值臂交裸序列号 46027.6046875。
# 收成一处三枚: vb6_fmtAmPmLenAt (在哪 / 几长 / 哪一形) + vb6_fmtHasAmPm (有没有, 同时是「小时
# 走 12 还是 24」的唯一判据) + vb6_fmtAppendAmPm (大小写跟随串里写的那一种)。
# 这里钉的是「那一问只许住一处、记号必须赶在令牌 switch 之前被认掉、两臂的邻居不许被牵连」。
#
# 规则:
#   R1 令牌表恰好 1 张; 三枚出口的站点数 = 3 / 3 / 2 (定义 + 各自的读者), 多一处就是有人又抄了一遍。
#   R2 小时那一问只答一次: 认记号那一行 1、use12 只出现 2 次 (声明 + 用)、旧的无条件 12 小时
#      形状 (if (h == 0) h = 12;) 归零。
#   R3 日期分支的门票必须问记号: vb6_fmtIsDateFormat 里 if (vb6_fmtHasAmPm(fmt)) return 1; 恰好 1
#      (③ 那一格就是这么补上的, 摘掉它「A/P」又会交序列号)。
#   R4 记号必须赶在令牌 switch 之前被认掉: 在 vb6_fmtDateSerial 函数体内, ampmLen 那一读的位次
#      早于 switch (lc); 函数体里跳过记号的 i += 那一行 1、交记号那一行 1 (② 的靶子就是「来不及认」)。
#   R5 大小写跟随串里写的那一种: towlower(fmt[i]) == fmt[i] 恰好 1。
#   R6 邻居不许被牵连: 月份档 case L'm': 与日档 case L'd': 各 1 (按 vb6_fmtDateSerial 体内数 ——
#      门票函数 vb6_fmtIsDateFormat 里也各写着一次同名 case)、prevHour = 1; (hh:mm 的 mm 是
#      分钟那一条规则) 1、月份名单 kMonLong[12] 1。
#   R7 判据两头都在: run_tests.ps1 里 b426 两台各 1、Test-FormatAmPm 被调用 1、本哨兵已登记;
#      两片的名字各钉成整行形状 —— x86 那一片必须以 -Arch "x86" 收尾, 且全文件不许出现
#      -Arch "x86\" (门 #700 实测: 多一个反斜杠就把实参传成 x86\, C3 报「--arch must be x86 or
#      x64」, 那一格就此 FAIL (compile) —— 用例名对、期望表对, 坏的是传参的形状)。
# ============================================================
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$F = Get-Content -Raw (Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_format.c')
$T = Get-Content -Raw (Join-Path $root 'tests\run_tests.ps1')
$fail = @()
function N([string]$pat, [string]$txt) {
    return ([regex]::Matches($txt, $pat, 'Multiline')).Count
}
# R1
$r1tbl = N 'static const wchar_t\* const kTokens\[\] = \{ L"AM/PM", L"AMPM", L"A/P" \};' $F
$r1len = N 'vb6_fmtAmPmLenAt\(' $F
$r1has = N 'vb6_fmtHasAmPm\(' $F
$r1app = N 'vb6_fmtAppendAmPm\(' $F
if ($r1tbl -ne 1) { $fail += ('R1 designator token table = ' + $r1tbl + ' (want 1: one owner for the shape question)') }
if ($r1len -ne 3) { $fail += ('R1 vb6_fmtAmPmLenAt sites = ' + $r1len + ' (want 3 = def + HasAmPm + the loop)') }
if ($r1has -ne 3) { $fail += ('R1 vb6_fmtHasAmPm sites = ' + $r1has + ' (want 3 = def + IsDateFormat + use12)') }
if ($r1app -ne 2) { $fail += ('R1 vb6_fmtAppendAmPm sites = ' + $r1app + ' (want 2 = def + its one reader)') }
# R2
$r2ask = N 'int use12 = vb6_fmtHasAmPm\(fmt\);' $F
$r2use = N 'use12' $F
$r2h12 = N 'int h = use12 \? \(\(st\.wHour % 12\) == 0 \? 12 : st\.wHour % 12\) : st\.wHour;' $F
$r2old = N 'if \(h == 0\) h = 12;' $F
if ($r2ask -ne 1) { $fail += ('R2 the 12/24 question is asked once = ' + $r2ask + ' (want 1)') }
if ($r2use -ne 2) { $fail += ('R2 use12 occurrences = ' + $r2use + ' (want 2 = declare + answer the hour)') }
if ($r2h12 -ne 1) { $fail += ('R2 hour shape asks use12 = ' + $r2h12 + ' (want 1)') }
if ($r2old -ne 0) { $fail += ('R2 the unconditional 12-hour shape is back = ' + $r2old + ' (want 0)') }
# R3
$r3gate = N 'if \(vb6_fmtHasAmPm\(fmt\)\) return 1;' $F
if ($r3gate -ne 1) { $fail += ('R3 the date branch admits a designator-only pattern = ' + $r3gate + ' (want 1: without it 「A/P」 prints the raw serial again)') }
# R4
$start = $F.IndexOf('static BSTR vb6_fmtDateSerial(')
$stop = $F.IndexOf('BSTR vb6_Format(', $start)
if ($start -lt 0 -or $stop -lt 0) {
    $fail += 'R4 cannot locate the vb6_fmtDateSerial body'
    $body = ''
} else {
    $body = $F.Substring($start, $stop - $start)
}
if ($body) {
    $r4read = N 'int ampmLen = vb6_fmtAmPmLenAt\(fmt, i, n\);' $body
    $r4skip = N 'i \+= \(ampmLen < 0\) \? 3 : ampmLen;' $body
    $r4emit = N 'vb6_fmtAppendAmPm\(out, 256, &o, fmt, i, ampmLen, st\.wHour < 12\);' $body
    $r4switch = $body.IndexOf('switch (lc) {')
    $r4order = $body.IndexOf('int ampmLen = vb6_fmtAmPmLenAt(')
    if ($r4read -ne 1) { $fail += ('R4 the loop reads the designator = ' + $r4read + ' (want 1)') }
    if ($r4skip -ne 1) { $fail += ('R4 the loop skips the whole token = ' + $r4skip + ' (want 1)') }
    if ($r4emit -ne 1) { $fail += ('R4 the loop emits the designator = ' + $r4emit + ' (want 1)') }
    if ($r4order -lt 0 -or $r4switch -lt 0 -or $r4order -gt $r4switch) {
        $fail += 'R4 the designator is not recognized BEFORE the token switch (that is how 「AM/PM」 got eaten as a month)'
    }
}
# R5
$r5case = N 'towlower\(fmt\[i\]\) == fmt\[i\]' $F
if ($r5case -ne 1) { $fail += ('R5 the case of the written token is echoed = ' + $r5case + ' (want 1)') }
# R6 邻居不许被牵连 (m/d 两档按 vb6_fmtDateSerial 体内数 —— 门票函数里也各写着一次同名 case)
if ($body) {
    $r6m = N "case L'm':" $body
    $r6d = N "case L'd':" $body
    if ($r6m -ne 1) { $fail += ('R6 month arm inside the date formatter = ' + $r6m + ' (want 1: the designator branch must not replace it)') }
    if ($r6d -ne 1) { $fail += ('R6 day arm inside the date formatter = ' + $r6d + ' (want 1)') }
}
$r6prev = N 'prevHour = 1;' $F
$r6tab = N 'static const wchar_t\* kMonLong\[12\]' $F
if ($r6prev -ne 1) { $fail += ('R6 the hh:mm minute rule = ' + $r6prev + ' (want 1)') }
if ($r6tab -ne 1) { $fail += ('R6 month-name table = ' + $r6tab + ' (want 1: the invariant-English caliber stays)') }
# R7
$r7cells = N "Test-Vbp 'b426_format_ampm" $T
$r7x86shape = N 'Test-Vbp ''b426_format_ampm_x86'' [^\r\n]*-Arch "x86"\r?$' $T
$r7badarch = N '-Arch "x86\\"' $T
$r7call = N '^\s+Test-FormatAmPm\s*$' $T
$r7sent = N 'check_format_ampm\.ps1' $T
if ($r7cells -ne 2) { $fail += ('R7 b426 run cells = ' + $r7cells + ' (want 2 = x64 + x86)') }
if ($r7call -ne 1) { $fail += ('R7 Test-FormatAmPm call site = ' + $r7call + ' (want 1: a defined-but-never-called cell is a silent non-running cell)') }
if ($r7sent -ne 1) { $fail += ('R7 this sentinel is not registered in run_tests.ps1 = ' + $r7sent + ' (want 1)') }
if ($r7x86shape -ne 1) { $fail += ('R7 the x86 cell asks -Arch "x86" exactly = ' + $r7x86shape + ' (want 1)') }
if ($r7badarch -ne 0) { $fail += ('R7 an arch literal carries a stray backslash = ' + $r7badarch + ' (want 0: CI read it as --arch x86 backslash and the cell died in compile)') }

if ($fail.Count -gt 0) {
    foreach ($m in $fail) { Write-Host ('FAIL ' + $m) }
    exit 1
}
Write-Host ('PASS check_format_ampm: table=' + $r1tbl + ' len=' + $r1len + ' has=' + $r1has + ' app=' + $r1app +
    ' use12=' + $r2use + ' gate=' + $r3gate + ' order-ok x86shape=' + $r7x86shape + ' cells=' + $r7cells + ' call=' + $r7call)
exit 0
