# ============================================================
# check_datepart_arms.ps1 - 账 #305 #422 (第九十七刀) 的结构性哨兵
#
# 为什么单开这一道: vb6_DatePart 里两条臂答错了问 —— case vb6_di_day 直接掉进「一年里第几天」
# 那条臂 (DatePart(「d」, 2026/3/1) 答 60，运行期答 1；答 60 的那一档叫「y」)，「ww」那一臂就地
# 抄了一条 (dayOfYear-1)/7+1 并把 firstDayOfWeek / firstWeekOfYear 用 (void) 丢掉 (与运行期逐日
# 比对 2026 年 365 天里 208 天不同、2024 年 366 天里 52 天不同，而语料那 4 处「ww」全都传了
# vbFirstFourDays ⇒ 不是纸面缺口)。本刀收成一条唯一出口 vb6_WeekOfYear，口径由本机 VB-family
# 运行期 14 年 × 21 组合 = 107394 天逐日拟合定死、零失配 (尺 .build/k321/k324/k330_ww_grid*.vbs，
# 拟合 .build/k323_ww_fit.py，报告 k329/k330_fit_report.txt)。
# 这里钉的是「那一问只许住一处、旧的抄形不许回来、两条形参必须真被读、DateDiff 那条整段除法不许误伤」，
# 外加一格本刀自己踩过的雷: 「一枚年月日折成 VB 序列号」这一问原本在本文件抄了四遍，而 vb6_date_to_serial 答的是儒略日数 (还要减 2415019 才是 VB 序列号) —— 新代码直接把它当序列号用，
# 第一发读数就成了 DatePart(「ww」) 答 -2415012 那一档 (尺 .build/k354_proj，逐值回声 k354_run.txt)。
#
# 规则:
#   R1 date.c 里 vb6_WeekOfYear 定义恰好 1 (第一百刀起**不再是 static**: 第二个读者 = Format 的「ww」)、
#      date.c 内站点恰好 3 (定义 + 年初那一次自调 + DatePart 的读者)、format.c 内读者恰好 1、
#      vb6rtl_builtin.h 里声明恰好 1;
#      周首那一问 vb6_LeadingPartialDays 定义 1、站点 3 (定义 + 本年 + 来年)。
#   R2 旧的抄形不许回来: 标识符 dayOfYear 在 date.c 里归零 (「ww」那臂原来就靠它)。
#   R3 「d」与「y」必须是两问: case vb6_di_day 单独成臂恰好 1、体里就一句 return dd 恰好 1、
#      dayofyear 那条臂恰好 1、两者之间不许再有裸 fallthrough。
#   R4 两条形参必须真被读: DatePart 交出 firstDayOfWeek / firstWeekOfYear 那一问恰好 1 处，
#      且 DatePart 函数体里不许再出现 (void)firstDayOfWeek / (void)firstWeekOfYear。
#   R5 另一问不许被牵连: DateDiff 的「ww」整段除法 `(int64_t)((d2 - d1) / 7)` 仍恰好 1
#      (它问「隔了几段七天」，与「年第几周」不同一条口径，本刀刻意不动、自成一格账)。
#   R6 判据两头都在: run_tests.ps1 里 b422 两台各 1、本哨兵已登记。
#   R7 序列号折算收成唯一出口 vb6_SerialOf: 定义恰好 1、站点恰好 11 (定义 + 10 处读者)，
#      裸儒略日 helper vb6_date_to_serial 的站点恰好 3 (它自己的定义 + 出口体内两枚)，
#      旧的 jd_date/jd_base/jd_this/jd_jan1 那四个手抄临时量各归零。
# ============================================================
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$D = Get-Content -Raw (Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_date.c')
$T = Get-Content -Raw (Join-Path $root 'tests\run_tests.ps1')
$fail = @()
function N([string]$pat, [string]$txt) {
    return ([regex]::Matches($txt, $pat, 'Multiline')).Count
}
# R1
$FMT = Get-Content -Raw (Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_format.c')
$HDR = Get-Content -Raw (Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_builtin.h')
$r1def = N '^int32_t vb6_WeekOfYear\(' $D
$r1static = N '^static int32_t vb6_WeekOfYear\(' $D
$r1sites = N 'vb6_WeekOfYear\(' $D
$r1leaddef = N '^static int32_t vb6_LeadingPartialDays\(' $D
$r1lead = N 'vb6_LeadingPartialDays\(' $D
if ($r1def -ne 1) { $fail += ('R1 vb6_WeekOfYear definition = ' + $r1def + ' (want 1)') }
if ($r1static -ne 0) { $fail += ('R1 vb6_WeekOfYear is still file-static = ' + $r1static + ' (want 0: Format is a second reader)') }
$r1fmt = N 'vb6_WeekOfYear\(\(int32_t\)serial, firstDayOfWeek, firstWeekOfYear\)' $FMT
$r1decl = N 'int32_t vb6_WeekOfYear\(int32_t serial, int32_t fdow, int32_t fwoy\);' $HDR
if ($r1fmt -ne 1) { $fail += ('R1 Format asks the same week outlet = ' + $r1fmt + ' (want 1: a second hand-copied formula is what this outlet exists to stop)') }
if ($r1decl -ne 1) { $fail += ('R1 the week outlet declaration in vb6rtl_builtin.h = ' + $r1decl + ' (want 1)') }
if ($r1sites -ne 3) { $fail += ('R1 vb6_WeekOfYear sites = ' + $r1sites + ' (want 3 = def + its own head call + DatePart)') }
if ($r1leaddef -ne 1) { $fail += ('R1 week-start helper definition = ' + $r1leaddef + ' (want 1)') }
if ($r1lead -ne 3) { $fail += ('R1 week-start helper sites = ' + $r1lead + ' (want 3 = def + this year + next year)') }
# R2
$r2 = N 'dayOfYear' $D
if ($r2 -ne 0) { $fail += ('R2 the in-place week formula is back (dayOfYear) = ' + $r2 + ' (want 0)') }
# R3
$r3day = N 'case vb6_di_day: \{' $D
$r3dd = N 'return dd;' $D
$r3doy = N 'case vb6_di_dayofyear: \{' $D
$r3fall = N 'case vb6_di_day:[\s]*\r?\n[\s]*case vb6_di_dayofyear: \{' $D
if ($r3day -ne 1) { $fail += ('R3 DatePart day arm = ' + $r3day + ' (want 1 = its own arm, not a fallthrough)') }
if ($r3dd -ne 1) { $fail += ('R3 day-of-month answers dd = ' + $r3dd + ' (want 1)') }
if ($r3doy -ne 1) { $fail += ('R3 day-of-year arm = ' + $r3doy + ' (want 1, still a separate question)') }
if ($r3fall -ne 0) { $fail += ('R3 bare day -> day-of-year fallthrough = ' + $r3fall + ' (want 0)') }
# R4
$r4ask = N 'return vb6_WeekOfYear\(d, firstDayOfWeek, firstWeekOfYear\);' $D
$dp = $D.Substring($D.IndexOf('int32_t vb6_DatePart('))
$dp = $dp.Substring(0, $dp.IndexOf('int32_t vb6_Weekday('))
$r4void = N '\(void\)first(DayOfWeek|WeekOfYear);' $dp
if ($r4ask -ne 1) { $fail += ('R4 DatePart asks the outlet with both parameters = ' + $r4ask + ' (want 1)') }
if ($r4void -ne 0) { $fail += ('R4 DatePart still voids one of the parameters = ' + $r4void + ' (want 0)') }
# R5
$r5 = N '\(int64_t\)\(\(d2 - d1\) / 7\)' $D
if ($r5 -ne 1) { $fail += ('R5 DateDiff whole-seven-day-span line = ' + $r5 + ' (want 1: a different question, must not be touched)') }
# R6
$r6cells = N "Test-Vbp 'b422_datepart_arms" $T
$r6sent = N 'check_datepart_arms\.ps1' $T
if ($r6cells -ne 2) { $fail += ('R6 b422 run cells = ' + $r6cells + ' (want 2 = x64 + x86)') }
if ($r6sent -ne 1) { $fail += ('R6 this sentinel is not registered in run_tests.ps1 = ' + $r6sent + ' (want 1)') }
# R8 星期序号只许问 vb6_Weekday (第一百刀: DatePart 的「w」臂原先是第二份 remap 抄写)
$r8ask = N 'return vb6_Weekday\(\(double\)d, 1\);' $D
$r8old = N 'if \(wd == 0\) return 7;' $D
$r8dp  = N '\(d % 7 \+ 7\) % 7' $D
if ($r8ask -ne 1) { $fail += ('R8 DatePart asks vb6_Weekday for the weekday number = ' + $r8ask + ' (want 1)') }
if ($r8old -ne 0) { $fail += ('R8 the hand-copied Saturday remap is back in DatePart = ' + $r8old + ' (want 0)') }
if ($r8dp -ne 1) { $fail += ('R8 copies of the mod-7 line = ' + $r8dp + ' (want 1: only vb6_Weekday itself keeps it)') }
# R7
$r7def = N '^static int32_t vb6_SerialOf\(' $D
$r7sites = N 'vb6_SerialOf\(' $D
$r7raw = N 'vb6_date_to_serial\(' $D
$r7jd = N 'jd_(date|base|this|jan1)' $D
if ($r7def -ne 1) { $fail += ('R7 vb6_SerialOf definition = ' + $r7def + ' (want 1)') }
if ($r7sites -ne 11) { $fail += ('R7 vb6_SerialOf sites = ' + $r7sites + ' (want 11 = def + ten readers)') }
if ($r7raw -ne 3) { $fail += ('R7 raw julian helper sites = ' + $r7raw + ' (want 3 = its own definition + the two inside the outlet)') }
if ($r7jd -ne 0) { $fail += ('R7 the hand-copied julian temporaries are back = ' + $r7jd + ' (want 0)') }

if ($fail.Count -gt 0) {
    foreach ($m in $fail) { Write-Host ('FAIL ' + $m) }
    exit 1
}
Write-Host ('PASS check_datepart_arms: outlet-def=' + $r1def + ' sites=' + $r1sites + ' weekstart-def=' + $r1leaddef +
    ' weekstart-sites=' + $r1lead + ' old-form=' + $r2 + ' day-arm=' + $r3day + ' dd=' + $r3dd +
    ' ask=' + $r4ask + ' void=' + $r4void + ' datediff=' + $r5 + ' cells=' + $r6cells + ' serial=' + $r7def + '/' + $r7sites +
 ' raw=' + $r7raw + ' jd=' + $r7jd + ' fmt=' + $r1fmt + ' decl=' + $r1decl + ' wday=' + $r8ask + '/' + $r8old + '/' + $r8dp)
exit 0
