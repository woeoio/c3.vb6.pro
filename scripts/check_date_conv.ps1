# ============================================================
# check_date_conv.ps1 - 账 #305 #417 (第九十四刀) 的结构性哨兵
#
# 为什么单开这一道: 「这串文本是哪一天」这一问在 RTL 里有三位各自抄写的主人 (vb6_DateValue、
# vb6_TimeValue、vb6rtl_conv.c 里那枚 CStr(日期形文本) 入口)，三处都自己调 VarDateFromStr、
# 三处都在解析不出时静默返 0；而 vb6_CDate 收文本那一支连日期解析都没问，转调的是**数值**解析
# (vb6_Val) —— 注释里却写着「字符串按日期解析」。实测 (本机 VB-family 运行期, 尺
# .build\k233\k234\k235\k236) CDate(2026/1/1) 答 2026 (那一年) 而不是那一天、CDate(abc) 答 0
# 且 Err 还是 0，而同一份程序里 DateValue(2026/1/1) 是对的。VB6 两句都是 13。
# 本刀把解析收成一条站点、把「认不出算几号」交给两位主人 (CDate 那一档还认纯数字串，
# DateValue/TimeValue 那一档不认 —— 这是实测出来的两档，不是随手兜底)。
# 发码 A/B 看不见这一族 (改的全在函数体，发出去的 C 文本一字不动 —— 本刀 BASE/NEW 两份
# --emit-c 逐字节相同)，语料普查也只说得出面有多大 (80 处文本形状的 CDate) => 行为判据住
# tests\b417 + Test-Vbp(两台)，「不许退回两遍」由这里钉形状。
#
# 规则:
#   R1  date.c 里解析那一句只许问 oleaut32 一次: VarDateFromStr 恰好 1 次 (住在唯一那条
#       vb6_TryDateFromStr 里，站点合计 4 = 定义 + 三位主人各一次转调 (DateParse / DateVal /
#       DateCanParse，第三枚是第九十五刀的 #418 给 IsDate 开的口))，交回 13 的那句恰好
#       1 次 (在 vb6_DateParse 里)，两位转换主人的定义各 1 次 (第三枚主人是判据，形状由
#       check_ispred_text 钉)，旧的「(int32_t) 砍成整天」不许回来。
#   R2  全仓 RTL 再不许有第二条日期解析路: src\rtl 整棵树里 VarDateFromStr 合计恰好 1；
#       conv.c 里日期那一族不许再出现 vb6_Val(v.bstrVal) 或 vb6_CStrDate(vb6_Val(...)。
#   R3  CDate 的文本分支问的是 vb6_DateVal；CStr(日期形文本) 入口问的是 vb6_DateParse。
#   R4  两位投影各归各的半边: DateValue = floor(解析)，TimeValue = 小数 (d - floor(d))。
#   R5  两位主人在 vb6rtl_builtin.h 各有声明 (发码那侧之外的人也得问得到它们)。
#   R6  发码那张表一字没动: cdate / datevalue / timevalue 三行各自还在 (本刀不动发码)。
#   R7  判据两头都在: run_tests.ps1 里 b417 两台各 1、本哨兵已登记。
# ============================================================
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$D = Get-Content -Raw (Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_date.c')
$C = Get-Content -Raw (Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_conv.c')
$H = Get-Content -Raw (Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_builtin.h')
$G = Get-Content -Raw (Join-Path $root 'src\backend\detail\expr\cgen_expr_ident_builtin_table.inc')
$T = Get-Content -Raw (Join-Path $root 'tests\run_tests.ps1')
# 整棵 RTL 树的日期解析站点计数 (R2 的面): 递归读每个 .c/.h, 只数 VarDateFromStr
$rtlFiles = Get-ChildItem -Path (Join-Path $root 'src\rtl') -Recurse -Include *.c,*.h -File
$r2sites = 0
foreach ($f in $rtlFiles) {
    $r2sites += ([regex]::Matches((Get-Content -Raw $f.FullName), 'VarDateFromStr\(')).Count
}
$fail = @()
function N([string]$pat, [string]$txt) {
    return ([regex]::Matches($txt, $pat, 'Multiline')).Count
}
# R1
$r1parse = N 'VarDateFromStr\(' $D
$r1try = N 'vb6_TryDateFromStr\(' $D
$r1raise = N 'vb6_RaiseError\(13, vb6_BSTR_FromStr\(L"Type mismatch"\)\)' $D
$r1dp = N '^double vb6_DateParse\(BSTR s\) \{' $D
$r1dv = N '^double vb6_DateVal\(BSTR s\) \{' $D
$r1old = N '\(double\)\(int32_t\)result' $D
if ($r1parse -ne 1) { $fail += ('R1 VarDateFromStr in date.c = ' + $r1parse + ' (want 1)') }
if ($r1try -ne 4) { $fail += ('R1 vb6_TryDateFromStr sites = ' + $r1try + ' (want 4 = def + the three owners)') }
if ($r1raise -ne 1) { $fail += ('R1 raise-13 outlet in date.c = ' + $r1raise + ' (want 1, inside vb6_DateParse)') }
if ($r1dp -ne 1) { $fail += ('R1 vb6_DateParse definition = ' + $r1dp + ' (want 1)') }
if ($r1dv -ne 1) { $fail += ('R1 vb6_DateVal definition = ' + $r1dv + ' (want 1)') }
if ($r1old -ne 0) { $fail += ('R1 the retired (int32_t) truncation is back = ' + $r1old + ' (want 0)') }
# R2
if ($r2sites -ne 1) { $fail += ('R2 date-parse sites across src\rtl = ' + $r2sites + ' (want 1: one authority only)') }
$r2lenient = N 'return vb6_Val\(v\.bstrVal\);' $C
$r2csd = N 'vb6_CStrDate\(vb6_Val\(' $C
if ($r2lenient -ne 0) { $fail += ('R2 CDate text branch still asks vb6_Val = ' + $r2lenient + ' (want 0)') }
if ($r2csd -ne 0) { $fail += ('R2 CStr(date-text) fallback still asks vb6_Val = ' + $r2csd + ' (want 0)') }
# R3
$r3branch = N 'return vb6_DateVal\(v\.bstrVal\);' $C
$r3entry = N 'return vb6_CStrDate\(vb6_DateParse\(s\)\);' $C
if ($r3branch -ne 1) { $fail += ('R3 vb6_CDate vtBSTR -> vb6_DateVal = ' + $r3branch + ' (want 1)') }
if ($r3entry -ne 1) { $fail += ('R3 vb6_CStrDateFromBSTR -> vb6_DateParse = ' + $r3entry + ' (want 1)') }
# R4
$r4dv = N 'return floor\(vb6_DateParse\(dateStr\)\);' $D
$r4tv = N 'double d = vb6_DateParse\(timeStr\);' $D
$r4frac = N 'return d - floor\(d\);' $D
if ($r4dv -ne 1) { $fail += ('R4 DateValue projection = ' + $r4dv + ' (want 1)') }
if ($r4tv -ne 1) { $fail += ('R4 TimeValue parses through the authority = ' + $r4tv + ' (want 1)') }
if ($r4frac -ne 1) { $fail += ('R4 TimeValue fraction projection = ' + $r4frac + ' (want 1)') }
# R5
$r5dp = N 'double vb6_DateParse\(BSTR s\);' $H
$r5dv = N 'double vb6_DateVal\(BSTR s\);' $H
if ($r5dp -ne 1) { $fail += ('R5 vb6_DateParse declaration = ' + $r5dp + ' (want 1)') }
if ($r5dv -ne 1) { $fail += ('R5 vb6_DateVal declaration = ' + $r5dv + ' (want 1)') }
# R6
$r6cd = N '\{"cdate",\s*"vb6_CDate"\}' $G
$r6dvl = N '\{"datevalue",\s*"vb6_DateValue"\}' $G
$r6tv = N '\{"timevalue",\s*"vb6_TimeValue"\}' $G
if (($r6cd + $r6dvl + $r6tv) -ne 3) {
    $fail += ('R6 codegen table rows cdate/datevalue/timevalue = ' + ($r6cd + $r6dvl + $r6tv) + ' (want 3, this knife moves no codegen)')
}
# R7
$r7cells = N 'b417_date_conv' $T
$r7sent = N 'check_date_conv\.ps1' $T
if ($r7cells -ne 2) { $fail += ('R7 b417 run cells = ' + $r7cells + ' (want 2 = x64 + x86)') }
if ($r7sent -ne 1) { $fail += ('R7 this sentinel is not registered in run_tests.ps1 = ' + $r7sent + ' (want 1)') }

if ($fail.Count -gt 0) {
    foreach ($m in $fail) { Write-Host ('FAIL ' + $m) }
    exit 1
}
Write-Host ('PASS check_date_conv: parse-sites=' + $r1parse + ' rtl-total=' + $r2sites +
    ' raise13=' + $r1raise + ' owners=2 projections=3 cgen-rows=3 cells=' + $r7cells)
exit 0
