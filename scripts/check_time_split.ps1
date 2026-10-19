# ============================================================
# check_time_split.ps1 - 账 #305 #419 (第九十六刀) 的结构性哨兵
#
# 为什么单开这一道: 「这一枚 DATE 的双半边怎么切」这一问在 RTL 里原本有六位各自抄写的主人
# (vb6_Hour / vb6_Minute / vb6_Second 三枚 getter 各一份截断，DateAdd 分解那一站又抄三份)，
# 全都 `(int32_t)(小数 * 24/1440/86400)` —— 而 DATE 的小数半边在二进制里落不整
# (14:30 = 0.6041666666666666)，截断就把这一格往下拽一秒。实测 (本机 VB-family 运行期，尺
# .build/k272_timeparts.vbs / k277_round.vbs / k278_roll.vbs) 运行期是**四舍五入到整秒**:
# Minute(CDate("2026/1/1 14:30")) 答 30 (本仓答 29)、Second(CDate("2026/1/1 23:59:59")) 答 59
# (本仓答 58)、+0.6 秒进位而 +0.4/+0.5 不进。一天 86400 秒按截断算有 43136 个落错一档。
# DateAdd 更疼: 拆完还要写回去，少的那一秒永久留在存下的值里。
# 本刀收成一条唯一出口 vb6_SecondsOfDay。所以这里钉的是「四舍五入那一问只许住一处，
# 旧的截断形不许回来，DateDiff 那三条差值取整不许被误伤」。
#
# 规则:
#   R1  date.c 里四舍五入那一问恰好 1 处 (`+ 0.5` 那一句)，且住在 vb6_SecondsOfDay 里
#       (它的定义恰好 1、站点恰好 5 = 定义 + 三枚 getter + DateAdd 分解站)。
#   R2  旧的截断形整体不许回来: `(int32_t)(t * 24.0)` / `(int32_t)(t * 1440.0)` /
#       `(int32_t)(t * 86400.0)` 各 0 (getter 那三枚)、`(int32_t)(timePart * ` 0 (DateAdd 那三行)。
#   R3  三枚 getter 各定义 1 次，且各自体里就是一行 `return vb6_SecondsOfDay(...) …` (合计 3)。
#   R4  另一问不许被牵连: DateDiff 那三条差值取整 `(int64_t)((date2 - date1) * …)` 仍恰好 3
#       (差值问的是「隔了几个完整区间」，VB6 就是朝零取整，与拆分那一问不同一条口径)。
#   R5  判据两头都在: run_tests.ps1 里 b419 两台各 1、本哨兵已登记。
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
$r1round = N '\+ 0\.5\);' $D
$r1def = N '^static int32_t vb6_SecondsOfDay\(double date\) \{' $D
$r1sites = N 'vb6_SecondsOfDay\(' $D
if ($r1round -ne 1) { $fail += ('R1 rounding site (+ 0.5) = ' + $r1round + ' (want 1: the split asks one outlet)') }
if ($r1def -ne 1) { $fail += ('R1 vb6_SecondsOfDay definition = ' + $r1def + ' (want 1)') }
if ($r1sites -ne 5) { $fail += ('R1 vb6_SecondsOfDay sites = ' + $r1sites + ' (want 5 = def + three getters + DateAdd)') }
# R2
$r2t24 = N '\(int32_t\)\(t \* 24\.0\)' $D
$r2t1440 = N '\(int32_t\)\(t \* 1440\.0\)' $D
$r2t86400 = N '\(int32_t\)\(t \* 86400\.0\)(?! \+ 0\.5)' $D
$r2tp = N '\(int32_t\)\(timePart \* ' $D
if ($r2t24 + $r2t1440 + $r2t86400 -ne 0) {
    $fail += ('R2 the old truncating getter bodies are back = ' + ($r2t24 + $r2t1440 + $r2t86400) + ' (want 0)')
}
if ($r2tp -ne 0) { $fail += ('R2 DateAdd re-copied the truncation = ' + $r2tp + ' (want 0)') }
# R3
$r3h = N '^int32_t vb6_Hour\(double time\) \{' $D
$r3m = N '^int32_t vb6_Minute\(double time\) \{' $D
$r3s = N '^int32_t vb6_Second\(double time\) \{' $D
$r3ret = N 'return \(?vb6_SecondsOfDay\(' $D
if (($r3h + $r3m + $r3s) -ne 3) { $fail += ('R3 getter definitions = ' + ($r3h + $r3m + $r3s) + ' (want 3)') }
if ($r3ret -ne 3) { $fail += ('R3 getters that ask the outlet = ' + $r3ret + ' (want 3)') }
# R4
$r4diff = N '\(int64_t\)\(\(date2 - date1\) \* ' $D
if ($r4diff -ne 3) { $fail += ('R4 DateDiff whole-interval lines = ' + $r4diff + ' (want 3: a different question, must not be touched)') }
# R5
$r5cells = N "Test-Vbp 'b419_time_split" $T
$r5sent = N 'check_time_split\.ps1' $T
if ($r5cells -ne 2) { $fail += ('R5 b419 run cells = ' + $r5cells + ' (want 2 = x64 + x86)') }
if ($r5sent -ne 1) { $fail += ('R5 this sentinel is not registered in run_tests.ps1 = ' + $r5sent + ' (want 1)') }

if ($fail.Count -gt 0) {
    foreach ($m in $fail) { Write-Host ('FAIL ' + $m) }
    exit 1
}
Write-Host ('PASS check_time_split: round=' + $r1round + ' outlet-def=' + $r1def + ' sites=' + $r1sites +
    ' old-trunc=' + ($r2t24 + $r2t1440 + $r2t86400 + $r2tp) + ' getters=' + $r3ret + ' datediff=' + $r4diff +
    ' cells=' + $r5cells)
exit 0
