# ============================================================
# check_ispred_text.ps1 - 账 #305 #418 (第九十五刀) 的结构性哨兵
#
# 为什么单开这一道: `IsDate` / `IsNumeric` 这两问以前只看 Variant 的 vt 标签，文本一律答否 ——
# 而 VB6 的这两枚判据吃的正是「这串折不折得出来」(实测 IsNumeric("123") 真、IsDate("2026/1/1")
# 真; 尺 .build/k253_isfam.vbs + k262_ispins.vbs)。第九十三/九十四刀已经把「认不认得出」与
# 「认不出报几号」拆成两层，本刀把判据接到「认不认得出」那一层上 —— 而这正是最容易退回
# 去的情形: 有人图省事会让 IsNumeric 直接调 vb6_NumVal(那会**报错**)，或者干脆再抄一份
# wcstod/Var*FromStr (那就长出第二三位主人)。所以这里钉的是「只许问成功、不许报错、
# 站点数不许长」。
#
# 规则:
#   R1  数值那一位: string.c 里 VarR8FromStr 恰好 1 次 (住在新出口 vb6_NumTryParse 里)、
#       vb6_NumTryParse 站点恰好 3 (定义 + vb6_NumVal + vb6_IsNumericText)、两枚主人各定义
#       1 次、而交回 13 的那句仍恰好 2 次 (只有 NumVal / BoolVal 两位主人有权报)。
#   R2  日期那一位: date.c 里 vb6_DateCanParse 定义 1 次、vb6_TryDateFromStr 站点恰好 4
#       (定义 + DateParse / DateVal / DateCanParse 三位主人)、交回 13 仍恰好 1 次。
#   R3  两处消费都改问新的那一层: conv.c 里 IsNumeric 的 vtBSTR 臂 1 处、IsDate 的文本臂 1 处，
#       而旧的 `IsDate(v) { return v.vt == vb6_vtDate ? -1 : 0; }` 那一行不许回来。
#   R4  头文件里三枚新声明各 1 处 (别人要问也问得到，不必再抄解析)。
#   R5  判据两头都在: run_tests.ps1 里 b418 两台各 1、本哨兵已登记。
# ============================================================
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$S = Get-Content -Raw (Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_string.c')
$D = Get-Content -Raw (Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_date.c')
$C = Get-Content -Raw (Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_conv.c')
$H = Get-Content -Raw (Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_builtin.h')
$T = Get-Content -Raw (Join-Path $root 'tests\run_tests.ps1')
$fail = @()
function N([string]$pat, [string]$txt) {
    return ([regex]::Matches($txt, $pat, 'Multiline')).Count
}
$DQ = [string][char]34
$RAISE = 'vb6_RaiseError\(13, vb6_BSTR_FromStr\(L' + $DQ + 'Type mismatch' + $DQ + '\)\)'
# R1
$r1site = N 'VarR8FromStr\(' $S
$r1try = N 'vb6_NumTryParse\(' $S
$r1def = N '^int32_t vb6_NumTryParse\(BSTR s, double\* out\) \{' $S
$r1txt = N '^int32_t vb6_IsNumericText\(BSTR s\) \{' $S
$r1num = N '^double vb6_NumVal\(BSTR s\) \{' $S
$r1raise = N $RAISE $S
if ($r1site -ne 1) { $fail += ('R1 VarR8FromStr in string.c = ' + $r1site + ' (want 1: the parse stays one site)') }
if ($r1try -ne 3) { $fail += ('R1 vb6_NumTryParse sites = ' + $r1try + ' (want 3 = def + NumVal + IsNumericText)') }
if ($r1def -ne 1) { $fail += ('R1 vb6_NumTryParse definition = ' + $r1def + ' (want 1)') }
if ($r1txt -ne 1) { $fail += ('R1 vb6_IsNumericText definition = ' + $r1txt + ' (want 1)') }
if ($r1num -ne 1) { $fail += ('R1 vb6_NumVal definition = ' + $r1num + ' (want 1)') }
if ($r1raise -ne 2) { $fail += ('R1 raise-13 outlets in string.c = ' + $r1raise + ' (want 2: only the two conversion owners raise)') }
# R2
$r2can = N '^int32_t vb6_DateCanParse\(BSTR s\) \{' $D
$r2try = N 'vb6_TryDateFromStr\(' $D
$r2raise = N $RAISE $D
if ($r2can -ne 1) { $fail += ('R2 vb6_DateCanParse definition = ' + $r2can + ' (want 1)') }
if ($r2try -ne 4) { $fail += ('R2 vb6_TryDateFromStr sites = ' + $r2try + ' (want 4 = def + three owners)') }
if ($r2raise -ne 1) { $fail += ('R2 raise-13 outlets in date.c = ' + $r2raise + ' (want 1: the predicate must not raise)') }
# R3
$r3num = N 'return vb6_IsNumericText\(v\.bstrVal\);' $C
$r3date = N 'return vb6_DateCanParse\(v\.bstrVal\);' $C
$r3old = N '^int32_t vb6_IsDate\(vb6_VARIANT v\) \{ return v\.vt == vb6_vtDate \? -1 : 0; \}' $C
if ($r3num -ne 1) { $fail += ('R3 IsNumeric text arm = ' + $r3num + ' (want 1)') }
if ($r3date -ne 1) { $fail += ('R3 IsDate text arm = ' + $r3date + ' (want 1)') }
if ($r3old -ne 0) { $fail += ('R3 the text-blind one-line IsDate is back = ' + $r3old + ' (want 0)') }
# R4
$r4a = N 'int32_t vb6_NumTryParse\(BSTR s, double\* out\);' $H
$r4b = N 'int32_t vb6_IsNumericText\(BSTR s\);' $H
$r4c = N 'int32_t vb6_DateCanParse\(BSTR s\);' $H
if (($r4a + $r4b + $r4c) -ne 3) {
    $fail += ('R4 header declarations of the three new outlets = ' + ($r4a + $r4b + $r4c) + ' (want 3)')
}
# R5
$r5cells = N 'b418_ispred_text' $T
$r5sent = N 'check_ispred_text\.ps1' $T
if ($r5cells -ne 2) { $fail += ('R5 b418 run cells = ' + $r5cells + ' (want 2 = x64 + x86)') }
if ($r5sent -ne 1) { $fail += ('R5 this sentinel is not registered in run_tests.ps1 = ' + $r5sent + ' (want 1)') }

if ($fail.Count -gt 0) {
    foreach ($m in $fail) { Write-Host ('FAIL ' + $m) }
    exit 1
}
Write-Host ('PASS check_ispred_text: num-site=' + $r1site + ' tryparse=3 text-outlets=' + ($r1txt + $r2can) +
    ' conv-arms=' + ($r3num + $r3date) + ' raise-outlets=' + ($r1raise + $r2raise) + ' decls=3 cells=' + $r5cells)
exit 0
