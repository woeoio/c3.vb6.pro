# ============================================================
# check_conv_strict_parse.ps1 - 账 #305 #416 (第九十三刀) 的结构性哨兵
#
# 为什么单开这一道: 「这串文本对一个严格转换是不是一个数」这一问, 过去在 RTL 里有三位
# 主人、两位各自抄了一份解析 (vb6_Val 一份手写的 &H/&O 十进制解析、vb6_NumVal 又一份,
# 十三枚 *V 薄壳直接吃 vb6_Val), 三条路都只答值不答错 => CInt(abc) 静默给 0、
# CInt(123abc) 吞尾巴, 而 VB6 两句都是 13 Type mismatch。发码 A/B 看不见它 (改的是函数体,
# 发出去的文本一字不动), 语料普查也只说得出面有多大 => 只有真跑看得见, 而只有结构哨兵能
# 保证这一问别再退回两遍。行为判据住 tests\b416 + Test-Vbp(两台); 这里钉的是形状。
#
# 规则:
#   R1  解析那一句只许问 oleaut32 一次: vb6rtl_string.c 里 VarR8FromStr / VarBoolFromStr
#       各恰好 1 次, 交回 13 的那句恰好 2 次 (两位主人各一), 而旧那两份手写解析 (wcstod /
#       wcstoul) 一个都不许回来。
#   R2  Variant 那一档改问同一句: vb6rtl_builtin.h 里 vb6_VariantConvNum 恰好 13 次
#       (定义 1 + 十二枚数值 *V 薄壳), vb6_VariantToDouble(v) 只许剩 2 次 (定义里的非文本
#       分支 + CBoolV 的非文本分支), CBoolV 必须走 vb6_BoolVal。
#   R3  十枚 BSTR 入口不许留宽松: vb6rtl_conv.c 里 vb6_NumVal 恰好 10 次, 且任何
#       vb6_*BSTR(BSTR s) 的单行体内都不许再出现 vb6_Val(。
#   R4  发码那侧两位桥分开: cgen_expr_call_conv_numeric.inc 里 convIsBool 恰好 2 次
#       (认档一次 + 选桥一次), 那条三元桥恰好 1 次, 而旧的「所有档共用 vb6_NumVal」那一行
#       必须为 0 (CBool 不能拿数值的答案当自己的答案)。
#   R5  判据两头都在: tests/run_tests.ps1 里 b416_conv_strict / _x86 / 发码形状针 /
#       负控 Absent 各恰好 1 次。
#   R6  宽松那一枚主人没被顺手改严: vb6_Val 的定义与它体内的 strtod 各恰好 1 次。
# ============================================================
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$S = Get-Content -Raw (Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_string.c')
$H = Get-Content -Raw (Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_builtin.h')
$C = Get-Content -Raw (Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_conv.c')
$G = Get-Content -Raw (Join-Path $root 'src\backend\detail\expr\cgen_expr_call_conv_numeric.inc')
$T = Get-Content -Raw (Join-Path $root 'tests\run_tests.ps1')
$fail = @()
function N([string]$pat, [string]$txt) {
    return ([regex]::Matches($txt, $pat, 'Multiline')).Count
}
$DQ = [string][char]34
# R1
$r1r8 = N 'VarR8FromStr\(' $S
$r1bool = N 'VarBoolFromStr\(' $S
$r1raise = N ('vb6_RaiseError\(13, vb6_BSTR_FromStr\(L' + $DQ + 'Type mismatch' + $DQ + '\)\)') $S
$r1numdef = N '^double vb6_NumVal\(BSTR s\) \{' $S
$r1def = N '^int16_t vb6_BoolVal\(BSTR s\) \{' $S
$r1old = (N 'wcstod\(' $S) + (N 'wcstoul\(' $S)
if ($r1r8 -ne 1) { $fail += ('R1 VarR8FromStr = ' + $r1r8 + ' (want 1)') }
if ($r1bool -ne 1) { $fail += ('R1 VarBoolFromStr = ' + $r1bool + ' (want 1)') }
if ($r1raise -ne 2) { $fail += ('R1 raise-13 outlets = ' + $r1raise + ' (want 2: NumVal + BoolVal)') }
if ($r1numdef -ne 1) { $fail += ('R1 vb6_NumVal definition = ' + $r1numdef + ' (want 1)') }
if ($r1def -ne 1) { $fail += ('R1 vb6_BoolVal definition = ' + $r1def + ' (want 1)') }
if ($r1old -ne 0) { $fail += ('R1 retired hand parser is back = ' + $r1old + ' (want 0)') }
# R2
$r2conv = N 'vb6_VariantConvNum\(' $H
$r2dbl = N 'vb6_VariantToDouble\(v\)' $H
$r2decl = N 'int16_t vb6_BoolVal\(BSTR s\);' $H
$r2boolv = N 'return vb6_BoolVal\(v\.bstrVal\);' $H
$r2bare = N 'vb6_C\w+V\(vb6_VARIANT v\) \{ return vb6_C\w+\(vb6_VariantToDouble\(v\)\)' $H
if ($r2conv -ne 13) { $fail += ('R2 VariantConvNum sites = ' + $r2conv + ' (want 13 = def + 12 numeric wrappers)') }
if ($r2dbl -ne 2) { $fail += ('R2 VariantToDouble(v) left in header = ' + $r2dbl + ' (want 2 = the two non-text branches)') }
if ($r2decl -ne 1) { $fail += ('R2 vb6_BoolVal declaration = ' + $r2decl + ' (want 1)') }
if ($r2boolv -ne 1) { $fail += ('R2 CBoolV asks BoolVal = ' + $r2boolv + ' (want 1)') }
if ($r2bare -ne 0) { $fail += ('R2 a *V wrapper still unwraps leniently = ' + $r2bare + ' (want 0)') }
# R3
$r3num = N 'vb6_NumVal\(' $C
$r3lenient = N 'vb6_C\w+BSTR\(BSTR s\)\s*\{[^\n]*vb6_Val\(' $C
if ($r3num -ne 10) { $fail += ('R3 BSTR entries asking the strict question = ' + $r3num + ' (want 10)') }
if ($r3lenient -ne 0) { $fail += ('R3 a BSTR entry is still on vb6_Val = ' + $r3lenient + ' (want 0)') }
# R4
$r4flag = N 'convIsBool' $G
$r4bridge = N ('convIsBool \? ' + $DQ + 'vb6_BoolVal\(' + $DQ + ' : ' + $DQ + 'vb6_NumVal\(' + $DQ) $G
$r4old = N ('= ' + $DQ + 'vb6_NumVal\(' + $DQ + ' \+ args\[0\]') $G
if ($r4flag -ne 2) { $fail += ('R4 convIsBool sites = ' + $r4flag + ' (want 2)') }
if ($r4bridge -ne 1) { $fail += ('R4 bridge split = ' + $r4bridge + ' (want 1)') }
if ($r4old -ne 0) { $fail += ('R4 one-bridge-for-all is back = ' + $r4old + ' (want 0)') }
# R5
$cells = @('b416_conv_strict\b', 'b416_conv_strict_x86\b', 'b416_conv_bridge_split',
           'b416_cbool_no_longer_asks_the_number')
foreach ($cell in $cells) {
    $got = N $cell $T
    if ($got -ne 1) { $fail += ('R5 cell ' + $cell + ' = ' + $got + ' (want 1)') }
}
# R6
$r6def = N '^double vb6_Val\(BSTR s\) \{' $S
$r6strtod = N 'strtod\(' $S
if ($r6def -ne 1) { $fail += ('R6 vb6_Val definition = ' + $r6def + ' (want 1)') }
if ($r6strtod -ne 1) { $fail += ('R6 Val keeps its own lenient parse = ' + $r6strtod + ' (want 1 strtod)') }
if ($fail.Count -gt 0) {
    Write-Host 'FAIL check_conv_strict_parse:' -ForegroundColor Red
    foreach ($f in $fail) { Write-Host ('  - ' + $f) -ForegroundColor Red }
    exit 1
}
Write-Host ('PASS check_conv_strict_parse: ole-outlets=2 raise13=2 hand-parser-retired=1 variant-sites=13 bstr-sites=10 cbool-bridge-split=1 val-still-lenient=1 cells=4')
exit 0
