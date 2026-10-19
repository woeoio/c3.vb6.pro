# ============================================================
# check_general_date_outlets.ps1 —— 账 #305 / §B203 / 第一百零一刀 (2026-10-17)
#
# 「general date 这一档到底给不给时间段」这一问，以前在 RTL 里有三位主人、三个答案:
#   vb6rtl_system.c 的 vb6_FormatDateTime (named=0)   —— 交 "2026/1/5 " (len 9, 一枚孤儿尾随空格)
#   vb6rtl_format.c 的 named "General Date" 那一臂    —— 交 "2026/1/5 0:00:00" (len 16)
#   vb6rtl_conv.c 的 vb6_CStrDate (Fix 175 手抄的容差) —— 交 "2026/1/5" (len 8, 只有这一条与引擎一字不差)
# 病机制: ① system.c 那句 buf[len - 1] = 空格 是**无条件**做的, 而 GetDateFormatW 交回的长度含结束符
#          ⇒ 那一句把结束符改成了空格; 只有拼时间那一段才判零头。
#        ④ 「locale 的短日期 / 长日期 / 长时间」这条问句在三个文件里各抄一份 Get*FormatW。
# 修法: 四问收成 vb6rtl_format.c 的四条出口 (不再是 file-static, 声明在 vb6rtl_builtin.h)，
#       「给不给时间段」只住 vb6_fmtLocaleGeneralDate 一条; 其余三家全部改问它。
#
# 引擎口径 (本刀逐点量出来的，尺 .build/k559_out.txt + .build/k564_out.txt，两台 oleaut32 同读):
#   零头四入到整秒为 0 就不给时间段 —— 1e-11 / 1e-10 / 1e-9 / 1e-8 日 与 50/200/400/499/500 毫秒一律只给日期，
#   600 毫秒起才给。所以那条闸既不是「零头 != 0」也不是任何 1e-9 级别的 epsilon。
#
# 判据面: tests/b431 十三针 + 一枚 DONE 证人 (两台 .build/k569_run_x64.txt / _x86.txt 逐字节相同)。
# 本文件里不写反斜杠字面量: 路径用 [IO.Path]::DirectorySeparatorChar，字面计数走 [regex]::Escape。
# ============================================================

param(
    [string]$RepoRoot = (Split-Path -Parent $PSScriptRoot)
)

$ErrorActionPreference = 'Stop'
$fail = @()
$SQ = [char]39
$DQ = [char]34

function C([string]$lit, [string]$txt) { return ([regex]::Matches($txt, [regex]::Escape($lit))).Count }
function Want([string]$rule, [int]$got, [int]$expect, [string]$why) {
    if ($got -ne $expect) { $script:fail += ($rule + ' = ' + $got + ' (want ' + $expect + '): ' + $why) }
}

$sep = [IO.Path]::DirectorySeparatorChar
$files = [ordered]@{
    fmt    = 'src/rtl/core/vb6rtl/vb6rtl_format.c'
    extr   = 'src/rtl/core/vb6rtl/detail/vb6rtl_format_extract.inc'
    sys    = 'src/rtl/core/vb6rtl/vb6rtl_system.c'
    conv   = 'src/rtl/core/vb6rtl/vb6rtl_conv.c'
    hdr    = 'src/rtl/core/vb6rtl/vb6rtl_builtin.h'
    run    = 'tests/run_tests.ps1'
    fx     = 'tests/b431/B431Main.bas'
}
$T = @{}
foreach ($k in $files.Keys) {
    $p = Join-Path $RepoRoot ($files[$k].replace('/', $sep))
    if (-not (Test-Path $p)) { Write-Output ('FAIL: missing ' + $p); exit 1 }
    $T[$k] = [System.IO.File]::ReadAllText($p)
}
$FM = $T['fmt']; $EX = $T['extr']; $SY = $T['sys']; $CV = $T['conv']; $HD = $T['hdr']; $RT = $T['run']; $FX = $T['fx']

$genCall = 'vb6_fmtLocaleGeneralDate('
$defShape = 'void vb6_fmtLocaleGeneralDate(const SYSTEMTIME* st, wchar_t* out, size_t cap) {'

# ---- R1 那条闸只住一条出口, 而且住的是量出来的那一个数 ----
Want 'R1 outlet defined' (C $defShape $FM) 1 'general date 那条出口恰好定义一处'
Want 'R1 rounded-zero gate' (C 'if (st->wHour == 0 && st->wMinute == 0 && st->wSecond == 0) return;' $FM) 1 '闸 = 四入之后当日时分秒全为零 (尺 k559_out.txt 二十行 + k564_out.txt 十九行, 两发合起来)'
Want 'R1 no epsilon in fmt' (C 'frac <' $FM) 0 '那条出口不再吃序列号本体: 任何浮点容差的形状都不许留'
Want 'R1 serial gone from outlet' (C 'vb6_fmtLocaleGeneralDate(double' ($FM + $EX + $SY + $HD)) 0 '这一问只要四入后的那枚 SYSTEMTIME, 形参不许再多一条'
Want 'R1 no epsilon in conv' (C 'frac < 1e-9' $CV) 0 'conv.c 里那条容差的代码形状不许留着 (它已交给出口)'
Want 'R1 no bare zero test' (C 'fracPart != 0.0' $SY) 0 'system.c 不许再自己判「零头是否为 0」'
Want 'R1 gate not duplicated' ((C $genCall $SY) + (C $genCall $EX) + (C $genCall $CV)) 2 'system.c 与 extract.inc 各问一次; conv.c 一律不许自己拼 (它只转调 vb6_Format)'

# ---- R2 四条 locale 出口: 一处定义、有声明、不再是 file-static ----
foreach ($n in @('ShortDate', 'LongDate', 'LongTime', 'GeneralDate')) {
    Want ('R2 definition ' + $n) (C ('void vb6_fmtLocale' + $n) $FM) 1 ('vb6_fmtLocale' + $n + ' 只住 format.c')
    Want ('R2 declared ' + $n) (C ('void vb6_fmtLocale' + $n) $HD) 1 ('跨文件读者要有声明 (第九十九刀那条地板)')
}
Want 'R2 outlets not static' (C 'static void vb6_fmtLocale' $FM) 0 '四条出口全部不再是 file-static (它们有跨文件的读者)'

# ---- R3 system.c 的四档全部改问出口, 手抄的 Get*FormatW 撤干净 ----
Want 'R3 general asks outlet' (C ('vb6_fmtLocaleGeneralDate(&st, buf, 256)') $SY) 1 'vbGeneralDate 那一档问出口'
Want 'R3 short asks outlet' (C ('vb6_fmtLocaleShortDate(&st, buf, 256)') $SY) 1 'vbShortDate 那一档问出口'
Want 'R3 long asks outlet' (C ('vb6_fmtLocaleLongDate(&st, buf, 256)') $SY) 1 'vbLongDate 那一档问出口'
Want 'R3 longtime asks outlet' (C ('vb6_fmtLocaleLongTime(&st, buf, 256)') $SY) 1 'vbLongTime 那一档问出口'
Want 'R3 no hand GetDateFormat' (C 'GetDateFormatW(LOCALE_USER_DEFAULT' $SY) 0 'system.c 不许再自己抄一份短/长日期 (只许留注释里那句回忆)'
Want 'R3 short-time keeps flag' (C 'GetTimeFormatW(LOCALE_USER_DEFAULT, TIME_NOSECONDS' $SY) 1 '短时间那一档 (TIME_NOSECONDS) 是另一问, 仍由 system.c 自己答'
Want 'R3 orphan space gone' (C ('len - 1] = L' + $SQ + ' ') $SY) 0 '那枚把结束符改成空格的孤儿尾随空格不许回来'

# ---- R4 conv.c 只剩转调, 不再自己判零头 ----
Want 'R4 conv delegates' (C 'v.vt = (vb6_vartype)VT_DATE; v.dblVal = x;' $CV) 1 'vb6_CStrDate 那一支只许装箱后转调 vb6_Format 的 !fmt 臂'
Want 'R4 conv no hand date fmt' (C 'GetDateFormatW(LOCALE_USER_DEFAULT' $CV) 0 'conv.c 不许再手抄一份短日期'
Want 'R4 conv no own frac test' (C 'frac' $CV) 0 'conv.c 里那条判零头整族撤干净 (这文件里连 frac 这个名字都不该再有)'
Want 'R4 extract asks outlet' (C $genCall $EX) 1 '无格式串那条兜底臂问出口, 不许再无条件拼「短 + 空格 + 长」'

# ---- R5 判据两头都在 (定义了没被调用 = 静默不跑那一格) ----
Want 'R5 b430 cells' (C ('Test-Vbp ' + $SQ + 'b431_general_date') $RT) 2 '两台各一行 (x64 + x86)'
Want 'R5 x86 cell shape' (C ('Test-Vbp ' + $SQ + 'b431_general_date_x86' + $SQ) $RT) 1 'x86 那一行恰好一处'
Want 'R5 stray arch backslash' (C ('-Arch ' + $DQ + 'x86' + $DQ + [char]92 + [char]92) $RT) 0 '不许有带反斜杠收尾的 -Arch'
Want 'R5 expectation table' (C '$b431Expected = @(' $RT) 1 '期望表恰好一张'
Want 'R5 expectation rows' (C ('"B431-') $RT) 16 '期望表里十五针加一枚证人, 一行不许少 (删掉一行 = 那一针静默不查)'
Want 'R5 fixture needles' (C ('Debug.Print ' + $DQ + 'B431-') $FX) 16 '十五针加一枚 DONE 证人都在'
Want 'R5 fixture witness' (C ('Debug.Print ' + $DQ + 'B431-DONE' + $DQ) $FX) 1 '跑到底的那枚证人'
Want 'R5 sentinel function' (C 'function Test-GeneralDateOutlets ' $RT) 1 '本哨兵的函数恰好定义一处'
Want 'R5 sentinel called' (C '    Test-GeneralDateOutlets' $RT) 1 '定义了必须被调用'
Want 'R5 self registered' (C 'check_general_date_outlets' $RT) 1 '本哨兵得写进 runner'

if ($fail.Count -gt 0) {
    foreach ($m in $fail) { Write-Output ('FAIL ' + $m) }
    Write-Output ('FAIL check_general_date_outlets: ' + $fail.Count + ' rule(s) broken')
    exit 1
}
Write-Output ('PASS check_general_date_outlets: gate=' + (C 'if (st->wHour == 0 && st->wMinute == 0 && st->wSecond == 0) return;' $FM) +
    ' outlets=' + (C $genCall $FM) + '/' + (C $genCall $EX) + '/' + (C $genCall $SY) +
    ' hand=' + (C 'GetDateFormatW(LOCALE_USER_DEFAULT' $SY) + '/' + (C 'GetDateFormatW(LOCALE_USER_DEFAULT' $CV) +
    ' delegates=' + (C 'return vb6_Format(v, NULL, 1, 1);' $CV) +
    ' cells=' + (C ('Test-Vbp ' + $SQ + 'b431_general_date') $RT) + ' needles=' + (C ('Debug.Print ' + $DQ + 'B431-') $FX))
exit 0
