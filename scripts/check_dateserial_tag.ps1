# ============================================================
# check_dateserial_tag.ps1 -- 账 #305 / §B205 / 第一百零二刀 (2026-10-17)
#
# 「DateSerial 交回来的那枚值进槽时贴哪枚标签」这一问只有一个主人: src/semantics/builtin/
# builtin_funcs.inc 那张内置函数表里它那一行的返回档。改前那一格写 Vb6Type::Double，而同表里
# Now / Date / Time / DateAdd / DateValue / TimeSerial / TimeValue / FileDateTime 全写
# Vb6Type::Date —— RTL 那九枚的 C 返回**都是 double**，所以那一个格子就是全部答案。
# 实测六档与引擎相反 (尺 .build/k607_run_x64.txt / k607_x86_run.txt 两台同读 vs 同机 oracle
# .build/k607_oracle_out.txt):
#     VarType 5 / TypeName "Double" / IsDate False / IsNumeric True  / CStr [46027]   / Len 5
#     引擎   7 / "Date"             / True         / False            / [2026/1/5]     / 8
#
# 第二格是顺带量出来的**旧洞**: 发码那侧 narrowCheckAssign (账 #248 的取整闸) 原来只把
# Single / Double / Currency 送去 vb6_FltToLng，换档之后的 Date 掉进「装得下, 不套」那一条，
# 把 double 裸递进 vb6_ChkLong(int64_t) => C 在调用边界上截断。这一格在改档之前就欠着
# (Now / DateAdd 交的本就是 Date)，判据 B432-17 拿一枚本来就答 Date 的值把它钉成 BASE 红针。
#
# 规则: R1 表里那九枚日期构造函数的返回档同归一处、DateSerial 不许回数值档、扫描行数要有地板;
#       R2 取整闸吃 Date 且那个函数体里 vb6_FltToLng 只一处;
#       R3 Format 的两条门票档号 (0 / 1) 不许并成一档，且那两份文件里不许再出现 VT_R8 这个说法
#          (兜底臂是留给真数值的，理由已经换了 —— 臂不撤);
#       R4 装箱出口与内置名字表各只一处 (改名或另抄一张表就会红);
#       R5 判据两头都在 (两片 cell、整行形状、期望表与针面计数、本哨兵自己已登记)。
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

$sep = [IO.Path]::DirectorySeparatorChar
$files = [ordered]@{
    tab    = 'src/semantics/builtin/builtin_funcs.inc'
    narrow = 'src/backend/cgen_util_type.cpp'
    parse  = 'src/rtl/core/vb6rtl/detail/vb6rtl_format_parse.inc'
    extr   = 'src/rtl/core/vb6rtl/detail/vb6rtl_format_extract.inc'
    box    = 'src/backend/cgen_util_ctrl.cpp'
    names  = 'src/backend/detail/expr/cgen_expr_ident_builtin_table.inc'
    run    = 'tests/run_tests.ps1'
    fx     = 'tests/b432/B432Main.bas'
}
$T = @{}
foreach ($k in $files.Keys) {
    $p = Join-Path $RepoRoot ($files[$k].replace('/', $sep))
    if (-not (Test-Path $p)) { Write-Output ('FAIL: missing ' + $p); exit 1 }
    $T[$k] = [System.IO.File]::ReadAllText($p).Replace("`r`n", "`n")
}
$TAB = $T['tab']; $NR = $T['narrow']; $PA = $T['parse']; $EX = $T['extr']
$BX = $T['box']; $NM = $T['names']; $RT = $T['run']; $FX = $T['fx']

# ---- R1 那张表是返回档的唯一主人，九枚日期构造同归一处 ----
$ctors = @('Now','Date','Time','DateAdd','DateValue','TimeSerial','TimeValue','FileDateTime','DateSerial')
$asDate = 0
foreach ($ctorName in $ctors) {
    $re = 'addBuiltinFunc(?:WithParams)?\(' + $DQ + $ctorName + $DQ + ',\s*Vb6Type::(\w+)'
    $mt = [regex]::Matches($TAB, $re)
    if ($mt.Count -ne 1) { $fail += ('R1 row count for ' + $ctorName + ' = ' + $mt.Count + ' (want 1): 那张表里这一枚必须恰好登记一次'); continue }
    $tag = $mt[0].Groups[1].Value
    if ($tag -ne 'Date') { $fail += ('R1 return tag of ' + $ctorName + ' = ' + $tag + ' (want Date): 日期构造函数的返回档只许问这张表、九枚同答 Date') }
    else { $asDate++ }
}
Want 'R1 nine constructors tagged Date' $asDate 9 '九枚日期构造函数(含 DateSerial)的返回档必须同为 Vb6Type::Date'
Want 'R1 DateSerial back to Double' (C 'addBuiltinFuncWithParams("DateSerial", Vb6Type::Double' $TAB) 0 'DateSerial 不许回数值档 (那六档就是它交的)'
$rows = [regex]::Matches($TAB, 'addBuiltinFunc').Count
if ($rows -lt 150) { $fail += ('R1 table floor: rows=' + $rows + ' (want >=150) —— 扫描瞎了却报干净') }

# ---- R2 取整闸吃 Date，且那个函数体里 vb6_FltToLng 只一处 ----
$bi = $NR.IndexOf('std::string CCodeGen::narrowCheckAssign(')
if ($bi -lt 0) {
    $fail += 'R2 cannot locate narrowCheckAssign (那道闸的定义被搬走或改名了)'
    $BODY = ''
} else {
    $rest = $NR.Substring($bi)
    $e = $rest.IndexOf($NL + '}')
    if ($e -lt 0) { $fail += 'R2 cannot find the end of narrowCheckAssign'; $BODY = '' }
    else { $BODY = $rest.Substring(0, $e) }
}
Want 'R2 FltToLng inside the gate' (C 'vb6_FltToLng(' $BODY) 1 '取整那一站在 narrowCheckAssign 里只许一处'
Want 'R2 gate eats Date' (C 'Vb6Type::Date' $BODY) 1 'Date 必须在那条取整闸的名单里 (换档之后裸递 int64_t 就是截断)'
Want 'R2 gate keeps Currency' (C 'Vb6Type::Currency' $BODY) 1 '同名单里的 Currency 不许被顺手删掉'
Want 'R2 gate asks the width table' (C 'narrowCheckFnForVbType(tt, &tgtBits)' $BODY) 1 '闸名与位宽仍问 §B184 那一张表，不许在这里另抄 switch'
Want 'R2 file keeps one rounding call' (C 'vb6_FltToLng(' $NR) 1 '整个文件里取整出口只此一处 (另开一站 = 第二个主人)'

# ---- R3 Format 的两条门票档号 + 那两份文件不许再写 VT_R8 ----
Want 'R3 numeric ticket' (C 'vb6_fmtIsDateFormat(fmt, 0)' $PA) 1 '数值臂那条门票仍是 0 档 (不许扩到 w/q/t)'
Want 'R3 date ticket' (C 'vb6_fmtIsDateFormat(fmt, 1)' $EX) 1 'VT_DATE 臂那条门票仍是 1 档'
Want 'R3 stale tag claim (parse)' (C 'VT_R8' $PA) 0 '那条兜底臂的理由已换 (留给真数值)，不许再用 DateSerial 交 R8 作论据'
Want 'R3 stale tag claim (extract)' (C 'VT_R8' $EX) 0 '同上: 日期那一臂里不该出现 R8 的说法'

# ---- R4 装箱出口与名字表各一处 ----
Want 'R4 boxing arm' (C 'return "vb6_VariantDate(" + cExpr + ")";' $BX) 1 'boxToVariant 的 Date 档只此一处 (改档靠的是它，不是别站的特判)'
Want 'R4 name map' (C '{"dateserial", "vb6_DateSerial"},' $NM) 1 '内置名字表那一行不许改名或另抄一张'

# ---- R5 判据两头都在 ----
Want 'R5 cells' (C ('Test-Vbp ' + $SQ + 'b432_dateserial_tag') $RT) 2 '两台各一片 (x64 + x86)'
$x86rows = [regex]::Matches($RT, '(?m)-Arch "x86"[ \t]*$').Count
if ($x86rows -lt 1) { $fail += 'R5 x86 cell row shape: 找不到以 -Arch "x86" 收尾的那一行' }
Want 'R5 stray escaped arch' (C '-Arch "x86\' $RT) 0 '那一行结尾多一个反斜杠会把 C3 的参数打成 "x86\" (第九十八刀真栽过)'
$tstart = $RT.IndexOf('$b432Expected = @(')
if ($tstart -lt 0) { $fail += 'R5 expectation table missing ($b432Expected)' ; $trows = 0 }
else {
    $tend = $RT.IndexOf(')', $tstart)
    $tbl = $RT.Substring($tstart, $tend - $tstart)
    $trows = [regex]::Matches($tbl, '"B432-').Count
}
Want 'R5 table rows' $trows 19 '期望表 18 针 + 一枚 DONE'
Want 'R5 fixture prints' (C ('Debug.Print ' + $DQ + 'B432-') $FX) 19 '夹具里 18 针 + DONE 证人'
Want 'R5 sentinel registered' (C 'check_dateserial_tag.ps1' $RT) 1 '本哨兵在 run_tests.ps1 里被调用一次'
Want 'R5 sentinel function' (C 'function Test-DateSerialTag' $RT) 1 '那道静态检查自己有一处定义'
# 上一行只认「定义」, 假改动 P12 拆掉那条**调用**时它照旧交 1 (定义还在) ⇒ 那道检查根本没进流水线。
Want 'R5 sentinel called' ([regex]::Matches($RT, '(?m)^    Test-DateSerialTag[ \t]*\r?$').Count) 1 '那条裸调用自己有一行 (没有它, 定义只是死码)'

if ($fail.Count -gt 0) {
    foreach ($m in $fail) { Write-Output ('FAIL ' + $m) }
    Write-Output ('FAIL check_dateserial_tag: ' + $fail.Count + ' rule(s) broken')
    exit 1
}
Write-Output ('PASS check_dateserial_tag: rows=' + $rows + ' date_tag=' + $asDate +
    ' gate_flt=' + (C 'vb6_FltToLng(' $BODY) + '/' + (C 'vb6_FltToLng(' $NR) +
    ' tickets=' + (C 'vb6_fmtIsDateFormat(fmt, 0)' $PA) + '/' + (C 'vb6_fmtIsDateFormat(fmt, 1)' $EX) +
    ' r8=' + (C 'VT_R8' $PA) + '/' + (C 'VT_R8' $EX) +
    ' cells=' + (C ('Test-Vbp ' + $SQ + 'b432_dateserial_tag') $RT) + ' needles=' + (C ('Debug.Print ' + $DQ + 'B432-') $FX))
exit 0
