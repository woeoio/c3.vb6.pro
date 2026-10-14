# check_implicit_cstr_station.ps1 —— 「数值实参交进 VB 的字符串形参 ⇒ 按 VB6 的隐式 CStr 折」这一条
# 规则的站点 / 名单 / 形状对账（账 #305 §B174，第六十五刀）
#
# 为什么要有这一道：这条规则在发码侧原本住**两家**——
#   Fix 175 (src/backend/detail/expr/cgen_expr_call_arg_emit.inc) 问**形参表** (ByVal String)，
#          形状走 wrapToBSTR；
#   Fix 108c (src/backend/detail/expr/cgen_expr_call_builtin_argtype.inc) 靠一张 13 枚名单，
#          自己写了一遍折算 if/else，而且做在 argList **已拼成一整串之后** —— 找回第 N 个实参
#          只能靠"从上次命中往后数子串"这种对齐。
# 两家的档位清单从不汇合（一家有 Byte / LongPtr，另一家有 SByte / UInteger / ULong / ULongLong /
# Decimal），Boolean 那一档还各自给出两种答案（wrapToBSTR 的 "True" 与名单自己写的 "-1"）。
# 第六十五刀收成一位，并补上这一位以前管不着的那一族（**故意没有形参表**的 InStr：位号答在
# builtinNoParamTableStringSlot 一处）。这一道钉的就是"别再长出第二位、别再把形状就地重写"。
#
# 判据（任何一条不齐就红）：
#   R1 站点唯一：typeTakesImplicitCStr 的**调用点**恰好 1 处，且住在 cgen_expr_call_arg_emit.inc
#      （声明 cgen_helpers.inc + 定义 cgen_util_type.cpp 各 1 处，是同一枚问句的两半，不算调用）。
#   R2 位号表唯一：builtinNoParamTableStringSlot 同一条，也必须恰好一处调用、同一文件。
#   R3 旧名单不许回来：strArgPos108c 在 src/ 里 0 处。
#   R4 那一站不再发形状：cgen_expr_call_builtin_argtype.inc 的**代码**里不许出现 vb6_CStr。
#   R5 名单自己的 cast 形状不许在别处重现：vb6_CStrLong((int32_t)( 在 src/backend 里 0 处。
#   R6 census 不许读空：档位闸的 case >= 14、位号表认到的内在函数名 >= 1、arg_emit 里
#      wrapToBSTR 的调用 >= 1（形状出口还在被问）。
#
# 用法: powershell -File scripts\check_implicit_cstr_station.ps1 [-RootOverride <目录>]  (PASS = exit 0)
#   -RootOverride 是给负控用的: 把 src\backend 抄到别处再判, 真树一字不动。

param([string]$RootOverride = '')
$ErrorActionPreference = 'Stop'
$Root = if ($RootOverride) { $RootOverride } else { Split-Path -Parent $PSScriptRoot }
$backend = Join-Path $Root 'src\backend'
$utilCpp = Join-Path $backend 'cgen_util_type.cpp'
$helpers = Join-Path $backend 'detail\util\cgen_helpers.inc'
$emitInc = Join-Path $backend 'detail\expr\cgen_expr_call_arg_emit.inc'
$argtypeInc = Join-Path $backend 'detail\expr\cgen_expr_call_builtin_argtype.inc'
foreach ($p in @($utilCpp, $helpers, $emitInc, $argtypeInc)) {
    if (-not (Test-Path -LiteralPath $p)) { Write-Host ('FAIL: 读不到 ' + $p); exit 1 }
}

function Strip-CComments([string]$t) {
    $t = [regex]::Replace($t, '/\*[\s\S]*?\*/', ' ')
    $t = [regex]::Replace($t, '(?m)//.*$', ' ')
    return $t
}

# 取一段花括号平衡的块体 (start 指向块首 '{' 的位置)
function Get-BlockBody([string]$t, [int]$start) {
    $depth = 0
    for ($j = $start; $j -lt $t.Length; $j++) {
        $ch = $t[$j]
        if ($ch -eq '{') { $depth++ }
        elseif ($ch -eq '}') {
            $depth--
            if ($depth -eq 0) { return $t.Substring($start + 1, $j - $start - 1) }
        }
    }
    return $null
}

$allFiles = @(Get-ChildItem -Path $backend -Recurse -File |
              Where-Object { $_.Extension -eq '.cpp' -or $_.Extension -eq '.inc' })
$codeByFile = @{}
foreach ($f in $allFiles) { $codeByFile[$f.FullName] = Strip-CComments ([IO.File]::ReadAllText($f.FullName)) }
$util = $codeByFile[$utilCpp]
$emit = $codeByFile[$emitInc]
$argtype = $codeByFile[$argtypeInc]
$allCode = ($codeByFile.Values -join [Environment]::NewLine)

$srcRaw = ''
foreach ($f in @(Get-ChildItem -Path (Join-Path $Root 'src') -Recurse -File |
                 Where-Object { $_.Extension -eq '.cpp' -or $_.Extension -eq '.inc' -or $_.Extension -eq '.h' })) {
    $srcRaw += [IO.File]::ReadAllText($f.FullName) + [Environment]::NewLine
}

# 调用点 = 名字后面紧跟 '(' ，且前面不是"定义 (CCodeGen::)"也不是"声明 (bool / std::string)"
function Get-CallSites([string]$name) {
    $rx = [regex]([regex]::Escape($name) + '\(')
    $hits = @()
    foreach ($path in $codeByFile.Keys) {
        $t = $codeByFile[$path]
        foreach ($m in $rx.Matches($t)) {
            # $m.Index 指的是**名字**那一处, 所以 $pre 以名字前的那个词收尾:
            # 定义 = "…bool CCodeGen::" / 声明 = "…bool " —— 两者都不算调用点。
            $pre = $t.Substring(0, $m.Index)
            if ($pre -match 'CCodeGen::\s*$') { continue }
            if ($pre -match '(bool|std::string)\s+$') { continue }
            $hits += (Split-Path -Leaf $path)
        }
    }
    return $hits
}

$viol = @()

# ---- R1 / R2 ----
foreach ($pair in @(@('typeTakesImplicitCStr', 'cgen_expr_call_arg_emit.inc'),
                    @('builtinNoParamTableStringSlot', 'cgen_expr_call_arg_emit.inc'))) {
    $nm = $pair[0]
    $want = $pair[1]
    $sites = @(Get-CallSites $nm)
    if ($sites.Count -ne 1) {
        $viol += ('R1/R2 ' + $nm + ' 的调用点数应为 1, 实为 ' + $sites.Count +
                  ' (' + (($sites | Sort-Object -Unique) -join ', ') +
                  ') —— 这条规则又长出第二位主人了')
    } elseif ($sites[0] -ne $want) {
        $viol += ('R1/R2 ' + $nm + ' 的唯一调用点应在 ' + $want + ', 现在在 ' + $sites[0])
    }
}

# ---- R3 ----
if ($srcRaw.Contains('strArgPos108c')) {
    $viol += 'R3 strArgPos108c 又出现了 (那张 13 枚名单已并入 Fix 175 + 位号表, 见 §B174)'
}

# ---- R4 ----
if ([regex]::IsMatch($argtype, 'vb6_CStr')) {
    $viol += ('R4 cgen_expr_call_builtin_argtype.inc 的代码里出现 vb6_CStr* —— 数值折串的形状' +
              '只许出自 wrapToBSTR 一处 (这一站现在只该管裸 COM 结果解包 / ObjPtr 取值)')
}

# ---- R5 ----
$castN = @([regex]::Matches($allCode, [regex]::Escape('vb6_CStrLong((int32_t)('))).Count
if ($castN -ne 0) {
    $viol += ('R5 vb6_CStrLong((int32_t)( 在 src/backend 出现 ' + $castN +
              ' 处, 应为 0 —— 那是旧名单自己的形状, 与 wrapToBSTR 不同口径')
}

# ---- R6 ----
$gateCases = -1
$gateDef = [regex]::Match($util, 'bool CCodeGen::typeTakesImplicitCStr[^\{]*\{')
if ($gateDef.Success) {
    $gateBody = Get-BlockBody $util ($gateDef.Index + $gateDef.Value.Length - 1)
    if ($gateBody) { $gateCases = @([regex]::Matches($gateBody, 'case Vb6Type::')).Count }
}
if ($gateCases -lt 14) {
    $viol += ('R6 档位闸读出 ' + $gateCases + ' 档 (< 14) —— 两位旧主人的清单并集是 14 档, ' +
              '少一档就是 Left/Replace 那类实参重新裸交给 BSTR')
}
$tableNames = @([regex]::Matches($util, 'callee (==|!=) "vb6_[A-Za-z]+"')).Count
if ($tableNames -lt 1) {
    $viol += 'R6 位号表里一枚内在函数名都没有 —— InStr 那一族会退回裸交 (x64 真跑 AV)'
}
$wrapN = @([regex]::Matches($emit, 'wrapToBSTR\(')).Count
if ($wrapN -lt 1) {
    $viol += 'R6 cgen_expr_call_arg_emit.inc 里不再问 wrapToBSTR —— 折算形状那一站断了'
}

Write-Host ('census: backendFiles=' + $allFiles.Count +
            ' gateCases=' + $gateCases +
            ' slotTableNames=' + $tableNames +
            ' wrapToBSTR@emit=' + $wrapN +
            ' castShapes=' + $castN)

if ($viol.Count -gt 0) {
    foreach ($v in $viol) { Write-Host ('  ' + $v) }
    Write-Host ('FAIL implicit_cstr_station: ' + $viol.Count +
                ' 条违规 —— 这条规则的站点在 cgen_expr_call_arg_emit.inc (Fix 175): ' +
                '位号问形参表 (缺表时问 builtinNoParamTableStringSlot), 档位问 typeTakesImplicitCStr, ' +
                '形状走 wrapToBSTR; 要加第二位主人请先说明这四个问为什么分不开')
    exit 1
}
Write-Host 'PASS implicit_cstr_station (R1..R6)'
exit 0
