# check_bitwise_authority.ps1 - 账 #216 的结构性哨兵 (只扫编译器源码, 不起 cl)
#
# 为什么单开这一道: VB6 的 And/Or/Xor/Eqv/Imp 与一元 Not 是**位运算** —— 只有操作数
# 本身是 Boolean 时结果才是 Boolean。这条口径全仓被写过两份 (语义层一份、发码层一份),
# 发码那份恒答 Boolean, 于是位运算数一进字符串上下文就打 True/False, COM 实参被装箱成
# VT_BOOL; 同一表达式先赋给 Long 变量再打印反倒是对的 —— 差的是"问类型"那一步, 不是算数。
# 这种缺陷逐字节 diff 抓不到 (发码本来就合法), 只能在结构上钉。
#
# 口径现在收在一处: src\semantics\type_system.cpp 的
#   TypeSystem::bitwiseResult(a, b) / TypeSystem::logicalNotResult(t)
# 两个消费点: src\semantics\semantic_analyzer_expr.cpp (lastExprType_) 与
#              src\backend\cgen_util_type.cpp  (inferExprType)
#
# 规则 (改坏了会红, 不是装饰 —— 同一套正则在 HEAD 那两份消费点里各命中一次):
#   B1  权威定义各只有一份 (type_system.cpp)
#   B2  声明各只有一份 (type_system.hpp)
#   B3  两层都真的走权威 (每文件每函数 >= 1 个调用点, 合计 >= 4)
#   B4  两份消费点里不许再手写作答: 位运算/Not 之后近距离内直接给 Vb6Type::Boolean = 红
#
# 用法:  pwsh -File scripts\check_bitwise_authority.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$bad = @()

function ReadSrc([string]$rel) {
    $p = Join-Path $root $rel
    if (-not (Test-Path -LiteralPath $p)) { return $null }
    return [System.IO.File]::ReadAllText($p)
}

$ts = ReadSrc 'src\semantics\type_system.cpp'
$th = ReadSrc 'src\semantics\type_system.hpp'
$se = ReadSrc 'src\semantics\semantic_analyzer_expr.cpp'
$be = ReadSrc 'src\backend\cgen_util_type.cpp'
foreach ($p in @('src\semantics\type_system.cpp', 'src\semantics\type_system.hpp',
                 'src\semantics\semantic_analyzer_expr.cpp', 'src\backend\cgen_util_type.cpp')) {
    if ($null -eq (ReadSrc $p)) { $bad += ('B0 missing file: ' + $p) }
}

# B1 + B2: 权威只写一次
if ($ts) {
    foreach ($fn in @('bitwiseResult', 'logicalNotResult')) {
        $n = @([regex]::Matches($ts, 'Vb6Type\s+TypeSystem::' + $fn + '\s*\(')).Count
        if ($n -ne 1) { $bad += ('B1 ' + $fn + ' defined ' + $n + ' times in type_system.cpp (want 1)') }
    }
}
if ($th) {
    foreach ($fn in @('bitwiseResult', 'logicalNotResult')) {
        $n = @([regex]::Matches($th, 'static\s+Vb6Type\s+' + $fn + '\s*\(')).Count
        if ($n -ne 1) { $bad += ('B2 ' + $fn + ' declared ' + $n + ' times in type_system.hpp (want 1)') }
    }
}

# B3: 两层都调权威
$calls = 0
foreach ($pair in @(@('semantics', $se), @('codegen', $be))) {
    $who = $pair[0]; $txt = $pair[1]
    if (-not $txt) { continue }
    foreach ($fn in @('bitwiseResult', 'logicalNotResult')) {
        $n = @([regex]::Matches($txt, 'TypeSystem::' + $fn + '\s*\(')).Count
        $calls += $n
        if ($n -lt 1) { $bad += ('B3 the ' + $who + ' side never calls TypeSystem::' + $fn + ' - the rule is being re-inlined') }
    }
}
if ($calls -lt 4) { $bad += ('B3 only ' + $calls + ' call sites into the authority (want >= 4)') }

# B4: 消费点里不许手写作答
$patBit = '(?:case\s+)?BinaryOp::(?:And|Or|Xor|Eqv|Imp)\b[\s\S]{0,200}?(?:return|lastExprType_ =)\s+Vb6Type::Boolean'
$patNot = 'UnaryOp::Not\b[\s\S]{0,200}?(?:return|lastExprType_ =)\s+Vb6Type::Boolean'
foreach ($pair in @(@('semantics\semantic_analyzer_expr.cpp', $se), @('backend\cgen_util_type.cpp', $be))) {
    $who = $pair[0]; $txt = $pair[1]
    if (-not $txt) { continue }
    foreach ($pr in @(@('And/Or/Xor/Eqv/Imp', $patBit), @('Not', $patNot))) {
        $ms = @([regex]::Matches($txt, $pr[1]))
        if ($ms.Count -gt 0) {
            $snip = ($ms[0].Value -replace '\s+', ' ')
            if ($snip.Length -gt 120) { $snip = $snip.Substring(0, 120) }
            $bad += ('B4 ' + $who + ' answers ' + $pr[0] + ' inline instead of via the authority: ' + $snip)
        }
    }
}

if ($bad.Count -eq 0) {
    Write-Host ('PASS bitwise authority: defs 1+1, callsites ' + $calls + ', inline answers 0') -ForegroundColor Green
    exit 0
}
foreach ($b in $bad) { Write-Host ('FAIL ' + $b) -ForegroundColor Red }
exit 1
