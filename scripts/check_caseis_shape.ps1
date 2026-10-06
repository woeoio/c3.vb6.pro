# check_caseis_shape.ps1 - 账 #217 第一刀的结构性哨兵 (只扫编译器源码, 不起 cl)
#
# 为什么单开这一道: `Case Is > 2` 里的 Is 站在测试表达式的位置, 不是标识符。改前
# parseCaseValue 为了借一次优先级解析, 把 IdentifierExpr("Is") 留在 AST 里交给语义层,
# 于是每条 `Case Is` 要么换一条 VB3001 (两份真工程 36 条, 全仓发码语料 138 条),
# 要么在没写 Option Explicit 的模块里换一枚没人引用的 vb6_VARIANT 局部 (探针实测)。
# 发码那侧从来只读比较符与右操作数, 值一直是对的 —— 逐字节 diff 抓不到, 只能钉结构。
#
# 口径现在收在 CaseClause::CaseValue 的两个字段上 (src\ast\detail\ast_stmt.hpp):
# relOp 装比较符, value 只装右操作数; 那枚占位标识符出了 parseCaseValue 就不存在。
#
# 规则 (改坏了会红, 不是装饰):
#   C1  CaseValue 里 relOp / hasRelOp 各声明 1 份
#   C2  parser 只设一次 hasRelOp, 且不得再把造出来的标识符存进 cv.value
#   C3  发码读 cv.relOp (三档: 字符串 / Single / 整数), 不得再把 cv.value 当 BinaryExpr 拆
#   C4  ast_clone 必须两个字段都抄 (漏一个 = 克隆过的 Case 静默退化成 `<> 0` 那一支)
#
# 用法:  pwsh -File scripts\check_caseis_shape.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$bad = @()

function ReadSrc([string]$rel) {
    $p = Join-Path $root $rel
    if (-not (Test-Path -LiteralPath $p)) { return $null }
    return [System.IO.File]::ReadAllText($p)
}

$astRel = 'src\ast\detail\ast_stmt.hpp'
$parserRel = 'src\parser\parser_expr_postfix.cpp'
$selectRel = 'src\backend\stmt\cgen_select.cpp'
$cloneRel = 'src\ast\ast_clone.cpp'

$ast = ReadSrc $astRel
$parser = ReadSrc $parserRel
$select = ReadSrc $selectRel
$clone = ReadSrc $cloneRel
foreach ($p in @($astRel, $parserRel, $selectRel, $cloneRel)) {
    if ($null -eq (ReadSrc $p)) { $bad += ('C0 missing file: ' + $p) }
}
if ($bad.Count -gt 0) {
    foreach ($b in $bad) { Write-Host ('FAIL ' + $b) -ForegroundColor Red }
    exit 1
}

# C1
$relOpDecl = @([regex]::Matches($ast, 'BinaryOp\s+relOp')).Count
$hasRelDecl = @([regex]::Matches($ast, 'bool\s+hasRelOp')).Count
if ($relOpDecl -ne 1) { $bad += ('C1 relOp declared ' + $relOpDecl + ' times in ast_stmt.hpp (want 1)') }
if ($hasRelDecl -ne 1) { $bad += ('C1 hasRelOp declared ' + $hasRelDecl + ' times in ast_stmt.hpp (want 1)') }

# C2: parseCaseValue 的函数体 (到下一个 Parser:: 为止)
$start = $parser.IndexOf('Parser::parseCaseValue')
if ($start -lt 0) {
    $bad += 'C2 Parser::parseCaseValue not found'
    $body = ''
} else {
    $next = $parser.IndexOf('Parser::', $start + 8)
    if ($next -lt 0) { $body = $parser.Substring($start) } else { $body = $parser.Substring($start, $next - $start) }
}
$sets = @([regex]::Matches($body, 'cv\.hasRelOp\s*=\s*true')).Count
if ($sets -ne 1) { $bad += ('C2 parseCaseValue sets hasRelOp ' + $sets + ' times (want 1)') }
$storedFake = @([regex]::Matches($body, 'cv\.value\s*=\s*std::make_unique<IdentifierExpr>')).Count
if ($storedFake -ne 0) { $bad += ('C2 parseCaseValue stores a fabricated identifier into cv.value ' + $storedFake + ' times (want 0) - that name is what VB3001 reports') }

# C3
$reads = @([regex]::Matches($select, 'mapBinaryOp\(cv\.relOp\)')).Count
if ($reads -ne 3) { $bad += ('C3 cgen_select reads cv.relOp ' + $reads + ' times (want 3: string/float/integer)') }
$recast = @([regex]::Matches($select, 'static_cast<BinaryExpr&>\(\*cv\.value\)')).Count
if ($recast -ne 0) { $bad += ('C3 cgen_select still unpacks cv.value as a BinaryExpr ' + $recast + ' times (want 0)') }

# C4
$cloneRel2 = @([regex]::Matches($clone, 'cv\.relOp\s*=\s*v\.relOp')).Count
$cloneHas = @([regex]::Matches($clone, 'cv\.hasRelOp\s*=\s*v\.hasRelOp')).Count
if ($cloneRel2 -ne 1) { $bad += ('C4 ast_clone copies relOp ' + $cloneRel2 + ' times (want 1)') }
if ($cloneHas -ne 1) { $bad += ('C4 ast_clone copies hasRelOp ' + $cloneHas + ' times (want 1)') }

if ($bad.Count -gt 0) {
    foreach ($b in $bad) { Write-Host ('FAIL ' + $b) -ForegroundColor Red }
    exit 1
}
Write-Host ('PASS case-is shape: fields ' + $relOpDecl + '+' + $hasRelDecl + ', parser sets ' + $sets +
    ', fabricated stores ' + $storedFake + ', cgen reads ' + $reads + ', clone copies ' + ($cloneRel2 + $cloneHas))
exit 0
