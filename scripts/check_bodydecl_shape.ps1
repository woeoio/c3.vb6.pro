# check_bodydecl_shape.ps1 - 账 #215 的结构性哨兵 (只扫编译器源码, 不起 cl)
#
# 为什么单开这一道: 体级声明以前有四条路、两种形状 —— parseDimStmt 自己手写一遍逗号展开
# (那份副本还漏了 parseVariableDecl 里的 WithEvents 与「后缀即类型」两步, 于是
#  `Dim a&, b&` 的第二枚落回 Variant), 而 Const / Static / 体级 Public 三条把
# MultiDecl 原样塞进 LocalDeclStmt, 语义层 visit(LocalDeclStmt) 的 switch 没有 MultiDecl
# 分支 ⇒ 一枚名字都不登记, 每条使用报一条 VB3001 (VBFlexGridDemo 一片就 276 条)。
# 发码那侧一直把每枚都发出来, 所以「值对、诊断错、类型偶尔也错」—— 逐字节 diff 与
# TypeName 读数都抓不到 (装箱后的 Long 打出来还是 Long), 只能在结构上钉「一种形状」。
#
# 口径现在收在一处: Parser::wrapBodyDecls (src\parser\stmt\parser_stmt_assign.cpp)
# —— 体级声明一律「一条声明符一条 LocalDeclStmt」。
#
# 规则 (改坏了会红, 不是装饰):
#   P1  wrapBodyDecls 定义 1 份、声明 1 份
#   P2  四条体级路全部走它 (调用点 = 4, 且 Dim/Const/Static/Access 四个函数各占一处)
#   P3  Dim 那份手写展开不许回来: parser_stmt_assign.cpp 里不得再有 expectName("expected variable name")
#       (共享的声明符解析在 parser_decl_var.cpp, 那边必须还有 >= 1 处)
#   P4  除了 wrapBodyDecls 与 Static Sub/Function 那两处, 别处不许把声明列表直接包成 LocalDeclStmt
#   P5  语义层不许再补一条 MultiDecl 分支 (那等于把「两种形状」再造一遍) —— 计数必须为 0
#
# 用法:  pwsh -File scripts\check_bodydecl_shape.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$bad = @()

function ReadSrc([string]$rel) {
    $p = Join-Path $root $rel
    if (-not (Test-Path -LiteralPath $p)) { return $null }
    return [System.IO.File]::ReadAllText($p)
}

$assignRel = 'src\parser\stmt\parser_stmt_assign.cpp'
$declvarRel = 'src\parser\parser_decl_var.cpp'
$hdrRel = 'src\parser\parser.hpp'
$semRel = 'src\semantics\semantic_analyzer_stmt.cpp'

$assign = ReadSrc $assignRel
$declvar = ReadSrc $declvarRel
$hdr = ReadSrc $hdrRel
$sem = ReadSrc $semRel
foreach ($p in @($assignRel, $declvarRel, $hdrRel, $semRel)) {
    if ($null -eq (ReadSrc $p)) { $bad += ('P0 missing file: ' + $p) }
}
if ($bad.Count -gt 0) {
    foreach ($b in $bad) { Write-Host ('FAIL ' + $b) -ForegroundColor Red }
    exit 1
}

# P1
$def = @([regex]::Matches($assign, 'StmtPtr\s+Parser::wrapBodyDecls\s*\(')).Count
if ($def -ne 1) { $bad += ('P1 wrapBodyDecls defined ' + $def + ' times (want 1) - the convention has more than one home') }
$dec = @([regex]::Matches($hdr, 'StmtPtr\s+wrapBodyDecls\s*\(')).Count
if ($dec -ne 1) { $bad += ('P1 wrapBodyDecls declared ' + $dec + ' times in parser.hpp (want 1)') }

# P2: 四个调用点, 分别落在四条体级路上
$calls = @([regex]::Matches($assign, 'wrapBodyDecls\s*\(\s*loc\s*,')).Count
if ($calls -ne 4) { $bad += ('P2 body paths call the authority ' + $calls + ' times (want 4: Dim/Const/Static/Access)') }
foreach ($fn in @('parseDimStmt', 'parseConstStmtInBody', 'parseStaticStmtInBody', 'parseAccessDeclInBody')) {
    $i = $assign.IndexOf('Parser::' + $fn + '(')
    if ($i -lt 0) { $bad += ('P2 ' + $fn + ' vanished from ' + $assignRel); continue }
    # 函数体的边界 = 下一个 Parser:: 成员函数 (停在第一个 '}' 会停在函数里的第一个 if 块)
    $j = $assign.IndexOf('Parser::', $i + 8)
    if ($j -lt 0) { $j = $assign.Length }
    $body = $assign.Substring($i, $j - $i)
    if ($fn -eq 'parseStaticStmtInBody') {
        # 这个函数还管 Static Sub / Static Function 两种形, 那两路本来就不该走展开
        if ($body -notmatch 'wrapBodyDecls\s*\(\s*loc\s*,\s*parseVariableDeclList') {
            $bad += 'P2 parseStaticStmtInBody no longer routes its variable list through wrapBodyDecls'
        }
    } elseif ($body -notmatch 'wrapBodyDecls\s*\(\s*loc\s*,') {
        $bad += ('P2 ' + $fn + ' no longer routes through wrapBodyDecls')
    }
}

# P3
$hand = @([regex]::Matches($assign, 'expectName\s*\(\s*' + '"expected variable name"' + '\s*\)')).Count
if ($hand -ne 0) {
    $bad += ('P3 the hand-rolled Dim expansion is back (' + $hand + ' sites in parser_stmt_assign.cpp) - ' +
             'it is the stale copy that lost WithEvents and the suffix-as-type step')
}
$shared = @([regex]::Matches($declvar, 'expectName\s*\(\s*' + '"expected variable name"' + '\s*\)')).Count
if ($shared -lt 1) { $bad += ('P3 the shared declarator parser lost its own expectName site (' + $shared + ')') }

# P4: 别处不许把声明列表直接包成 LocalDeclStmt
$direct = @([regex]::Matches($assign, 'make_unique<LocalDeclStmt>\s*\(\s*loc\s*,\s*std::move\(\s*(decl|varDecl)\s*\)\s*\)')).Count
if ($direct -ne 1) {
    $bad += ('P4 raw MultiDecl -> LocalDeclStmt sites = ' + $direct + ' (want 1, the one inside wrapBodyDecls)')
}
$allDirect = 0
foreach ($f in @($assignRel, $declvarRel, 'src\parser\parser.cpp', 'src\parser\parser_decl.cpp')) {
    $t = ReadSrc $f
    if ($t) { $allDirect += @([regex]::Matches($t, 'make_unique<LocalDeclStmt>\s*\(\s*loc\s*,\s*std::move\(\s*(decl|varDecl)\s*\)\s*\)')).Count }
}
if ($allDirect -ne 1) { $bad += ('P4 parser-wide raw-wrap sites = ' + $allDirect + ' (want 1)') }

# P5
$semMulti = @([regex]::Matches($sem, 'case\s+ASTNodeKind::MultiDecl')).Count
if ($semMulti -ne 0) {
    $bad += ('P5 semantics grew a MultiDecl case (' + $semMulti + ') - that recreates the two-shape problem this account removed')
}

if ($bad.Count -eq 0) {
    Write-Host ('PASS body-decl shape: defs 1, callsites ' + $calls + ', hand expansion 0, raw wraps 1, semantics MultiDecl 0') -ForegroundColor Green
    exit 0
}
foreach ($b in $bad) { Write-Host ('FAIL ' + $b) -ForegroundColor Red }
exit 1
