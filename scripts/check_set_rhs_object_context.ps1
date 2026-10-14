# check_set_rhs_object_context.ps1 - 账 #258 (§B87) 的结构性哨兵 (只扫源码, 不起 cl)
#
# 为什么单开这一道: 「Set 的右值该不该折默认属性」这一问改前有两份答案 —— 环境闸
# suppressDefaultProp_ (With 块 / As Object 形参设它) 与 P16 的事后手术 (在已经发好的 value
# 文本里 find("vb6_hwnd_") 再截一段，只认 typed 控件变量那一种目标)。两份各盖一部分形状，
# 于是 `Set o = <控件名>` 那一种交出的是控件的**默认属性读数**(一枚 BSTR / 一个 int)，
# 而编译与运行都不响。现在这一问只有 CCodeGen::emitSetObjectRhs 一处回答，
# 「作为对象交出去发什么 C」只有 CCodeGen::ctrlObjectRefExpr 一处回答。
#
# 规则 (改坏了会红, 不是装饰):
#   S1  emitSetObjectRhs( 在 src/backend 里恰好 7 处 = 声明 + 定义 + **五个** Set 右值发码点
#       (MyBase 属性写 / 控件属性原生 setter / 跨类 prop_set_ / COM SetRef / 主路)。
#       少一处 = 又有一个 Set 出口没问它 (本账的形状)；多一处 = 有人在抄第二份
#   S2  cgen_setlet_set_prop.inc 里 emitExpr(*node.value) 必须 0 处 —— 那正是"绕过唯一出口"的形
#   S3  同一文件里 value.find("vb6_hwnd_") 必须 0 处 —— P16 那截事后手术不许回来
#   S4  ctrlObjectRefExpr( 恰好 3 处 = 声明 + 定义 + 那一个消费点 (suppressDefaultProp_ 那一支)
#   S5  消费点里不许再手拼 "vb6_hwnd_" + cIdent(node.name) —— 那拼不出 ListView 槽变量 /
#       ImageList 的 vb6_com_ / WithEvents 控件变量那三档 (改前它们发的是 C 里没有的名字)
#   S6  那道闸只许在右值**整枚是一枚标识符**时开 (bareIdent 那一行必须在，且只有一处) ——
#       闸是无环境的，顺着子树开就会把嵌套实参位 (`Set o = f(Text1)`) 也改成交对象
#
# 用法:  pwsh -File scripts\check_set_rhs_object_context.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$bad = @()

$backend = Join-Path $root "src\backend"
if (-not (Test-Path -LiteralPath $backend)) {
    Write-Host ("FAIL S0 missing " + $backend) -ForegroundColor Red
    exit 1
}

$textOf = @{}
Get-ChildItem -LiteralPath $backend -Recurse -File |
    Where-Object { $_.Extension -in ".c", ".h", ".cpp", ".inc", ".hpp" } |
    ForEach-Object { $textOf[$_.FullName] = [System.IO.File]::ReadAllText($_.FullName) }

function Count-Pat([string]$pattern) {
    $n = 0
    foreach ($k in $textOf.Keys) { $n += @([regex]::Matches($textOf[$k], $pattern)).Count }
    return $n
}

# ---- S1: 唯一出口的站点计数 ----
$s1 = Count-Pat 'emitSetObjectRhs\s*\('
if ($s1 -ne 7) { $bad += ("S1 emitSetObjectRhs sites = " + $s1 + " (need exactly 7: decl + def + 5 Set RHS sites)") }

# ---- S2: Set 的右值不许再绕过唯一出口 ----
$setPropFile = Join-Path $root "src\backend\detail\stmt\cgen_setlet_set_prop.inc"
if (-not (Test-Path -LiteralPath $setPropFile)) {
    Write-Host ("FAIL S2 missing " + $setPropFile) -ForegroundColor Red
    exit 1
}
$setPropText = [System.IO.File]::ReadAllText($setPropFile)
$s2 = @([regex]::Matches($setPropText, 'emitExpr\(\s*\*node\.value\s*\)')).Count
if ($s2 -ne 0) { $bad += ("S2 Set RHS emitted outside the single authority: emitExpr(*node.value) x" + $s2 +
                          " in cgen_setlet_set_prop.inc (each one re-opens account 258)") }

# ---- S3: P16 的事后字符串手术不许回来 ----
$s3 = @([regex]::Matches($setPropText, 'value\.find\(\s*"vb6_hwnd_"\s*\)')).Count
if ($s3 -ne 0) { $bad += ("S3 the after-the-fact hwnd substring surgery is back (x" + $s3 +
                          ") -- that is the second copy of the answer this account removed") }

# ---- S4/S5: 句柄拼法那一份权威 ----
$s4 = Count-Pat 'ctrlObjectRefExpr\s*\('
if ($s4 -ne 3) { $bad += ("S4 ctrlObjectRefExpr sites = " + $s4 + " (need exactly 3: decl + def + the one consumer)") }

$identFile = Join-Path $root "src\backend\detail\expr\cgen_expr_ident_symbol.inc"
if (-not (Test-Path -LiteralPath $identFile)) {
    Write-Host ("FAIL S5 missing " + $identFile) -ForegroundColor Red
    exit 1
}
$identText = [System.IO.File]::ReadAllText($identFile)
$s5 = @([regex]::Matches($identText, '"vb6_hwnd_"\s*\+\s*cIdent\(\s*node\.name\s*\)')).Count
if ($s5 -ne 0) { $bad += ("S5 the suppress branch hand-concats vb6_hwnd_+node.name (x" + $s5 +
                          ") instead of asking ctrlObjectRefExpr") }
$s5b = @([regex]::Matches($identText, 'lastExpr_\s*=\s*ctrlObjectRefExpr\(')).Count
if ($s5b -ne 1) { $bad += ("S5 the suppress branch asks ctrlObjectRefExpr " + $s5b + " times (need exactly 1)") }

# ---- S6: 闸只开在整枚标识符上 ----
$setletFile = Join-Path $root "src\backend\stmt\cgen_setlet.cpp"
if (-not (Test-Path -LiteralPath $setletFile)) {
    Write-Host ("FAIL S6 missing " + $setletFile) -ForegroundColor Red
    exit 1
}
$setletText = [System.IO.File]::ReadAllText($setletFile)
$s6 = @([regex]::Matches($setletText, 'bareIdent\s*=\s*value\s*&&\s*value->kind\s*==\s*ASTNodeKind::IdentifierExpr')).Count
if ($s6 -ne 1) { $bad += ("S6 the suppress gate is keyed on the bare-identifier RHS " + $s6 +
                          " times (need exactly 1, otherwise it leaks into nested argument positions)") }

if ($bad.Count -ne 0) {
    foreach ($b in $bad) { Write-Host ("FAIL " + $b) -ForegroundColor Red }
    exit 1
}
Write-Host "PASS set_rhs_object_context (S1 7 sites, S2 no bypass, S3 no P16 surgery, S4 3 hwnd-authority sites, S5 suppress branch asks it, S6 gate is identifier-only)"
