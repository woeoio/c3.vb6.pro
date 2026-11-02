#!/usr/bin/env powershell
# check_fileout_bstr_gate.ps1 - 结构哨兵: Print # / Write # 的「实参该以什么形状交给出口」只有一位主人
# (账 #305 第329刀 + 第331刀 / 账 #608 + 账 #609 / 台账 §B11)
#
# 背景 (实测): `knownBstrVars_` 在每个过程入口被 clear 后**只由局部声明与形参**重新填
# (src/backend/decl/cgen_decl_proc.cpp:51 / cgen_decl_func.cpp:47)，模块级那句
# `Public gAll As String` 进不了表 ⇒ 旧的闸门把 `Print #f, gAll` 判成"非 BSTR"，落进数值兜底
# `vb6_Print(fn, vb6_Str((int32_t)(gAll)))` —— 把 BSTR 指针当数打印，文件里只剩一枚指针数。
# 而这道闸门原先在同一个文件里**抄了两份** (PrintStmt 与 WriteStmt 各一份名单 + 各一份 lambda)，
# 两半都不问符号表，所以修一处漏一处是这条族的固有形状。
#
# 现在的形状: 两位主人各管一问，而且只有一问——
#   `fileOutArgIsBstr(expr, argNode)` 答「这枚实参是不是已经是 BSTR」；名单只有一份，
#   局部表问完之后**只对裸标识符**补问一次符号表 (账 #113 同款那把尺)。名单匹配按
#   **顶层函数名精确**：裸前缀会把工程自己的函数蒙进来 (`Direction()` 命中 `vb6_Dir`、
#   `Stripped()` 命中 `vb6_Str`)，实测那种蒙法直接 C2440「无法从 double 转换为 BSTR」。
#   `fileOutArgBstrExpr(expr, argNode)` 答「该以什么形状交给写文件的出口」= 上面那问 +
#   Form Print 早已在用的 `wrapToBSTR` 按档分诊；三户 (窗体面 Print / Print # / Write #)
#   全走它，`vb6_Str((int32_t)(...))` 那条数值兜底整个撤掉 —— 它同时塌三档 (返回 String 的
#   工程函数打成指针数 / 返回 Variant 的整棵工程编不过 / Date·Double·Boolean 截成整数)。
# 本哨兵钉: 定义 1+1 / 声明 1+1 / 出口调用点 3 / 闸门被直接调 0 / 旧 lambda 0 / 名单 1 /
# 符号表那一问 1 / 只认 IdentifierExpr 那一拦 1 / 精确名匹配 1 / 裸前缀 0 / wrapToBSTR 1 /
# 数值兜底 0 (本文件 + src/ 其余文件)。与 check_form_icon.ps1 同一类静态用例，
# 由 run_tests.ps1 的 [STATIC] 项调用。
param(
    [string]$Root = (Split-Path -Parent $PSScriptRoot)
)
$ErrorActionPreference = "Continue"
$viol = @()

function Read-Rel([string]$rel) {
    $p = Join-Path $Root $rel
    if (-not (Test-Path $p)) { return $null }
    return (Get-Content -Raw -LiteralPath $p)
}

$drvRel = "src/backend/stmt/cgen_file_io.cpp"
$apiRel = "src/backend/detail/util/cgen_api.inc"

$drv = Read-Rel $drvRel
$api = Read-Rel $apiRel
foreach ($pair in @(@($drvRel, $drv), @($apiRel, $api))) {
    if ($null -eq $pair[1]) { $viol += ("MISSING " + $pair[0]) }
}

$defRel = "bool CCodeGen::fileOutArgIsBstr\(const std::string& expr, Expr& argNode\) const"
$declRel = "bool fileOutArgIsBstr\(const std::string& expr, Expr& argNode\) const;"
$shapeDefRel = "std::string CCodeGen::fileOutArgBstrExpr\(const std::string& expr, Expr& argNode\)"
$shapeDeclRel = "std::string fileOutArgBstrExpr\(const std::string& expr, Expr& argNode\);"
$shapeRel = "fileOutArgBstrExpr\(val, \*expr\)"
$gateRel = "fileOutArgIsBstr\(val, \*expr\)"

if ($drv) {
    # --- 1) 主人只有一位，而且只有这一份名单 ---
    $nDef = ([regex]::Matches($drv, $defRel)).Count
    if ($nDef -ne 1) { $viol += ("R1 cgen_file_io.cpp: fileOutArgIsBstr 定义 = " + $nDef + " (want 1)") }
    $nShapeDef = ([regex]::Matches($drv, $shapeDefRel)).Count
    if ($nShapeDef -ne 1) { $viol += ("R1 cgen_file_io.cpp: fileOutArgBstrExpr 定义 = " + $nShapeDef + " (want 1: 形状这一问也要一位主人)") }
    $nList = ([regex]::Matches($drv, "static const std::vector<std::string> bstrFuncs")).Count
    if ($nList -ne 1) { $viol += ("R1 cgen_file_io.cpp: bstrFuncs 名单 = " + $nList + " (want 1: 两份抄本收成一处)") }

    # --- 2) 三个调用点 (窗体面 Print / Print # / Write #) 全走形状出口；旧抄本不许回潮 ---
    $nShape = ([regex]::Matches($drv, $shapeRel)).Count
    if ($nShape -ne 3) { $viol += ("R2 cgen_file_io.cpp: 形状出口调用点 = " + $nShape + " (want 3: 三户同一条)") }
    $nGate = ([regex]::Matches($drv, $gateRel)).Count
    if ($nGate -ne 0) { $viol += ("R2 cgen_file_io.cpp: 住户直接调闸门 = " + $nGate + " (want 0: 绕开形状出口自成一条 = 抄本回潮)") }
    $nGateUse = ([regex]::Matches($drv, "fileOutArgIsBstr\(expr, argNode\)")).Count
    if ($nGateUse -ne 1) { $viol += ("R2 cgen_file_io.cpp: 闸门被形状出口调 = " + $nGateUse + " (want 1)") }
    $nOld = ([regex]::Matches($drv, "auto isBstrExpr")).Count
    if ($nOld -ne 0) { $viol += ("R2 cgen_file_io.cpp: 旧 isBstrExpr lambda = " + $nOld + " (want 0: 抄本回潮)") }

    # --- 3) 符号表那一问只许一次，而且只吃裸标识符 ---
    $nAsk = ([regex]::Matches($drv, "inferExprType\(argNode\) == Vb6Type::String")).Count
    if ($nAsk -ne 1) { $viol += ("R3 cgen_file_io.cpp: 符号表那一问 = " + $nAsk + " (want 1)") }
    $nGuard = ([regex]::Matches($drv, "argNode\.kind == ASTNodeKind::IdentifierExpr")).Count
    if ($nGuard -ne 1) { $viol += ("R3 cgen_file_io.cpp: 只认裸标识符那一拦 = " + $nGuard + " (want 1)") }

    # --- 4) 非 BSTR 那半只许走 wrapToBSTR；那条 int32 数值兜底整个撤掉 (账 #609) ---
    $nWrap = ([regex]::Matches($drv, "return wrapToBSTR\(expr, argNode\);")).Count
    if ($nWrap -ne 1) { $viol += ("R4 cgen_file_io.cpp: wrapToBSTR 那半 = " + $nWrap + " (want 1: 与 Form Print 同一位主人)") }
    $nNum = ([regex]::Matches($drv, 'vb6_Str\(\(int32_t\)\(" \+ val')).Count
    if ($nNum -ne 0) { $viol += ("R4 cgen_file_io.cpp: 数值兜底 = " + $nNum + " (want 0: 三户一条口，兜底已撤)") }
}
if ($api) {
    $nDecl = ([regex]::Matches($api, $declRel)).Count
    if ($nDecl -ne 1) { $viol += ("R1 cgen_api.inc: fileOutArgIsBstr 声明 = " + $nDecl + " (want 1)") }
    $nShapeDecl = ([regex]::Matches($api, $shapeDeclRel)).Count
    if ($nShapeDecl -ne 1) { $viol += ("R1 cgen_api.inc: fileOutArgBstrExpr 声明 = " + $nShapeDecl + " (want 1)") }
}

# --- 6) 名单匹配按「顶层函数名精确」，裸前缀那把尺不许回潮 (账 #609 附带那格) ---
if ($drv) {
    $nExact = ([regex]::Matches($drv, "if \(fnName == name\) return true;")).Count
    if ($nExact -ne 1) { $viol += ("R6 cgen_file_io.cpp: 精确名匹配 = " + $nExact + " (want 1)") }
    $nPrefix = ([regex]::Matches($drv, "expr\.compare\(0, prefix\.size\(\), prefix\)")).Count
    if ($nPrefix -ne 0) { $viol += ("R6 cgen_file_io.cpp: 裸前缀匹配 = " + $nPrefix + " (want 0: 会蒙进工程自己的函数)") }
}

# --- 5) 数值兜底这条形状在 src/ 里只住这一个文件 (别处再抄一份 = 第二位主人) ---
$elsewhere = 0
Get-ChildItem -LiteralPath (Join-Path $Root "src") -Recurse -Include *.cpp, *.inc -File | ForEach-Object {
    if ($_.FullName -replace "\\", "/" -like "*cgen_file_io.cpp") { return }
    $t = Get-Content -Raw -LiteralPath $_.FullName
    if ($t) { $elsewhere += ([regex]::Matches($t, 'vb6_Str\(\(int32_t\)\(" \+ val')).Count }
}
if ($elsewhere -ne 0) {
    $viol += ("R5 src/ 其余文件里的数值兜底 = " + $elsewhere + " (want 0: 这条口径只住 cgen_file_io.cpp)")
}

if ($viol.Count -gt 0) {
    Write-Host ("FAIL fileout_bstr_gate ({0} 处)" -f $viol.Count)
    $viol | ForEach-Object { Write-Host ("  " + $_) }
    exit 1
}
Write-Host "OK fileout_bstr_gate"
exit 0
