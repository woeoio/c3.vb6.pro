#!/usr/bin/env powershell
# check_fileout_bstr_gate.ps1 - 结构哨兵: Print # / Write # 的「实参该以什么形状交给出口」只有一位主人
# (账 #305 第329刀 + 第331刀 / 账 #608 + 账 #609 / 台账 §B11 / 第二步 §B146 收口)
#
# 背景 (实测): `knownBstrVars_` 在每个过程入口被 clear 后**只由局部声明与形参**重新填
# (src/backend/decl/cgen_decl_proc.cpp:51 / cgen_decl_func.cpp:47)，模块级那句
# `Public gAll As String` 进不了表 ⇒ 旧的闸门把 `Print #f, gAll` 判成"非 BSTR"，落进数值兜底
# `vb6_Print(fn, vb6_Str((int32_t)(gAll)))` —— 把 BSTR 指针当数打印，文件里只剩一枚指针数。
# 而这道闸门原先在同一个文件里**抄了两份** (PrintStmt 与 WriteStmt 各一份名单 + 各一份 lambda)。
#
# 现在的形状 (第二步之后): 一位主人 + 两问，**没有名字名单**。
#   `fileOutArgIsBstr(expr, argNode)` 答「这枚实参是不是已经是一枚 BSTR 槽」：
#     ① 语义类型 String (inferExprType; builtin 返回档由 cgen_builtin_return_type_table.inc 答)
#     ② 发射形状 (BSTR 构造/读取族 / vb6_IIfBSTR / VB6_SA_AT(BSTR, / 静态 String 变量表)
#   34 名 C 出口名名录 (bstrFuncs) 已删 —— 类型答不出时**不许猜**。
#   `fileOutArgBstrExpr(expr, argNode)` 答「该以什么形状交给写文件的出口」= 上面那问 +
#   非 BSTR 那半按档 wrapToBSTR; 三户里的窗体面 Print 走它，Print#/Write# 的数值那半走
#   printWriteTextExpr (VB6 Str 字节, 见 vb6rtl_file.c)。
#
# 本哨兵钉的是**退化**，不是"出现即过"：
#   R1 两位主人各 1 份定义 + 各 1 行声明; 名单 bstrFuncs 0; 两份 lambda 抄本 0
#   R2 调用点: Print#/Write# 各问一次闸门 (=2), 窗体面 Print 走形状出口 (=1)
#   R3 「符号表那一问」只许住在主人里一次; 「只认裸标识符」那一拦 0 (回潮即丢非标识符形状)
#   R4 数值那半只许 printWriteTextExpr (false/true 各 1); wrapToBSTR 只许在形状出口里 1 次;
#      旧的 vb6_Str((int32_t)( 兜底 本文件 0 / src 其余文件 0
#   R6 名字名单那套机器 (fnName 精确匹配 / 裸前缀循环) 0
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
function N([string]$text, [string]$lit) {
    if ($null -eq $text) { return 0 }
    return ([regex]::Matches($text, [regex]::Escape($lit))).Count
}

$drvRel = "src/backend/stmt/cgen_file_io.cpp"
$apiRel = "src/backend/detail/util/cgen_api.inc"
$drv = Read-Rel $drvRel
$api = Read-Rel $apiRel
foreach ($pair in @(@($drvRel, $drv), @($apiRel, $api))) {
    if ($null -eq $pair[1]) { $viol += ("MISSING " + $pair[0]) }
}

$defRel       = "bool CCodeGen::fileOutArgIsBstr(const std::string& expr, Expr& argNode) const"
$declRel      = "bool fileOutArgIsBstr(const std::string& expr, Expr& argNode) const;"
$shapeDefRel  = "std::string CCodeGen::fileOutArgBstrExpr(const std::string& expr, Expr& argNode)"
$shapeDeclRel = "std::string fileOutArgBstrExpr(const std::string& expr, Expr& argNode);"

if ($drv) {
    # --- R1) 两位主人各一份, 而且不再有名字名单 ---
    if ((N $drv $defRel) -ne 1) { $viol += ("R1 cgen_file_io.cpp: fileOutArgIsBstr 定义 = " + (N $drv $defRel) + " (want 1)") }
    if ((N $drv $shapeDefRel) -ne 1) { $viol += ("R1 cgen_file_io.cpp: fileOutArgBstrExpr 定义 = " + (N $drv $shapeDefRel) + " (want 1: 形状这一问也要一位主人)") }
    if ((N $drv "static const std::vector<std::string> bstrFuncs") -ne 0) {
        $viol += ("R1 cgen_file_io.cpp: bstrFuncs 名单 = " + (N $drv "static const std::vector<std::string> bstrFuncs") + " (want 0: 第二步已删, 回潮 = 又按名字猜)")
    }
    if ((N $drv "bstrFuncs") -ne 0) { $viol += ("R1 cgen_file_io.cpp: bstrFuncs 残名 = " + (N $drv "bstrFuncs") + " (want 0)") }
    if ((N $drv "auto isBstrExpr") -ne 0) { $viol += ("R1 cgen_file_io.cpp: 旧 isBstrExpr lambda = " + (N $drv "auto isBstrExpr") + " (want 0)") }
    if ((N $drv "auto isStringItem") -ne 0) { $viol += ("R1 cgen_file_io.cpp: 抄本 isStringItem lambda = " + (N $drv "auto isStringItem") + " (want 0: 两份抄本回潮)") }

    # --- R2) 三户的调用点: Print#/Write# 问闸门, 窗体面 Print 走形状出口 ---
    if ((N $drv "fileOutArgIsBstr(val, *expr)") -ne 2) {
        $viol += ("R2 cgen_file_io.cpp: Print#/Write# 闸门调用点 = " + (N $drv "fileOutArgIsBstr(val, *expr)") + " (want 2)")
    }
    if ((N $drv "fileOutArgBstrExpr(val, *expr)") -ne 1) {
        $viol += ("R2 cgen_file_io.cpp: 形状出口调用点 = " + (N $drv "fileOutArgBstrExpr(val, *expr)") + " (want 1: 窗体面 Print)")
    }
    if ((N $drv "fileOutArgIsBstr(expr, argNode)") -ne 1) {
        $viol += ("R2 cgen_file_io.cpp: 闸门被形状出口调 = " + (N $drv "fileOutArgIsBstr(expr, argNode)") + " (want 1)")
    }

    # --- R3) 语义那一问只住在主人里一次; 不许再退回「只认裸标识符」 ---
    if ((N $drv "inferExprType(argNode) == Vb6Type::String") -ne 1) {
        $viol += ("R3 cgen_file_io.cpp: 符号表那一问 = " + (N $drv "inferExprType(argNode) == Vb6Type::String") + " (want 1)")
    }
    if ((N $drv "argNode.kind == ASTNodeKind::IdentifierExpr") -ne 0) {
        $viol += ("R3 cgen_file_io.cpp: 只认裸标识符那一拦 = " + (N $drv "argNode.kind == ASTNodeKind::IdentifierExpr") + " (want 0: 非标识符的 BSTR 形状会全掉进数值档)")
    }

    # --- R4) 数值那半走 printWriteTextExpr; wrapToBSTR 只许在形状出口; 旧 int32 兜底 0 ---
    if ((N $drv "printWriteTextExpr(*expr, val, false)") -ne 1) {
        $viol += ("R4 cgen_file_io.cpp: Print# 数值档 = " + (N $drv "printWriteTextExpr(*expr, val, false)") + " (want 1)")
    }
    if ((N $drv "printWriteTextExpr(*expr, val, true)") -ne 1) {
        $viol += ("R4 cgen_file_io.cpp: Write# 数值档 = " + (N $drv "printWriteTextExpr(*expr, val, true)") + " (want 1)")
    }
    if ((N $drv "return wrapToBSTR(expr, argNode);") -ne 1) {
        $viol += ("R4 cgen_file_io.cpp: wrapToBSTR 那半 = " + (N $drv "return wrapToBSTR(expr, argNode);") + " (want 1: 只住在形状出口)")
    }
    # 旧 int32 兜底现在只许活在这一个函数体内 (类型判不准的形态, 文档在册), 且必须恰好两处:
    #   printWriteTextExpr 的 Integer 档 + default 档; 第三个出现 = 有人又在 visitor 里就地抄了一份。
    $nFallback = N $drv 'vb6_Str((int32_t)("'
    if ($nFallback -ne 2) {
        $viol += ("R4 cgen_file_io.cpp: int32 兜底出现 = " + $nFallback + " (want 2: 只在 printWriteTextExpr 的 Integer/default 两档)")
    }

    # --- R6) 名字名单那套机器整族不许回潮 ---
    if ((N $drv "if (fnName == name) return true;") -ne 0) {
        $viol += ("R6 cgen_file_io.cpp: 名字精确匹配 = " + (N $drv "if (fnName == name) return true;") + " (want 0: 名单已删)")
    }
    if ((N $drv "expr.compare(0, prefix.size(), prefix)") -ne 0) {
        $viol += ("R6 cgen_file_io.cpp: 裸前缀循环 = " + (N $drv "expr.compare(0, prefix.size(), prefix)") + " (want 0)")
    }
    if ((N $drv "fnName") -ne 0) { $viol += ("R6 cgen_file_io.cpp: fnName 残名 = " + (N $drv "fnName") + " (want 0)") }
}

if ($api) {
    if ((N $api $declRel) -ne 1) { $viol += ("R1 cgen_api.inc: fileOutArgIsBstr 声明 = " + (N $api $declRel) + " (want 1)") }
    if ((N $api $shapeDeclRel) -ne 1) { $viol += ("R1 cgen_api.inc: fileOutArgBstrExpr 声明 = " + (N $api $shapeDeclRel) + " (want 1)") }
}

# --- R5) 旧数值兜底这条形状在 src/ 里只许住零处 ---
$elsewhere = 0
Get-ChildItem -LiteralPath (Join-Path $Root "src") -Recurse -Include *.cpp, *.inc -File | ForEach-Object {
    if ($_.FullName -replace "\\", "/" -like "*cgen_file_io.cpp") { return }
    $t = Get-Content -Raw -LiteralPath $_.FullName
    if ($t) { $elsewhere += ([regex]::Matches($t, [regex]::Escape('vb6_Str((int32_t)("')).Count) }
}
if ($elsewhere -ne 0) {
    $viol += ("R5 src/ 其余文件里的旧数值兜底 = " + $elsewhere + " (want 0)")
}

if ($viol.Count -gt 0) {
    Write-Host ("FAIL fileout_bstr_gate ({0} 处)" -f $viol.Count)
    $viol | ForEach-Object { Write-Host ("  " + $_) }
    exit 1
}
Write-Host "OK fileout_bstr_gate"
exit 0
