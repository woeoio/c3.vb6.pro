# check_udt_callret_noparens.ps1 —— 「零参函数/方法不写括号」返回工程 UDT 的唯一判据（账 §56.38 F2，C2440 族）
#
# 为什么要有这一道：VB6 允许零参调用不写括号 ——
#   `FakeCreate 0, 0, GUID_NULL, 1`            （裸名 = IdentifierExpr）
#   `m_Pen.SetPenColorRGBA m_Brush.GetBrushColorRGBA`（成员链 = MemberAccessExpr）
#   `With srcImg.GetBoundaryRectF`             （With 目标同一形状）
# 这两种 AST 形状都**不是** IndexOrCallExpr，而 inferUdtTypeOfExpr 以前只认 IndexOrCallExpr
# 里的 §56.27/§56.30（裸函数名、Mod.Fn、obj.Method(...)）。于是：
#   - ByRef UDT 实参掉回 `(intptr_t)(vb6_type_X)` 兜底 ⇒ C2440
#     （PhotoDemon WIC.c 415/432、pdPSPShape.c 501、Tools_Crop/Tools_Fill 一族）；
#   - With 目标被当普通对象 (void*)，判成 UDT 后又对右值取址 ⇒ C2102
#     （Drawing.c 1248/1378）。
# 判据收在 udtReturnOfClassMember 这一枚：有括号 (IndexOrCallExpr) 与无括号
# (MemberAccessExpr / IdentifierExpr) 两个形状共用同一张表、同一口径。
#
# 判据面（任何一条不齐就红）：
#   R1 udtReturnOfClassMember 声明 1 处（cgen_helpers.inc）、定义 1 处（cgen_util_classtype.cpp）、
#      调用 3 处（IndexOrCallExpr 档 / MemberAccessExpr 档 / cgen_with.cpp 的右值判定）。
#   R2 class 符号 memberReturnTypes 的读点**恰好 1 处**（在 udtReturnOfClassMember 体内）——
#      第二份扫描 = 两份名单会走散。
#   R3 IdentifierExpr 档（裸名零参函数）必须同时问「无参」与「返回 UserDefinedType」。
#   R4 MemberAccessExpr 档必须调 udtReturnOfClassMember（不许再手拼一份成员扫描）。
#   R5 cgen_with.cpp 的「右值」判定必须由 memberIsUdtCall 提供（裸名 / 成员两形状），
#      且裸名那一支仍要认 Function 与 PropertyGet。
#   R6 census 不许读空（文件长度 + 每条命中数）。
#
# 用法: powershell -File scripts\check_udt_callret_noparens.ps1   (PASS = exit 0)

$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot
$clsCpp  = Join-Path $Root 'src\backend\cgen_util_classtype.cpp'
$helpers = Join-Path $Root 'src\backend\detail\util\cgen_helpers.inc'
$withCpp = Join-Path $Root 'src\backend\stmt\cgen_with.cpp'
foreach ($p in @($clsCpp, $helpers, $withCpp)) {
    if (-not (Test-Path -LiteralPath $p)) { Write-Host ('FAIL: 读不到 ' + $p); exit 1 }
}

$cls = [IO.File]::ReadAllText($clsCpp)
$hlp = [IO.File]::ReadAllText($helpers)
$wth = [IO.File]::ReadAllText($withCpp)

$viol = @()

# ---- R1: 一枚定义 / 一枚声明 / 三个消费点 ----
$def = ([regex]::Matches($cls, 'std::string CCodeGen::udtReturnOfClassMember\s*\(')).Count
if ($def -ne 1) { $viol += ('R1 udtReturnOfClassMember 定义 ' + $def + ' 处 (钉死 1 处)') }
$decl = ([regex]::Matches($hlp, 'std::string udtReturnOfClassMember\s*\(')).Count
if ($decl -ne 1) { $viol += ('R1 udtReturnOfClassMember 声明 ' + $decl + ' 处 (钉死 1 处)') }
$useCls = ([regex]::Matches($cls, 'udtReturnOfClassMember\s*\(')).Count - 1
if ($useCls -ne 2) { $viol += ('R1 classtype 内消费 ' + $useCls + ' 处 (钉死 2 处: IndexOrCallExpr 档 + MemberAccessExpr 档)') }
$useWth = ([regex]::Matches($wth, 'udtReturnOfClassMember\s*\(')).Count
if ($useWth -ne 1) { $viol += ('R1 With 右值判定消费 ' + $useWth + ' 处 (钉死 1 处)') }

# ---- R2: memberReturnTypes 只有一个读点 ----
$scan = ([regex]::Matches($cls, 'memberReturnTypes\.find')).Count
if ($scan -ne 1) { $viol += ('R2 memberReturnTypes 读点 ' + $scan + ' 处 (钉死 1 处: udtReturnOfClassMember 体内)') }

# ---- R3: 裸名零参函数那一档 ----
if ($cls -notmatch 'fnSymF2->P\(\)\.params\.empty\(\)') {
    $viol += 'R3 IdentifierExpr 档没问「无参」—— 带参同名函数会被当返回值槽'
}
if ($cls -notmatch 'udtSymF2->kind == SymbolKind::UserDefinedType') {
    $viol += 'R3 IdentifierExpr 档没问「返回 UserDefinedType」—— 类/标量返回会被当 UDT 打包'
}

# ---- R4: 成员无括号那一档走同一枚出口 ----
if ($cls -notmatch 'udtReturnOfClassMember\(clsF2, ma\.memberName\)') {
    $viol += 'R4 MemberAccessExpr 档没走 udtReturnOfClassMember'
}

# ---- R5: cgen_with.cpp 的右值判定 ----
$lambdaDef = ([regex]::Matches($wth, 'auto memberIsUdtCall = \[')).Count
if ($lambdaDef -ne 1) { $viol += ('R5 memberIsUdtCall 定义 ' + $lambdaDef + ' 处 (钉死 1 处)') }
$lambdaUse = ([regex]::Matches($wth, 'memberIsUdtCall\(')).Count
if ($lambdaUse -ne 2) { $viol += ('R5 memberIsUdtCall 消费 ' + $lambdaUse + ' 处 (钉死 2 处: 有括号成员调用 / 无括号成员链)') }
if ($wth -notmatch 'fn160w->kind == SymbolKind::Function[\s\S]{0,120}SymbolKind::PropertyGet') {
    $viol += 'R5 裸名右值判定丢了 Property/Function 两档 —— Property Get 当 With 目标会对右值取址 (C2102)'
}

# ---- R6: census ----
$census = [pscustomobject]@{
    def = $def; decl = $decl; useCls = $useCls; useWth = $useWth
    scan = $scan; lambdaDef = $lambdaDef; lambdaUse = $lambdaUse
}
if ($cls.Length -lt 1000 -or $hlp.Length -lt 1000 -or $wth.Length -lt 1000) {
    $viol += 'R6 源文件读出过短 —— 扫描读空, 不是一条绿着的判据'
}

Write-Host ('census: def=' + $census.def + ' decl=' + $census.decl + ' useCls=' + $census.useCls +
            ' useWth=' + $census.useWth + ' memberScan=' + $census.scan +
            ' lambdaDef=' + $census.lambdaDef + ' lambdaUse=' + $census.lambdaUse)

if ($viol.Count -gt 0) {
    foreach ($v in $viol) { Write-Host ('  ' + $v) }
    Write-Host ('FAIL udt_callret_noparens: ' + $viol.Count + ' 条违规 —— 恢复口径是让两个 AST 形状回到 udtReturnOfClassMember 这一枚出口, 不是放宽判据')
    exit 1
}
Write-Host 'PASS udt_callret_noparens (R1..R6)'
exit 0
