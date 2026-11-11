# check_uc_default_member.ps1 —— 工程内 UserControl 设计期子控件「裸名当值/当对象」的唯一出口（账 §42，族① C2039）
#
# 为什么要有这一道：VB6 里 `.ctl` 中 `Begin PhotoDemon.pdSpinner tudWidth` 这样放下的子控件，
# 代码里写 `x = tudWidth` / `tudWidth = v` / `TakeIt tudWidth`，语义都是**该 UC 的默认成员**
# （`.ctl` 里 `Attribute Value.VB_UserMemId = 0` 那枚 Property Get）或**该 UC 的实例**。
# 改前只有 13 个内置 FrmControlType 有默认属性表（getDefaultPropertyName），工程内 UC 一律
# Unknown ⇒ folds=false ⇒ 落到「类模块变量」那一格发 `me->tudWidth`（C2039：结构体没这枚成员）。
# PhotoDemon 的 pdResize.ctl 三枚（tudWidth/tudHeight/tudResolution）实测 32 条 C2039。
#
# 判定链只有一份答案、三个发码面：
#   A 解析   `.ctl` 代码段里 `Attribute <Prop>.VB_UserMemId = 0` + `Property Get <Prop>() As <型>`
#            —— 只许在 src/backend/cgen_util_ctrl.cpp 里认（第二份解析 = 两份名单会走散）。
#   B 值/标量  ctrlDefaultPropOf 回落 → `vb6_ComGet*Prop(句柄, L"<默认成员>")`（带 castTo 收口）。
#   C 写     裸名赋值 → `vb6_ComSetProp(句柄, L"<默认成员>", 打包(右值))`。
#   D 对象    实参位（`ByRef <UC>` 形参）→ `(vb6_cls_<UC>*)vb6_UC_InstanceOf(句柄)`。
#            With / 成员接收者那一族保持 HWND（那里临时量就是 HWND，换成实例指针再强回转是错的）。
#
# 判据面（任何一条不齐就红）：
#   R1 全仓 src 里认 `.VB_UserMemId` 的文件**恰好一个**（且是 cgen_util_ctrl.cpp）。
#   R2 `findUcDefaultMember` / `ucDefaultReadPlan` 各定义 1 处、各被调用 1 处。
#   R3 `isUcDefault = true` 恰好 1 处；读它的地方恰好 1 处且与 `folds` 同段（写腿）。
#   R4 三个发码面各恰好 1 处：读（castTo + readFn）、写（`/* UC default prop let */`）、
#      对象（实参位调 ucInstanceThisPtrExpr 一次；形状 `vb6_UC_InstanceOf(` 只许住
#      cgen_util_classcall.cpp 的 ucInstanceThisPtrExpr 体内 —— 账 #191 续把形状收成这一枚
#      出口，expr/ 不再手拼）。
#   R5 写腿两枚守卫必须同段在：返回槽（isRetSlotUc）与局部变量（knownLocalVars_）。
#   R6 读腿必须走 propText()（属性名是运行时读出来的，不许再拼 const char* 的 prop）。
#   R7 census 不许读空：上面每条都要真的读出数（文件存在 + 关键字命中）。
#
# 用法: powershell -File scripts\check_uc_default_member.ps1   (PASS = exit 0)

$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot
$ctrlCpp   = Join-Path $Root 'src\backend\cgen_util_ctrl.cpp'
$helpers   = Join-Path $Root 'src\backend\detail\util\cgen_helpers.inc'
$identInc  = Join-Path $Root 'src\backend\detail\expr\cgen_expr_ident_symbol.inc'
$assignInc = Join-Path $Root 'src\backend\detail\stmt\cgen_assign_stmt_special.inc'
$argInc    = Join-Path $Root 'src\backend\detail\expr\cgen_expr_call_arg_emit.inc'
$clsCall   = Join-Path $Root 'src\backend\cgen_util_classcall.cpp'
foreach ($p in @($ctrlCpp, $helpers, $identInc, $assignInc, $argInc, $clsCall)) {
    if (-not (Test-Path -LiteralPath $p)) { Write-Host ('FAIL: 读不到 ' + $p); exit 1 }
}

function Strip-CComments([string]$t) {
    $t = [regex]::Replace($t, '/\*[\s\S]*?\*/', ' ')
    $t = [regex]::Replace($t, '(?m)//.*$', ' ')
    return $t
}

$ctrl = [IO.File]::ReadAllText($ctrlCpp)
$hlp  = [IO.File]::ReadAllText($helpers)
$idt  = [IO.File]::ReadAllText($identInc)
$asg  = [IO.File]::ReadAllText($assignInc)
$arg  = [IO.File]::ReadAllText($argInc)
$cls  = [IO.File]::ReadAllText($clsCall)

$viol = @()

# ---- R1: 默认成员属性的解析只许一个文件认 ----
$owner = @(Get-ChildItem -Path (Join-Path $Root 'src') -Recurse -File -Include *.cpp,*.hpp,*.inc,*.h |
           Where-Object { (Strip-CComments ([IO.File]::ReadAllText($_.FullName))) -match 'VB_UserMemId' } |
           ForEach-Object { $_.Name } | Sort-Object -Unique)
if ($owner.Count -ne 1 -or $owner[0] -ne 'cgen_util_ctrl.cpp') {
    $viol += ('R1 默认成员解析有 ' + $owner.Count + ' 个主人: ' + ($owner -join ', ') +
              ' —— 只许 cgen_util_ctrl.cpp 认 VB_UserMemId（第二份解析 = 两份名单会走散）')
}

# ---- R2: 解析/落型各一枚定义、一枚调用 ----
$defFind = ([regex]::Matches($ctrl, 'static bool findUcDefaultMember\s*\(')).Count
$useFind = ([regex]::Matches($ctrl, 'findUcDefaultMember\s*\(\s*ucSpec->codeSection')).Count
if ($defFind -ne 1) { $viol += ('R2 findUcDefaultMember 定义 ' + $defFind + ' 处 (钉死 1 处)') }
if ($useFind -ne 1) { $viol += ('R2 findUcDefaultMember 调用 ' + $useFind + ' 处 (钉死 1 处, 只许在 ctrlDefaultPropOf 的回落里)') }
$defPlan = ([regex]::Matches($ctrl, 'static bool ucDefaultReadPlan\s*\(')).Count
$usePlan = ([regex]::Matches($ctrl, 'ucDefaultReadPlan\s*\(\s*ucType')).Count
if ($defPlan -ne 1) { $viol += ('R2 ucDefaultReadPlan 定义 ' + $defPlan + ' 处 (钉死 1 处)') }
if ($usePlan -ne 1) { $viol += ('R2 ucDefaultReadPlan 调用 ' + $usePlan + ' 处 (钉死 1 处)') }

# ---- R3: isUcDefault 一处置位、一处消费 ----
$setUc = ([regex]::Matches($ctrl, 'r\.isUcDefault\s*=\s*true')).Count
if ($setUc -ne 1) { $viol += ('R3 isUcDefault 置位 ' + $setUc + ' 处 (钉死 1 处, 只在 ctrlDefaultPropOf 的回落里)') }
$readUc = ([regex]::Matches($asg, 'foldUc\.isUcDefault')).Count
if ($readUc -ne 1) { $viol += ('R3 写腿消费 isUcDefault ' + $readUc + ' 处 (钉死 1 处)') }
if ($asg -notmatch 'foldUc\.folds\s*&&\s*foldUc\.isUcDefault') {
    $viol += 'R3 写腿没有同时问 folds 与 isUcDefault —— 两枚旗标各说一遍就会把普通控件也改道'
}

# ---- R4: 三个发码面各恰好一处 ----
$readSite = ([regex]::Matches($idt, 'fold256\.castTo\s*\+\s*fold256\.readFn')).Count
if ($readSite -ne 1) { $viol += ('R4 读腿发码 ' + $readSite + ' 处 (钉死 1 处: castTo + readFn)') }
$writeSite = ([regex]::Matches($asg, '/\* UC default prop let \*/')).Count
if ($writeSite -ne 1) { $viol += ('R4 写腿发码 ' + $writeSite + ' 处 (钉死 1 处)') }
# 实参位那一格: 只数**设计槽 → 实例指针**这一发 (同文件另有 ByRef 实参那一格也调本出口,
# 那是别的判据的站点, 不在这一条里)。
$objSite = ([regex]::Matches($arg, 'ucInstanceThisPtrExpr\("vb6_cls_" \+ cIdent\(ucParamCls\) \+ "\*", argVal\)')).Count
if ($objSite -ne 1) {
    $viol += ('R4 对象上下文发码 ' + $objSite + ' 处 (钉死 1 处: 设计槽实参调 ucInstanceThisPtrExpr 交实例指针)')
}
$ucShape = [regex]::Match($cls, 'std::string CCodeGen::ucInstanceThisPtrExpr[\s\S]{0,400}?\r?\n\}')
if (-not $ucShape.Success -or $ucShape.Value -notmatch 'vb6_UC_InstanceOf\(') {
    $viol += 'R4 ucInstanceThisPtrExpr 不再拼 vb6_UC_InstanceOf —— 实参位会交回裸句柄'
}
if (([regex]::Matches($ctrl, 'vb6_UC_InstanceOf\(')).Count -ne 0) {
    $viol += 'R4 ctrlObjectRefExpr 又长出了实例指针分支 —— With/成员接收者那一族会退化成 instance→HWND 的来回强转'
}

# ---- R5: 写腿两枚守卫 ----
if ($asg -notmatch 'isRetSlotUc') { $viol += 'R5 写腿少了返回槽守卫 (isRetSlotUc) —— Property Get 里的 `Value = …` 会被改道' }
if ($asg -notmatch '!knownLocalVars_\.count\(tgtLowerUc\)') { $viol += 'R5 写腿少了局部变量守卫 (knownLocalVars_) —— 同名局部变量会被当成控件' }

# ---- R6: 读腿属性名必须走 propText() ----
if ($idt -notmatch 'fold256\.propText\(\)') { $viol += 'R6 读腿没用 propText() —— UC 的属性名是运行时读出来的, prop 是 nullptr' }
$defPropText = ([regex]::Matches($hlp, 'std::string propText\(\) const')).Count
if ($defPropText -ne 1) { $viol += ('R6 propText() 定义 ' + $defPropText + ' 处 (钉死 1 处)') }

# ---- R7: census 不许读空 ----
$census = [pscustomobject]@{
    ownerFiles = $owner.Count
    findDef = $defFind; findUse = $useFind; planDef = $defPlan; planUse = $usePlan
    ucSet = $setUc; ucRead = $readUc
    readSites = $readSite; writeSites = $writeSite; objSites = $objSite
}
if ($ctrl.Length -lt 1000 -or $hlp.Length -lt 1000 -or $idt.Length -lt 1000 -or $asg.Length -lt 1000) {
    $viol += 'R7 源文件读出过短 —— 扫描读空, 不是一条绿着的判据'
}

Write-Host ('census: ownerFiles=' + $census.ownerFiles + ' findDef=' + $census.findDef +
            ' findUse=' + $census.findUse + ' planDef=' + $census.planDef + ' planUse=' + $census.planUse +
            ' ucSet=' + $census.ucSet + ' ucRead=' + $census.ucRead +
            ' readSites=' + $census.readSites + ' writeSites=' + $census.writeSites +
            ' objSites=' + $census.objSites)

if ($viol.Count -gt 0) {
    foreach ($v in $viol) { Write-Host ('  ' + $v) }
    Write-Host ('FAIL uc_default_member: ' + $viol.Count + ' 条违规 —— 恢复口径是让三个发码面回到 ctrlDefaultPropOf 这一枚出口, 不是放宽判据')
    exit 1
}
Write-Host 'PASS uc_default_member (R1..R7)'
exit 0
