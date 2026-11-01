# check_prop_write_value_slot.ps1 - 账 #305 第一百九十七刀的结构性哨兵 (只扫源码, 不起 cl)
#
# 为什么单开这一道: 「Pattern C/D2 那一发属性写的**值槽该发哪一档**」同一条问句原本住着两半 ——
# Let 臂 (Fix 090w/156-B/169/159-B) 带着那张解箱表, Set 臂 (Fix 268) 只有一条 `type == Variant
# ⇒ 装箱`, 而两半都只问**语义** Vb6Type、谁都没问**槽的 C 型**。两个方向的漏各打出一处真红,
# 且是同一声 cl C2172「实参不是指针: 参数 2」:
#   · 多装 —— VbEclipse ucButton.ctl:196 `Property Set Picture(ByVal New_Picture As Picture)`
#     的语义档折成 Variant 而 C 槽是 void* ⇒ vb6_VariantFromValue(...) 递给 void*;
#   · 少装 —— 类体内裸写自己的对象属性 (夹具 tests/k3122_own_prop_set 的 `Set Pic = GetV()`)
#     槽是 void*、RHS 是整只 vb6_VARIANT ⇒ 解箱那一发没人问。
# 现在这一问只有一处回答 (writeValArgForParamSlot), 两站都是消费者。哨兵钉的是这六件事
# 不许再长出第二份、也不许把已经撤掉的手抄抄回来。
#
# 规则 (改坏了会红, 不是装饰):
#   R1  唯一主人: `writeValArgForParamSlot(` 在 cgen_util_comwrite.cpp 恰好 3 处 = 定义 1 +
#       站点 2 (Let 臂 + Set 臂); 全 src/backend 恰好 4 处 = 再加 cgen_helpers.inc 那份声明。
#       多一处 = 别处又开始自己答这一问。
#   R2  站点不许绕过主人直接装箱: `valArg = packLetValueArg` 恰好 0; 本文件里
#       `packLetValueArg(` 恰好 2 = 它自己的定义 + 主人内部那一句调用。
#   R3  解箱表只有一份: `ext = "vb6_VariantTo` 恰好 7 档, 且 `vb6_VariantToSafeArray1D`
#       在本文件恰好 1 处 (第二张表就是从这一行开始的)。
#   R4  装箱必须问 C 槽型: `slotC == "vb6_VARIANT"` 恰好 1; 而旧的「只看语义档」那道闸
#       `setP268.back().type == Vb6Type::Variant` 必须 0 处 (它就是把 Picture 那发装错的尺)。
#   R5  顺序: 泛对象槽「交槽地址」那一发套在两臂**之后** —— `if (wrapSetSlotAddr9) valArg`
#       恰好 1, 且旧形 `std::string valArg = wrapSetSlotAddr9` 恰好 0 (旧序里 Fix 268 一旦
#       装箱就把 rev9 的 &(...) 整个丢掉, 两问互相覆盖)。
#   R6  夹具与判据齐件: k3122 三件套在; run_tests.ps1 里 `k3122_own_prop_set` 恰好 8 处
#       (两架构真跑 + 两架构路径 + 两枚发码格名 + 两枚发码路径), 两头发码针各恰好 1 处 ——
#       正针 = vb6_VariantToObjectVal(vb6_P3122_GetV, 反面对象 = vb6_P3122_prop_set_Pic((void*)me, o);,
#       反面针 (旧病灶) = vb6_P3122_prop_set_Pic((void*)me, vb6_P3122_GetV。
# 用法: powershell -File scripts\check_prop_write_value_slot.ps1   (PASS = exit 0)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$script:bad = @()

function Text-Of([string]$rel) {
    $p = Join-Path $root $rel
    if (-not (Test-Path -LiteralPath $p)) { $script:bad += ("缺文件 " + $rel); return "" }
    return (Get-Content -Raw -Encoding UTF8 $p)
}

$ownerRel = "src\backend\cgen_util_comwrite.cpp"
$declRel  = "src\backend\detail\util\cgen_helpers.inc"
$rtRel    = "tests\run_tests.ps1"
$owner = Text-Of $ownerRel
$decl  = Text-Of $declRel
$rt    = Text-Of $rtRel

function Count-Of([string]$hay, [string]$needle) {
    if (-not $hay) { return 0 }
    return @([regex]::Matches($hay, [regex]::Escape($needle))).Count
}

# ---- R1 唯一主人 ----
$nOwner = Count-Of $owner "writeValArgForParamSlot("
if ($nOwner -ne 3) { $script:bad += ("R1 本文件里 writeValArgForParamSlot = $nOwner 处 (要 3: 定义 1 + Let/Set 两站)") }
$nDecl = Count-Of $decl "writeValArgForParamSlot("
if ($nDecl -ne 1) { $script:bad += ("R1 cgen_helpers.inc 里那份声明 = $nDecl 处 (要 1)") }
$backendFiles = @(Get-ChildItem -Path (Join-Path $root "src\backend") -Recurse -File -Include *.cpp,*.inc,*.hpp)
$nAll = 0
foreach ($f in $backendFiles) { $nAll += (Count-Of (Get-Content -Raw -Encoding UTF8 $f.FullName) "writeValArgForParamSlot(") }
if ($nAll -ne 4) { $script:bad += ("R1 全 src/backend 里 writeValArgForParamSlot = $nAll 处 (要 4: 声明 + 定义 + 站点 2; 多一处就是第二份答案)") }

# ---- R2 站点不许绕过主人 ----
if ((Count-Of $owner "valArg = packLetValueArg") -ne 0) {
    $script:bad += "R2 站点里又出现 valArg = packLetValueArg (装箱必须走主人)"
}
$nPack = Count-Of $owner "packLetValueArg("
if ($nPack -ne 2) { $script:bad += ("R2 本文件里 packLetValueArg = $nPack 处 (要 2: 定义 1 + 主人调用 1)") }

# ---- R3 解箱表只有一份 ----
$nArms = Count-Of $owner 'ext = "vb6_VariantTo'
if ($nArms -ne 7) { $script:bad += ("R3 解箱档位数 = $nArms (要 7: String/Long 族/Double 族/LongPtr 族/Object/数组/泛对象)") }
$nArr = Count-Of $owner "vb6_VariantToSafeArray1D"
if ($nArr -ne 1) { $script:bad += ("R3 vb6_VariantToSafeArray1D 在本文件 = $nArr 处 (要 1: 第二张表就从这里长出来)") }

# ---- R4 装箱必须问 C 槽型 ----
if ((Count-Of $owner 'slotC == "vb6_VARIANT"') -ne 1) {
    $script:bad += "R4 装箱那一问不再问 C 槽型 (slotC == vb6_VARIANT 那一判据必须恰好 1)"
}
if ((Count-Of $owner "setP268.back().type == Vb6Type::Variant") -ne 0) {
    $script:bad += "R4 Set 臂的旧闸回来了 (只看语义档的 setP268.back().type == Vb6Type::Variant)"
}

# ---- R5 取址套在值档之后 ----
if ((Count-Of $owner 'if (wrapSetSlotAddr9) valArg') -ne 1) {
    $script:bad += "R5 泛对象槽取址那一发不在两臂之后 (要恰好 1 处 if (wrapSetSlotAddr9) valArg)"
}
if ((Count-Of $owner "std::string valArg = wrapSetSlotAddr9") -ne 0) {
    $script:bad += "R5 旧形回来了 (valArg 初值就套 &(...), 装箱会把它整个丢掉)"
}

# ---- R6 夹具与判据齐件 ----
foreach ($f in @("tests\k3122_own_prop_set\K3122.vbp", "tests\k3122_own_prop_set\P3122.cls",
                 "tests\k3122_own_prop_set\mdl3122.bas")) {
    if (-not (Test-Path -LiteralPath (Join-Path $root $f))) { $script:bad += ("R6 缺夹具 " + $f) }
}
$nCells = Count-Of $rt "k3122_own_prop_set"
if ($nCells -ne 8) { $script:bad += ("R6 run_tests.ps1 里 k3122_own_prop_set = $nCells 处 (要 8: 真跑 2 + 路径 2 + 发码格名 2 + 发码路径 2)") }
foreach ($c in @("vb6_VariantToObjectVal(vb6_P3122_GetV",
                 "vb6_P3122_prop_set_Pic((void*)me, o);",
                 "vb6_P3122_prop_set_Pic((void*)me, vb6_P3122_GetV")) {
    $k = Count-Of $rt $c
    if ($k -ne 1) { $script:bad += ("R6 发码针 " + $c + " = $k 处 (要恰好 1)") }
}

if ($script:bad.Count -eq 0) {
    $summary = "PASS owner-sites=$nOwner/backend=$nAll pack-bypass=0 pack-local=$nPack unbox-arms=$nArms"
    $summary += " array-table=$nArr slotC-ask=1 semantic-gate=0 addr-wrap=1 cells=$nCells"
    Write-Host $summary -ForegroundColor Green
    exit 0
}
foreach ($b in $script:bad) { Write-Host ("FAIL " + $b) -ForegroundColor Red }
exit 1
