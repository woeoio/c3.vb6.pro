# check_com_variant_out_ref.ps1 - [out] Variant 出参通路 (FIX-011/012) 的形状哨兵
#
# 为什么必须单开这一道 (而不是靠形状门): 形状门只哈希 --emit-c 的 **stdout**,
# 看不见**头文件里有没有原型**。而 FIX-011 的发码会调 vb6_ComVariantOuts(...) ——
# 生成码只 include "vb6com.h", 若该函数没在那个头里声明, 就是 C4013 隐式声明:
# 返回值按 int 截断指针 (x64 崩雷), 且 cl 19.51 (GA runner) 会把后续 argcount
# 判成 C2172/C2197/C2198/C2224 (VTableHandle.c 458/461 一次爆 30+ 条同源)。
# 这条隐患在 FIX-011 首轮真编译时**实测出现过** (ModProbe.c(19) C4013), 修法是
# 把声明补进 vb6com.h —— 没有针的话, 下一个人挪动/删除声明就会静默复发。
#
# 规则 (改坏了会红, 不是装饰):
#   R1  vb6com.h 必须声明 vb6_ComVariantOuts (生成码唯一 include 的 COM 头)。
#   R2  vb6com.h 必须声明 vb6_ComPackVariantRef 与 vb6_VariantComSlotToRealArray
#       (前者被生成码调, 后者被 vb6com_pack.c 调; 缺任一 ⇒ 隐式声明)。
#   R3  vb6rtl_variant.h 必须声明 vb6_SlotHasRealSafeArray 与
#       vb6_VariantNormalizeComSlot (vb6_VariantArrayGet 惰性兜底调用它们;
#       前者定义在 vb6rtl_com.c, 声明必须能被那个编译单元看到)。
#   R4  实现处不得有**与声明不一致**的签名: 三个函数各只允许一处定义。
#   R5  vb6rtl_com.c 里的 memcpy 必须按**目标**类型取长度 (sizeof(VARIANT)),
#       不许按 sizeof(vb6_VARIANT) —— 两档当前相等 (探针 layout_probe 实测
#       x64 24/24, x86 16/16), 但按源取长度就是 MEMORY 那条 "越界写堆" 的形状。
#   R6  「槽里挂的是不是**真** SAFEARRAY」那一问的判据必须是上游口径:
#       **cDims == 1 且 fFeatures != 0xA1D1**。写成 `!vb6_SaIs1D(parray)` 会把
#       **C3 的 ND 描述符 (二维数组)** 也判成真 SAFEARRAY ⇒ 归一化对它调
#       VariantClear ⇒ 把 ND 描述符当真 SAFEARRAY SafeArrayDestroy ⇒ 堆损坏。
#       实测 (2026-11-13 GA 43347652260): 三个 bas 用例 x64+x86 六枚全红 + 段错误。
#
# 用法:  powershell -File scripts\check_com_variant_out_ref.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$bad = @()

$comH  = Join-Path $root "src\rtl\core\vb6com\vb6com.h"
$varH  = Join-Path $root "src\rtl\core\vb6rtl\vb6rtl_variant.h"
$comC  = Join-Path $root "src\rtl\core\vb6rtl\vb6rtl_com.c"

foreach ($f in @($comH, $varH, $comC)) {
    if (-not (Test-Path -LiteralPath $f)) { $bad += ("M0 missing file: " + $f) }
}

if ($bad.Count -eq 0) {
    $h1 = [System.IO.File]::ReadAllText($comH)
    $h2 = [System.IO.File]::ReadAllText($varH)
    $c1 = [System.IO.File]::ReadAllText($comC)

    # --- R1: vb6_ComVariantOuts 必须在 vb6com.h 声明 ---
    if ($h1 -notmatch 'void\s*\*\s*vb6_ComVariantOuts\s*\(\s*void\s*\*\s*result\s*,\s*int32_t\s+n\s*,\s*\.\.\.\s*\)') {
        $bad += 'R1 vb6com.h does not declare vb6_ComVariantOuts(void*, int32_t, ...) -- generated code would hit C4013 implicit declaration (x64 pointer truncation)'
    }

    # --- R2: 另两个 COM 头的声明 ---
    if ($h1 -notmatch 'void\s*\*\s*vb6_ComPackVariantRef\s*\(\s*void\s*\*\s*pVar\s*\)') {
        $bad += 'R2 vb6com.h does not declare vb6_ComPackVariantRef(void*)'
    }
    if ($h1 -notmatch 'void\s+vb6_VariantComSlotToRealArray\s*\(\s*void\s*\*\s*pVar\s*\)') {
        $bad += 'R2 vb6com.h does not declare vb6_VariantComSlotToRealArray(void*)'
    }

    # --- R3: 变体头的两个声明 ---
    if ($h2 -notmatch 'int\s+vb6_SlotHasRealSafeArray\s*\(\s*void\s*\*\s*pVar\s*\)') {
        $bad += 'R3 vb6rtl_variant.h does not declare vb6_SlotHasRealSafeArray(void*)'
    }
    if ($h2 -notmatch 'void\s+vb6_VariantNormalizeComSlot\s*\(\s*void\s*\*\s*pVar\s*\)') {
        $bad += 'R3 vb6rtl_variant.h does not declare vb6_VariantNormalizeComSlot(void*)'
    }

    # --- R4: 每个函数只允许一处定义 ---
    $defs = @(
        @('vb6_ComVariantOuts', ([regex]::Matches($c1, '(?m)^void\s*\*\s*vb6_ComVariantOuts\s*\(')).Count),
        @('vb6_SlotHasRealSafeArray', ([regex]::Matches($c1, '(?m)^int\s+vb6_SlotHasRealSafeArray\s*\(')).Count),
        @('vb6_VariantNormalizeComSlot', ([regex]::Matches($c1, '(?m)^void\s+vb6_VariantNormalizeComSlot\s*\(')).Count),
        @('vb6_VariantComSlotToRealArray', ([regex]::Matches($c1, '(?m)^void\s+vb6_VariantComSlotToRealArray\s*\(')).Count)
    )
    foreach ($d in $defs) {
        if ($d[1] -ne 1) { $bad += ("R4 " + $d[0] + " has " + $d[1] + " definition(s) in vb6rtl_com.c (want exactly 1)") }
    }

    # --- R5: memcpy 长度必须按目标类型 ---
    if ($c1 -match 'memcpy\s*\(\s*pv\s*,\s*&conv\s*,\s*sizeof\s*\(\s*vb6_VARIANT\s*\)\s*\)') {
        $bad += 'R5 vb6rtl_com.c memcpy(pv, &conv, sizeof(vb6_VARIANT)) -- must use sizeof(VARIANT) (target type)'
    }

    # --- R6: 「槽里是不是真 SAFEARRAY」那一问的判据 ---
    # ⚠ 必须按上游 FIX-011 §6.2 的口径: **cDims == 1 且 fFeatures != 0xA1D1**。
    #   写成 `!vb6_SaIs1D(parray)` 会把 **C3 的 ND 描述符 (二维数组)** 也判成真
    #   SAFEARRAY ⇒ 归一化对它调 VariantClear ⇒ 把 ND 描述符当真 SAFEARRAY
    #   SafeArrayDestroy ⇒ 堆损坏 / 段错误。
    #   实测 (2026-11-13 GA 43347652260): b175_num_bstr / test_registry_settings /
    #   test_redim_elem 三个用例 x64+x86 六枚全红, b175 直接段错误 (139) ——
    #   崩在 GetAllSettings 之后读 v(0,1) (二维访问走 vb6_VariantArrayGetND)。
    if ($c1 -match 'return\s+!\s*vb6_SaIs1D\s*\(\s*pv->parray\s*\)') {
        $bad += 'R6 vb6_SlotHasRealSafeArray must test cDims==1 && fFeatures!=0xA1D1 (upstream FIX-011 6.2) -- `!vb6_SaIs1D(parray)` also matches the ND descriptor => VariantClear destroys it => heap corruption'
    }
    if ($c1 -notmatch 'a->cDims\s*!=\s*1' -or $c1 -notmatch 'a->fFeatures\s*!=\s*\(uint16_t\)VB6_SA1D_FFEATURES') {
        $bad += 'R6 vb6_SlotHasRealSafeArray lost the cDims/fFeatures two-step test (one-dimensional AND not the C3 1D descriptor magic)'
    }
}

if ($bad.Count -gt 0) {
    Write-Host "FAIL check_com_variant_out_ref"
    $bad | ForEach-Object { Write-Host ("  " + $_) }
    exit 1
}
Write-Host "OK check_com_variant_out_ref (R1..R6)"
exit 0
