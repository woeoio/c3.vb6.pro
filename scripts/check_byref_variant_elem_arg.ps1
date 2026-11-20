# check_byref_variant_elem_arg.ps1 - Account #611: `ByRef x As Variant` 实参必须整枚交过去
#
# 三处同源口径, 全在「ByRef Variant 形参这一格收的是什么」:
#   R1  Variant 数组的**裸元素** (`VB6_SA_AT(vb6_VARIANT, arr, i)`) 本身就是可寻址的槽
#       ⇒ 取址直传 (问句归 byRefElementAddressExpr 收口), 不许落进 VT_BSTR 装箱。
#   R2  槽元素读的**右值** (`vb6_VariantArrayGet` / `vb6_VariantArrayGetND`) 交的是整枚
#       VARIANT 值 ⇒ 用数组复合字面量整枚交出去, 不许 `vb6_VariantToString` 压成串。
#   R3  `UBound`/`LBound` 的首参是 Variant 槽时, 秩是**运行期**属性 ⇒ 必须保留派发型
#       (vb6_UBound/vb6_LBound 自带 vb6_SaIs1D 判别), 不许被「同一过程里另有一处
#       UBound(x,2)」那条按名字猜秩的判据改道成 ND 版。
#
# 改坏即红 (实测读数, 夹具 tests\jsonfast_611\VariantSlotElement.bas):
#   只坏 R1/R2 -> ub=2 d2= | vt8 | vt8 | vt8       (Boolean/嵌套数组的 vt 全丢)
#   只坏 R3    -> ub=0 d2= | vt11                  (1D 描述符被按 ND 布局读, 只剩一枚)
#   三处都在   -> ub=2 d2= | vt11 | vt8 | vt8204   (VB6 口径)
#
# Usage:  powershell -File scripts\check_byref_variant_elem_arg.ps1   (PASS = exit 0)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot

function Read-Ascii([string]$rel) {
    $p = Join-Path $root $rel
    if (-not (Test-Path -LiteralPath $p)) { Write-Host ("FAIL missing: " + $rel); exit 1 }
    return [System.IO.File]::ReadAllText($p)
}

# 判据落在**逐字串**上, 所以先把空白折成一格 —— 排版改动不该让哨兵假红,
# 但任何一个 token 变了都必须红。
function Squash([string]$s) {
    return ([regex]::Replace($s, '\s+', ' '))
}

$emit = Squash (Read-Ascii 'src\backend\detail\expr\cgen_expr_call_arg_emit.inc')
$fix  = Squash (Read-Ascii 'src\backend\detail\expr\cgen_expr_call_builtin_fixup.inc')
$fixT = Read-Ascii 'tests\jsonfast_611\VariantSlotElement.bas'
$bad = @()

# ---- R1: 裸元素走取址, 且必须排在「槽元素读」那条**之前** (否则先被装箱吃掉) ----
$r1 = 'const std::string addrElem611 = byRefElementAddressExpr(argVal, "vb6_VARIANT"); if (!addrElem611.empty()) { argVal = addrElem611; }'
$iR1 = $emit.IndexOf($r1)
$r2 = 'if (isVariantSlotElemRead(argVal)) { argVal = "&((vb6_VARIANT[]){" + argVal + "})[0]"; } else { argVal = "(&(vb6_VARIANT){.vt=VT_BSTR, .bstrVal=vb6_VariantToString(" + argVal + ")})"; }'
$iR2 = $emit.IndexOf($r2)
if ($iR1 -lt 0) {
    $bad += "R1 裸元素取址那一问 (byRefElementAddressExpr + addrElem611) 不见了"
} elseif ($iR2 -ge 0 -and $iR1 -gt $iR2) {
    $bad += "R1 裸元素取址排在槽元素读那一段之后"
}
# 旧形状必须缺席: isVariantSlot 那一臂**直接**压成 VT_BSTR (= 本刀修掉的那一形)。
# 注意 160w 那一处同名条件后面接的是 vb6_VariantToLongPtr, 不会撞上这条。
if ($emit.IndexOf('isVariantSlot(argVal, nullptr)) { argVal = "(&(vb6_VARIANT){.vt=VT_BSTR') -ge 0) {
    $bad += "R1 isVariantSlot 那一臂又直接压成 VT_BSTR 了"
}

# ---- R2: 槽元素读右值 → 数组复合字面量; else 那一支才是 VT_BSTR ----
if ($iR2 -lt 0) {
    $bad += "R2 槽元素读右值不再整枚交出去 (isVariantSlotElemRead → 复合字面量)"
}

# ---- R3: Variant 槽作 UBound/LBound 首参时保留派发型 ----
$r3 = 'if (useND && node.positional.size() >= 1 && node.positional[0] && isVariantSlot(args.empty() ? std::string() : args[0], node.positional[0].get())) { useND = false; }'
if ($fix.IndexOf($r3) -lt 0) {
    $bad += "R3 Variant 槽作 UBound/LBound 首参时不再退回派发型 (ND 版会把 1D 当 2D 读)"
}

# ---- R4: 夹具保留真实形状 (同一过程里既问 dim 2 又问 dim 1; 元素过 ByRef Variant) ----
if (-not $fixT.Contains('Walk a, 1'))          { $bad += "R4 夹具丢了调用点" }
if (-not $fixT.Contains('UBound(v, 2)'))       { $bad += "R4 夹具丢了 dim-2 那一问 (R3 的触发条件)" }
if (-not $fixT.Contains('UBound(v, 1)'))       { $bad += "R4 夹具丢了 dim-1 那一问" }
if (-not $fixT.Contains('ByRef x As Variant')) { $bad += "R4 夹具丢了 ByRef Variant 那道交接" }
if (-not $fixT.Contains('a(2) = Array(1, 2)')) { $bad += "R4 夹具丢了嵌套数组元素" }

if ($bad.Count -gt 0) {
    Write-Host "FAIL byref_variant_elem_arg" -ForegroundColor Red
    $bad | ForEach-Object { Write-Host ("  " + $_) -ForegroundColor Red }
    exit 1
}
Write-Host "PASS byref_variant_elem_arg" -ForegroundColor Green
exit 0
