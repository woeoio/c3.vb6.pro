# check_nd_variant_com_roundtrip.ps1 - Account #611 knife 349: ND Variant SAFEARRAY <-> COM
#
# Why: a 2D Variant array passed through a real external COM object (Scripting.Dictionary)
# came back as VT_EMPTY. Two independent holes: (1) vb6_ComPackArrayOut's ND arm only
# accepted vb6_ComSaVtBlittable element kinds, so vb6_sa_variant returned with out=VT_EMPTY;
# (2) both inbound arms (vb6_ComVariantIntoSlot / vb6_VariantArrayIntoSlot) hard-rejected
# cDims != 1. C3's vb6_SafeArrayND is NOT an OLE SAFEARRAY: bounds[] field order is
# {lBound,cElements} vs SAFEARRAYBOUND {cElements,lLbound}, and psa->rgsabound[] is
# stored in reverse. So the only correct bridge is per-dimension SafeArrayGetLBound/
# GetUBound + per-element copy, never a struct cast.
#
# Dimension order was measured (not guessed) with .temp/nd349/sa_dim_probe.c:
#   SafeArrayCreate(vt, 2, rg) takes rg[d] in FORWARD order (SafeArrayGetBound(dim=d+1)),
#   and linear element order has d=0 varying fastest - the same as vb6_SaNdElemPtr.
#   So no transpose is needed, but a struct cast would silently transpose + mis-size.
#
# Rules (each one is a real red when broken, not decoration):
#   R1  the ND variant outbound arm exists and builds a VT_VARIANT SAFEARRAY
#   R2  that arm packs each element with vb6_VariantPackCopyInto (never memcpy)
#   R3  exactly one definition + one forward decl of the ND conversion outlet
#   R4  the outlet uses SafeArrayGetLBound/UBound(d+1) + vb6_SafeArrayCreateND, no cast
#   R5  VT_VARIANT elements are deep-copied (VariantCopy -> steal -> VariantClear)
#   R6  the COM result arm dispatches cDims 2..16 to that single outlet
#   R7  the server inbound arm no longer early-returns on cDims != 1
#   R8  vb6_VariantSlotRelease picks the ND destructor for ND descriptors
#   R9  1D variant outbound + the blittable whitelist survive (zero-regression guard)
#   R10 harness registers the fixture + discriminating needles; fixture has them
# Usage: powershell -File scripts\check_nd_variant_com_roundtrip.ps1  (PASS = exit 0)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$com = Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_com.c'
$hdr = Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_array.h'
$harness = Join-Path $root 'tests\run_tests.ps1'
$fixture = Join-Path $root 'tests\jsonfast_611\NDVariantComRoundTrip.bas'
foreach ($p in @($com, $hdr, $harness, $fixture)) {
    if (-not (Test-Path -LiteralPath $p)) { Write-Host "FAIL missing path: $p"; exit 1 }
}
$s = [System.IO.File]::ReadAllText($com)
$a = [System.IO.File]::ReadAllText($hdr)
$h = [System.IO.File]::ReadAllText($harness)
$f = [System.IO.File]::ReadAllText($fixture)
$bad = @()

# Slice between two EXACT ASCII anchors. Every end anchor below is a function
# signature, so a comment mentioning a name cannot shorten the block.
function Slice-Between([string]$text, [string]$from, [string]$to, [string]$label) {
    $i = $text.IndexOf($from)
    if ($i -lt 0) { $script:bad += "$label cannot locate start anchor: $from"; return '' }
    $j = $text.IndexOf($to, $i + $from.Length)
    if ($j -le $i) { $script:bad += "$label cannot locate end anchor: $to"; return '' }
    return $text.Substring($i, $j - $i)
}

$ndOut  = Slice-Between $s 'if (nd->elemType == vb6_sa_variant) {' 'vt = vb6_ComSaVtBlittable(nd->elemType);' 'R1'
$ndIn   = Slice-Between $s 'static void vb6_ComNDArrayIntoSlot(SAFEARRAY* psa, UINT dim, VARTYPE baseVt,' 'static void vb6_VariantArrayIntoSlot' 'R4'
$resArm = Slice-Between $s 'UINT cDims = psa ? SafeArrayGetDim(psa) : 0;' 'if (steal) pv->parray = NULL;' 'R6'
$srvArm = Slice-Between $s 'static void vb6_VariantArrayIntoSlot' 'void vb6_VariantSlotRelease' 'R7'
$rel    = Slice-Between $s 'void vb6_VariantSlotRelease(void *slot) {' 'static void vb6_ComPackArrayOut' 'R8'

# ---- R1: ND variant outbound arm ----
if ($ndOut -ne '') {
    foreach ($needle in @('SafeArrayCreate(VT_VARIANT, (ULONG)nd->dimCount, rg)',
                          'out->vt = (VARTYPE)(VT_ARRAY | VT_VARIANT);',
                          'out->parray = psa;')) {
        if ($ndOut.IndexOf($needle) -lt 0) { $bad += "R1 outbound arm missing $needle" }
    }
}
foreach ($needle in @('rg[d].lLbound = (ULONG)nd->bounds[d].lBound;',
                      'rg[d].cElements = (ULONG)nd->bounds[d].cElements;')) {
    if ($s.IndexOf($needle) -lt 0) { $bad += "R1 ND bounds loop missing $needle" }
}

# ---- R2: per-element OWNED copy, never a raw memcpy ----
if ($ndOut -ne '') {
    if (@([regex]::Matches($ndOut, 'vb6_VariantPackCopyInto\(')).Count -ne 1) {
        $bad += 'R2 ND variant arm must pack elements through exactly one vb6_VariantPackCopyInto'
    }
    if ($ndOut.IndexOf('memcpy(') -ge 0) {
        $bad += 'R2 ND variant arm must NOT memcpy VARIANT slots (ownership + layout)'
    }
    if (@([regex]::Matches($ndOut, 'SafeArrayAccessData\(')).Count -ne 1 -or
        @([regex]::Matches($ndOut, 'SafeArrayUnaccessData\(')).Count -ne 1) {
        $bad += 'R2 ND variant arm must lock/unlock the destination SAFEARRAY exactly once'
    }
}

# ---- R3: one definition + one forward decl, two call sites, void+out-param ----
$decl = 'static void vb6_ComNDArrayIntoSlot(SAFEARRAY* psa, UINT dim, VARTYPE baseVt,'
if (@([regex]::Matches($s, [regex]::Escape($decl))).Count -ne 2) {
    $bad += 'R3 ND conversion outlet needs exactly one forward decl + one definition'
}
if (@([regex]::Matches($s, '(?m)^static vb6_VARIANT\s+vb6_ComNDArrayIntoSlot')).Count -ne 0) {
    $bad += 'R3 ND outlet must be void+out-param (a vb6_VARIANT return leaks into cgen_variant_slot_prefixes.inc)'
}
if (@([regex]::Matches($s, 'vb6_ComNDArrayIntoSlot\(psa, cDims')).Count -ne 2) {
    $bad += 'R3 both inbound arms must call the same ND conversion outlet'
}

# ---- R4/R5: the single ND conversion outlet ----
if ($ndIn -ne '') {
    foreach ($needle in @('SafeArrayGetLBound(psa, d + 1, &lo)',
                          'SafeArrayGetUBound(psa, d + 1, &hi)',
                          'vb6_SafeArrayCreateND(etNd, (int32_t)dim, bounds)',
                          'SafeArrayAccessData(psa, &srcData)',
                          's->vt = (vb6_vartype)(vb6_vtArray | allowNd);',
                          's->parray = (struct vb6_SafeArray1D*)nd;')) {
        if ($ndIn.IndexOf($needle) -lt 0) { $bad += "R4 ND outlet missing $needle" }
    }
    if ($ndIn.IndexOf('->rgsabound[0].cElements') -ge 0) {
        $bad += 'R4 ND outlet must not read psa->rgsabound directly (reverse order + field swap)'
    }
    foreach ($needle in @('VariantCopy(&elem, srcElem)',
                          'vb6_ComVariantIntoSlot(&elem, 1)',
                          'VariantClear(&elem)')) {
        if ($ndIn.IndexOf($needle) -lt 0) { $bad += "R5 ND outlet missing deep-copy step $needle" }
    }
    if (@([regex]::Matches($ndIn, 'vb6_ComSaVtBlittable\(etNd\)')).Count -ne 1) {
        $bad += 'R5 ND outlet must consult the one blittable whitelist exactly once'
    }
}

# ---- R6: COM result arm ----
if ($resArm -ne '') {
    foreach ($needle in @('psa && cDims >= 2 && cDims <= 16',
                          'vb6_ComNDArrayIntoSlot(psa, cDims,')) {
        if ($resArm.IndexOf($needle) -lt 0) { $bad += "R6 result arm missing $needle" }
    }
}

# ---- R7: server inbound arm ----
if ($srvArm -ne '') {
    if ($srvArm.IndexOf('SafeArrayGetDim(psa) != 1') -ge 0) {
        $bad += 'R7 server inbound arm still hard-rejects cDims != 1'
    }
    foreach ($needle in @('if (cDims >= 2) {', 'if (cDims <= 16) vb6_ComNDArrayIntoSlot(psa, cDims, base, s);',
                          'if (cDims != 1) return;')) {
        if ($srvArm.IndexOf($needle) -lt 0) { $bad += "R7 server inbound arm missing $needle" }
    }
}

# ---- R8: slot release ----
if ($rel -ne '') {
    foreach ($needle in @('if (vb6_SaIs1D(s->parray)) {',
                          'vb6_SafeArrayDestroy1D(s->parray);',
                          'vb6_SafeArrayDestroyND((vb6_SafeArrayND*)s->parray);')) {
        if ($rel.IndexOf($needle) -lt 0) { $bad += "R8 slot release missing $needle" }
    }
}

# ---- R9: zero-regression guards + the two descriptor layouts still differ ----
foreach ($needle in @('if (a1->elemType == vb6_sa_variant) {',
                      'static VARTYPE vb6_ComSaVtBlittable(vb6_safearray_elemtype et)',
                      'vb6_ComSaVtBlittable(a1->elemType)',
                      'vb6_ComSaVtBlittable(nd->elemType)',
                      'memcpy(pvData, data, (size_t)n * (size_t)esz);')) {
    if ($s.IndexOf($needle) -lt 0) { $bad += "R9 zero-regression guard disappeared: $needle" }
}
foreach ($needle in @('typedef struct vb6_SafeArrayBound {',
                      'int32_t lBound;', 'int32_t cElements;',
                      'typedef struct vb6_SafeArrayND {')) {
    if ($a.IndexOf($needle) -lt 0) { $bad += "R9 ND descriptor layout anchor missing: $needle" }
}

# ---- R10: harness + fixture needles ----
foreach ($needle in @('NDVariantComRoundTrip.bas', 'NC-vt=8204', 'NC-3e=ah',
                      'NC-lvt=8195', 'check_nd_variant_com_roundtrip.ps1',
                      'Test-NDVariantComRoundTrip')) {
    if ($h.IndexOf($needle) -lt 0) { $bad += "R10 harness missing $needle" }
}
foreach ($needle in @('NC-vt=', 'NC-b1=', 'NC-b2=', 'NC-e00=', 'NC-e12=',
                      'NC-3e=', 'NC-lvt=', 'NC-DONE')) {
    if ($f.IndexOf($needle) -lt 0) { $bad += "R10 fixture missing $needle" }
}

if ($bad.Count -gt 0) {
    $bad | ForEach-Object { Write-Host "FAIL $_" }
    exit 1
}
Write-Host 'PASS nd_variant_com_roundtrip'
exit 0
