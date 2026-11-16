# check_created_object_typename.ps1 - Account #611: CreateObject product answers its coclass name
#
# VB6 TypeName(CreateObject("Scripting.Dictionary")) = "Dictionary", not "Object".
# JSonWrite.Serialize gates on TypeName(objRef) = "Dictionary"; answering "Object"
# made it raise "Unsupported object type: Object" (JSonFast button 7 died).
# The fix deliberately tags ONLY objects returned by vb6_CreateObject, using the
# last segment of the ProgID (the coclass name). A generic ITypeInfo lookup would
# change TypeName for every external IDispatch and break existing corpus.
#
# Usage: pwsh -File scripts\check_created_object_typename.ps1   (PASS = exit 0)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot

function Read-Ascii([string]$rel) {
    $p = Join-Path $root $rel
    if (-not (Test-Path -LiteralPath $p)) { Write-Host ("FAIL missing: " + $rel); exit 1 }
    return [System.IO.File]::ReadAllText($p)
}

$com  = Read-Ascii 'src\rtl\core\vb6com\vb6com.c'
$conv = Read-Ascii 'src\rtl\core\vb6rtl\vb6rtl_conv.c'
$fix  = Read-Ascii 'tests\jsonfast_611\CreatedObjectTypeName.bas'
$bad = @()

# ---- R1: one tagger definition, exactly two tag sites (both CreateObject returns) ----
$nTagDef = @([regex]::Matches($com, 'static void vb6_ComRememberProgID\s*\(')).Count
if ($nTagDef -ne 1) { $bad += ("R1 tagger definitions = " + $nTagDef + " (want 1)") }
# Needles copied verbatim from both CreateObject success returns:
#   vb6_ComRememberProgID(p, progId);            (local ComLib activation)
#   vb6_ComRememberProgID((void*)pDisp, progId); (CoCreateInstance)
#   vb6_ComRememberProgID(re, progId);            (C30-10 内置 VBScript.RegExp 抢先拦截)
# Count the bare identifier occurrences: 1 definition + 3 tag sites = 4.
# 第三处是 2026-11-16 加的内置正则对象：它同样是从 CreateObject 返回的，不记 ProgID 的话
# TypeName 会退化成通用 IDispatch，与真组件的 "RegExp" 不一致 —— 所以它是同一类 tag site，
# 不是漏改。事实是 4 而不是 3，判据跟着事实改，但只改这一条计数，R2/R3 两条不动。
$nTagUse = @([regex]::Matches($com, 'vb6_ComRememberProgID\s*\(')).Count
if ($nTagUse -ne 4) {
    $bad += ("R1 CreateObject tag sites = " + ($nTagUse - 1) + " (want 3: local ComLib + CoCreateInstance + 内置 RegExp)")
}
# ---- R2: the ProgID's last segment (the coclass name), not the whole ProgID ----
if (-not $com.Contains("wcsrchr(progId, L'.')")) {
    $bad += "R2 tagger no longer takes the last ProgID segment"
}

# ---- R3: one lookup definition + one consumer in vb6_TypeName ----
$nGetDef = @([regex]::Matches($com, 'BSTR vb6_ComCreatedTypeNameBStr\s*\(')).Count
if ($nGetDef -ne 1) { $bad += ("R3 lookup definitions = " + $nGetDef + " (want 1)") }
$nGetUse = @([regex]::Matches($conv, 'vb6_ComCreatedTypeNameBStr\s*\(')).Count
if ($nGetUse -ne 2) {
    $bad += ("R3 lookup references in vb6rtl_conv.c = " + $nGetUse + " (want 2: extern declaration + call site)")
}

# ---- R4: the generic "Object" answer must still be there (no broad rewrite) ----
if (-not $conv.Contains('case vb6_vtDispatch: return L"Object";')) {
    $bad += "R4 generic VT_DISPATCH -> Object answer was removed"
}

# ---- R5: fixture keeps the real shape ----
if (-not $fix.Contains('CreateObject("Scripting.Dictionary")')) { $bad += "R5 fixture lost CreateObject" }
if (-not $fix.Contains('TypeName(d)')) { $bad += "R5 fixture lost TypeName(d)" }

if ($bad.Count -gt 0) {
    Write-Host "FAIL created_object_typename" -ForegroundColor Red
    $bad | ForEach-Object { Write-Host ("  " + $_) -ForegroundColor Red }
    exit 1
}
Write-Host "PASS created_object_typename" -ForegroundColor Green
exit 0
