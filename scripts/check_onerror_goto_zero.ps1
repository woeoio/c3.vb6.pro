# check_onerror_goto_zero.ps1 - Account #611: `On Error GoTo 0` must restore the CALLER error frame
#
# Before the fix, cgen_jumps.cpp emitted two global clears for GoToZero:
#     vb6_err_jmp_active = 0;  vb6_error_jmp_set = 0;
# That disabled the caller's setjmp frame as well, so an Err.Raise later in the
# callee became an unhandled error -> ExitProcess (JSonFast button 5).
# VB6 semantics: GoTo 0 closes only the CURRENT procedure's handler; the error
# still propagates to the caller's active handler.
#
# The verdict is pinned on the generated string AND the RTL restore site, so a
# "fix" that only changes one side (or a broad relaxation) turns this red.
# Usage: pwsh -File scripts\check_onerror_goto_zero.ps1   (PASS = exit 0)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot

function Read-Ascii([string]$rel) {
    $p = Join-Path $root $rel
    if (-not (Test-Path -LiteralPath $p)) { Write-Host ("FAIL missing: " + $rel); exit 1 }
    return [System.IO.File]::ReadAllText($p)
}

$jumps = Read-Ascii 'src\backend\stmt\cgen_jumps.cpp'
$rtl   = Read-Ascii 'src\rtl\core\vb6rtl\vb6rtl.c'
$hdr   = Read-Ascii 'src\rtl\core\vb6rtl\vb6rtl_array.h'
$fix   = Read-Ascii 'tests\jsonfast_611\OnErrorGotoZero.bas'
$bad = @()

# ---- R1: GoToZero emits the RTL restore, not the old global clears ----
if (-not $jumps.Contains('c_.emitLine("vb6_ErrGoToZero();')) {
    $bad += "R1 GoToZero no longer emits vb6_ErrGoToZero()"
}
if ($jumps.Contains('vb6_err_jmp_active = 0;') -or $jumps.Contains('vb6_error_jmp_set = 0;')) {
    $bad += "R1 old global-clear emission is back"
}

# ---- R2: exactly one RTL definition, and it restores the caller frame ----
$nDef = @([regex]::Matches($rtl, 'void vb6_ErrGoToZero\s*\(void\)')).Count
if ($nDef -ne 1) { $bad += ("R2 vb6_ErrGoToZero definitions = " + $nDef + " (want 1)") }
if (-not $rtl.Contains('vb6_err_stack[vb6_err_stack_top - 1]')) {
    $bad += "R2 RTL definition does not restore the caller frame from the stack"
}
if (-not $rtl.Contains('void vb6_ErrGoToZero(void) {')) {
    $bad += "R2 RTL definition body missing"
}

# ---- R3: declaration in the header the generated C includes ----
if (-not $hdr.Contains('void vb6_ErrGoToZero(void);')) {
    $bad += "R3 declaration missing from vb6rtl_array.h (C4013 risk)"
}

# ---- R4: fixture keeps the REAL shape (GoTo0 immediately followed by a raise,
#      with a handler in the caller) ----
if (-not $fix.Contains('On Error GoTo 0')) { $bad += "R4 fixture lost 'On Error GoTo 0'" }
if (-not [regex]::IsMatch($fix, 'On Error GoTo 0[\s\S]{0,200}?Err\.Raise')) {
    $bad += "R4 fixture no longer raises AFTER GoTo 0"
}
if (-not $fix.Contains('On Error GoTo Failed')) { $bad += "R4 fixture lost the caller handler" }

if ($bad.Count -gt 0) {
    Write-Host "FAIL onerror_goto_zero" -ForegroundColor Red
    $bad | ForEach-Object { Write-Host ("  " + $_) -ForegroundColor Red }
    exit 1
}
Write-Host "PASS onerror_goto_zero" -ForegroundColor Green
exit 0
