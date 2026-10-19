# check_box_temp_reclaim.ps1 - ledger 305 B217 (knife 121) structural sentinel.
#
# The ByRef-Variant actual-argument station mints a named temp Variant that holds TWO references
# per call (one taken by vb6_ComObject_From*Instance, a second one taken by the boxing helper
# vb6_VariantObject itself). This knife hands the second one back at the end of the same
# statement. Rules below pin that the decision stays in exactly one place and that the reclaim
# really lands AFTER the call, which is what the runtime reading alone cannot prove.
#
# R1 single authority   : push_back exactly 1, declaration exactly 1, drain loop exactly 1
# R2 drain placement    : the reclaim loop sits in the same per-statement tail as ansiTempsToFree_
# R3 indent guard       : the drain only fires on the emitter's current indent level
# R4 fixture shape      : 10 station temps / 9 reclaims, every reclaim after its last use,
#                        both reclaims of the two-argument statement after that call line,
#                        the If-header temp deliberately NOT reclaimed (registered limit)
# R5 no over-release    : every reclaimed temp cleared exactly once, and no reclaim uses
#                        vb6_VariantAssign (that would deep-copy instead of handing one back)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$bad = @()

function ReadFile([string]$p) { return [IO.File]::ReadAllText($p) }

$station = Join-Path $root 'src\backend\detail\expr\cgen_expr_call_arg_emit.inc'
$state   = Join-Path $root 'src\backend\detail\util\cgen_state.inc'
$drainF  = Join-Path $root 'src\backend\stmt\cgen_stmt.cpp'
$fx      = Join-Path $root 'tests\b444_vararg\B444.vbp'

# ---- R1: one registration site, one declaration, one drain -------------------------------
$pushHits = @(Get-ChildItem -Recurse (Join-Path $root 'src') -Include *.inc,*.cpp |
              Select-String -Pattern 'boxTempsToClear_\.push_back').Count
$declHits = @(Select-String -Path $state -Pattern 'std::vector<std::pair<std::string, int>> boxTempsToClear_').Count
$drainHits = @(Select-String -Path $drainF -Pattern 'for \(auto& bt : boxTempsToClear_\)').Count
if ($pushHits -ne 1) { $bad += "R1 push-back-sites=$pushHits want=1" }
if ($declHits -ne 1) { $bad += "R1 declaration-sites=$declHits want=1" }
if ($drainHits -ne 1) { $bad += "R1 drain-sites=$drainHits want=1" }

# ---- R2: same per-statement tail as the ANSI temp release --------------------------------
$d = ReadFile $drainF
$iFree = $d.IndexOf('ansiTempsToFree_.clear();')
$iDrain = $d.IndexOf('for (auto& bt : boxTempsToClear_)')
if ($iFree -lt 0 -or $iDrain -lt 0 -or $iDrain -lt $iFree) { $bad += 'R2 drain-not-in-statement-tail' }
$iClearVec = $d.IndexOf('boxTempsToClear_.clear();')
if ($iClearVec -lt $iDrain) { $bad += 'R2 vector-not-cleared-after-drain' }

# ---- R3: the indent guard must be there --------------------------------------------------
$drainTxt = $d.Substring($iDrain, [Math]::Min(700, $d.Length - $iDrain))
if ($drainTxt -notmatch 'bt\.second\s*==\s*c_\.indentLevel\(\)') { $bad += 'R3 indent-guard-missing' }
if ($drainTxt -notmatch 'vb6_VariantClear\(&"') { $bad += 'R3 drain-does-not-clear-the-slot' }

# ---- R4/R5: the emitted shape of the fixture ---------------------------------------------
# NOTE: the temp counter restarts per procedure, so the SAME name (_vb6_vararg_0) legitimately
# appears in more than one routine. The walk below therefore pairs each reclaim with the most
# recent still-open declaration of that name instead of keying by name alone.
$c3 = Join-Path $root '.build\C3.exe'
if (-not (Test-Path $c3)) { $c3 = Join-Path $root '.build\Release\C3.exe' }
# OS-level redirection, same as check_data_recordset_shape.ps1: C3.exe puts cache maintenance
# ("cleanupOldSessions: removing ...") on stderr whenever a stale session ages out, and with
# ErrorActionPreference=Stop a merged stream turns that noise into a terminating error.
$tmpOut = Join-Path $env:TEMP ('k444_emit_' + [System.Guid]::NewGuid().ToString('N') + '.c')
$tmpErr = $tmpOut + '.err'
$pr = Start-Process -FilePath $c3 -ArgumentList @('"' + $fx + '"', '--emit-c') `
        -RedirectStandardOutput $tmpOut -RedirectStandardError $tmpErr -NoNewWindow -Wait -PassThru
$emitRc = $pr.ExitCode
$emit = [IO.File]::ReadAllText($tmpOut)
Remove-Item $tmpOut, $tmpErr -ErrorAction SilentlyContinue
if ($emitRc -ne 0) { $bad += "R4 emit-c rc=$emitRc" }
$lines = $emit -split "`r?`n"
$declCount = 0
$clearCount = 0
$open = @{}
$useOf = @{}
$clearOf = @{}
$callTwo = -1
$openAtTwo = @()
for ($i = 0; $i -lt $lines.Count; $i++) {
    $ln = $lines[$i]
    if ($ln -match 'vb6_VARIANT (_vb6_vararg_\d+) = ') {
        $declCount++
        $open[$Matches[1]] = $i
        $useOf[$Matches[1]] = -1
    } elseif ($ln -match 'vb6_VariantClear\(&(_vb6_vararg_\d+)\)') {
        $clearCount++
        $nm = $Matches[1]
        $clearOf[$nm] = $i
        if (-not $open.ContainsKey($nm)) { $bad += "R4 reclaim-without-decl $nm" }
        else {
            if ($i -le $open[$nm]) { $bad += "R4 reclaim-not-after-decl $nm" }
            if ($useOf[$nm] -ge 0 -and $i -lt $useOf[$nm]) { $bad += "R4 reclaim-before-last-use $nm clear=$i use=$($useOf[$nm])" }
            $open.Remove($nm) | Out-Null
        }
    } else {
        # a use line can carry TWO of these temps (the two-argument statement), so collect every
        # name on the line; keying on the first match alone would leave the second arg un-checked.
        foreach ($m in [regex]::Matches($ln, '&(_vb6_vararg_\d+)')) {
            $nm = $m.Groups[1].Value
            if ($open.ContainsKey($nm)) { $useOf[$nm] = $i }
        }
    }
    if ($callTwo -lt 0 -and $ln -match 'vb6_Main_TakeTwo\(&') {
        $callTwo = $i
        $openAtTwo = @($open.Keys)
    }
}
if ($declCount -ne 10) { $bad += "R4 decls=$declCount want=10" }
if ($clearCount -ne 9) { $bad += "R4 clears=$clearCount want=9" }
if ($emit.Contains('vb6_VariantAssign(&(_vb6_vararg_')) { $bad += 'R5 reclaim-uses-VariantAssign' }
# the registered limit: exactly one station temp is left unreclaimed, and it must be the one
# declared right before an If header (declaration and reclaim would sit on different levels).
if ($open.Count -ne 1) { $bad += "R4 unreclaimed-count=$($open.Count) want=1 (names: $($open.Keys -join ','))" }
foreach ($nm in $open.Keys) {
    $di = $open[$nm]
    $next = ''
    for ($j = $di + 1; $j -lt $lines.Count; $j++) { if ($lines[$j].Trim() -ne '') { $next = $lines[$j].Trim(); break } }
    if ($next -notmatch '^if \(') { $bad += "R4 unreclaimed-site-is-not-the-if-header $nm next='$next'" }
}
if ($callTwo -lt 0) { $bad += 'R4 two-arg-call-missing' }
else {
    # the reclaim must land AFTER the call, not in the generic post-line queue: the two-argument
    # statement is the shape that proves it (both temps are still open at the call line).
    if ($openAtTwo.Count -ne 2) { $bad += "R4 two-arg-statement-opens=$($openAtTwo.Count) want=2" }
    foreach ($nm in $openAtTwo) {
        if (-not $clearOf.ContainsKey($nm)) { $bad += "R4 two-arg temp never reclaimed $nm" }
        elseif ($clearOf[$nm] -le $callTwo) { $bad += "R4 two-arg reclaim not after the call $nm clear=$($clearOf[$nm]) call=$callTwo" }
    }
}
# ---- PASS/FAIL ---------------------------------------------------------------------------
if ($bad.Count -eq 0) {
    Write-Host ('check_box_temp_reclaim: PASS decls=' + $declCount + ' clears=' + $clearCount +
                ' unreclaimed=' + $open.Count + ' push=' + $pushHits + ' drain=' + $drainHits + ' guarded=1')
    exit 0
}
Write-Host 'check_box_temp_reclaim: FAIL'
foreach ($b in $bad) { Write-Host ('  - ' + $b) }
exit 1
