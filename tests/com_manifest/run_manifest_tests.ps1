# com_manifest fixture - registration-free COM (manifest = EXTERNAL manifest file)
#
# The toolkit reference manifests (D:\vb_yqt4qPac\...\清单\MSDatGrd.ocx.manifest,
# TabCtl32.Ocx.manifest) have NO <assemblyIdentity> and NO requestedExecutionLevel:
# they are <assembly><file name=".."><comClass clsid tlbid/></file> + proxy stubs.
# Runtime reading (this machine, Win10.0.26200): a bare <component>.manifest sitting
# next to the DLL is NOT consumed by SxS (CoCreateInstance = 0x80040154
# REGDB_E_CLASSNOTREG), but the SAME <file>+<comClass> block inside an activated
# process manifest resolves fine (CreateActCtx=OK, no 14001).
#
# Cases below (all with a NEVER-REGISTERED test_activex_dll.dll):
#   A  explicit client manifest (<file>+comClass) -> CreateObject("TestAXDLL.Calc") OK
#   A' same exe in -ComMode direct               -> must FAIL (proves A used SxS)
#   B  <component>.manifest beside the DLL, no -ComManifest, ComLib= -> runtime
#      composer reuses its <file> block; CreateObject("ExtManifest.Calc") OK
#   B' same exe in -ComMode direct               -> must FAIL (proves B consumed it)
#   C  no external <component>.manifest -> composer synthesizes <file>+comClass
#      from the component table; with C3_COM_PREFER_REG=1 CreateObject OK via SxS
#   C' direct + PREFER_REG=1                     -> must FAIL
#   D  client asks dependentAssembly version 2.0.0.0, private assembly is 1.0.0.0
#      -> manifest mode must fail loudly naming the SxS code (no false green)
#   E  direct regression: ComLib= + -ComMode direct -> direct-create=OK
#   F  unusable client manifest -> manifest mode fails naming what is missing
#
# Windows caches a parsed activation context per client manifest PATH: every run
# uses a fresh dir + a run-tagged manifest name so stale mappings cannot leak.
#
param(
    [string]$C3 = "",
    [string]$OutDir = ""
)

$ErrorActionPreference = 'Continue'
$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
if (-not $C3)  { $C3 = Join-Path $root '.build\C3.exe' }
if (-not $OutDir) { $OutDir = Join-Path $PSScriptRoot 'output' }
New-Item -ItemType Directory -Force $OutDir | Out-Null
$runTag = (Get-Date -Format 'yyyyMMddHHmmssfff')
$base = Join-Path $OutDir ("run-$runTag")
New-Item -ItemType Directory -Force $base | Out-Null
$env:TMP = (Join-Path $base 'tmp')
$env:TEMP = $env:TMP
New-Item -ItemType Directory -Force $env:TMP | Out-Null

$script:fail = 0
function Check($name, $cond, $detail) {
    if ($cond) { Write-Host "PASS  $name" -ForegroundColor Green }
    else { Write-Host "FAIL  $name  $detail" -ForegroundColor Red; $script:fail++ }
}

function Find-Mt {
    $c = @()
    if ($env:MT -and (Test-Path $env:MT)) { $c += $env:MT }
    $kits = @('C:\Program Files (x86)\Windows Kits\10\bin', 'D:\Windows Kits\10\bin')
    foreach ($k in $kits) {
        if (Test-Path $k) {
            $c += Get-ChildItem $k -Recurse -Filter 'mt.exe' -ErrorAction SilentlyContinue |
                  Where-Object { $_.FullName -match '\\x64\\' } | Select-Object -ExpandProperty FullName
        }
    }
    foreach ($x in $c) { if (Test-Path $x) { return $x } }
    return $null
}

$mt = Find-Mt
Check 'mt.exe available' ($null -ne $mt) 'Windows SDK mt.exe not found'

# --- server DLL: never registered in either registry view --------------------
$buildDir = Join-Path $base 'build'
& $C3 (Join-Path $root 'tests\test_activex_dll\test_activex_dll.vbp') --dll --output-dir $buildDir | Out-Null
$srcDll = Join-Path $buildDir 'test_activex_dll.dll'
Check 'build test_activex_dll.dll' (Test-Path $srcDll) 'no output'
$clsKey = 'CLSID\{D84F362F-8EF1-D16D-8814-C16ADB700BAB}\InprocServer32'
$reg64 = Test-Path (Join-Path 'HKLM:\SOFTWARE\Classes' $clsKey)
$reg32 = Test-Path (Join-Path 'HKLM:\SOFTWARE\Classes\WOW6432Node' $clsKey)
Check 'fixture DLL is unregistered' ((-not $reg64) -and (-not $reg32)) "reg64=$reg64 reg32=$reg32"

function New-CaseDir($name) {
    $d = Join-Path $base $name
    New-Item -ItemType Directory -Force $d | Out-Null
    Copy-Item $srcDll (Join-Path $d 'test_activex_dll.dll') -Force
    return $d
}

# Module sources are generated into the (git-ignored, corpus-excluded) output dir
# instead of living in tests/*.bas: every .bas under tests/ is a registered input
# of the emit-manifest shape oracle, and new fixture modules would red-gate it as
# "added rows" without a CI-produced hash.
$ModuleCreate = @'
Attribute VB_Name = "Module1"
Option Explicit

Sub Main()
    Dim d As Object
    On Error Resume Next
    Set d = CreateObject("TestAXDLL.Calc")
    If Err.Number = 0 Then
        Debug.Print "manifest-create=OK"
    Else
        Debug.Print "manifest-create=FAIL"
    End If
End Sub
'@

$ModuleDirect = @'
Attribute VB_Name = "ModuleDirect"
Option Explicit

Sub Main()
    Dim d As Object
    On Error Resume Next
    Set d = CreateObject("TestAXDLL.Calc")
    If Err.Number = 0 Then
        Debug.Print "direct-create=OK"
    Else
        Debug.Print "direct-create=FAIL"
    End If
End Sub
'@

$ModuleExt = @'
Attribute VB_Name = "ModuleExt"
Option Explicit

' External per-file manifest reuse: the ProgID below exists ONLY in the
' <component>.manifest block that the runtime composer pulls in. It is not in
' the ComLib= table, so this call must go through SxS or fail.
Sub Main()
    Dim d As Object
    On Error Resume Next
    Set d = CreateObject("ExtManifest.Calc")
    If Err.Number = 0 Then
        Debug.Print "ext-manifest-create=OK"
    Else
        Debug.Print "ext-manifest-create=FAIL"
    End If
End Sub
'@

function Write-Module($dir, $module) {
    $src = switch ($module) {
        'Module1.bas'      { $ModuleCreate }
        'ModuleDirect.bas' { $ModuleDirect }
        'ModuleExt.bas'    { $ModuleExt }
        default            { $null }
    }
    if (-not $src) { throw "unknown module $module" }
    Set-Content -Path (Join-Path $dir $module) -Encoding ASCII -Value $src
}

function Setup-Project($dir, $module, $exeName, $comLib) {
    Write-Module $dir $module
    $vbp = Join-Path $dir 'Case.vbp'
    $lines = @('Type=Exe', ("Module={0}; {0}" -f [System.IO.Path]::GetFileNameWithoutExtension($module)),
               'Startup="Sub Main"', ('ExeName32="{0}"' -f $exeName))
    if ($comLib) { $lines += ("ComLib={0}" -f $comLib) }
    Set-Content -Path $vbp -Encoding ASCII -Value $lines
    return $vbp
}

function New-ManifestCopy($src, $dir) {
    $dst = Join-Path $dir ("client-{0}.manifest" -f $runTag)
    Copy-Item $src $dst -Force
    return $dst
}

# --- A: explicit client manifest, <file> + <comClass>, no ComLib ---------------
$a = New-CaseDir 'a'
$aVbp = Setup-Project $a 'Module1.bas' 'ComA.exe' $null
$aMan = New-ManifestCopy (Join-Path $PSScriptRoot 'client-file-comclass.manifest') $a
& $C3 $aVbp --output-dir $a -ComMode manifest -ComManifest $aMan | Out-Null
$ax = Join-Path $a 'ComA.exe'
Check 'A build (manifest)' (Test-Path $ax) 'no output'
if (Test-Path $ax) {
    $o = & $ax 2>&1 | Out-String
    Check 'A manifest positive: CreateObject succeeds unregistered' ($LASTEXITCODE -eq 0 -and $o -match 'manifest-create=OK') "rc=$LASTEXITCODE out=$o"
}

# --- A': same project, direct mode must fail (control for A) -----------------
$a2 = Join-Path $base 'a-direct'
New-Item -ItemType Directory -Force $a2 | Out-Null
Copy-Item $srcDll (Join-Path $a2 'test_activex_dll.dll') -Force
$a2Vbp = Setup-Project $a2 'Module1.bas' 'ComA2.exe' $null
& $C3 $a2Vbp --output-dir $a2 -ComMode direct | Out-Null
$a2x = Join-Path $a2 'ComA2.exe'
if (Test-Path $a2x) {
    $o = & $a2x 2>&1 | Out-String
    Check 'A direct control: unregistered ProgID fails' ($o -match 'manifest-create=FAIL') "out=$o"
}

# --- B: external <component>.manifest reused by the runtime composer ---------
$b = New-CaseDir 'b'
$bPerFile = Join-Path $b 'test_activex_dll.dll.manifest'
@"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
    <file name="test_activex_dll.dll" hashalg="SHA1">
        <comClass clsid="{D84F362F-8EF1-D16D-8814-C16ADB700BAB}" threadingModel="Apartment" progid="ExtManifest.Calc"></comClass>
    </file>
</assembly>
"@ | Set-Content -Path $bPerFile -Encoding ASCII
$bDll = Join-Path $b 'test_activex_dll.dll'
$bVbp = Setup-Project $b 'ModuleExt.bas' 'ComB.exe' $bDll
& $C3 $bVbp --output-dir $b -ComMode manifest | Out-Null
$bx = Join-Path $b 'ComB.exe'
Check 'B build (manifest, external per-file manifest)' (Test-Path $bx) 'no output'
if (Test-Path $bx) {
    $env:C3_COM_TRACE = '1'
    $o = & $bx 2>&1 | Out-String
    Remove-Item Env:C3_COM_TRACE -ErrorAction SilentlyContinue
    Check 'B external <file>.manifest consumed: ExtManifest.Calc created' ($LASTEXITCODE -eq 0 -and $o -match 'ext-manifest-create=OK') "rc=$LASTEXITCODE out=$o"
    Check 'B ActCtx path is reported' ($o -match 'using manifest \(ActCtx activated\)') "out=$o"
}

# --- B': direct mode must fail (control for B) -------------------------------
$b2 = New-CaseDir 'b-direct'
Copy-Item $bPerFile (Join-Path $b2 'test_activex_dll.dll.manifest') -Force
$b2Dll = Join-Path $b2 'test_activex_dll.dll'
$b2Vbp = Setup-Project $b2 'ModuleExt.bas' 'ComB2.exe' $b2Dll
& $C3 $b2Vbp --output-dir $b2 -ComMode direct | Out-Null
$b2x = Join-Path $b2 'ComB2.exe'
if (Test-Path $b2x) {
    $o = & $b2x 2>&1 | Out-String
    Check 'B direct control: ExtManifest.Calc fails' ($o -match 'ext-manifest-create=FAIL') "out=$o"
}

# --- C: synthesized manifest (no external <component>.manifest) --------------
$c = New-CaseDir 'c'
$cDll = Join-Path $c 'test_activex_dll.dll'
$cVbp = Setup-Project $c 'Module1.bas' 'ComC.exe' $cDll
& $C3 $cVbp --output-dir $c -ComMode manifest | Out-Null
$cx = Join-Path $c 'ComC.exe'
Check 'C build (manifest, synthesized manifest)' (Test-Path $cx) 'no output'
if (Test-Path $cx) {
    $env:C3_COM_PREFER_REG = '1'
    $env:C3_COM_TRACE = '1'
    $o = & $cx 2>&1 | Out-String
    Remove-Item Env:C3_COM_PREFER_REG -ErrorAction SilentlyContinue
    Remove-Item Env:C3_COM_TRACE -ErrorAction SilentlyContinue
    Check 'C synthesized comClass resolves via SxS (registry path forced)' ($LASTEXITCODE -eq 0 -and $o -match 'manifest-create=OK') "rc=$LASTEXITCODE out=$o"
    Check 'C ActCtx path is reported' ($o -match 'using manifest \(ActCtx activated\)') "out=$o"
}

# --- C': direct + forced registry must fail (control for C) ------------------
$c2 = Join-Path $base 'c-direct'
New-Item -ItemType Directory -Force $c2 | Out-Null
Copy-Item $srcDll (Join-Path $c2 'test_activex_dll.dll') -Force
$c2Dll = Join-Path $c2 'test_activex_dll.dll'
$c2Vbp = Setup-Project $c2 'Module1.bas' 'ComC2.exe' $c2Dll
& $C3 $c2Vbp --output-dir $c2 -ComMode direct | Out-Null
$c2x = Join-Path $c2 'ComC2.exe'
if (Test-Path $c2x) {
    $env:C3_COM_PREFER_REG = '1'
    $o = & $c2x 2>&1 | Out-String
    Remove-Item Env:C3_COM_PREFER_REG -ErrorAction SilentlyContinue
    Check 'C direct control: unregistered ProgID fails' ($o -match 'manifest-create=FAIL') "out=$o"
}

# --- D: reverse case - identity version mismatch ----------------------------
# Private assembly manifest (embedded in the DLL) says 1.0.0.0; the client asks
# for 2.0.0.0. SxS must refuse to build the activation context; manifest mode
# must name that instead of silently falling back.
$d = New-CaseDir 'd'
$dDll = Join-Path $d 'test_activex_dll.dll'
if ($mt) {
    & $mt '-manifest' (Join-Path $PSScriptRoot 'assembly.manifest') "-outputresource:$dDll;#2" 2>&1 | Out-Null
}
Copy-Item (Join-Path $PSScriptRoot 'assembly.manifest') (Join-Path $d 'C3.FixtureCom.manifest') -Force
$dVbp = Setup-Project $d 'Module1.bas' 'ComD.exe' $dDll
$dMan = New-ManifestCopy (Join-Path $PSScriptRoot 'client-rev-version.manifest') $d
& $C3 $dVbp --output-dir $d -ComMode manifest -ComManifest $dMan | Out-Null
$dx = Join-Path $d 'ComD.exe'
Check 'D build (manifest, version mismatch)' (Test-Path $dx) 'no output'
if (Test-Path $dx) {
    $o = & $dx 2>&1 | Out-String
    $rc = $LASTEXITCODE
    Check 'D identity version mismatch fails loudly' ($rc -ne 0 -and $o -match 'manifest mode FAILED' -and $o -match 'error 14') "rc=$rc out=$o"
    Check 'D no silent fallback' ($o -match 'does NOT fall back') "out=$o"
}

# --- E: direct regression ---------------------------------------------------
$e = New-CaseDir 'e'
$eDll = Join-Path $e 'test_activex_dll.dll'
$eVbp = Setup-Project $e 'ModuleDirect.bas' 'ComE.exe' $eDll
& $C3 $eVbp --output-dir $e -ComMode direct | Out-Null
$ex = Join-Path $e 'ComE.exe'
Check 'E build (direct)' (Test-Path $ex) 'no output'
if (Test-Path $ex) {
    $o = & $ex 2>&1 | Out-String
    Check 'E direct regression: direct-load CreateObject succeeds' ($LASTEXITCODE -eq 0 -and $o -match 'direct-create=OK') "rc=$LASTEXITCODE out=$o"
}

# --- F: unusable client manifest must be named ------------------------------
$f = New-CaseDir 'f'
$fVbp = Setup-Project $f 'Module1.bas' 'ComF.exe' $null
$fMan = New-ManifestCopy (Join-Path $PSScriptRoot 'client-no-dependency.manifest') $f
& $C3 $fVbp --output-dir $f -ComMode manifest -ComManifest $fMan | Out-Null
$fx = Join-Path $f 'ComF.exe'
if (Test-Path $fx) {
    $o = & $fx 2>&1 | Out-String
    $rc = $LASTEXITCODE
    Check 'F unusable manifest fails naming the missing part' ($rc -ne 0 -and $o -match 'manifest mode FAILED' -and $o -match 'dependentAssembly') "rc=$rc out=$o"
}

if ($script:fail -ne 0) { Write-Host "$($script:fail) check(s) failed" -ForegroundColor Red; exit 1 }
Write-Host 'com_manifest: all passed' -ForegroundColor Green
exit 0
