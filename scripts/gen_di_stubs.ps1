# gen_di_stubs.ps1 - generate Win32 forwarding stubs for `Declare` symbols, one file per Lib family.
#
# Why: C3 maps every `Declare ... Lib "x"` to `extern <ret> __stdcall vb6_di_<alias>(...)`
# (Fix 076) and the RTL must provide the implementation. Hand-writing them one by one is
# what kept the link phase red, so this tool derives each stub from C3's own generated
# header prototype (which is the ABI the caller actually uses) and forwards 1:1.
#
# Family-aware + re-runnable (2026-09-17):
#   C3 now emits `/* vb6_di_lib: <lib> */` immediately before every generated `vb6_di_*`
#   prototype, so this tool groups the stubs by DLL family and writes one .c per family
#   (vb6_di_<family>_stubs.c). It no longer needs the external unresolved-symbol list
#   (.temp/unresolved_syms.txt, long gone) -- the input is the generated header set
#   itself, so a re-run only needs a session dir. Symbols already implemented by hand in
#   vb6_di_stubs.c are detected and skipped, and stale generated files are removed.
#
# Usage:
#   powershell -File scripts\gen_di_stubs.ps1 [-SessionDir <dir>] [-OutDir <dir>] [-HandFile <c>]
#   -SessionDir defaults to the newest %TEMP%\C3C\<id> (where C3 puts generated .c/.h)
#   -OutDir     defaults to src\rtl\core\di
#   -HandFile   defaults to <OutDir>\vb6_di_stubs.c (hand-written stubs, skipped here)
#
# ASCII-only file (PowerShell 5.1 reads .ps1 as ANSI without BOM).
param(
    [string]$SessionDir = '',
    [string]$OutDir = '',
    [string]$HandFile = ''
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$utf8 = [System.Text.Encoding]::UTF8

if ($SessionDir -eq '') {
    $base = Join-Path $env:LOCALAPPDATA 'Temp\C3C'
    if (-not (Test-Path $base)) { $base = Join-Path $env:TEMP 'C3C' }
    $SessionDir = (Get-ChildItem $base -Directory | Sort-Object LastWriteTime -Descending |
        Select-Object -First 1).FullName
}
if ($OutDir -eq '') { $OutDir = Join-Path $root 'src\rtl\core\di' }
if ($HandFile -eq '') { $HandFile = Join-Path $OutDir 'vb6_di_stubs.c' }

# ---------- 0. family table: one output file vb6_di_<name>_stubs.c per family ----------
# Grouping follows the DLL the VB6 `Declare` actually targets. Libs that are too small to
# carry a file of their own are merged into the closest domain family (documented above).
$families = @(
    @{ name = 'win32';   libs = @('kernel32', 'winmm') },
    @{ name = 'user32';  libs = @('user32', 'gdi32', 'dwmapi', 'comctl32', 'uxtheme') },
    @{ name = 'gdiplus'; libs = @('gdiplus') },
    @{ name = 'crypto';  libs = @('crypt32', 'bcrypt', 'ncrypt') },
    @{ name = 'com';     libs = @('ole32', 'oleaut32', 'advapi32', 'comdlg32') },
    @{ name = 'net';     libs = @('ws2_32', 'iphlpapi') },
    @{ name = 'shell';   libs = @('shell32', 'shlwapi', 'imagehlp') }
)
$libToFamily = @{}
foreach ($fam in $families) { foreach ($l in $fam.libs) { $libToFamily[$l] = $fam.name } }

# ---------- 0b. gdiplus 二级分组 (2026-09-20) ----------
# gdiplus 家族单文件已达 110 桩/828 行, 按对象域拆 5 份输出到子目录 OutDir\gdiplus\。
# 路由规则与对旧 vb6_di_gdiplus_stubs.c 的一次性拆分保持一致, 顺序敏感:
# 先 text/image/brush (PathGradient 在此吃掉), 再 Clip (拦截 SetClipPath), 再 Path, 兜底 draw。
$gdiplusSubs = @('draw', 'text', 'image', 'brush', 'path')
function Get-GdiplusSub([string]$api) {
    if ($api -match 'Font|String')                { return 'text' }
    if ($api -match 'Image|Bitmap|Effect')        { return 'image' }
    if ($api -match 'Brush|SolidFill|LineBrush|Texture|PathGradient|Pen') { return 'brush' }
    if ($api -match 'Clip')                       { return 'draw' }
    if ($api -match 'Path|PointCount')            { return 'path' }
    return 'draw'
}

Write-Output "session: $SessionDir"
Write-Output "outdir : $OutDir"

# ---------- 1. index prototypes (+ their Lib family) from generated headers ----------
$protos = @{}
$order = New-Object System.Collections.Generic.List[string]
$noLib = New-Object System.Collections.Generic.List[string]
$headers = Get-ChildItem $SessionDir -File | Where-Object { $_.Extension -eq '.h' }
foreach ($f in $headers) {
    $ls = [System.IO.File]::ReadAllLines($f.FullName, $utf8)
    $curLib = ''
    for ($i = 0; $i -lt $ls.Count; $i++) {
        $l = $ls[$i]
        if ($l -match 'vb6_di_lib:\s*([^\s*]*)\s*\*/') {
            $curLib = $matches[1].ToLower()
            if ($curLib.EndsWith('.dll')) { $curLib = $curLib.Substring(0, $curLib.Length - 4) }
            continue
        }
        if ($l -notmatch '^\s*extern\b') { continue }
        if ($l -notmatch 'vb6_di_') { continue }
        $buf = $l.Trim()
        $j = $i
        while ((($buf.ToCharArray() | Where-Object { $_ -eq '(' }).Count) -gt
               (($buf.ToCharArray() | Where-Object { $_ -eq ')' }).Count) -and $j + 1 -lt $ls.Count) {
            $j++
            $buf = $buf + ' ' + $ls[$j].Trim()
        }
        $m = [regex]::Match($buf, '^extern\s+(.+?)\s+(?:__stdcall\s+)?(vb6_di_[A-Za-z0-9_]+)\s*\(')
        if (-not $m.Success) { continue }
        $name = $m.Groups[2].Value
        if ($protos.ContainsKey($name)) { continue }
        $ret = $m.Groups[1].Value.Trim()
        $pOpen = $buf.IndexOf('(', $m.Index + $m.Length - 1)
        $pClose = $buf.LastIndexOf(')')
        if ($pOpen -lt 0 -or $pClose -le $pOpen) { continue }
        $params = $buf.Substring($pOpen + 1, $pClose - $pOpen - 1).Trim()
        if ($params -eq 'void') { $params = '' }
        if ($curLib -eq '') { $noLib.Add($name) }
        $protos[$name] = @{ ret = $ret; params = $params; lib = $curLib }
        $order.Add($name)
    }
}
Write-Output ("prototypes found: " + $protos.Count)

# ---------- 2. symbols already implemented by hand (must not be emitted again) ----------
$hand = @{}
if (Test-Path $HandFile) {
    foreach ($ln in [System.IO.File]::ReadAllLines($HandFile, $utf8)) {
        $m = [regex]::Match($ln, '\b(vb6_di_[A-Za-z0-9_]+)\s*\(')
        if ($m.Success) { $hand[$m.Groups[1].Value] = 1 }
    }
}
Write-Output ("hand-written symbols skipped: " + $hand.Count)

# ---------- 3. split parameter list on top-level commas ----------
function Split-Params([string]$p) {
    if ($p.Trim() -eq '') { return @() }
    $out = @()
    $depth = 0
    $cur = ''
    foreach ($ch in $p.ToCharArray()) {
        if ($ch -eq '(' -or $ch -eq '[') { $depth++ }
        if ($ch -eq ')' -or $ch -eq ']') { $depth-- }
        if ($ch -eq ',' -and $depth -eq 0) { $out += $cur; $cur = ''; continue }
        $cur += $ch
    }
    if ($cur.Trim() -ne '') { $out += $cur }
    return $out
}

# ---------- 4. emit one body list per family ----------
$bodies = @{}
$usedLibs = @{}
foreach ($fam in $families) {
    $bodies[$fam.name] = New-Object System.Collections.Generic.List[string]
    $usedLibs[$fam.name] = @{}
}
$bodies['unknown'] = New-Object System.Collections.Generic.List[string]
$usedLibs['unknown'] = @{}
foreach ($sub in $gdiplusSubs) { $bodies['gdiplus_' + $sub] = New-Object System.Collections.Generic.List[string] }
$skipped = New-Object System.Collections.Generic.List[string]
$noProto = New-Object System.Collections.Generic.List[string]
$badUdt = New-Object System.Collections.Generic.List[string]
$seen = @{}
$counts = @{}

foreach ($s in $order) {
    if ($seen.ContainsKey($s)) { continue }
    $seen[$s] = 1
    $api = $s.Substring(7)
    if ($api -like 'ord_*') { $skipped.Add($s + '  (ordinal import)'); continue }
    if ($hand.ContainsKey($s)) { $skipped.Add($s + '  (hand-written in vb6_di_stubs.c)'); continue }

    $famName = $libToFamily[$protos[$s].lib]
    if (-not $famName) { $famName = 'unknown' }
    # gdiplus 桩落到 gdiplus_<sub> 桶; usedLibs 仍按家族级收集
    $bodyKey = $famName
    if ($famName -eq 'gdiplus') { $bodyKey = 'gdiplus_' + (Get-GdiplusSub $api) }
    if (-not $counts.ContainsKey($bodyKey)) { $counts[$bodyKey] = 0 }

    $ret = $protos[$s].ret
    $paramList = Split-Params $protos[$s].params
    $declParts = @()
    $typeParts = @()
    $nameParts = @()
    $n = 0
    $udtBad = $false
    foreach ($p in $paramList) {
        $pt = $p.Trim()
        $n++
        $mm = [regex]::Match($pt, '([A-Za-z_][A-Za-z0-9_]*)\s*$')
        if ($mm.Success -and $mm.Index -gt 0) {
            $pname = $mm.Groups[1].Value
            $ptype = $pt.Substring(0, $mm.Index).Trim()
        } else {
            $pname = 'p' + $n
            $ptype = $pt
        }
        if ($ptype -eq '') { $ptype = 'intptr_t' }
        # VB UDTs (vb6_type_xxx) are project-generated structs the RTL cannot see.
        # Every occurrence here is by-pointer, so void* is ABI-identical.
        if ($ptype -match 'vb6_type_[A-Za-z0-9_]*\s*\*') {
            $ptype = [regex]::Replace($ptype, 'vb6_type_[A-Za-z0-9_]*\s*\*', 'void*')
        } elseif ($ptype -match 'vb6_type_[A-Za-z0-9_]*') {
            $udtBad = $true
        }
        $declParts += ($ptype + ' ' + $pname)
        $typeParts += $ptype
        $nameParts += $pname
    }
    if ($udtBad) { $badUdt.Add($s + '  (by-value VB UDT parameter)'); continue }
    $decl = ($declParts -join ', ')
    $types = ($typeParts -join ', ')
    if ($types -eq '') { $types = 'void' }
    $args = ($nameParts -join ', ')

    # gdiplus: no C header in the SDK, so resolve the flat API lazily by name.
    $dyn = ($api -match '^(Gdip|Gdiplus)')
    if ($dyn) { $usedLibs[$famName]['__dynamic__'] = 1 }
    if ($protos[$s].lib -ne '') { $usedLibs[$famName][$protos[$s].lib] = 1 }

    $out = $bodies[$bodyKey]
    $out.Add('/* ' + ($s -replace '^vb6_di_', '') + ' */')
    if ($dyn) {
        if ($ret -eq 'void') {
            $out.Add('void __stdcall ' + $s + '(' + $decl + ') {')
            $out.Add('    void (WINAPI *fn)(' + $types + ') = (void (WINAPI *)(' + $types + '))vb6_di_gdiplus_proc("' + $api + '");')
            $out.Add('    if (fn != NULL) { fn(' + $args + '); }')
            $out.Add('}')
        } else {
            $out.Add($ret + ' __stdcall ' + $s + '(' + $decl + ') {')
            $out.Add('    ' + $ret + ' (WINAPI *fn)(' + $types + ') = (' + $ret + ' (WINAPI *)(' + $types + '))vb6_di_gdiplus_proc("' + $api + '");')
            $out.Add('    if (fn == NULL) { return (' + $ret + ')2; /* GpStatus InvalidParameter */ }')
            $out.Add('    return fn(' + $args + ');')
            $out.Add('}')
        }
    } elseif ($ret -eq 'void') {
        $out.Add('void __stdcall ' + $s + '(' + $decl + ') {')
        $out.Add('    ((void (WINAPI *)(' + $types + '))' + $api + ')(' + $args + ');')
        $out.Add('}')
    } else {
        $out.Add($ret + ' __stdcall ' + $s + '(' + $decl + ') {')
        $out.Add('    return ((' + $ret + ' (WINAPI *)(' + $types + '))' + $api + ')(' + $args + ');')
        $out.Add('}')
    }
    $out.Add('')
    $counts[$bodyKey]++
}

# ---------- 5. write one file per family ----------
function New-Banner([string]$family, [string[]]$libs, [bool]$needDynamic, [int]$stubs) {
    $b = @()
    $b += ('// vb6_di_' + $family + '_stubs.c - generated by scripts/gen_di_stubs.ps1')
    $b += '//'
    $b += '// Win32 forwarding stubs for VB6 `Declare ... Lib "x"` (Fix 076 scheme: C3 emits'
    $b += '// `extern <ret> __stdcall vb6_di_<alias>(...)` and the RTL implements it).'
    $b += '//'
    $b += ('// Family: ' + $family + '   (libs: ' + ($libs -join ', ') + ')   stubs: ' + $stubs)
    $b += '//'
    $b += '// Each stub reproduces C3''s own generated prototype verbatim (that is the ABI the'
    $b += '// caller uses: ByVal Long is widened to intptr_t, ByVal Single stays float, ByRef'
    $b += '// Long stays int32_t*) and forwards 1:1 to the real API through a function-pointer'
    $b += '// cast. The cast keeps the compiler from complaining about unrelated API parameter'
    $b += '// types while preserving the register/memory passing class of every argument.'
    $b += '//'
    $b += '// generated from: a C3 compile session (pass -SessionDir to regenerate)'
    $b += ('// date: ' + (Get-Date -Format 'yyyy-MM-dd HH:mm'))
    $b += '//'
    $b += '// Hand-maintained special cases stay in vb6_di_stubs.c (ordinals, msvbvm60 runtime,'
    $b += '// dynamically loaded DLLs). Re-run the generator after a build exposes new symbols.'
    $b += ''
    $b += '/* VB6 code uses the classic winsock names (gethostbyname, inet_addr, ...) */'
    $b += '#define _WINSOCK_DEPRECATED_NO_WARNINGS'
    $b += '/* winsock2.h must come before windows.h */'
    $b += '#include <winsock2.h>'
    $b += '#include <windows.h>'
    # dwmapi family: windows.h does not declare Dwm*WindowAttribute (e.g. czUI)
    $b += '#include <dwmapi.h>'
    # Rtl*Memory are function-like macros in winnt.h; C3's cast-call stops macro
    # expansion, so drop the macros and declare the real kernel32 exports instead.
    $b += '#undef RtlMoveMemory'
    $b += '#undef RtlCopyMemory'
    $b += '#undef RtlFillMemory'
    $b += '#undef RtlZeroMemory'
    $b += 'void WINAPI RtlMoveMemory(void*, const void*, size_t);'
    $b += 'void WINAPI RtlCopyMemory(void*, const void*, size_t);'
    $b += 'void WINAPI RtlFillMemory(void*, size_t, unsigned char);'
    $b += 'void WINAPI RtlZeroMemory(void*, size_t);'
    $b += '#include <stdint.h>'
    $b += '#include <shlwapi.h>'
    $b += '#include <shlobj.h>'
    $b += '#include <mmsystem.h>'
    $b += '#include <dbghelp.h>'
    $b += '#include <iphlpapi.h>'
    $b += '#include <bcrypt.h>'
    $b += '#include <ncrypt.h>'
    $b += '#include <ole2.h>'
    $b += '#include <oleauto.h>'
    $b += '#include <olectl.h>'
    $b += '/* NOTE: gdiplus.h is C++-only (class Gdiplus...). The GDI+ flat API is'
    $b += ' * resolved with GetProcAddress instead - see below. */'
    $b += ''
    $b += '/* Import libs for the DLLs below. Adding one that is not needed is harmless:'
    $b += ' * a static import lib is only pulled when a symbol from it is referenced. */'
    foreach ($l in $libs) { $b += ('#pragma comment(lib, "' + $l + '.lib")') }
    # Fix 111b: ChooseColorA (Charts 2020 ppProgressCircular) needs comdlg32.lib;
    # without it any project not otherwise referencing comdlg32 gets LNK2019 __imp_ChooseColorA.
    $b += '#pragma comment(lib, "comdlg32.lib")'
    $b += ''
    if ($needDynamic) {
        $b += '/* GDI+ flat API lives in gdiplus.dll but its header is C++-only, so the'
        $b += ' * symbols below are resolved by name at first use. */'
        $b += 'static void* vb6_di_gdiplus_proc(const char* name) {'
        $b += '    static HMODULE mod = NULL;'
        $b += '    if (mod == NULL) { mod = LoadLibraryA("gdiplus.dll"); }'
        $b += '    return (mod != NULL) ? (void*)GetProcAddress(mod, name) : NULL;'
        $b += '}'
        $b += ''
    }
    return $b
}

# gdiplus 子文件 banner: 只需 windows.h/stdint.h (桩参数全是 C3 生成类型, 不需要 GDI+ 头),
# 每份自带 static vb6_di_gdiplus_proc (RTL 硬约束: 文件级 static 跨编译单元不可见)
function New-GdiplusBanner([string]$sub, [int]$stubs) {
    $b = @()
    $b += ('// vb6_di_gdiplus_' + $sub + '_stubs.c - generated by scripts/gen_di_stubs.ps1')
    $b += '//'
    $b += '// Win32 forwarding stubs for VB6 `Declare ... Lib "gdiplus"` (Fix 076 scheme).'
    $b += '//'
    $b += ('// Family: gdiplus/' + $sub + '   (routed by object domain, see Get-GdiplusSub)   stubs: ' + $stubs)
    $b += '//'
    $b += '// Each stub reproduces C3''s own generated prototype verbatim and forwards 1:1 to'
    $b += '// the real flat API through a function-pointer cast (gdiplus.h is C++-only, so the'
    $b += '// flat API is resolved by name at first use via the loader below).'
    $b += ('//')
    $b += ('// date: ' + (Get-Date -Format 'yyyy-MM-dd HH:mm'))
    $b += ''
    $b += '#include <windows.h>'
    $b += '#include <stdint.h>'
    $b += '#pragma comment(lib, "gdiplus.lib")'
    $b += ''
    $b += 'static void* vb6_di_gdiplus_proc(const char* name) {'
    $b += '    static HMODULE mod = NULL;'
    $b += '    if (mod == NULL) { mod = LoadLibraryA("gdiplus.dll"); }'
    $b += '    return (mod != NULL) ? (void*)GetProcAddress(mod, name) : NULL;'
    $b += '}'
    $b += ''
    return $b
}

# ---------- 2c. carry-over helpers (2026-09-20): regeneration is add-only ----------
# Wipe-on-regen bit us once: the 2026-09-19 czUI re-run replaced the accumulated
# 213-stub set with czUI's own subset, and the vbman DLL then failed on CI with ~90
# LNK2019 (its Declare stubs were gone with the uncommitted 09-18 state). A stub is a
# pure forwarder: an extra one is a few hundred bytes of dead code, a missing one is a
# link failure in any project that declares it. So family files are now merged:
# stubs present in the old file but not produced by the current session are carried
# over verbatim, appended after the fresh ones.
function Get-ExistingStubBlocks([string]$path) {
    $blocks = @{}
    if (-not (Test-Path $path)) { return $blocks }
    $lines = [System.IO.File]::ReadAllLines($path)
    $i = 0
    while ($i -lt $lines.Count) {
        # Locate the block by its DEFINITION line, not by the comment above it.
        # 2026-10-06: matching `^/\* <name> \*/$` only saw generator-emitted stubs (the
        # generator writes a lone `/* Api */` line, while hand-added restorations carry a
        # richer comment such as `/* ChooseFontA - Task #44 SSTabEx cDlg.cls:33 */`).
        # Those hand-added stubs were therefore invisible to the carry-over below and got
        # rewritten away on the next run - exactly the silent-drop this section exists to
        # prevent. Keying on `__stdcall vb6_di_X(` covers both shapes.
        if ($lines[$i] -match '^\s*[A-Za-z_][A-Za-z0-9_]*\s+__stdcall\s+(vb6_di_[A-Za-z0-9_]+)\s*\(') {
            $nm = $Matches[1]
            $end = $i
            while ($end -lt $lines.Count -and $lines[$end] -ne '}') { $end++ }
            if ($end -lt $lines.Count) {
                # Pull in the contiguous comment block directly above (ASCII only), so a
                # carried-over stub keeps its rationale. Stops at a blank line or at any
                # line that is not part of a comment.
                $start = $i
                while ($start -gt 0) {
                    $prev = $lines[$start - 1].Trim()
                    if ($prev -eq '') { break }
                    if ($prev -match '^\*' -or $prev -match '^//' -or $prev -match '^/\*' -or $prev -match '\*/$') {
                        $start--
                    } else { break }
                }
                $blocks[$nm] = $lines[$start..$end]
                $i = $end + 1
                continue
            }
        }
        $i++
    }
    return $blocks
}

function Get-StubNamesFromLines([string[]]$lines) {
    $names = @{}
    foreach ($ln in $lines) {
        if ($ln -match '__stdcall vb6_di_([A-Za-z0-9_]+)\(') { $names[$Matches[1]] = $true }
    }
    return $names
}

# names produced by the current session across ALL families (+ hand file): a carried
# stub must not duplicate any of them (two definitions of one symbol = LNK2005)
$freshAll = @{}
foreach ($k in $bodies.Keys) {
    $freshNames = Get-StubNamesFromLines $bodies[$k]
    foreach ($nm in $freshNames.Keys) { $freshAll[$nm] = $true }
}
foreach ($nm in $hand.Keys) { $freshAll[$nm] = $true }

if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }

$written = New-Object System.Collections.Generic.List[string]
foreach ($fam in $families) {
    $name = $fam.name
    if ($bodies[$name].Count -eq 0) {
        # gdiplus 主体不再直接输出, 走下面的子族分支
        if ($name -ne 'gdiplus') { continue }
    }
    if ($name -eq 'gdiplus') {
        $gdir = Join-Path $OutDir 'gdiplus'
        if (-not (Test-Path $gdir)) { New-Item -ItemType Directory -Path $gdir | Out-Null }
        foreach ($sub in $gdiplusSubs) {
            $key = 'gdiplus_' + $sub
            if ($bodies[$key].Count -eq 0) { continue }
            $path = Join-Path $gdir ('vb6_di_' + $key + '_stubs.c')
            $carried = @()
            $oldBlocks = Get-ExistingStubBlocks $path
            foreach ($nm in ($oldBlocks.Keys | Sort-Object)) {
                if (-not $freshAll.ContainsKey($nm)) { $carried += $oldBlocks[$nm] }
            }
            if ($carried.Count -gt 0) {
                Write-Output ('  ' + $key.PadRight(14) + '    carried over ' + $carried.Count + ' stub(s) from previous file')
            }
            $content = ((New-GdiplusBanner $sub ($counts[$key] + $carried.Count)) + $bodies[$key] + $carried) -join "`r`n"
            [System.IO.File]::WriteAllText($path, $content, (New-Object System.Text.UTF8Encoding($false)))
            $written.Add($path)
            Write-Output ('  ' + $key.PadRight(14) + ' -> gdiplus\' + (Split-Path $path -Leaf) +
                          '   stubs ' + $counts[$key])
        }
        continue
    }
    $libs = @()
    # czUI fix: 始终包含全族 lib (会话没引用的 lib 里的 API 也可能被 stub 直接调用,
    # 如 dwmapi 的 DwmSetWindowAttribute — 按需裁剪会漏 lib 导致 LNK2019)
    foreach ($l in $fam.libs) { $libs += $l }
    $needDynamic = $usedLibs[$name].ContainsKey('__dynamic__')
    $path = Join-Path $OutDir ('vb6_di_' + $name + '_stubs.c')
    $carried = @()
    $oldBlocks = Get-ExistingStubBlocks $path
    foreach ($nm in ($oldBlocks.Keys | Sort-Object)) {
        if (-not $freshAll.ContainsKey($nm)) { $carried += $oldBlocks[$nm] }
    }
    if ($carried.Count -gt 0) {
        Write-Output ('  ' + $name.PadRight(9) + '    carried over ' + $carried.Count + ' stub(s) from previous file')
    }
    $content = ((New-Banner $name $libs $needDynamic ($counts[$name] + $carried.Count)) + $bodies[$name] + $carried) -join "`r`n"
    [System.IO.File]::WriteAllText($path, $content, (New-Object System.Text.UTF8Encoding($false)))
    $written.Add($path)
    Write-Output ('  ' + $name.PadRight(9) + ' -> ' + (Split-Path $path -Leaf) +
                  '   stubs ' + $counts[$name] + '   libs ' + ($libs -join ','))
}

if ($bodies['unknown'].Count -gt 0) {
    $path = Join-Path $OutDir 'vb6_di_unknown_stubs.c'
    $carried = @()
    $oldBlocks = Get-ExistingStubBlocks $path
    foreach ($nm in ($oldBlocks.Keys | Sort-Object)) {
        if (-not $freshAll.ContainsKey($nm)) { $carried += $oldBlocks[$nm] }
    }
    $content = ((New-Banner 'unknown' @() $false ($counts['unknown'] + $carried.Count)) + $bodies['unknown'] + $carried) -join "`r`n"
    [System.IO.File]::WriteAllText($path, $content, (New-Object System.Text.UTF8Encoding($false)))
    $written.Add($path)
    Write-Output ('  unknown   -> vb6_di_unknown_stubs.c   stubs ' + $counts['unknown'] + '   !!! no vb6_di_lib marker')
}

# ---------- 6. (retired 2026-09-20) stale-file removal ----------
# This section used to delete any vb6_di_*_stubs.c not produced by the current run.
# Together with whole-file rewrites that made every re-run able to silently shrink the
# stub set (see carry-over note above). Removal is retired: a family file left over
# from an earlier run keeps its stubs and merges on the next run that produces it.
# If a file truly needs to go away, delete it by hand with git rm.

Write-Output ("files written: " + $written.Count)
Write-Output ''
Write-Output '=== skipped (hand-written or ordinal) ==='
$skipped | ForEach-Object { Write-Output ('  ' + $_) }
Write-Output ''
Write-Output '=== vb6_di_* prototype without a vb6_di_lib marker ==='
$noLib | ForEach-Object { Write-Output ('  ' + $_) }
Write-Output ''
Write-Output '=== skipped: by-value VB UDT parameter ==='
$badUdt | ForEach-Object { Write-Output ('  ' + $_) }
