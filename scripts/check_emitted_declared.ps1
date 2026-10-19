# ============================================================
# check_emitted_declared.ps1 —— 账 #305 / §B202 / 第九十九刀 (2026-10-17)
#
# 一条结构性规则: 内置表点名的每一枚 vb6_*（会以**字面量**发进用户 TU）
# 必须在某个 RTL 头文件里有声明。缺声明时 C 只报一声忠告(C4013)、接着把它当「返回 int」
# 的隐式声明 —— 发码形状一个字节都没错，所以形状门永远看不见这一格；坏的是运行期:
# x64 把返回的 BSTR 指针折成 32 位再交 vb6_BSTR_AssignMove ⇒ 解引用即 AV(0xC0000005)，
# x86 指针本来 32 位 ⇒ 看着全对（五档答案还与本机 VBScript 的 FormatDateTime 逐档相同）。
# 实物 = vb6_FormatDateTime（尺 .build/k423_proj -> k424_run.txt / k426_new_x86_run.txt）；本哨兵就是
# 那把 census（.build/k428_emit_gap.py）的 PowerShell 版。
#
# 口径边界（改规则先改这段话）:
#   - 只要**字面量**出现的名字。靠拼接现拼的那一族（例如 vb6_arr_<名>）不在名单里。
#   - 名字必须在 src/rtl 的 .c/.inc 里**真有定义**才参与比对（宏与纯声明不算）。
#   - 后端文件的**注释**也会被这个扫描看见（按字面量取名，不分语法）⇒ 它只会把名单变大，不会漏报。
#   - R2 的三道下限是监「扫描变盲」的: 目录改名或正则写断时 gap 会因为三张名单都空而看看像
#     零缺口，下限把那种情形监成红。
#
# 判据面: tests/b429 两台各钉五档之外还钉一枚 VarType —— 缺声明时答 3(VT_I4)、有了才答 8，
# 所以这一格在 x86 上也会移，不只靠「x64 能跑到 DONE」当证人。
# 注: 本文件一个反斜杠都不写——所有正则里皀转义符都由 [char] 现拼，针对的是「工具层先把反斜杠折掉」这一类坑。
# ============================================================

param(
    [string]$RepoRoot = (Split-Path -Parent $PSScriptRoot)
)

$ErrorActionPreference = 'Stop'
$fail = 0
$B = [char]92   # 反斜杠
$Q = [char]39   # 单引号
$D = [char]34   # 双引号

$sep = [IO.Path]::DirectorySeparatorChar
$backendDir = Join-Path $RepoRoot ('src' + $sep + 'backend')
$rtlDir = Join-Path $RepoRoot ('src' + $sep + 'rtl')
$headerRel = 'src/rtl/core/vb6rtl/vb6rtl_builtin.h'
$systemRel = 'src/rtl/core/vb6rtl/vb6rtl_system.c'
$fixupRel = 'src/backend/detail/expr/cgen_expr_call_builtin_fixup.inc'
$runnerRel = 'tests/run_tests.ps1'

foreach ($d0 in @($backendDir, $rtlDir)) {
    if (-not (Test-Path $d0)) { Write-Output ('FAIL: missing ' + $d0); exit 1 }
}

# ---- 三张名单（R1 的底料） --------------------------------------------------------
$emitPat = $D + '(vb6_[A-Za-z0-9_]+)' + $D
$declPat = $B + 'b(vb6_[A-Za-z0-9_]+)' + $B + 's*' + $B + '('
$defnPat = '(?m)^(?:static' + $B + 's+)?[A-Za-z_][A-Za-z0-9_ ' + $B + 't' + $B + '*]*?' + $B + 'b(vb6_[A-Za-z0-9_]+)' + $B + 's*' + $B + '('

$emit = New-Object 'System.Collections.Generic.HashSet[string]'
$decl = New-Object 'System.Collections.Generic.HashSet[string]'
$defn = New-Object 'System.Collections.Generic.HashSet[string]'

foreach ($f in Get-ChildItem -Path $backendDir -Recurse -File) {
    $t = [System.IO.File]::ReadAllText($f.FullName)
    foreach ($m in [regex]::Matches($t, $emitPat)) { [void]$emit.Add($m.Groups[1].Value) }
}
foreach ($f in Get-ChildItem -Path $rtlDir -Recurse -File -Filter '*.h') {
    $t = [System.IO.File]::ReadAllText($f.FullName)
    foreach ($m in [regex]::Matches($t, $declPat)) { [void]$decl.Add($m.Groups[1].Value) }
}
foreach ($f in Get-ChildItem -Path $rtlDir -Recurse -File) {
    if ($f.Extension -ne '.c' -and $f.Extension -ne '.inc') { continue }
    $t = [System.IO.File]::ReadAllText($f.FullName)
    foreach ($m in [regex]::Matches($t, $defnPat)) { [void]$defn.Add($m.Groups[1].Value) }
}

$gap = New-Object 'System.Collections.Generic.List[string]'
foreach ($n in $emit) {
    if ($defn.Contains($n) -and -not $decl.Contains($n)) { [void]$gap.Add($n) }
}
$gap.Sort()

# R1: 这一问的唯一答案
if ($gap.Count -ne 0) {
    Write-Output ('FAIL R1: ' + $gap.Count + ' emitted name(s) have an RTL definition but no header declaration: ' + ($gap -join ', '))
    $fail++
}

# R2: 下限 —— 名单空了不许读成「零缺口」
if ($emit.Count -lt 800)  { Write-Output ('FAIL R2a: emitted-name floor is 800, saw ' + $emit.Count); $fail++ }
if ($defn.Count -lt 2400) { Write-Output ('FAIL R2b: RTL-definition floor is 2400, saw ' + $defn.Count); $fail++ }
if ($decl.Count -lt 1500) { Write-Output ('FAIL R2c: header-declaration floor is 1500, saw ' + $decl.Count); $fail++ }

# R3/R4/R5: 本刀那枚的声明与定义同形，且那个名字确实还被发码点名（可达）
$hdrText = [System.IO.File]::ReadAllText((Join-Path $RepoRoot $headerRel))
$sysText = [System.IO.File]::ReadAllText((Join-Path $RepoRoot $systemRel))
$fxText = [System.IO.File]::ReadAllText((Join-Path $RepoRoot $fixupRel))

$declLine = 'BSTR vb6_FormatDateTime(double dateSerial, int32_t namedFormat);'
$nDecl = ([regex]::Matches($hdrText, [regex]::Escape($declLine))).Count
if ($nDecl -ne 1) { Write-Output ('FAIL R3: want exactly 1 declaration of vb6_FormatDateTime in ' + $headerRel + ', saw ' + $nDecl); $fail++ }

$defLine = [regex]::Escape('BSTR vb6_FormatDateTime(double dateSerial, int32_t namedFormat) {')
$nDef = ([regex]::Matches($sysText, $defLine)).Count
if ($nDef -ne 1) { Write-Output ('FAIL R4: the definition in ' + $systemRel + ' no longer matches the declared shape, saw ' + $nDef); $fail++ }

$rowLine = [regex]::Escape('{ ' + $D + 'vb6_FormatDateTime' + $D + ',')
$nRow = ([regex]::Matches($fxText, $rowLine)).Count
if ($nRow -ne 1) { Write-Output ('FAIL R5: the arity row that emits vb6_FormatDateTime is gone (saw ' + $nRow + ') - if this census should stop covering it, delete the rule here instead of letting it pass silently'); $fail++ }

# R6..R9: 夹具真的进了门，而 x86 那一片是一整行形状
$runText = [System.IO.File]::ReadAllText((Join-Path $RepoRoot $runnerRel))
$nCell = ([regex]::Matches($runText, 'Test-Vbp ' + $Q + 'b429_formatdatetime')).Count
if ($nCell -ne 2) { Write-Output ('FAIL R6: want 2 Test-Vbp cells for b429 (x64 + x86), saw ' + $nCell); $fail++ }

$nTable = ([regex]::Matches($runText, [regex]::Escape('$b429Expected = @('))).Count
if ($nTable -ne 1) { Write-Output ('FAIL R7: the b429 expectation table is missing (saw ' + $nTable + ')'); $fail++ }

$x86Pat = 'Test-Vbp ' + $Q + 'b429_formatdatetime_x86' + $Q + '[^' + $B + 'rn]*-Arch ' + $D + 'x86' + $D + $B + 'r?$'
$nX86 = ([regex]::Matches($runText, $x86Pat, 'Multiline')).Count
if ($nX86 -ne 1) { Write-Output ('FAIL R8: the x86 cell for b429 is not one whole line ending in -Arch x86 (saw ' + $nX86 + ')'); $fail++ }

$badPat = '-Arch ' + $D + 'x86' + $D + $B + $B
$nBad = ([regex]::Matches($runText, $badPat)).Count
if ($nBad -ne 0) { Write-Output ('FAIL R9: -Arch x86 with a trailing backslash survived in run_tests.ps1 (saw ' + $nBad + ') - PowerShell hands that backslash to C3 as a literal'); $fail++ }

if ($fail -eq 0) {
    Write-Output ('PASS check_emitted_declared: emitted=' + $emit.Count + ' defined=' + $defn.Count + ' declared=' + $decl.Count + ' gap=0 decl=1 def=1 row=1 cells=2 table=1 x86shape=1')
    exit 0
}
Write-Output ('FAIL check_emitted_declared: ' + $fail + ' rule(s) red')
exit 1
