# ============================================================
# check_emitted_declared.ps1 —— 账 #305 / §B202 / 第九十九刀 (2026-10-17)，第二百六十四刀 (2026-10-28) 改问「闭包可达」
#
# 一条结构性规则: 内置表点名的每一枚 vb6_*（会以**字面量**发进用户 TU）
# 必须在**生成代码看得见的那批 RTL 头**里有声明。缺声明时 C 只报一声忠告(C4013)、接着把它当「返回 int」
# 的隐式声明 —— 发码形状一个字节都没错，所以形状门永远看不见这一格；坏的是运行期:
# x64 把返回的 BSTR / 集合指针折成 32 位再符号扩展 ⇒ 解引用即 AV(0xC0000005)，
# x86 指针本来 32 位 ⇒ 看着全对（五档答案还与本机 VBScript 的 FormatDateTime 逐档相同）。
# 实物两枚 = vb6_FormatDateTime（尺 .build/k423_proj -> k424_run.txt）与 vb6_UC_ParentControls
# （ComCtlsDemo 的野接收者：av read target=0xfffffffff8cc8958，低 32 位落在 .data 的单例那格）。
#
# 为什么要改这一道（第二百六十三刀「量」的读数）: 旧口径三处都不问对的问题，于是那一次真崩溃它一条没报 ——
#   (a) emit 只认「整枚字面量恰好是裸名」那形（"vb6_X"），"vb6_X()" 带括号尾巴的一律不进名单；
#   (b) decl 问的是「src/rtl 下**任意** .h 里有没有」，而 ParentControls 的声明住在内部头
#       vb6forms_uc_internal.h ⇒ 它算「有声明」，可用户 TU 经 vb6forms.h 伞头根本看不见那一份；
#   (c) 不剥注释 ⇒ 注释里提一嘴也算声明。
# 现在三张名单各问各的那一问（下面 R1..R2 的口径边界）。
#
# 口径边界（改规则先改这段话）:
#   - emit = 后端源（剥注释）里以双引号起头的 vb6_ 名 ⇒ 认 "vb6_X" 也认 "vb6_X()"、"vb6_X(" 拼接尾巴。
#     这一张名单**故意偏大**（拼名字的前缀片段也进来），因为 R1 还要它同时是 RTL 定义。
#   - decl = 从四枚伞头（vb6rtl.h / vb6com.h / vb6comserver.h / vb6forms.h）沿**引号 include** 走到的
#     那批头（剥注释）里出现过的 `名字 (`。内部头不在闭包内 ⇒ 闭包外的声明不算数，正是这一刀的靶子。
#     这一张名单**故意偏松**（头里一处调用也认）: 只许漏报「头里没有原型」那一格，由 R10..R12 三枚
#     按名 + 计数针补住（本仓那两枚实物都是「压根没声明」或「声明全在内部头」，两种都点名得动）。
#   - defn = RTL 的 .c/.inc（剥注释）里定义形状的名字（宏与纯声明不算）。
#   - R2 的四道下限是监「扫描变盲」的: 目录改名或正则写断时 gap 会因为名单都空而看着像
#     零缺口，下限把那种情形监成红；闭包那份下限监的是「伞头改名后闭包走空」这一种。
#   - R12 监的是「闭包这一问还分不分得开内部头」: 若哪天两份名单完全相同，要么全公共化了、
#     要么闭包走漏了 —— 两种都该重新看一眼，不许它静默变成旧口径。
#
# 判据面: tests/b429 两台各钉五档之外还钉一枚 VarType —— 缺声明时答 3(VT_I4)、有了才答 8，
# 所以这一格在 x86 上也会移，不只靠「x64 能跑到 DONE」当证人。
# 注: 本文件一个反斜杠都不写——所有正则里的转义符都由 [char] 现拼，针对的是「工具层先把反斜杠折掉」这一类坑。
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
$coreDir = Join-Path $rtlDir ('core' + $sep)
$headerRel = 'src/rtl/core/vb6rtl/vb6rtl_builtin.h'
$systemRel = 'src/rtl/core/vb6rtl/vb6rtl_system.c'
$fixupRel = 'src/backend/detail/expr/cgen_expr_call_builtin_fixup.inc'
$runnerRel = 'tests/run_tests.ps1'
$ctrlHdrRel = 'src/rtl/core/vb6forms/vb6forms_controls.h'
$internalRel = 'src/rtl/core/vb6forms/vb6forms_uc_internal.h'

foreach ($d0 in @($backendDir, $rtlDir)) {
    if (-not (Test-Path $d0)) { Write-Output ('FAIL: missing ' + $d0); exit 1 }
}

# ---- 读码 + 剥注释（三张名单共用同一把尺） ---------------------------------------
function Read-Text([string]$p) { return [System.IO.File]::ReadAllText($p) }
function Strip-Comments([string]$t) {
    $block = '/' + $B + '*' + '[' + $B + 's' + $B + 'S' + ']*?' + $B + '*' + $B + '/'
    $line = '//' + '[^' + $B + 'r' + $B + 'n' + ']' + '*'
    $u = [regex]::Replace($t, $block, ' ')
    return [regex]::Replace($u, $line, ' ')
}

# ---- 伞头闭包（decl 那一张名单的搜索域） ------------------------------------------
$umbrellas = @('core/vb6rtl/vb6rtl.h', 'core/vb6com/vb6com.h',
               'core/vb6comserver/vb6comserver.h', 'core/vb6forms/vb6forms.h')
foreach ($u in $umbrellas) {
    $p = Join-Path $rtlDir ($u -replace '/', $sep)
    if (-not (Test-Path -LiteralPath $p)) { Write-Output ('FAIL: missing umbrella ' + $u); exit 1 }
}
$subDirs = @(Get-ChildItem -Path $coreDir -Directory | ForEach-Object { $_.FullName })
$inclRe = [regex]::new($B + 's*#' + $B + 's*include' + $B + 's*' + $D + '([^' + $D + ']+)' + $D, 'Multiline')
$nameParenRe = [regex]::new($B + 'b(vb6_[A-Za-z0-9_]+)' + $B + 's*' + $B + '(')

$seen = New-Object 'System.Collections.Generic.HashSet[string]'
$closure = New-Object 'System.Collections.Generic.List[string]'
$stack = New-Object 'System.Collections.Generic.List[string]'
foreach ($u in $umbrellas) { $stack.Add((Resolve-Path (Join-Path $rtlDir ($u -replace '/', $sep))).Path) }
while ($stack.Count -gt 0) {
    $cur = $stack[$stack.Count - 1]; $stack.RemoveAt($stack.Count - 1)
    if ($seen.Contains($cur)) { continue }
    [void]$seen.Add($cur)
    $closure.Add($cur)
    $body = Strip-Comments (Read-Text $cur)
    foreach ($m in $inclRe.Matches($body)) {
        $inc = $m.Groups[1].Value
        $cands = New-Object 'System.Collections.Generic.List[string]'
        $cands.Add((Join-Path (Split-Path -Parent $cur) $inc))
        foreach ($s in $subDirs) { $cands.Add((Join-Path $s $inc)) }
        $cands.Add((Join-Path $rtlDir $inc))
        $hit = $cands | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
        if ($hit) { $stack.Add((Resolve-Path -LiteralPath $hit).Path) }
    }
}

$decl = New-Object 'System.Collections.Generic.HashSet[string]'
foreach ($p in $closure) {
    foreach ($m in $nameParenRe.Matches((Strip-Comments (Read-Text $p)))) { [void]$decl.Add($m.Groups[1].Value) }
}

# ---- 三张名单（R1 的底料） --------------------------------------------------------
$emitPat = $D + '(vb6_[A-Za-z0-9_]+)'
$defnPat = '(?m)^(?:static' + $B + 's+)?[A-Za-z_][A-Za-z0-9_ ' + $B + 't' + $B + '*]*?' + $B + 'b(vb6_[A-Za-z0-9_]+)' + $B + 's*' + $B + '('

$emit = New-Object 'System.Collections.Generic.HashSet[string]'
$defn = New-Object 'System.Collections.Generic.HashSet[string]'

foreach ($f in Get-ChildItem -Path $backendDir -Recurse -File) {
    $t = Strip-Comments (Read-Text $f.FullName)
    foreach ($m in [regex]::Matches($t, $emitPat)) { [void]$emit.Add($m.Groups[1].Value) }
}
foreach ($f in Get-ChildItem -Path $rtlDir -Recurse -File) {
    if ($f.Extension -ne '.c' -and $f.Extension -ne '.inc') { continue }
    $t = Strip-Comments (Read-Text $f.FullName)
    foreach ($m in [regex]::Matches($t, $defnPat)) { [void]$defn.Add($m.Groups[1].Value) }
}

# 旧口径那份（任意 .h、剥注释）只为 R12 的「闭包这一问还分不分得开」留底料
$declAll = New-Object 'System.Collections.Generic.HashSet[string]'
foreach ($f in Get-ChildItem -Path $rtlDir -Recurse -File -Filter '*.h') {
    foreach ($m in $nameParenRe.Matches((Strip-Comments (Read-Text $f.FullName)))) { [void]$declAll.Add($m.Groups[1].Value) }
}

$gap = New-Object 'System.Collections.Generic.List[string]'
foreach ($n in $emit) {
    if ($defn.Contains($n) -and -not $decl.Contains($n)) { [void]$gap.Add($n) }
}
$gap.Sort()

# R1: 这一问的唯一答案
if ($gap.Count -ne 0) {
    Write-Output ('FAIL R1: ' + $gap.Count + ' emitted name(s) have an RTL definition but no declaration visible from the umbrella headers: ' + ($gap -join ', '))
    $fail++
}

# R2: 下限 —— 名单空了不许读成「零缺口」
if ($closure.Count -lt 20) { Write-Output ('FAIL R2d: the umbrella closure must reach at least 20 headers, saw ' + $closure.Count); $fail++ }
if ($emit.Count -lt 1100)  { Write-Output ('FAIL R2a: emitted-name floor is 1100, saw ' + $emit.Count); $fail++ }
if ($defn.Count -lt 2400)  { Write-Output ('FAIL R2b: RTL-definition floor is 2400, saw ' + $defn.Count); $fail++ }
if ($decl.Count -lt 1400)  { Write-Output ('FAIL R2c: closure-visible declaration floor is 1400, saw ' + $decl.Count); $fail++ }

# R10/R11: 第二百六十四刀那两枚声明 —— 按名钉在**伞头可达**的那份公共头里，且整条闭包认得它们
# （缺声明时 cl 只报 C4013，x64 把返回的指针折成 32 位：这两枚一枚是崩溃实物、一枚是规则活样本）
$ctrlText = [System.IO.File]::ReadAllText((Join-Path $RepoRoot $ctrlHdrRel))
$nParentDecl = ([regex]::Matches($ctrlText, [regex]::Escape('void* vb6_UC_ParentControls(void);'))).Count
if ($nParentDecl -ne 1) { Write-Output ('FAIL R10: want exactly 1 declaration of vb6_UC_ParentControls in ' + $ctrlHdrRel + ' (the umbrella-visible one), saw ' + $nParentDecl); $fail++ }
if (-not $decl.Contains('vb6_UC_ParentControls')) { Write-Output ('FAIL R10: vb6_UC_ParentControls is not visible from the umbrella closure'); $fail++ }

$builtinText = [System.IO.File]::ReadAllText((Join-Path $RepoRoot $headerRel))
$nPrevDecl = ([regex]::Matches($builtinText, [regex]::Escape('int32_t vb6_App_PrevInstance(void);'))).Count
if ($nPrevDecl -ne 1) { Write-Output ('FAIL R11: want exactly 1 declaration of vb6_App_PrevInstance in ' + $headerRel + ', saw ' + $nPrevDecl); $fail++ }
if (-not $decl.Contains('vb6_App_PrevInstance')) { Write-Output ('FAIL R11: vb6_App_PrevInstance is not visible from the umbrella closure'); $fail++ }

# R12: 闭包这一问必须还分得开内部头 —— 两份名单相同 = 这一问退化成了旧口径
$internalOnly = 0
foreach ($n in $declAll) { if (-not $decl.Contains($n)) { $internalOnly++ } }
if ($internalOnly -lt 1) {
    Write-Output ('FAIL R12: no declaration is internal-header-only any more (closure=' + $decl.Count + ' of all=' + $declAll.Count + ') - either everything went public or the closure walk stopped distinguishing; the old blind oracle is back')
    $fail++
}

# R13: 内部头那份重复声明仍在（RTL 的 .c 只 include 它；公共头这一份是给生成代码看的）
$internalText = [System.IO.File]::ReadAllText((Join-Path $RepoRoot $internalRel))
$nInternal = ([regex]::Matches($internalText, [regex]::Escape('void* vb6_UC_ParentControls(void);'))).Count
if ($nInternal -ne 1) { Write-Output ('FAIL R13: the RTL-side declaration in ' + $internalRel + ' is gone (saw ' + $nInternal + ') - uc_controls.c then compiles without a prototype'); $fail++ }

# R3/R4/R5: 第九十九刀那枚的声明与定义同形，且那个名字确实还被发码点名（可达）
$sysText = [System.IO.File]::ReadAllText((Join-Path $RepoRoot $systemRel))
$fxText = [System.IO.File]::ReadAllText((Join-Path $RepoRoot $fixupRel))

$declLine = 'BSTR vb6_FormatDateTime(double dateSerial, int32_t namedFormat);'
$nDecl = ([regex]::Matches($builtinText, [regex]::Escape($declLine))).Count
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
if ($nX86 -ne 1) { Write-Output ('FAIL R8: the x86 cell for b429 is not one whole line ending in -Arch x86 (saw ' + $nX86 + ') - a regex escape built by hand is the fragile part here'); $fail++ }

$badPat = '-Arch ' + $D + 'x86' + $D + $B + $B
$nBad = ([regex]::Matches($runText, $badPat)).Count
if ($nBad -ne 0) { Write-Output ('FAIL R9: -Arch x86 with a trailing backslash survived in run_tests.ps1 (saw ' + $nBad + ') - PowerShell hands that backslash to C3 as a literal'); $fail++ }

if ($fail -eq 0) {
    Write-Output ('PASS check_emitted_declared: closure=' + $closure.Count + ' emitted=' + $emit.Count + ' defined=' + $defn.Count + ' declared=' + $decl.Count + ' internalOnly=' + $internalOnly + ' gap=0 parentDecl=1 prevDecl=1 decl=1 def=1 row=1 cells=2 table=1 x86shape=1')
    exit 0
}
Write-Output ('FAIL check_emitted_declared: ' + $fail + ' rule(s) red')
exit 1
