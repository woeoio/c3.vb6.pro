# check_owned_temp_reclaim.ps1 - 账 #464 §B164 回收刀：「现造的带权临时量」那一族发码形状的机械检查
#
# 这一刀做的事：一条语句里现造、**整枚**就是一个 mint 调用、并且**没人持有**的带权值，
#   提成具名临时量 `_own_N`，语句末按**档**收回（BSTR ⇒ vb6_BSTR_Free / VARIANT ⇒ vb6_VariantClear）。
#
# 本哨兵钉六件事（都是**机械**的，不是断言）：
#   R1 **档对不对**：`_own_N` 的声明档必须等于 **mint 出口自己的 RTL 返回档**
#      （[own] 节那一列；那张表由 scripts/check_owned_temp_keeps.ps1 从 RTL 推导 ⇒ 双向对账）。
#      —— 这一条正是 #461 折过的地方：那一版拿**消费者的 VB 形参档**当临时量的 C 档，
#         于是一枚 double 被声明成 BSTR ⇒ cl C2440（Charts 2020 ucTreeMaps + b433/b434 各四处）。
#   R2 **放几次**：每一枚 `_own_N` 必须**恰好**被放一次，而且放它的调用要对得上它的档。
#      —— 放两次 = 双放 ⇒ 堆损坏；放零次 = 还在漏。
#   R3 **用完才放**：`_own_N` 的每一次**读**都必须落在「声明之后、放掉之前」。
#      —— 放早了 = 悬空。
#   R4 **拼接站那条暗道没被动过**：拼接站判的是**左操作数的发码文本**
#      （"左文本里含 vb6_BSTR_Concat 子串" ⇒ 发 ConcatFree，而 ConcatFree 放掉左操作数）
#      ⇒ 左操作数里**不许**出现 `_own_`（提了名那条判据会翻面 ⇒ 行为变了）。
#   R5 **判据没翻面**：`vb6_BSTR_Concat(`（不带 Free）那一站的左操作数里
#      **不许**含 "vb6_BSTR_Concat" 子串（含了就说明它本该发 ConcatFree ⇒ 判据翻了面）。
#   R6 **同层**：声明与收回必须落在**同一个过程**里（`_own_N` 是 C 里的块级变量，
#      跨过程 ⇒ 引用越界）。本哨兵按**过程**记账（过程 = 发码里顶格 `... ) {` 那一行）。
#
# 用法: powershell -File scripts/check_owned_temp_reclaim.ps1
param(
    [string]$Root = ""
)
$ErrorActionPreference = 'Continue'
if (-not $Root) { $Root = (Split-Path -Parent $PSScriptRoot) }
$Root = (Resolve-Path $Root).Path
$c3 = Join-Path $Root ".build\C3.exe"
if (-not (Test-Path -LiteralPath $c3)) { Write-Host ("MISSING " + $c3); exit 2 }

$fail = @()

# ---- [own] 表（mint 出口的档：BSTR / VARIANT）----
$own = @{}
$sec = ''
foreach ($ln in [System.IO.File]::ReadAllLines((Join-Path $Root 'scripts\owned_temp_keeps_164.txt'))) {
    if ($ln -like '#*返回侧所有权：owned*') { $sec = 'own'; continue }
    if ($ln -like '#*') { continue }
    if ($ln.StartsWith('[')) {
        $p = $ln.Split(']')
        $rest = $p[1].Trim()
        $f = $rest.Split(' ')
        if ($p[0] -eq '[own' -and $f.Count -ge 2) { $own[$f[0]] = $f[1] }
        $sec = ''
        continue
    }
    if ($sec -ne 'own') { continue }
}

# ---- 工具：在一段 C 文本里找配平的 ')'（认字符串字面量）----
function Find-CloseParen([string]$s, [int]$open) {
    $depth = 0
    $inStr = $false
    for ($i = $open; $i -lt $s.Length; $i++) {
        $ch = $s[$i]
        if ($inStr) {
            if ($ch -eq '\') { $i++; continue }
            if ($ch -eq '"') { $inStr = $false }
            continue
        }
        if ($ch -eq '"') { $inStr = $true; continue }
        if ($ch -eq '(') { $depth++ }
        elseif ($ch -eq ')') {
            $depth--
            if ($depth -eq 0) { return $i }
        }
    }
    return -1
}
# ---- 工具：在实参表里找**顶层**逗号（认字符串字面量）----
function Split-TopLevel([string]$s) {
    $depth = 0
    $inStr = $false
    for ($i = 0; $i -lt $s.Length; $i++) {
        $ch = $s[$i]
        if ($inStr) {
            if ($ch -eq '\') { $i++; continue }
            if ($ch -eq '"') { $inStr = $false }
            continue
        }
        if ($ch -eq '"') { $inStr = $true; continue }
        if ($ch -eq '(' -or $ch -eq '[' -or $ch -eq '{') { $depth++ }
        elseif ($ch -eq ')' -or $ch -eq ']' -or $ch -eq '}') { $depth-- }
        elseif ($ch -eq ',' -and $depth -eq 0) {
            return @($s.Substring(0, $i), $s.Substring($i + 1))
        }
    }
    return @($s, '')
}
# ---- 工具：把一段 C 文本里**每一个**拼接站（vb6_BSTR_Concat / vb6_BSTR_ConcatFree）找出来 ----
#   每个站记: 名字（Concat / ConcatFree）+ 左操作数文本 + 右操作数文本
#   （**递归**进两个实参里 —— 嵌套的拼接站也在其内）
$script:concatSites = @()
function Scan-ConcatSites([string]$text) {
    $i = 0
    $tag = 'vb6_BSTR_Concat'
    $tagLen = $tag.Length
    while ($true) {
        $p = $text.IndexOf($tag, $i)
        if ($p -lt 0) { break }
        $i = $p + $tagLen
        $name = $tag
        if (($text.Length - $i) -ge 4 -and $text.Substring($i, 4) -eq 'Free') {
            $name = $tag + 'Free'
            $i += 4
        }
        $op = $text.IndexOf('(', $i)
        if ($op -lt 0) { break }
        $close = Find-CloseParen $text $op
        if ($close -lt 0) { break }
        $inner = $text.Substring($op + 1, $close - $op - 1)
        $sp = Split-TopLevel $inner
        $script:concatSites += ,@{ Name = $name; Left = $sp[0]; Right = $sp[1] }
        Scan-ConcatSites $sp[0]
        Scan-ConcatSites $sp[1]
        $i = $close + 1
    }
}

# ---- 工具：这一段文本**整枚**是不是一个调用？返回那个被调的名字（不是调用 ⇒ 空）----
function Get-TopLevelCallee([string]$text) {
    $t = $text.Trim()
    if ($t -notmatch '^([A-Za-z_][A-Za-z0-9_]*)\s*\(') { return '' }
    $name = $Matches[1]
    # 括号必须**配平到末尾**（f(a)(b) / f(a) + g(b) 都不算「整枚就是一个调用」）
    $depth = 0
    for ($i = $t.IndexOf('('); $i -lt $t.Length; $i++) {
        if ($t[$i] -eq '(') { $depth++ }
        elseif ($t[$i] -eq ')') {
            $depth--
            if ($depth -eq 0) { if ($i -eq $t.Length - 1) { return $name } else { return '' } }
        }
    }
    return ''
}

# ---- 被检查的语料（够覆盖那一族，又不至于慢到不能进 CI）----
# 语料里**两档**都要有样例：档 1 = BSTR（多数）、档 2 = VARIANT（少数）。
#   VARIANT 那一档的样例是 test_builtin_arity.bas（vb6_VariantDate / vb6_VariantLong 当整枚实参）
#   与 b432（vb6_VariantDate）⇒ 本哨兵下面那道「VARIANT 枚数 > 0」才立得住：
#   否则那一支就是**死的**（语料走不到 ⇒ 没人证它对）。
$samples = @(
    'tests\test_b164_owned_temp.bas',
    'tests\test_builtin_arity.bas',
    'tests\acc\acc_fam_main.bas',
    'tests\arr_cls\ArrCls.vbp',
    'tests\asm\asm_ok.vbp',
    'tests\b176_var_slot.bas',
    'tests\b432\B432.vbp',
    'tests\Charts 2020\ucTreeMaps\Proyecto1.vbp',
    'tests\VBFlexGridDemo\VBFlexGridDemo.vbp'
)

$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ("own464_" + [guid]::NewGuid().ToString("N") + ".c")
$totalDecl = 0; $totalFree = 0; $totalSites = 0
$declB = 0; $declV = 0
foreach ($rel in $samples) {
    $src = Join-Path $Root $rel
    if (-not (Test-Path -LiteralPath $src)) { $fail += ("样例不存在: " + $rel); continue }
    Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
    $p = Start-Process -FilePath $c3 -ArgumentList @('"' + $src + '"', '--emit-c') -NoNewWindow -Wait -PassThru `
        -RedirectStandardOutput $tmp -RedirectStandardError ([System.IO.Path]::ChangeExtension($tmp, '.err'))
    if ($p.ExitCode -ne 0) { $fail += ($rel + " 发码失败 rc=" + $p.ExitCode); continue }
    $lines = [System.IO.File]::ReadAllLines($tmp)
    # 临时量名是**过程内**的（C 里是块级变量）⇒ 这一格的账本必须**按过程**记,
    # 否则两个过程的同名临时量会互相算成「放完了还在读」。过程 = 顶格 `... ) {` 那一行。
    $decl = @{}; $free = @{}
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $ln = $lines[$i]
        # R4 / R5: 拼接站（vb6_BSTR_Concat / vb6_BSTR_ConcatFree）逐站查
        # ⚠ **#define 那一行不算**：编译期**常量**是**另一条路**折出来的
        #   （cgen_decl.cpp 的常量折叠：把常量的表达式折成一段文本再发成 #define），
        #   那条路**不走**表达式生成器 ⇒ 拼接站那一条规则压根没经过它。
        #   ⇒ 常量里那一串拼接的中间量**仍然漏**（另一族，另记），本哨兵不替它背书。
        if ($ln.StartsWith('#')) { $script:concatSites = @() } else {
            $script:concatSites = @()
            Scan-ConcatSites $ln
        }
        foreach ($st in $script:concatSites) {
            $totalSites++
            $leftShow = $st.Left.Trim()
            if ($leftShow.Length -gt 100) { $leftShow = $leftShow.Substring(0, 100) + " ..." }
            # R4: 拼接站**自己**放掉左操作数（ConcatFree）⇒ 若左操作数**整枚就是**一枚
            #     被语句末收回的临时量, 就会被放**两次** ⇒ 双放 ⇒ 堆损坏。
            #     （左操作数**里面**含一枚临时量是安全的: 那一站放掉的是**外层**那个表达式的值,
            #       不是里面那一枚 ⇒ 里面那一枚仍只由语句末放一次。）
            if ($st.Left.Trim() -match '^_own_\d+$') {
                $fail += ($rel + " 第 " + ($i + 1) + " 行: 拼接站的左操作数**整枚就是**一枚临时量 ⇒ 拼接站会放它 + 语句末也放它 ⇒ 双放: " + $leftShow)
            }
            # R5: 那一站的**发法**必须与**那条规则**一致（账 #468 §B164）：
            #     ConcatFree ⇔（左操作数文本含 "vb6_BSTR_Concat" 子串）**或**
            #                （左操作数整枚是一个 owned 且档是 BSTR 的 mint ⇒ 拼接站自己放掉它）
            #     少了那一支 ⇒ 那一枚没人放（漏）；多了 ⇒ 放了一枚不该放的（双放/悬空）。
            $leftMint = Get-TopLevelCallee $st.Left
            $leftIsOwnedBstr = ($leftMint -ne '' -and $own.ContainsKey($leftMint) -and $own[$leftMint] -eq 'BSTR')
            $wantFree = ($st.Left.Contains('vb6_BSTR_Concat') -or $leftIsOwnedBstr)
            $isFree = ($st.Name -eq 'vb6_BSTR_ConcatFree')
            if ($wantFree -ne $isFree) {
                $fail += ($rel + " 第 " + ($i + 1) + " 行: 拼接站发的是 " + $st.Name +
                          " 而按规则该发 " + $(if ($wantFree) { 'vb6_BSTR_ConcatFree' } else { 'vb6_BSTR_Concat' }) +
                          "（左操作数含拼接子串=" + $st.Left.Contains('vb6_BSTR_Concat') +
                          " / 左操作数整枚是 owned BSTR mint=" + $leftIsOwnedBstr + "）: " + $leftShow)
            }
        }
        # 新过程入口: 顶格、以 ) { 结尾
        if ($ln -match '^\S.*\)\s*\{\s*$') {
            # 上一过程收尾: 声明了却没被放掉的, 就是在漏
            foreach ($n in $decl.Keys) {
                if (-not $free.ContainsKey($n)) {
                    $fail += ($rel + ": " + $n + " 声明了却**没有**被放掉（还在漏）")
                }
            }
            $decl = @{}; $free = @{}
            continue
        }
        # 声明: BSTR _own_N = ... 或 vb6_VARIANT _own_N = ...
        foreach ($m in [regex]::Matches($ln, '(BSTR|vb6_VARIANT)\s+(_own_\d+)\s*=\s*([A-Za-z_][A-Za-z0-9_]*)\s*\(')) {
            $kind = $m.Groups[1].Value
            $name = $m.Groups[2].Value
            $callee = $m.Groups[3].Value
            $decl[$name] = @{ line = $i; kind = $kind; callee = $callee }
            $totalDecl++
            if ($kind -eq 'BSTR') { $declB++ } else { $declV++ }
            # R1: 档必须等于 mint 出口自己的返回档
            if (-not $own.ContainsKey($callee)) {
                $fail += ($rel + " 第 " + ($i + 1) + " 行: " + $name + " 的 mint 出口 " + $callee + " 不在 [own] 表里")
            } else {
                $want = if ($own[$callee] -eq 'VARIANT') { 'vb6_VARIANT' } else { 'BSTR' }
                if ($want -ne $kind) {
                    $fail += ($rel + " 第 " + ($i + 1) + " 行: " + $name + " 声明成 " + $kind + "，但 " + $callee + " 的返回档是 " + $own[$callee])
                }
            }
        }
        # 放掉: vb6_BSTR_Free(_own_N) 或 vb6_VariantClear(&_own_N)
        foreach ($m in [regex]::Matches($ln, 'vb6_BSTR_Free\(\s*(_own_\d+)\s*\)')) {
            $n = $m.Groups[1].Value
            $totalFree++
            if (-not $decl.ContainsKey($n)) {
                $fail += ($rel + " 第 " + ($i + 1) + " 行: 放掉未声明的 " + $n)
                continue
            }
            if ($free.ContainsKey($n)) {
                $fail += ($rel + " 第 " + ($i + 1) + " 行: " + $n + " 被放掉两次（双放）")
                continue
            }
            # R2: 放法必须对得上声明的档（BSTR ⇒ BSTR_Free; VARIANT ⇒ VariantClear）
            if ($decl[$n].kind -ne 'BSTR') {
                $fail += ($rel + " 第 " + ($i + 1) + " 行: " + $n + " 声明成 " + $decl[$n].kind + "，却用 BSTR_Free 放掉")
            }
            $free[$n] = @{ line = $i; kind = 'BSTR' }
        }
        foreach ($m in [regex]::Matches($ln, 'vb6_VariantClear\(\s*&?\s*(_own_\d+)\s*\)')) {
            $n = $m.Groups[1].Value
            $totalFree++
            if (-not $decl.ContainsKey($n)) {
                $fail += ($rel + " 第 " + ($i + 1) + " 行: 放掉未声明的 " + $n)
                continue
            }
            if ($free.ContainsKey($n)) {
                $fail += ($rel + " 第 " + ($i + 1) + " 行: " + $n + " 被放掉两次（双放）")
                continue
            }
            if ($decl[$n].kind -ne 'vb6_VARIANT') {
                $fail += ($rel + " 第 " + ($i + 1) + " 行: " + $n + " 声明成 " + $decl[$n].kind + "，却用 VariantClear 放掉")
            }
            $free[$n] = @{ line = $i; kind = 'VARIANT' }
        }
        # 读取（声明行和放掉行本身不算读）
        foreach ($m in [regex]::Matches($ln, '_own_\d+')) {
            $n = $m.Groups[0].Value
            if (-not $decl.ContainsKey($n)) {
                # 本过程里没有这一枚的声明 ⇒ 要么读在声明之前, 要么声明在别的过程里
                $fail += ($rel + " 第 " + ($i + 1) + " 行: 读取本过程里没有声明的 " + $n)
                continue
            }
            if ($decl[$n].line -eq $i) { continue }          # 声明那一行本身
            if ($free.ContainsKey($n) -and $free[$n].line -eq $i) { continue }   # 放它那一行本身
            # R3: 读都在放之前。只有当 free 已经发生且读在 free 之后才报错
            if ($free.ContainsKey($n) -and $i -gt $free[$n].line) {
                $fail += ($rel + " 第 " + ($i + 1) + " 行: " + $n + " 在被放掉之后还被读（悬空）")
            }
        }
    }
    # 最后一个过程收尾
    foreach ($n in $decl.Keys) {
        if (-not $free.ContainsKey($n)) {
            $fail += ($rel + ": " + $n + " 声明了却**没有**被放掉（还在漏）")
        }
    }
}
Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue

if ($totalDecl -eq 0) { $fail += '一样本里一枚 _own_ 都没有 ⇒ 这一刀根本没生效（哨兵不许空转）' }
if ($totalDecl -ne $totalFree) { $fail += ("声明 " + $totalDecl + " 枚，放 " + $totalFree + " 次 ⇒ 数目对不上") }
# R7: **两档都得有样例** —— VARIANT 那一档（vb6_VariantClear）若语料走不到，
#     那一支就是**死的**（没人证它对）⇒ 语料里必须真的有一枚 VARIANT 档的临时量。
if ($declV -eq 0) { $fail += '语料里一枚 VARIANT 档的临时量都没有 ⇒ VARIANT 那一支（vb6_VariantClear）没被走到 ⇒ 那一支是死的' }
if ($declB -eq 0) { $fail += '语料里一枚 BSTR 档的临时量都没有 ⇒ BSTR 那一支（vb6_BSTR_Free）没被走到' }

if ($fail.Count -gt 0) {
    Write-Host "OWNED-TEMP-RECLAIM FAIL"
    foreach ($f in $fail) { Write-Host ("  - " + $f) }
    exit 1
}
Write-Host ("OWNED-TEMP-RECLAIM OK (样例 " + $samples.Count + " 份 / 声明 " + $totalDecl +
    " 枚 [BSTR " + $declB + " / VARIANT " + $declV + "] / 收回 " + $totalFree + " 次 / 档全对 / 放一次 / 用完才放 / 拼接站 " + $totalSites +
    " 站暗道未动 / 判据未翻面)")
exit 0
