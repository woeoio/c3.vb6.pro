# check_byval_str_copy.ps1 — 账 #610 第343刀: ByVal String 副本的**释放**闭环哨兵。
#
# 为什么要有它 (形状门看不见的那部分):
#   · 形状门只哈希 --emit-c 的 stdout ⇒ 「副本变量是不是每个 ByVal String 形参都有」
#     「释放段是不是带判等」「Function 侧是不是漏了 free」这几格它一个字都不管;
#   · 出口释放是**运行期**行为 (堆), 形状门跑不出来 ⇒ 必须有静态哨兵钉住逐字形态;
#   · 「Function 侧不得有副本变量」是**按过程作用域**的判据 —— 同一份产物里
#     Sub 侧必须有、Function 侧必须没有 ⇒ 整文件 grep 区分不了, 必须切函数体。
#
# 判据:
#   R1  源码: Sub 侧的副本发码必须**另立副本变量** (不得是裸的 `p = vb6_BSTR_FromBSTR(p);`)
#   R2  源码: 副本释放必须**带判等** (if (形参 == 副本)) —— 无条件 free = 双释放
#   R3  源码: 副本变量声明必须**两轮发** (先全部声明, 再全部赋值) ⇒ C89 块首约束
#   R4  源码: 释放段必须在 vb6_proc_exit 标签**之后** ⇒ Exit Sub 的 goto 也走到
#   R5  产物: vb6_LeakFuncC (Function) 体内**不得**有副本变量/释放 ⇒ 零回归
#   R6  产物: vb6_LeakSinkA/B/D (Sub) 体内必须有副本变量 + 带判等的释放
#   R7  产物: vb6_LeakSinkD 体内, 释放段必须在 `vb6_proc_exit:;` **之后** (Exit Sub 档)
param(
    [string]$Root = (Split-Path -Parent $PSScriptRoot),
    [string]$FixtureVbp = ""
)
$ErrorActionPreference = "Stop"
$problems = @()

$Root = (Resolve-Path $Root).Path
if (-not $FixtureVbp) { $FixtureVbp = Join-Path $Root "tests\fgz_fix013\FgzFix013.vbp" }
if (-not (Test-Path $FixtureVbp)) { Write-Output "SKIP: 夹具不存在 ($FixtureVbp)"; exit 0 }

# ---------- R1..R4: 源码级 ----------
$srcPath = Join-Path $Root "src\backend\decl\cgen_decl_proc.cpp"
if (-not (Test-Path $srcPath)) { $problems += "A: 找不到 src/backend/decl/cgen_decl_proc.cpp" }
else {
    $src = Get-Content $srcPath -Raw
    # R1: 另立副本变量 —— 必须有 `BSTR vb6_byvalcopy_` 的发码, 且不得只剩裸赋值
    if ($src -notmatch 'BSTR " \+ copyVar012|"\+ copyVar012 \+ " = vb6_BSTR_FromBSTR' -and
        $src -notmatch '"BSTR " \+ copyVar012') {
        $problems += "R1: Sub 侧副本发码不再另立副本变量 (找不到 vb6_byvalcopy_ 变量的发码) —— FIX-012 的副本语义丢了"
    }
    # R2: 带判等的释放
    if ($src -notmatch 'if \(" \+ bc343\.first \+ " == " \+ bc343\.second') {
        $problems += "R2: 副本释放不带判等 (必须 `if (形参 == 副本)`) —— 无条件 free 会双重释放"
    }
    # R3: 两轮发 (先声明后赋值) —— 判据必须**逐字**钉住两轮的形态:
    #   第一轮: 只 emitLine 声明 + push_back 收集 (不许 emit 赋值)
    #   第二轮: 才 emitLine "形参 = 副本;"
    # 只查 "copies343" 这个标识符存在是不够的 (把赋值挪回第一轮它照样过 => 假绿)。
    # 行尾: 源码是 CRLF => 逐行比对时先归一化再比。
    $srcLines = @($src -split "`r?`n")
    function Find-Line([string]$needle) {
        # ⚠ 两个坑都在这里:
        #   1) .Contains() 的返回值必须吞掉 ($hit = ...) —— 否则混进输出流,
        #      返回 @(False, index) 数组 ⇒ 后面 -lt 0 的比较全乱 ⇒ 假绿;
        #   2) 循环变量**不能叫 $i** —— PowerShell 函数默认读写父作用域,
        #      $i 会把调用方的 $i 覆盖掉 ⇒ 调用方的 for 循环范围错乱 ⇒ 假绿。
        $fl = 0
        while ($fl -lt $srcLines.Count) {
            $hit = $srcLines[$fl].Contains($needle)
            if ($hit) { return $fl }
            $fl++
        }
        return -1
    }
    $iDecl   = Find-Line 'copyVar012 + " = vb6_BSTR_FromBSTR('
    $iPush   = Find-Line 'copies343.push_back({pn012, copyVar012});'
    $iAssign = Find-Line 'cp343.first + " = " + cp343.second + ";'
    $iSecond = Find-Line 'for (auto& cp343 : copies343) {'
    if ($iDecl -lt 0 -or $iPush -lt 0 -or $iAssign -lt 0 -or $iSecond -lt 0) {
        $problems += "R3: 找不到两轮发的四个关键点 (声明 emit / push_back / 第二轮 for / 赋值 emit)"
    } else {
        if ($iPush -le $iDecl) { $problems += "R3: 第一轮顺序不对 (push_back 应在声明 emit 之后)" }
        if ($iSecond -le $iPush) { $problems += "R3: 第二轮 for 应在第一轮 (push_back) 之后" }
        if ($iAssign -le $iSecond) { $problems += "R3: 赋值 emit 应在第二轮 for 之内" }
        for ($i = $iDecl; $i -le $iPush; $i++) {
            if ($srcLines[$i].Contains('c_.emitLine(pn012 + " = " + copyVar012 + ";')) {
                $problems += "R3: 赋值 emit 出现在第一轮 (收集之前) => 单轮发 => 多形参时第二个声明夹在语句之后"
                break
            }
        }
    }

    # R4: 释放段在标签之后 —— 源码里 emitLine("vb6_proc_exit:;") 必须早于释放段的 emitLine
    $tagPos = $src.IndexOf('vb6_proc_exit:;')
    $freePos = $src.IndexOf('bc343.first')
    if ($tagPos -lt 0 -or $freePos -lt 0) {
        $problems += "R4: 找不到出口标签或释放段的 emit 点"
    } elseif ($freePos -lt $tagPos) {
        $problems += "R4: 释放段在 vb6_proc_exit 标签**之前** ⇒ Exit Sub 的 goto 走不到 ⇒ 副本仍泄漏"
    }
}

# ---------- R5..R7: 产物级 (按过程作用域切函数体) ----------
$c3 = Join-Path $Root ".build\C3.exe"
if (-not (Test-Path $c3)) { Write-Output "SKIP: 找不到 C3.exe"; exit 0 }
$out = & $c3 $FixtureVbp --emit-c 2>$null | Out-String
if (-not $out) { $problems += "B: C3 --emit-c 没有输出" }
else {
    $lines = $out -split "`r?`n"
    $bodies = @{}
    $cur = $null
    foreach ($ln in $lines) {
        if ($ln -match '^\s*(?:void|BSTR|int32_t|int16_t)\s+vb6_(\w+)\([^)]*\)\s*\{') {
            $cur = $Matches[1]; $bodies[$cur] = @(); continue
        }
        if ($cur -ne $null) {
            if ($ln -match '^\}') { $cur = $null; continue }
            $bodies[$cur] += $ln
        }
    }
    # R5: Function 侧零回归
    if (-not $bodies.ContainsKey("LeakFuncC")) {
        $problems += "R5: 产物里找不到过程 LeakFuncC"
    } else {
        $fbody = ($bodies["LeakFuncC"] -join "`n")
        if ($fbody -match 'vb6_byvalcopy') {
            $problems += "R5: Function 侧出现了副本变量/释放 ⇒ 返回值可能就是该形参 ⇒ free 它会让返回值悬垂"
        }
        if ($fbody -match 'vb6_BSTR_Free') {
            $problems += "R5: Function 侧出现了 vb6_BSTR_Free ⇒ 同上 (返回值悬垂)"
        }
    }
    # R6: Sub 侧三档都要副本变量 + 带判等的释放
    foreach ($fn in @("LeakSinkA", "LeakSinkB", "LeakSinkD")) {
        if (-not $bodies.ContainsKey($fn)) { $problems += "R6: 产物里找不到过程 $fn"; continue }
        $body = ($bodies[$fn] -join "`n")
        if ($body -notmatch 'BSTR vb6_byvalcopy_s = vb6_BSTR_FromBSTR\(s\);') {
            $problems += "R6: $fn 体内没有副本变量的发码"
        }
        if ($body -notmatch 'if \(s == vb6_byvalcopy_s\) \{ vb6_BSTR_Free\(s\); \}') {
            $problems += "R6: $fn 体内没有带判等的副本释放"
        }
    }
    # R7: Exit Sub 档 —— 释放段必须在标签之后
    if ($bodies.ContainsKey("LeakSinkD")) {
        $body = $bodies["LeakSinkD"]
        $tagIdx = -1; $freeIdx = -1
        for ($i = 0; $i -lt $body.Count; $i++) {
            if ($body[$i] -match 'vb6_proc_exit:;' -and $tagIdx -lt 0) { $tagIdx = $i }
            if ($body[$i] -match 'vb6_BSTR_Free\(s\)' -and $freeIdx -lt 0) { $freeIdx = $i }
        }
        if ($tagIdx -lt 0) { $problems += "R7: LeakSinkD 体内找不到 vb6_proc_exit 标签 (Exit Sub 档没走到出口)" }
        elseif ($freeIdx -lt 0) { $problems += "R7: LeakSinkD 体内找不到释放段" }
        elseif ($freeIdx -lt $tagIdx) { $problems += "R7: LeakSinkD 的释放段在 vb6_proc_exit 标签之前 ⇒ Exit Sub 路径漏释放" }
    }
}

if ($problems.Count -eq 0) {
    Write-Output "OK: ByVal String 副本释放闭环 (R1..R7) 全部通过"
    exit 0
} else {
    Write-Output "FAIL: ByVal String 副本释放闭环哨兵"
    foreach ($p in $problems) { Write-Output "  - $p" }
    exit 1
}
