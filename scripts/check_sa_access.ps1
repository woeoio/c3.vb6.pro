# 账 #209 + 账 #214 的结构性哨兵 (只扫源码, 不起 cl)
#
# 为什么单开这一道: 一维动态数组的**元素**访问在 RTL 里只有一个入口 —— VB6_SA_AT。
# 这一刀之前它是裸指针算术 (`((type*)((arr)->data))[idx - (arr)->lBound]`), 描述符是 NULL
# (动态数组从未 ReDim / 已被 Erase) 或下标越界时它一声不吭地读 NULL+0xc → 原生 0xC0000005。
# 实测 Charts 2020 ucChartBar demo 点 Random: 三次崩在同一偏移 0x1abf2, 符号化落在
# 生成代码 ucChartBar.c:591 的 `With m_Serie(Index)`。VB6 在这一条是运行时错误 9, 和
# vb6_UBound/vb6_LBound 的 rev2 同一族 —— 那两处已经抛 9, 这一处漏了。
#
# 现在两支都收在同一形状的出口上: 一维走 vb6_SaElemPtr, 多维走 vb6_SaNdElemPtr -- (各一条 inline 热路径 + 一条冷路径)
# 账 #214 是这一族的下一半: VB6_SA_ND_AT1/2/3 当初与 #209 修法前的一维一模一样 (描述符在不在、
# 每一维在不在范围, 一个都不问), 而 4 秩以上另有一条 `_ndoff_` 兜底 —— 那条更糟: 发码侧把实参写成
# `(int[]){i0, i1}` 只塞两枚下标却按实际秩数交出去, 于是 `Dim a4(1 To 2,1 To 2,1 To 2,1 To 2)`
# **在范围内也** 0xC0000005 (x86/x64 同形)。两支合起来存量 1468 处 / 4 份工程, 大头是 VBFlexGridDemo。
# 所以 A4 现在改成“一条都不许有”, 另加 A6 (多维那一支的唯一入口) 与 A7 (发码侧的形状)。
#
# 规则 (改坏了会红, 不是装饰):
#   A1  VB6_SA_AT 的宏定义恰好 1 处, 且宏体走 vb6_SaElemPtr (不许再退回裸算术)
#   A2  vb6_SaElemPtr 的 inline 定义恰好 1 处, 体内两条比较(下界/上界)与 NULL 那条都还在
#   A3  vb6_SaElemFail 声明 1 + 定义 1, 定义体里 vb6_ErrRaise(9 那条还在
#   A4  src/backend 里非注释的 `->data` 元素寻址行 = 0 (两支都不许再手算元素地址)
#   A5  src/rtl 与 src/backend 里 `->data))[` 这种"宏里手算下标"的形状 = 0
#   A6  多维那一支: AT1/AT2/AT3/ATN 四条宏各 1 处且宏体都经 vb6_SaNdElemPtr;
#       vb6_SaNdElemPtr 定义 1 处, 体内形状问句(dimCount)+两条上下界比较+冷路径调用都在;
#       vb6_SaNdElemFail 声明 1 + 定义 1, 定义体真的抛 9
#   A7  发码侧 4 秩以上那一条: 发 VB6_SA_ND_ATN, 下标实参按 `((const int32_t[]){` 拼(最外层括号
#       不是装饰, 少了它就是 cl C4002), 且不再手拼 `(int[]){` / 不再直接发 vb6_SafeArrayND_Offset(
#
# 用法:  pwsh -File scripts\check_sa_access.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$bad = @()

function Lines([string]$path) {
    ([System.IO.File]::ReadAllText($path) -split "`r?`n")
}

$rtlDir = Join-Path $root "src\rtl"
$beDir = Join-Path $root "src\backend"
$arrH = Join-Path $root "src\rtl\core\vb6rtl\vb6rtl_array.h"
$arrC = Join-Path $root "src\rtl\core\vb6rtl\vb6rtl_array.c"
foreach ($p in @($arrH, $arrC)) {
    if (-not (Test-Path -LiteralPath $p)) {
        Write-Host ("FAIL A0 file missing: " + $p) -ForegroundColor Red
        exit 1
    }
}

# A1: 宏只有定义点这一处写法, 而且必须经 vb6_SaElemPtr
$macroDefs = @()
$macroBodyCallsHelper = $false
foreach ($f in (Get-ChildItem -LiteralPath $rtlDir -Recurse -File | Where-Object { $_.Extension -in ".c", ".h" })) {
    $ln = 0
    foreach ($line in (Lines $f.FullName)) {
        $ln++
        if ($line.Trim() -match '^#\s*define\s+VB6_SA_AT\(') {
            $macroDefs += ($f.Name + ":" + $ln)
            $body = (Lines $f.FullName)[$ln..([Math]::Min($ln + 2, (Lines $f.FullName).Count - 1))] -join " "
            if ($body -match "vb6_SaElemPtr") { $macroBodyCallsHelper = $true }
        }
    }
}
if ($macroDefs.Count -ne 1) {
    $bad += ("A1 VB6_SA_AT macro definitions = " + $macroDefs.Count + " (want exactly 1) -> " + ($macroDefs -join " | "))
}
if ($macroDefs.Count -eq 1 -and -not $macroBodyCallsHelper) {
    $bad += "A1 VB6_SA_AT body no longer routes through vb6_SaElemPtr (the check moved back to raw math)"
}

# A2: 那个唯一的 inline, 三条检查一条不许哑
$h = [System.IO.File]::ReadAllText($arrH)
$defs = @([regex]::Matches($h, 'static\s+inline\s+void\*\s+vb6_SaElemPtr\s*\('))
if ($defs.Count -ne 1) {
    $bad += ("A2 vb6_SaElemPtr defined " + $defs.Count + " times (want 1)")
} else {
    $blk = [regex]::Match($h, 'static\s+inline\s+void\*\s+vb6_SaElemPtr[\s\S]{0,800}?\r?\n\}')
    if ($blk.Value -notmatch '!arr') { $bad += "A2 NULL-descriptor check is gone from the helper" }
    if ($blk.Value -notmatch 'idx\s*<\s*arr->lBound') { $bad += "A2 lower-bound compare is gone from the helper" }
    if ($blk.Value -notmatch 'idx\s*>\s*arr->uBound') { $bad += "A2 upper-bound compare is gone from the helper" }
    if ($blk.Value -notmatch 'vb6_SaElemFail') { $bad += "A2 the cold-path call is gone from the helper" }
}

# A3: 冷路径声明 + 定义各 1, 且真的抛 9
$decl = @([regex]::Matches($h, 'void\s+vb6_SaElemFail\s*\(')).Count
$csrc = [System.IO.File]::ReadAllText($arrC)
$defn = @([regex]::Matches($csrc, '\r?\nvoid\s+vb6_SaElemFail\s*\(')).Count
if ($decl -ne 1) { $bad += ("A3 vb6_SaElemFail 声明 = " + $decl + " (want 1)") }
if ($defn -ne 1) { $bad += ("A3 vb6_SaElemFail 定义 = " + $defn + " (want 1)") }
$cblk = [regex]::Match($csrc, 'void\s+vb6_SaElemFail[\s\S]{0,1200}?\r?\n\}')
if (-not $cblk.Success -or $cblk.Value -notmatch 'vb6_ErrRaise\(\s*9\b') {
    $bad += "A3 cold path no longer raises runtime error 9 (silent return would spare sites without On Error)"
}

# A4: 后端不再许有任何一处手算元素地址的行 (两支都收进出口了)
$rawData = @()
foreach ($f in (Get-ChildItem -LiteralPath $beDir -Recurse -File | Where-Object { $_.Extension -in ".cpp", ".inc", ".hpp" })) {
    $ln = 0
    foreach ($line in (Lines $f.FullName)) {
        $ln++
        $t = $line.Trim()
        if ($t.StartsWith("//")) { continue }
        if ($line.Contains('->data')) { $rawData += ($f.Name + ":" + $ln) }
    }
}
if ($rawData.Count -ne 0) {
    $bad += ("A4 后端手算元素地址的行 = " + $rawData.Count +
             " (want 0: 一维走 VB6_SA_AT, 多维走 VB6_SA_ND_AT*, 都不该再碰 ->data) -> " + ($rawData -join " | "))
}

# A5: "宏里手算下标"这个形状在整个仓里不许复活
$shape = 0
foreach ($dir in @($rtlDir, $beDir)) {
    foreach ($f in (Get-ChildItem -LiteralPath $dir -Recurse -File | Where-Object { $_.Extension -in ".c", ".h", ".cpp", ".inc", ".hpp" })) {
        foreach ($line in (Lines $f.FullName)) {
            if ($line.Trim().StartsWith("//")) { continue }
            if ($line.Contains('->data))[')) { $shape++ }
        }
    }
}
if ($shape -ne 0) {
    $bad += ("A5 裸 `->data))[` 元素寻址 = " + $shape + " 处 (want 0; 一维那支必须只有宏这一个入口)")
}

# A6: 多维那一支的入口与热路径问句
$ndMacros = @()
$ndViaHelper = 0
foreach ($f in (Get-ChildItem -LiteralPath $rtlDir -Recurse -File | Where-Object { $_.Extension -in ".c", ".h" })) {
    $fl = Lines $f.FullName
    $ln = 0
    foreach ($line in $fl) {
        $ln++
        if ($line -match '^\s*#\s*define\s+(VB6_SA_ND_AT\d|VB6_SA_ND_ATN)\b') {
            $ndMacros += ($Matches[1] + "@" + $f.Name + ":" + $ln)
            $hi = [Math]::Min($ln + 2, $fl.Count - 1)
            $body = $fl[$ln..$hi] -join " "
            if ($body -match "vb6_SaNdElemPtr") { $ndViaHelper++ }
        }
    }
}
if ($ndMacros.Count -ne 4) {
    $bad += ("A6 多维元素访问宏 = " + $ndMacros.Count + " 处 (want 4: AT1/AT2/AT3/ATN) -> " + ($ndMacros -join " | "))
}
if ($ndMacros.Count -eq 4 -and $ndViaHelper -ne 4) {
    $bad += ("A6 宏体经 vb6_SaNdElemPtr 的条数 = " + $ndViaHelper + " (want 4; 有一支退回裸算术了)")
}
$ndDefs = @([regex]::Matches($h, 'static\s+inline\s+void\*\s+vb6_SaNdElemPtr\s*\('))
if ($ndDefs.Count -ne 1) {
    $bad += ("A6 vb6_SaNdElemPtr defined " + $ndDefs.Count + " times (want 1)")
} else {
    $ndblk = [regex]::Match($h, 'static\s+inline\s+void\*\s+vb6_SaNdElemPtr[\s\S]{0,1200}?\r?\n\}')
    if ($ndblk.Value -notmatch '!a') { $bad += "A6 NULL-descriptor check is gone from the ND helper" }
    if ($ndblk.Value -notmatch 'dimCount') { $bad += "A6 the ND helper no longer asks the shape (rank vs dimCount)" }
    if ($ndblk.Value -notmatch 'idx\[d\]\s*<\s*lb') { $bad += "A6 lower-bound compare is gone from the ND helper" }
    if ($ndblk.Value -notmatch 'idx\[d\]\s*>=\s*lb\s*\+\s*cnt') { $bad += "A6 upper-bound compare is gone from the ND helper" }
    if ($ndblk.Value -notmatch 'vb6_SaNdElemFail') { $bad += "A6 the cold-path call is gone from the ND helper" }
}
$ndDecl = @([regex]::Matches($h, 'void\s+vb6_SaNdElemFail\s*\(')).Count
$ndDefn = @([regex]::Matches($csrc, '\r?\nvoid\s+vb6_SaNdElemFail\s*\(')).Count
if ($ndDecl -ne 1) { $bad += ("A6 vb6_SaNdElemFail 声明 = " + $ndDecl + " (want 1)") }
if ($ndDefn -ne 1) { $bad += ("A6 vb6_SaNdElemFail 定义 = " + $ndDefn + " (want 1)") }
$ndCblk = [regex]::Match($csrc, '\r?\nvoid\s+vb6_SaNdElemFail[\s\S]{0,1400}?\r?\n\}')
if (-not $ndCblk.Success -or $ndCblk.Value -notmatch 'vb6_ErrRaise\(\s*9\b') {
    $bad += "A6 ND cold path no longer raises runtime error 9 (silent return would spare sites without On Error)"
}

# A7: 发码侧 4 秩以上那一支的形状
$atnEmit = 0; $atnLiteral = 0; $handIdxList = 0; $offsetEmit = 0
foreach ($f in (Get-ChildItem -LiteralPath $beDir -Recurse -File | Where-Object { $_.Extension -in ".cpp", ".inc", ".hpp" })) {
    $ln = 0
    foreach ($line in (Lines $f.FullName)) {
        $ln++
        $t = $line.Trim()
        if ($t.StartsWith("//")) { continue }
        if ($t.Contains("VB6_SA_ND_ATN(")) { $atnEmit++ }
        if ($t.Contains('((const int32_t[])')) { $atnLiteral++ }
        if ($t.Contains('(int[]){')) { $handIdxList++ }
        if ($t.Contains('vb6_SafeArrayND_Offset(')) { $offsetEmit++ }
    }
}
if ($atnEmit -lt 1) { $bad += "A7 后端不再发 VB6_SA_ND_ATN (4 秩以上那支又回到手算偏移了)" }
if ($atnLiteral -ne 1) {
    $bad += ("A7 ATN 的下标实参不是按 ((const int32_t[]){ 拼的 = " + $atnLiteral +
             " 处 (want 1; 少最外层那对括号 ⇒ cl C4002 参数过多)")
}
if ($handIdxList -ne 0) {
    $bad += ("A7 手拼 (int[]){ 的下标列表又出现 " + $handIdxList +
             " 处 (那正是「只塞两枚下标却按实际秩数交出去」那个 bug 的形状)")
}
if ($offsetEmit -ne 0) {
    $bad += ("A7 后端直接发 vb6_SafeArrayND_Offset( 的行 = " + $offsetEmit +
             " (want 0: 元素访问只许走 AT* 那几个出口)")
}


if ($bad.Count -eq 0) {
    Write-Host ("PASS SA element access: macro " + $macroDefs.Count +
                " / inline " + $defs.Count +
                " / cold-path decl " + $decl + " def " + $defn +
                " / backend raw-data lines " + $rawData.Count +
                " / raw-index shape " + $shape + " / ND macros " + $ndMacros.Count + " / ND inline " + $ndDefs.Count + " / ATN emit " + $atnEmit + " / ATN literal " + $atnLiteral) -ForegroundColor Green
    exit 0
}
foreach ($b in $bad) { Write-Host ("FAIL " + $b) -ForegroundColor Red }
exit 1
