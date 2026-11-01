# check_radix_tier_sites.ps1 - 账 #305 第313刀 (§B239) 的结构性哨兵 (只扫源码, 不起 cl)
#
# 为什么单开这一道: 无后缀进制字面量的「档位与值」那一问, 从前在 &H / &O / &B 三支里**各抄一份**
# (三处同一个比较 + 三处同一个 static_cast)。第312刀量到的缺口 (缺 16 位补码那一档) 因此一次性
# 存在于三处 —— 加一档要改三遍, 漏一遍就是**静默错值**: &HBF32 被当成 +48946, C 侧只警告不报红,
# 整片工程照编照跑, 直到用户点下 Command2 才由 vb6_ChkInt 抛出运行期错误 6 打死进程。
#
# 数法说明 (两条都是被负控逼出来的):
#   - 每行先剥掉 `//` 之后的部分再数。本哨兵钉的那些形状**也写在代码注释里** (解释旧阶梯为什么撤掉),
#     不剥就会把散文数成站点 (第300刀 check_control_dc 的 D11 撞过同一件事)。
#   - 「定义」与「调用点」分开钉, 且调用点必须**逐字**带 allowNarrowFold= —— 只数裸名会让一支
#     漏问这一格的改动照样过针 (第280/279 刀同一条教训)。
#
# 规则:
#   R1  主人三位一体: struct RadixTier 1 / radixTierOf 定义 1 / Lexer::makeRadixToken 定义 1 + 声明 1
#   R2  三条进制支路各问一次: "return makeRadixToken(" 恰 3, 且每一行都当场答 allowNarrowFold=
#   R3  旧的两份抄写不许回来: 阶梯比较 <= 0xFFFFFFFFu 在代码里恰 1 处 (主人内), 旧截断
#       scanHexNumber 之后任何 tok.intValue = 赋值恰 0 处 (十进制那支不在切片里)
#   R4  16 位那一支的三个必要件都在场: 0xFFFFu / 0x8000u / 0x10000 各恰 1
#   R5  判据与哨兵自己: 夹具在场 + 两台 Add-BasTest 各 1 + 发码形状针 1 + 本脚本被调 1 + 调用点 1
#
# 用法:  pwsh -File scripts\check_radix_tier_sites.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$bad = @()

function CodeRows([string]$path) {
    # 剥掉行内 // 注释与整行注释, 换行与行号不动
    $t = [System.IO.File]::ReadAllText($path)
    $out = @()
    foreach ($ln in ($t -split "\r?\n")) {
        $c = $ln.IndexOf("//")
        if ($c -ge 0) { $out += $ln.Substring(0, $c) } else { $out += $ln }
    }
    return $out
}

function CountMatches($rows, [string]$pattern) {
    $n = 0
    foreach ($r in $rows) { $n += [regex]::Matches($r, $pattern).Count }
    return $n
}

$lexPath = Join-Path $root "src\lexer\lexer_number.cpp"
$hppPath = Join-Path $root "src\lexer\lexer.hpp"
$rtPath  = Join-Path $root "tests\run_tests.ps1"
$fixPath = Join-Path $root "tests\b591_radixtier.bas"
foreach ($f in @($lexPath, $hppPath, $rtPath, $fixPath)) {
    if (-not (Test-Path -LiteralPath $f)) {
        Write-Host ("FAIL S0 missing file " + $f) -ForegroundColor Red
        exit 1
    }
}
$lex = CodeRows $lexPath
$hpp = CodeRows $hppPath
$rt  = [System.IO.File]::ReadAllText($rtPath)
$fix = [System.IO.File]::ReadAllText($fixPath)

# --- R1 主人三位一体 ---
$nStruct = CountMatches $lex "struct\s+RadixTier\s*\{"
$nDef    = CountMatches $lex "^\s*RadixTier\s+radixTierOf\s*\("
$nTokDef = CountMatches $lex "^\s*Token\s+Lexer::makeRadixToken\s*\("
$nTokDec = CountMatches $hpp "\bmakeRadixToken\s*\("
$nCallOf = CountMatches $lex "(?<!::\b)\bradixTierOf\s*\("
if ($nStruct -ne 1) { $bad += ("R1 struct RadixTier found " + $nStruct + " times (want 1)") }
if ($nDef -ne 1)    { $bad += ("R1 radixTierOf definition found " + $nDef + " times (want exactly 1 - the tier question has one owner)") }
if ($nTokDef -ne 1) { $bad += ("R1 Lexer::makeRadixToken definition found " + $nTokDef + " times (want 1)") }
if ($nTokDec -ne 1) { $bad += ("R1 makeRadixToken declaration in lexer.hpp found " + $nTokDec + " times (want 1 - no definition = cl C2146/C2039 in the driver)") }
if ($nCallOf -lt 1) { $bad += "R1 makeRadixToken no longer asks radixTierOf (a second copy of the ladder is back)" }

# --- R2 三条支路各问一次, 且当场答 allowNarrowFold ---
$nRet = CountMatches $lex "return\s+makeRadixToken\s*\("
# 调用跨两行: 每个 return 那行与它的下一行拼起来再看这一格答没答
$foldSites = 0
for ($i = 0; $i -lt $lex.Count; $i++) {
    if ($lex[$i] -match "return\s+makeRadixToken\s*\(") {
        $win = ($lex[$i..([Math]::Min($i + 1, $lex.Count - 1))] -join " ")
        if ($win -match "allowNarrowFold\s*=") { $foldSites++ }
    }
}
if ($nRet -ne 3) { $bad += ("R2 radix scanners return through makeRadixToken " + $nRet + " times (want 3: &H + &O + &B)") }
if ($foldSites -ne 3) { $bad += ("R2 only " + $foldSites + " of the 3 call sites answer allowNarrowFold= (want 3) - an un-asked suffix arm reads the wrong sign") }

# --- R3 旧的抄写不许回来 ---
$nLadder = CountMatches $lex "<=\s*0xFFFFFFFFu"
# 十进制那支 (scanNumber) 本来就有自己的 intValue 装法, 不许拿它当红; 所以先把切片
# 起在 scanHexNumber 那一行 —— 三条进制支路里再出现任何 tok.intValue = 都是第二份抄写。
$hexStart = -1
for ($i = 0; $i -lt $lex.Count; $i++) { if ($lex[$i] -match "^Token\s+Lexer::scanHexNumber") { $hexStart = $i; break } }
if ($hexStart -lt 0) { $bad += "R3 scanHexNumber header not found - the radix slice cannot be located" }
$nOldCast = CountMatches $lex[([Math]::Max($hexStart, 0))..($lex.Count - 1)] "tok\.intValue\s*="
$nOldCond = CountMatches $lex "intSuffix\s*\|\|"
if ($nLadder -ne 1) { $bad += ("R3 the 32-bit ladder test appears " + $nLadder + " times (want 1, inside radixTierOf only)") }
if ($nOldCast -ne 0) { $bad += ("R3 the old per-branch truncation is back " + $nOldCast + " time(s)") }
if ($nOldCond -ne 0) { $bad += ("R3 a branch decides the tier itself again (" + $nOldCond + " `intSuffix ||` left)") }

# --- R4 16 位那一支的三个必要件 ---
foreach ($tok in @("0xFFFFu", "0x8000u", "0x10000")) {
    $n = CountMatches $lex ([regex]::Escape($tok))
    if ($n -ne 1) { $bad += ("R4 the 16-bit arm needs " + $tok + " exactly once (found " + $n + ")") }
}

# --- R5 判据在场 + 哨兵自己登记 ---
if (-not $fix.Contains("RT-store=") -or -not $fix.Contains("&HBF32")) {
    $bad += "R5 tests\b591_radixtier.bas no longer carries the &HBF32 store probe"
}
$nBase = ([regex]::Matches($rt, 'Add-BasTest "b591_radixtier"')).Count
$nX86  = ([regex]::Matches($rt, 'Add-BasTest "b591_radixtier_x86"')).Count
$nNote = ([regex]::Matches($rt, 'Test-CodegenNote "radix_tier_shape"')).Count
# 「本脚本在场」与「本哨兵被调用」是两格: 前者写在 Test-RadixTierSites 的**体内** (那一行才是真去起
# powershell 的地方), 后者是 Enter-StaticShard 那一站。只数脚本名的话, 撤掉调用点还剩体内那一处,
# 于是「写完却没人调的哨兵」会报绿 —— 负控 F/E 撞出来的 (定义与调用点分开钉, 同 R1)。
$nRun  = ([regex]::Matches($rt, 'check_radix_tier_sites\.ps1')).Count
$nCall = ([regex]::Matches($rt, 'Enter-StaticShard\)\s*\{\s*Test-RadixTierSites\s*\}')).Count
if ($nBase -ne 1) { $bad += ("R5 default-arch fixture registered " + $nBase + " times (want 1)") }
if ($nX86 -ne 1)  { $bad += ("R5 x86 twin registered " + $nX86 + " times (want 1 - the gate only runs the default arch)") }
if ($nNote -ne 1) { $bad += ("R5 emit-shape needle radix_tier_shape registered " + $nNote + " times (want 1)") }
if ($nRun -ne 1)  { $bad += ("R5 the wrapper runs this script " + $nRun + " times (want 1)") }
if ($nCall -ne 1) { $bad += ("R5 this sentinel's call site `Enter-StaticShard { Test-RadixTierSites }` found " + $nCall + " times (want 1 - a sentinel nobody calls is a sentinel that cannot fail)") }

if ($bad.Count -eq 0) {
    Write-Host ("PASS radix tier: one owner (struct+radixTierOf+makeRadixToken), " + $nRet +
                " scanners all answer allowNarrowFold, ladder text 1 copy, 16-bit arm armed, " +
                "fixture + both arches + emit needle + call site in place") -ForegroundColor Green
    exit 0
}
foreach ($b in $bad) { Write-Host ("FAIL " + $b) -ForegroundColor Red }
exit 1
