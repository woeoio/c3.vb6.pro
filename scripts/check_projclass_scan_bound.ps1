# check_projclass_scan_bound.ps1 - 工程类登记表: 「扫到哪儿」只许一处答, 且本表现在没人读
#
# 起因 (2026-10-15, 账 #305 §B187 / 第七十九刀): vb6comserver_obj.c 里那张 512 格的
# g_vb6_projClass 上, 同一个「扫到哪儿」有两个答案 —— 登记腿那圈只按 g_vb6_projClassCount
# 走, 查表腿那圈另外夹住 VB6_PROJCLASS_MAX; 而计数在夹界检查之前自增、放弃登记那一路不回退
# ⇒ 计数无上界, 从第 514 枚实例起登记腿就在读数组外面的相邻静态数据 (巧合命中还会往外面写)。
# 本刀把两圈收成一处 vb6_ProjClassScanLimit()。另两条在册事实也一并钉住: 那句指向
# vb6_ProjClassVerify 的注释描述的机制从来不存在, 而当时查表的读者是 0 枚 ——
# 「撤掉整格 vs 把读者接上」是口径待拍 (§B187 出路 a/b), 所以哪天有人接上读者,
# 这一道要红一声, 逼他先回台账把那一格读完。
# 2026-10-22 第一百七十一刀 (新账 #407 / §B190) 接上了第一枚读者, 红声按原话起过作用:
# 台账那两条读数读完的答案是走 (b) 档 —— 视图进装箱站按指针查身份 vb6_VariantFromProjectInstance。
# 它只读登记腿本来就发的那一档 (P5 的 isDll_ 夹一个字没动), 所以这张表从今天起不再是
# 「写了没人读」。 (b) 档原话还配一句「注销腿」, 至今没有: 表是弱映射, 实例释放后地址留在表里
# —— 读者为 0 时那条是惰性缺陷, 读者为 1 起它是活风险, 欠项已回台账记着。
#
# 规则 (改坏了会红, 不是装饰):
#   P1  两圈扫描都问 vb6_ProjClassScanLimit(), 且那个形状在文件里恰好 2 处。
#   P2  任何 for 条件都不许再直接拿 g_vb6_projClassCount 当边界 (代码里 0 处)。
#   P3  那枚夹界出口: 定义恰好 1 处; VB6_PROJCLASS_MAX 在代码里恰好 5 处(住在 4 行上:
#       #define / 数组尺寸 / 出口里那行夹界(同一行点了它两次) / 放弃登记那一句), 后者恰好 1 处。
#   P5  登记腿只许 ActiveX 那一档发: cgen_com.cpp 里恰好 1 处, 且它上面一行就是 if (isDll_) {
#   P4  读者普查: vb6_FindProjectClassDesc( 在 .c 的代码里恰好 3 处 (定义 + ...Raw 的转调 +
#       第一百七十一刀那枚真读者)、在 .h 里至多 1 处 (就那一枚声明); vb6_ProjClassVerify 在代码里 0 处 (它不存在,
#       只许住在注释里)。
#
# 用法: powershell -File scripts\check_projclass_scan_bound.ps1 [-Root <副本目录>]  (PASS = exit 0)

param([string]$Root = "")

$ErrorActionPreference = "Stop"
if ($Root -eq "") { $Root = Split-Path -Parent $PSScriptRoot }

$obj = Join-Path $Root "src\rtl\core\vb6comserver\vb6comserver_obj.c"
$hdr = Join-Path $Root "src\rtl\core\vb6comserver\vb6comserver.h"
foreach ($p in @($obj, $hdr)) {
    if (-not (Test-Path -LiteralPath $p)) { Write-Host ("FAIL missing path: " + $p); exit 1 }
}

# 按**行**筛出代码文本: 以 // 、/* 、* 开头的行整行不算代码, 其余行裁掉行内 // 之后才算。
# (整块 /**/ 拿正则删会把行合并 —— 这一族在别的哨兵上红过一次, 所以这里不用那种写法;
#  而不裁注释 ⇒ 一句解释机制的注释会被数成一次调用, 两条都得躲开。)
function Get-CodeText187([string]$text) {
    $kept = New-Object System.Collections.Generic.List[string]
    foreach ($line in ($text -split "\r?\n")) {
        $t = $line.TrimStart()
        if ($t.StartsWith("//")) { continue }
        if ($t.StartsWith("/*")) { continue }
        if ($t.StartsWith("*")) { continue }
        $idx = $line.IndexOf("//")
        if ($idx -ge 0) { $line = $line.Substring(0, $idx) }
        $kept.Add($line)
    }
    return ($kept -join "`n")
}

$objCode = Get-CodeText187 (Get-Content -Raw -Encoding UTF8 $obj)
$hdrCode = Get-CodeText187 (Get-Content -Raw -Encoding UTF8 $hdr)

$bad = @()

# ---- P1: 两圈都问那一枚出口 ----
$loopRx = 'for \(i = 0; i < vb6_ProjClassScanLimit\(\); i\+\+\)'
$loops = @([regex]::Matches($objCode, $loopRx)).Count
if ($loops -ne 2) {
    $bad += ("P1 scan loops that ask vb6_ProjClassScanLimit() = " + $loops + " (want exactly 2: register + lookup)")
}

# ---- P2: 边界不许再直接问计数 ----
$rawRx = 'i\s*<\s*g_vb6_projClassCount'
$raw = @([regex]::Matches($objCode, $rawRx)).Count
if ($raw -ne 0) {
    $bad += ("P2 a scan loop still bounds itself on the raw g_vb6_projClassCount (" + $raw +
             ") - that counter has no upper bound, so the bound reads past the array")
}

# ---- P3: 夹界出口自己只住一处 ----
$defn = @([regex]::Matches($objCode, 'static LONG vb6_ProjClassScanLimit\(void\) \{')).Count
if ($defn -ne 1) { $bad += ("P3 vb6_ProjClassScanLimit definitions = " + $defn + " (want 1)") }
$capUse = @([regex]::Matches($objCode, 'VB6_PROJCLASS_MAX')).Count
if ($capUse -ne 5) {
    $bad += ("P3 VB6_PROJCLASS_MAX uses in code = " + $capUse +
             " (want 5 on 4 lines: #define + array size + the clamp line (it names the cap twice) + the bail-out)")
}
$giveUp = @([regex]::Matches($objCode, 'i < 0 \|\| i >= VB6_PROJCLASS_MAX')).Count
if ($giveUp -ne 1) { $bad += ("P3 the table-full bail-out = " + $giveUp + " (want 1)") }

# ---- P4: 读者普查 + 那句不存在的机制 ----
$callRx = 'vb6_FindProjectClassDesc\s*\('
$objCalls = @([regex]::Matches($objCode, $callRx)).Count
$hdrCalls = @([regex]::Matches($hdrCode, $callRx)).Count
# .c 里那 3 处 = 定义本身 + ...Raw 里那句转调 (它自己至今 0 调用者) + 第一百七十一刀那枚真读者。
# 从 <=2 收成 ==3 是这一刀接读者换来的: 再有人接第二枚读者, 这一道照样红。
if ($objCalls -ne 3 -or $hdrCalls -gt 1) {
    $bad += ("P4 the project-class lookup gained readers (code mentions: .c=" + $objCalls +
             " .h=" + $hdrCalls + ", want exactly 3 / <=1: 定义 + 那枚至今 0 调用者的 ...Raw + 第一百七十一刀那枚读者) - §B187 那条「接读者」按台账分档走 (b),")
    $bad += ("     接之前先回台账读那两条读数 (2026-10-15 拍定的分档: ActiveX DLL/Control 走 (b) = 接读者 + 配注销腿,")
    $bad += ("     其余走 (a) = 连登记都不发; 而「EXE 里没有 CoClass 块 ⇒ 实例没进表」那句旧解释已被 com_entry.c 推翻")
}
$verify = @([regex]::Matches($objCode, 'vb6_ProjClassVerify')).Count
if ($verify -ne 0) {
    $bad += ("P4 vb6_ProjClassVerify appears " + $verify +
             " times in code text - that function does not exist; the claim belongs in a comment at most")
}

# ---- P5 (账 #305 §B187 第二格, 2026-10-15 拍定的分档) ----
# 登记腿只许在 ActiveX 那一档发: 那一句 emit 必须恰好 1 处, 且它上面一行就是 if (isDll_) { 。
# 撤掉这层夹 = 非 ActiveX 工程又开始喂那张表 (读者 0 枚 ⇒ 白付每实例一格), 或者两处各发一遍。
$cgenRel = 'src\backend\module\cgen_com.cpp'
$cgen = Join-Path $Root $cgenRel
if (-not (Test-Path -LiteralPath $cgen)) {
    $bad += ("P5 missing " + $cgenRel + " - the registration leg's project-type gate is not where B187 says it is")
} else {
    $cgenAll = Get-CodeText187 (Get-Content -Raw -Encoding UTF8 $cgen)
    $regAll = @([regex]::Matches($cgenAll, 'vb6_RegisterProjectClassInstance')).Count
    if ($regAll -ne 1) {
        $bad += ("P5 registration-emit sites = " + $regAll + " (want exactly 1)")
    }
    $cl = @($cgenAll -split "`n")
    $at = -1
    for ($j = 0; $j -lt $cl.Count; $j++) { if ($cl[$j] -match 'vb6_RegisterProjectClassInstance') { $at = $j; break } }
    $gated = ($at -ge 2 -and ($cl[$at - 1] -match 'if \(isDll_\) \{' -or $cl[$at - 2] -match 'if \(isDll_\) \{'))
    if ($at -ge 1 -and -not $gated) {
        $bad += ("P5 the registration leg is not gated by the single project-type answer (want if (isDll_) { within the 2 lines above) - reads: " + $cl[$at - 1].Trim())
    }
}

if ($bad.Count -gt 0) {
    foreach ($b in $bad) { Write-Host ("FAIL " + $b) }
    exit 1
}
Write-Host ("PASS check_projclass_scan_bound: loops=2 raw-bound=0 helper=1 cap=5/4lines bail=1 c-mentions=3(readers=1) verify-in-code=0 reg-site=1 gated=isDll_")
exit 0
