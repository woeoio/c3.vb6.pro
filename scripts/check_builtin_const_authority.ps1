# check_builtin_const_authority.ps1 - 账 #218 第三刀的结构哨兵 (只扫源码, 不起 cl)
#
# 为什么单开这一道: 这一格的缺陷形状还是"答案的来源没被守住"，而且比 #245 更隐蔽 ——
# 同一个内在常量在仓里有**三处**各自的答案：语义层的表 (两份 .inc)、发码层逐名折叠 (.inc)、
# RTL 头里的 #define。谁都不知道另外两个存在，于是:
#   · 只在 RTL 格里那 14 枚 (vbPicType*/vbHitResult*/vbAsync*) 语义层不认识 ⇒ 源文件里
#     `p = vbPicTypeBitmap` 被当未声明标识符建**隐式 Variant 局部**，运行期交出 0 (实测)；
#   · vbUseSystem 在 ext 表里记 -1、在折叠里记 0，而 VB6 手册 (docs/vb6-manual/09-常数/Date 常数.md)
#     写的是 0 ⇒ 同一份源文件里 `Const c = vbUseSystem` 交出 -1、直接读交出 0 (两台编译都实测)。
# 编译不响、链接不响、语料 A/B 只响一半，所以要结构钉。
#
# 规则 (改坏了会红，不是装饰):
#   C1  RTL 里**一枚** `#define vb*` 都不许有 (vb6_ 前缀不算) —— 头文件重新变成某枚常量的唯一
#       来源，就又是"语义层不认识 ⇒ 隐式变量"那一格。
#   C2  发码折叠里每枚**数值**名字都必须在表里 (两份 .inc 的并集) —— 折叠只许是"第二份答案"，
#       不许再是"唯一答案"。这条是 #218 的口径本身。
#   C3  棘轮: 数值折叠的**出现次数**钉死 = 56（= 54 枚不同名字，其中 vbObject 与 vbUseSystemDayOfWeek
#       各被折了两次，是历史遗留的重复行）。要减是改进 (回这里把数改小)，要加是长回第二答案源 (必须先补表)。
#   C4  两份表片段之间不许有重名 —— 本刀第一轮就是插重了 6 枚，每份工程编译都响 VB3002。
#   C5  三枚有实测读数钉在这里: vbPicTypeBitmap=1 / vbHitResultHit=2 / vbUseSystem=0 (手册值)。
#       值被人改了必须同时改这条，不许静默漂。
#   C6  非数值折叠 (交出 BSTR/NULL 表达式那一类) 数目钉死 = 11，口径同上: 加名字必须先让它进表。
#
# 用法: powershell -File scripts/check_builtin_const_authority.ps1   (PASS => exit 0)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot

$foldPath = Join-Path $root 'src\backend\detail\expr\cgen_expr_ident_dispatch.inc'
$rtlDir   = Join-Path $root 'src\rtl'
$tblPaths = @(
    (Join-Path $root 'src\semantics\builtin\builtin_consts.inc'),
    (Join-Path $root 'src\semantics\builtin\builtin_consts_ext.inc')
)
foreach ($p in (@($foldPath) + $tblPaths)) {
    if (-not (Test-Path $p)) { Write-Host "FAIL missing file: $p"; exit 1 }
}

$fails = @()

# 先脱掉 //… 与 /*…*/: 判定之前必须只看代码，注释里就写着 "vbUseSystem" 这类名字。
function Strip-Comments([string]$s) {
    $noBlock = [regex]::Replace($s, '/\*[\s\S]*?\*/', '')
    return [regex]::Replace($noBlock, '//[^\r\n]*', '')
}

# ---- 表: 两份 .inc 的并集, 名字 -> 值 (按出现次数一起记, C4 要看重复) ----
$tableAll = @{}
$perFile  = @{}
foreach ($tp in $tblPaths) {
    $txt = Strip-Comments (Get-Content -Raw -Encoding UTF8 $tp)
    $ms = [regex]::Matches($txt, 'add\w*Const\w*\(\s*"([A-Za-z_][A-Za-z0-9_]*)"\s*,\s*Vb6Type::\w+\s*,\s*(-?\d+)(?:LL|ull|UI|ui)?\s*\)')
    $seen = @{}
    foreach ($m in $ms) {
        $k = $m.Groups[1].Value.ToLower()
        $v = [int64]$m.Groups[2].Value
        if ($seen.ContainsKey($k)) { $fails += ('C4 ' + (Split-Path -Leaf $tp) + ' declares ' + $k + ' twice') }
        $seen[$k] = $v
        if ($tableAll.ContainsKey($k)) {
            if ($tableAll[$k] -ne $v) {
                $fails += ('C4 name ' + $k + ' has two values across the table fragments: ' + $tableAll[$k] + ' / ' + $v)
            }
        } else { $tableAll[$k] = $v }
    }
    $perFile[(Split-Path -Leaf $tp)] = $seen
}
$dupAcross = 0
$leaves = @($tblPaths | ForEach-Object { (Split-Path -Leaf $_) })
$setA = $perFile[$leaves[0]]
$setB = $perFile[$leaves[1]]
if ($setA -and $setB) {
    $common = @($setA.Keys | Where-Object { $setB.ContainsKey($_) })
    $dupAcross = @($common).Count
    if ($dupAcross -gt 0) {
        $fails += ('C4 the two table fragments share ' + $dupAcross + ' name(s): ' + (($common | Select-Object -First 6) -join ', ') +
                   ' - one name must have one home (a duplicate is a VB3002 on every project build)')
    }
} else {
    $fails += ('C4 could not read both table fragments: ' + ($leaves -join ', '))
}

# ---- C1: RTL 里不许有 #define vb* (vb6_ 前缀不算) ----
$rtlMacros = @()
Get-ChildItem -Path $rtlDir -Recurse -File -Include *.h,*.c,*.inc | ForEach-Object {
    $t = Strip-Comments (Get-Content -Raw -Encoding UTF8 $_.FullName)
    foreach ($m in [regex]::Matches($t, '(?m)^\s*#\s*define\s+(vb(?!6_)[A-Za-z0-9_]+)\s')) {
        $rtlMacros += ($_.Name + ':' + $m.Groups[1].Value)
    }
}
if ($rtlMacros.Count -gt 0) {
    $fails += ('C1 RTL defines vb* constants again (' + $rtlMacros.Count + '): ' + ($rtlMacros -join ', ') +
               ' - the table in src\semantics\builtin is the only answer source')
}

# ---- C2/C3/C6: 发码折叠 ----
$foldTxt = Strip-Comments (Get-Content -Raw -Encoding UTF8 $foldPath)
$msNum = [regex]::Matches($foldTxt, 'lower\s*==\s*"(vb[A-Za-z0-9_]+)"\s*\)\s*\{\s*lastExpr_\s*=\s*"\(?(-?\d+)\)?";')
$msAny = [regex]::Matches($foldTxt, 'lower\s*==\s*"(vb[A-Za-z0-9_]+)"\s*\)\s*\{\s*lastExpr_\s*=')
$orphan = @()
foreach ($m in $msNum) {
    $k = $m.Groups[1].Value.ToLower()
    if (-not $tableAll.ContainsKey($k)) { $orphan += $k }
    elseif ($tableAll[$k] -ne [int64]$m.Groups[2].Value) {
        $fails += ('C2 fold for ' + $k + ' says ' + $m.Groups[2].Value + ' but the table says ' + $tableAll[$k])
    }
}
if ($orphan.Count -gt 0) {
    $fails += ('C2 fold-only intrinsic constant(s) (the table never heard of them): ' + ($orphan -join ', ') +
               ' - put them in src\semantics\builtin\builtin_consts.inc first')
}
$PIN_NUM_FOLD_OCCURRENCES = 56
$PIN_NUM_FOLD_NAMES = 54
$distinctNum = @($msNum | ForEach-Object { $_.Groups[1].Value.ToLower() } | Sort-Object -Unique)
if ($msNum.Count -ne $PIN_NUM_FOLD_OCCURRENCES) {
    $fails += ('C3 ratchet: numeric fold occurrences are ' + $msNum.Count + ', pinned at ' +
               $PIN_NUM_FOLD_OCCURRENCES + ' (fewer = good, come update the number; more = a second ' +
               'answer source grew back)')
}
if ($distinctNum.Count -ne $PIN_NUM_FOLD_NAMES) {
    $fails += ('C3 ratchet: distinct numeric fold names are ' + $distinctNum.Count + ', pinned at ' +
               $PIN_NUM_FOLD_NAMES + ' (vbObject / vbUseSystemDayOfWeek are the two folded twice)')
}
$PIN_STR_FOLDS = 11
$cntStr = $msAny.Count - $msNum.Count
if ($cntStr -ne $PIN_STR_FOLDS) {
    $fails += ('C6 ratchet: expression-shaped folds (BSTR/NULL) are ' + $cntStr + ', pinned at ' + $PIN_STR_FOLDS)
}

# ---- C5: 三枚有手册/实测读数的值 ----
$must = @{ 'vbpictypebitmap' = 1; 'vbhitresulthit' = 2; 'vbusesystem' = 0 }
foreach ($k in ($must.Keys | Sort-Object)) {
    if (-not $tableAll.ContainsKey($k)) { $fails += ('C5 ' + $k + ' is not in the constant table at all'); continue }
    if ($tableAll[$k] -ne $must[$k]) {
        $fails += ('C5 ' + $k + ' = ' + $tableAll[$k] + ' but the measured/manual answer is ' + $must[$k])
    }
}

if ($fails.Count -gt 0) {
    $fails | ForEach-Object { Write-Host "FAIL $_" }
    exit 1
}
Write-Host ('OK check_builtin_const_authority: C1..C5 (table=' + $tableAll.Count +
             ' numeric-folds=' + $msNum.Count + ' expr-folds=' + $cntStr + ' rtl-vb-defines=0)')
exit 0
