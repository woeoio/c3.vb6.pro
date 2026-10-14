# check_com_sig_collision_policy.ps1 - 账 #245 / §B94 最后一刀的结构性哨兵 (只扫源码, 不起 cl)
#
# 为什么单开这一道: `Symbol::comMethods` 是全仓唯一一枚「类型库里这枚成员是什么型」的答案表,
# 而它的写侧只有一条政策 (getter 永远赢 over setter)。剩下的那一格是「**其他**覆盖谁赢」——
# 赢家决定两件事: 发 vb6_ComGet* 的哪一种 (返回档) 与 ByRef 对象出参探测用哪份形参表。
# 实测 (临时工作树插 stderr 探针跑全 396 份输入): 覆盖共两类 —— blessed 那一向 put->get 21930 次、
# 同类 put->put 994 次 (put_Item 与 putref_Item 折成同一个小写键, 读形三项全同) —— 两类都不许响;
# 「读形真的变了又不是 blessed」那一种 0 次, 所以产品侧只欠一声哨子 + 把政策写在出口旁边。
# 这一格形状特殊: 没有实物 ⇒ 行为面钉不住, 只能钉「政策住在哪、响声的条件是什么、写侧没长第二处」。
#
# 规则 (改坏了会红, 不是装饰):
#   W1  写侧唯一: 全 src\ 里 `comMethods[...] =` 的赋值恰好 1 处, 就在 insertComMethod 内;
#       同一份文件里 clear / erase / emplace / insert 对这枚表必须一处都没有。
#   W1b 第二张表在册: `comSourceMethods[...] =` 的赋值恰好 1 处 (事件表今天不吃这条政策 ——
#       哪天它也要, 必须先回来把口径并进来再登记, 不许悄悄多一处写)。
#   W2  政策①还在: insertComMethod 体内必须有「incoming 是 setter 且表里已有 getter ⇒ return」那道闸。
#   W3  哨子的条件: 必须问 comSigReadShapeDiffers(it->second, sig), 且同一条件里带 !blessed,
#       且机器可读的标签就是 "C3: COMSIG-AMBIG (改名字要先想清楚谁在 grep 它)。
#   W4  「读形」三项真问: 助手体内必须比 isPropertyGet / returnType / params.size(),
#       形参循环里必须比 type 与 isByVal (ByRef 出参探测读的就是这两样)。
#   W5  不许退化成无条件响: 助手必须以 return false; 收尾, 且 blessed 必须真用 incomingGetter。
#   W6  调用点 census: insertComMethod(*sym, 的调用恰好 3 处 (coclass 默认接口 / ComInterface /
#       ComGlobalNs 提升) —— 多一处就得回来问「这张表还归谁写」。
#
# 用法: powershell -File scripts/check_com_sig_collision_policy.ps1   (PASS ⇒ exit 0)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$wPath = Join-Path $root 'src\driver\driver_semantics.cpp'
if (-not (Test-Path $wPath)) { Write-Host "FAIL missing file: $wPath"; exit 1 }

$w = Get-Content -Raw -Encoding UTF8 $wPath

function StripComments([string]$s) {
    (($s -replace '//[^\r\n]*', '') -replace '/\*[\s\S]*?\*/', '')
}

$fails = @()

# ---- W1 / W1b: 两张表的写侧各恰好一处 ----
$srcAll = @(Get-ChildItem -Path (Join-Path $root 'src') -Recurse -File -Include *.cpp,*.inc,*.hpp)
$mainFiles = @{}; $srcFiles = @{}
$mainCount = 0; $srcCount = 0
foreach ($f in $srcAll) {
    $txt = StripComments (Get-Content -Raw -Encoding UTF8 $f.FullName)
    $cm = @([regex]::Matches($txt, 'comMethods\s*\[[^\]]*\]\s*=[^=]'))
    $cs = @([regex]::Matches($txt, 'comSourceMethods\s*\[[^\]]*\]\s*=[^=]'))
    $mainCount += $cm.Count; $srcCount += $cs.Count
    if ($cm.Count -gt 0) { $mainFiles[$f.Name] = $cm.Count }
    if ($cs.Count -gt 0) { $srcFiles[$f.Name] = $cs.Count }
}
if ($mainCount -ne 1 -or $mainFiles.Keys.Count -ne 1 -or @($mainFiles.Keys)[0] -ne 'driver_semantics.cpp') {
    $fails += ('W1 comMethods must be written by exactly ONE statement, in insertComMethod; ' +
               'found ' + $mainCount + ' assignment(s) across ' + (($mainFiles.GetEnumerator() |
               ForEach-Object { $_.Name + '=' + $_.Value }) -join ', '))
}
if ($srcCount -ne 1 -or $srcFiles.Keys.Count -ne 1 -or @($srcFiles.Keys)[0] -ne 'driver_semantics.cpp') {
    $fails += ('W1b the event table comSourceMethods must keep exactly one writer; found ' + $srcCount +
               ' assignment(s) across ' + (($srcFiles.GetEnumerator() |
               ForEach-Object { $_.Name + '=' + $_.Value }) -join ', '))
}
$wCode = StripComments $w
if ($wCode -match '\.comMethods\.(clear|erase|emplace|insert)\b') {
    $fails += 'W1 a second mutation of comMethods appeared (clear/erase/emplace/insert) - the single-writer rule is gone'
}

# ---- 取 insertComMethod 的函数体 (懒配到第一个顶格 }) ----
$iIdx = $w.IndexOf('void insertComMethod(Symbol& sym,')
if ($iIdx -lt 0) {
    $fails += 'W2 insertComMethod (the single writer) is gone from driver_semantics.cpp'
} else {
    $iTail = $w.IndexOf("`n}", $iIdx)
    $iBlk = if ($iTail -gt 0) { $w.Substring($iIdx, $iTail - $iIdx) } else { $w.Substring($iIdx) }
    $iCode = StripComments $iBlk
    # ---- W2: getter 永远赢 over setter 那道闸 ----
    if ($iCode -notmatch 'if\s*\(\s*incomingSetter\s*&&\s*it\s*!=\s*sym\.comMethods\.end\(\)\s*&&\s*it->second\.isPropertyGet\s*\)') {
        $fails += 'W2 the getter-beats-setter gate is gone (a put signature must not replace a getter - that is how BSTR props lost their type)'
    }
    if ($iCode -notmatch 'return;') {
        $fails += 'W2 the getter-beats-setter gate no longer returns (it only looks at the flags)'
    }
    # ---- W3: 哨子的条件与标签 ----
    if ($iCode -notmatch 'comSigReadShapeDiffers\(\s*it->second,\s*sig\s*\)') {
        $fails += 'W3 the whistle no longer asks whether the read shape changed (any overwrite would fire - 22924 lines of noise measured on the corpus)'
    }
    if ($iCode -notmatch '!blessed') {
        $fails += 'W3 the blessed put->get direction is no longer excluded from the whistle'
    }
    if ($iCode -notmatch 'C3: COMSIG-AMBIG') {
        $fails += 'W3 the machine-readable tag "C3: COMSIG-AMBIG" is gone (renamed or dropped)'
    }
    # ---- W5: blessed 真用 incomingGetter ----
    if ($iCode -notmatch 'bool incomingGetter = \(kind == ComMemberKind::PropertyGet\)') {
        $fails += 'W5 blessed is no longer computed from the incoming kind'
    }
}

# ---- W4 / W5: 助手本身 ----
$hIdx = $w.IndexOf('bool comSigReadShapeDiffers(')
if ($hIdx -lt 0) {
    $fails += 'W4 comSigReadShapeDiffers is gone - the "does this overwrite change anything a reader uses?" question must live in one place'
} else {
    $hTail = $w.IndexOf("`n}", $hIdx)
    $hBlk = if ($hTail -gt 0) { $w.Substring($hIdx, $hTail - $hIdx) } else { $w.Substring($hIdx) }
    $hCode = StripComments $hBlk
    foreach ($need in @('isPropertyGet\s*!=', 'returnType\s*!=', 'params\.size\(\)\s*!=', 'type\s*!=', 'isByVal\s*!=')) {
        if ($hCode -notmatch $need) {
            $fails += ('W4 comSigReadShapeDiffers no longer compares ' + $need + ' (the whistle would then fire on, or stay silent for, the wrong set of overwrites)')
        }
    }
    $tailLine = (($hCode -split "`r?`n" | Where-Object { $_ -match '\S' } | Select-Object -Last 1)).Trim()
    if ($tailLine -notmatch 'return false;') {
        $fails += 'W5 comSigReadShapeDiffers no longer ends in return false (an unconditional true makes the whistle pure noise)'
    }
}

# ---- W6: 写侧调用点 census ----
$callCount = ([regex]::Matches($wCode, 'insertComMethod\(\s*\*sym,')).Count
if ($callCount -ne 3) {
    $fails += ('W6 insertComMethod(*sym, appears ' + $callCount + ' times; exactly 3 typelib ingestion sites are registered ' +
               '(coclass default interface / ComInterface / ComGlobalNs promotion)')
}

if ($fails.Count -gt 0) {
    $fails | ForEach-Object { Write-Host "FAIL $_" }
    exit 1
}
Write-Host 'OK check_com_sig_collision_policy: W1 W1b W2 W3 W4 W5 W6'
exit 0
