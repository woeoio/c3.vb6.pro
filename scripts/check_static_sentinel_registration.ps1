# check_static_sentinel_registration.ps1 - census: 哨兵的存在面 == 登记面（§B95 欠的那一句）
#
# 起因（2026-10-08）：把 scripts/ 里所有 check_*.ps1 逐道跑了一遍，**35 绿 1 红，而门是全绿的**。
# 数下去发现 scripts/ 有 36 道，而 tests/run_tests.ps1 只引用了 32 道 —— 那四道从写下那天起
# 就没被门跑过一次。"哨兵住在 scripts/ 里" 从来不是覆盖面，**登记那一行才是**，
# 而登记这件事本身没人管 —— 这一道就是把那句"没人管"变成一句可判红的话。
#
# 规则（改坏了会红，不是装饰）：
#   R1  scripts/check_*.ps1 的每一道都必须被 tests/run_tests.ps1 按**文件名**引用一次以上。
#       新做哨兵忘记登记 = 这一格当场红，并点名是哪一道。
#   R1b 引用某道哨兵的 Test-* 函数必须还有一处**调用**（只有定义没调用 = 门永远不跑那一格）。
#   R2  run_tests.ps1 引用的每个 check_*.ps1 都必须真在 scripts/ 里存在。
#       登记了一个不存在的文件 = 那一格永远看不见失败（与 Test-Compile 重名遮蔽那一族同形）。
#   R3  两边的**条数**也要报出来（红的时候好定位：是漏登记还是引用打错了名字）。
#
#   R4  每一份 check_*.ps1 的**文件头**必须能被 PS 5.1 当注释读: UTF-8 BOM 恰好一次,
#       去掉 BOM 后首行以 '#' 开头。(本轮实测撞过: 补丁脚本把 BOM 写了两遍 ⇒ 首行成了
#       "?# …" ⇒ PS 报 CommandNotFound 却**继续跑到底、exit 0** —— 一道绿着的坏哨兵,
#       比红更难发现; R1/R2/R3 只看"有没有被登记", 看不见"登记了却跑不起来"。)
# 用法: powershell -File scripts\check_static_sentinel_registration.ps1   (PASS = exit 0)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$scriptsDir = Join-Path $root "scripts"
$harness = Join-Path $root "tests\run_tests.ps1"

foreach ($p in @($scriptsDir, $harness)) {
    if (-not (Test-Path -LiteralPath $p)) { Write-Host ("FAIL missing path: " + $p); exit 1 }
}

$onDisk = @((Get-ChildItem -Path $scriptsDir -File -Filter "check_*.ps1" | ForEach-Object { $_.Name }) | Sort-Object)
$harnessText = Get-Content -Raw -Encoding UTF8 $harness
$referenced = @(([regex]::Matches($harnessText, 'check_[A-Za-z0-9_]+\.ps1') |
                 ForEach-Object { $_.Value }) | Sort-Object -Unique)

$bad = @()

# ---- R1: 盘上的每一道都得被引用 ----
$missing = @($onDisk | Where-Object { $referenced -notcontains $_ })
if ($missing.Count -gt 0) {
    $bad += ("R1 sentinels exist in scripts/ but are NOT registered in tests/run_tests.ps1 (the gate never runs them): " +
             ($missing -join ", "))
}

# ---- R1b: 光在 function 体里出现一次不算登记 —— 引用它的那个 Test-* 函数还得有调用点 ----
#（起因：写这一道时拿「删掉调用行」当假改动，R1 居然还是绿的 —— 因为函数体里那句
#  `-File ...\check_x.ps1` 本身就是一次引用。**覆盖面是跑不跑，不是提没提**。）
$fnRx = [regex]'(?m)^function\s+(Test-[A-Za-z0-9_]+)\s*\{[\s\S]*?\r?\n\}'
$covered = @{}
foreach ($m in $fnRx.Matches($harnessText)) {
    $name = $m.Groups[1].Value
    $refs = @([regex]::Matches($m.Value, 'check_[A-Za-z0-9_]+\.ps1') | ForEach-Object { $_.Value })
    if ($refs.Count -eq 0) { continue }
    $callRx = '(?m)^\s+' + [regex]::Escape($name) + '\s*$'
    if (@([regex]::Matches($harnessText, $callRx)).Count -lt 1) {
        $bad += ('R1b ' + $name + ' runs ' + ($refs -join ', ') + ' but the harness never calls it (dead cell)')
        continue
    }
    foreach ($r in $refs) { $covered[$r] = $true }
}
$uncalled = @($onDisk | Where-Object { -not $covered.ContainsKey($_) })
if ($uncalled.Count -gt 0) {
    $bad += ('R1b sentinels not covered by any CALLED Test-* function: ' + ($uncalled -join ', '))
}

# ---- R2: 引用的每一道都得真在盘上 ----
$dangling = @($referenced | Where-Object { $onDisk -notcontains $_ })
if ($dangling.Count -gt 0) {
    $bad += ("R2 tests/run_tests.ps1 references check_*.ps1 files that do not exist (that cell can never fail): " +
             ($dangling -join ", "))
}

# ---- R3: 条数对齐（前两条已经各自会红，这里只把数报出来）----
if ($bad.Count -eq 0 -and $onDisk.Count -ne $referenced.Count) {
    $bad += ("R3 count mismatch: scripts/ has " + $onDisk.Count +
             " check_*.ps1 while the harness references " + $referenced.Count)
}

# ---- R4: 文件头读不成注释 = 哨兵自己废了还会绿 ----
$badHdr = @()
foreach ($f in (Get-ChildItem -Path $scriptsDir -File -Filter "check_*.ps1")) {
    $b = [System.IO.File]::ReadAllBytes($f.FullName)
    if ($b.Length -lt 4 -or $b[0] -ne 0xEF -or $b[1] -ne 0xBB -or $b[2] -ne 0xBF) {
        $badHdr += ($f.Name + " (no UTF-8 BOM)")
        continue
    }
    if ($b[3] -eq 0xEF -and $b[4] -eq 0xBB -and $b[5] -eq 0xBF) {
        $badHdr += ($f.Name + " (BOM twice - line 1 is not a comment)")
        continue
    }
    if ($b[3] -ne 0x23) { $badHdr += ($f.Name + " (line 1 does not start with #)") }
}
if ($badHdr.Count -gt 0) {
    $bad += ("R4 sentinel file headers unreadable as comments (PS errors, then exits 0 anyway): " +
             ($badHdr -join ", "))
}

if ($bad.Count -gt 0) {
    $bad | ForEach-Object { Write-Host ("FAIL " + $_) }
    Write-Host ("DETAIL onDisk=" + $onDisk.Count + " referenced=" + $referenced.Count)
    exit 1
}
Write-Host ("PASS static_sentinel_registration: " + $onDisk.Count +
            " sentinels on disk, all registered in the harness")
exit 0
