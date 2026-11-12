# pd_readout.ps1 - 取 PD 整编 GA 工件并读出这一刀的全部关键数。
#
# §56.53 起 PD 整编有五组读数，缺一不可：
#   1) pd_build.meta            —— BUILD_SECONDS / MIN_FREE_PHYS_MB / PEAK_TREE_WS_MB（总账）
#   2) pd_build.err 里的 C3: MEM 行 —— 发码后释放 AST 的内存账（§56.55 起）
#   3) pd_build.err 里的 C3: TIMING cl:<label> —— cl 四段耗时（§56.54 起）
#   4) pd_build.mem.csv         —— 内存曲线（cl 阶段 c3_ws_mb 应掉下来）
#   5) c3-error.log 逐码比对   —— 错误面（判据唯一口径在 scripts\diff_c3error.ps1）
#
# 只读，不改任何东西；下载落到 .temp\art_<run>\（可随时删）。
#
# 用法：
#   pwsh -NoProfile -File scripts\pd_readout.ps1 -Run 43224386890
#   pwsh -NoProfile -File scripts\pd_readout.ps1 -Run 43224386890 -Baseline 43218740715
#   pwsh -NoProfile -File scripts\pd_readout.ps1            # 不给 -Run 就读 .temp\pd_run.txt
#
# 注意：本脚本含中文，必须用 pwsh（PowerShell 7）跑。
param(
    [string]$Run = "",
    [string]$Baseline = "",
    [string]$Root = "",
    [string]$ArtifactRoot = "",
    [string]$Artifact = "photodemon-full-build-logs",
    [switch]$NoWait,
    [switch]$SkipDownload,
    [int]$WaitMinutes = 90,
    [int]$Tries = 8,
    [int]$SleepSeconds = 10
)
$ErrorActionPreference = 'Continue'
if (-not $Root) { $Root = (Split-Path -Parent $PSScriptRoot) }
$Root = (Resolve-Path $Root).Path
if (-not $ArtifactRoot) { $ArtifactRoot = Join-Path $Root ".temp" }
Set-Location $Root

$runFile = Join-Path $ArtifactRoot 'pd_run.txt'
if (-not $Run -and (Test-Path $runFile)) {
    $t = Get-Content $runFile | Select-Object -First 1
    if ($t) { $Run = $t.Trim() }
}
if (-not $Run) {
    Write-Host "NO RUN: 给了 -Run，或让 scripts\ga_dispatch_pd.ps1 先写下 .temp\pd_run.txt"
    exit 2
}
Write-Host ("PD run = " + $Run)
if ($Baseline) { Write-Host ("基线 run = " + $Baseline) }

function Invoke-Native {
    param([scriptblock]$Sb, [int]$Tries = 8, [int]$Sleep = 10)
    $res = @{ Ok = $false; Rc = -1; Out = @() }
    for ($i = 1; $i -le $Tries; $i++) {
        $out = & $Sb 2>&1
        $rc = $LASTEXITCODE
        $res.Rc = $rc
        $res.Out = @($out)
        if ($rc -eq 0) { $res.Ok = $true; return $res }
        Write-Host ("  [重试 {0}/{1}] rc={2} :: {3}" -f $i, $Tries, $rc, (($out | Select-Object -Last 1) -join ' '))
        Start-Sleep -Seconds $Sleep
    }
    return $res
}

if (-not $NoWait) {
    Write-Host ''
    Write-Host '=== 1) 等 run 结束 ==='
    $deadline = (Get-Date).AddMinutes($WaitMinutes)
    while ((Get-Date) -lt $deadline) {
        $r = Invoke-Native { gh run view $Run --json status,conclusion } $Tries $SleepSeconds
        if ($r.Ok -and $r.Out) {
            try {
                $o = (($r.Out -join "`n") | ConvertFrom-Json)
                Write-Host ("  [{0}] status={1} conclusion={2}" -f (Get-Date -Format HH:mm:ss), $o.status, $o.conclusion)
                if ($o.status -eq 'completed') { break }
            } catch {
                Write-Host ("  [{0}] 状态解析失败: {1}" -f (Get-Date -Format HH:mm:ss), (($r.Out | Select-Object -Last 1) -join ' '))
            }
        }
        Start-Sleep -Seconds 20
    }
}

$art = Join-Path $ArtifactRoot ("art_" + $Run)
$errf = Join-Path $art 'pd_build.err'
$memf = Join-Path $art 'pd_build.mem.csv'
$meta = Join-Path $art 'pd_build.meta'

if (-not $SkipDownload) {
    Write-Host ''
    Write-Host '=== 2) 下载工件 ==='
    if (Test-Path $art) { Remove-Item $art -Recurse -Force }
    New-Item -ItemType Directory -Force -Path $art | Out-Null
    $r = Invoke-Native { gh run download $Run --name $Artifact --dir $art } $Tries $SleepSeconds
    if (-not $r.Ok) {
        Write-Host ("工件下载失败（rc=" + $r.Rc + "）—— 工件可能还没生成，或名字不对：" + $Artifact)
    }
}

Write-Host ''
Write-Host '=== 3) pd_build.meta（总账） ==='
if (Test-Path $meta) {
    Get-Content $meta | ForEach-Object { Write-Host ('  ' + $_) }
} else {
    Write-Host '  (无 meta —— 工件没下来？)'
}

Write-Host ''
Write-Host '=== 4) ★ MEM 行（§56.55 发码后释放 AST 的内存账） ==='
if (Test-Path $errf) {
    $sr = New-Object IO.StreamReader($errf)
    try {
        $n = 0
        while (($line = $sr.ReadLine()) -ne $null) {
            if ($line -like 'C3: MEM*') { Write-Host ('  ' + $line); $n++ }
        }
        if ($n -eq 0) { Write-Host '  (0 条 MEM 行 —— 这一刀没有释放，或这一刀之前的旧二进制)' }
    } finally { $sr.Close() }
} else {
    Write-Host '  (无 pd_build.err)'
}

Write-Host ''
Write-Host '=== 5) ★ cl 四段耗时（§56.54 装的那只表） ==='
# 读数口径（形状已用 §56.55 工件核对过，别凭印象改）：
#   C3: TIMING cl:pch-c3_pch_std 1883 ms (1 TU, 1882.6 ms/TU)
#   C3: TIMING gen <TU名> 18 ms              ← 发码侧，每个 TU 一行
#   C3: TIMING <相位名> 943 ms               ← 顶层相位（名字里可含 + 与 -，**不含 :**）
#   C3: TIMING g:prologue 1727 ms            ← gen 的**子相位**（名字里带 :），已含在 gen 里
# 顶层相位与 g: 子相位分开统计，否则 emit 合计会把 gen 算两遍。
$clMs = @{}
$clOrder = New-Object System.Collections.Generic.List[string]
$phaseMs = @{}
$subMs = @{}
$genMs = 0; $genN = 0
$parseBad = 0
if (Test-Path $errf) {
    $sr = New-Object IO.StreamReader($errf)
    try {
        while (($line = $sr.ReadLine()) -ne $null) {
            if (-not $line.StartsWith('C3: TIMING')) { continue }
            # 逐行容错：解析坏一行不许打断整轮（§56.55 的教训：一行读错就全盘读数空掉）
            try {
                if ($line -match '^C3: TIMING cl:(\S+) ([\d.]+) ms') {
                    $l = $Matches[1]
                    if (-not $clMs.ContainsKey($l)) { $clMs[$l] = 0.0; $clOrder.Add($l) }
                    $clMs[$l] += [double]$Matches[2]
                } elseif ($line -match '^C3: TIMING gen \S+ (\d+) ms$') {
                    $genMs += [int]$Matches[1]; $genN++
                } elseif ($line -match '^C3: TIMING g:(\S+) (\d+) ms$') {
                    $k = $Matches[1]
                    if (-not $subMs.ContainsKey($k)) { $subMs[$k] = 0 }
                    $subMs[$k] += [int]$Matches[2]
                } elseif ($line -match '^C3: TIMING ([A-Za-z0-9+_.-]+) (\d+) ms$') {
                    $k = $Matches[1]
                    if (-not $phaseMs.ContainsKey($k)) { $phaseMs[$k] = 0 }
                    $phaseMs[$k] += [int]$Matches[2]
                } else {
                    $parseBad++
                }
            } catch {
                $parseBad++
            }
        }
    } finally { $sr.Close() }
}
$sum = 0.0
foreach ($l in $clOrder) {
    Write-Host ("  cl:{0,-18} {1,10:N1} ms" -f $l, $clMs[$l])
    $sum += $clMs[$l]
}
Write-Host ("  {0,-24} {1,10:N1} ms  ({2:N1} s)" -f '== cl 四段合计', $sum, ($sum / 1000))
$emitSum = 0
foreach ($k in $phaseMs.Keys) { $emitSum += $phaseMs[$k] }
$subSum = 0
foreach ($k in $subMs.Keys) { $subSum += $subMs[$k] }
Write-Host ("  {0,-24} {1,10:N1} ms  (gen 段 {2} 个 TU, 其中 g: 子相位 {3:N1} ms 已含在 gen 内)" -f `
            'emit 相位合计', ($emitSum + $genMs), $genN, $subSum)
foreach ($k in ($phaseMs.Keys | Sort-Object)) {
    Write-Host ("    顶层相位 {0,-28} {1,8:N0} ms" -f $k, $phaseMs[$k])
}
if ($parseBad -gt 0) { Write-Host ("  ★ 有 " + $parseBad + " 行 TIMING 没认出来（不是上面几类之一）") }

Write-Host ''
Write-Host '=== 6) ★ 内存曲线（cl 阶段 c3_ws_mb 应掉下来） ==='
if (Test-Path $memf) {
    $rows = Import-Csv $memf
    Write-Host ("  采样 " + $rows.Count + " 行；打印前 30 s、每 20 s、以及 c3_ws_mb 变动 >= 500 MB 的点")
    $prev = -1
    foreach ($r in $rows) {
        $s = [int]$r.seconds; $c = [int]$r.c3_ws_mb; $t = [int]$r.tree_ws_mb
        if ($s -le 30 -or $s % 20 -eq 0 -or [Math]::Abs($c - $prev) -ge 500) {
            Write-Host ("    {0,5}s  tree={1,6}MB  c3={2,6}MB" -f $s, $t, $c)
        }
        $prev = $c
    }
    $peakC = ($rows | ForEach-Object { [int]$_.c3_ws_mb } | Measure-Object -Maximum).Maximum
    Write-Host ("  c3_ws_mb 峰值 = " + $peakC + " MB")
    $after = @($rows | Where-Object { [int]$_.seconds -gt 80 } | ForEach-Object { [int]$_.c3_ws_mb })
    if ($after.Count -gt 0) {
        Write-Host ("  80 s 之后 c3_ws_mb: 最大 " + (($after | Measure-Object -Maximum).Maximum) + " MB / 最小 " + (($after | Measure-Object -Minimum).Minimum) + " MB")
    }
} else {
    Write-Host '  (无 mem.csv)'
}

if ($Baseline) {
    Write-Host ''
    Write-Host ('=== 7) 错误面逐码 vs 基线 ' + $Baseline + ' ===')
    & pwsh -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'diff_c3error.ps1') $Baseline $Run
    Write-Host ("  (错误面判据退出码 = " + $LASTEXITCODE + "；0 = 逐码相等)")
} else {
    Write-Host ''
    Write-Host '=== 7) 错误面逐码比对：跳过（没给 -Baseline） ==='
}
Write-Host '=== DONE ==='
