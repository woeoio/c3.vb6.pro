# ga_dispatch_pd.ps1 - 派发 PD 整编 GA（必要时连带 dev 回归），并核对 headSha。
#
# 干的事（顺序固定）：
#   1) 把探针分支（.temp\probe_sync）同步到 github/dev 并推上 github
#      —— **必须推 github，不是 gitcode**（GA workflow 挂在 github 那个远端上）
#   2) 派发 "PhotoDemon x86 full build" 到探针分支
#   3) 按 headSha 认领这一次的 run id（不看"最新那个"，避免认错别人的 run）
#   4) 等 run 跑完，打印结论
#   5) 可选：再派发 "VB6 C3 Regression" 到 dev（形状门那 22 格）
#
# 为什么必须核对 headSha：CI 队列里可能同时有好几个 run；只看"最新"会认到上一次
# 刀留下的 run，读数就串了。headSha 是"这一刀到底编的是哪份代码"的唯一硬证据。
#
# 为什么每一步都带重试：这台机器的 hosts 把 github.com 全家指向 127.0.0.1，
# 本地转发服务时有时无 ⇒ gh/git 会间歇 dial tcp 127.0.0.1:443 被拒或 TLS timeout。
# 所以所有网络动作一律重试循环，不要一次失败就下结论。
#
# 用法：
#   pwsh -NoProfile -File scripts\ga_dispatch_pd.ps1
#   pwsh -NoProfile -File scripts\ga_dispatch_pd.ps1 -AlsoDevRegression
#   pwsh -NoProfile -File scripts\ga_dispatch_pd.ps1 -NoWait
#
# 注意：本脚本含中文，必须用 pwsh（PowerShell 7）跑。
param(
    [string]$Root = "",
    [string]$Ref = "probe/photodemon-emit",
    [string]$DevRef = "dev",
    [string]$ProbeWorktree = "",
    [string]$Workflow = "PhotoDemon x86 full build",
    [string]$DevWorkflow = "VB6 C3 Regression",
    [switch]$SkipProbeSync,
    [switch]$NoWait,
    [switch]$AlsoDevRegression,
    [int]$WaitMinutes = 120,
    [int]$Tries = 8,
    [int]$SleepSeconds = 10
)
$ErrorActionPreference = 'Continue'
if (-not $Root) { $Root = (Split-Path -Parent $PSScriptRoot) }
$Root = (Resolve-Path $Root).Path
if (-not $ProbeWorktree) { $ProbeWorktree = Join-Path $Root ".temp\probe_sync" }
Set-Location $Root

function Invoke-Native {
    param([scriptblock]$Sb, [int]$Tries = 8, [int]$Sleep = 10, [string]$Tag = "")
    $res = @{ Ok = $false; Rc = -1; Out = @() }
    for ($i = 1; $i -le $Tries; $i++) {
        $out = & $Sb 2>&1
        $rc = $LASTEXITCODE
        $res.Rc = $rc
        $res.Out = @($out)
        if ($rc -eq 0) { $res.Ok = $true; return $res }
        Write-Host ("  [{0} 重试 {1}/{2}] rc={3} :: {4}" -f $Tag, $i, $Tries, $rc, (($out | Select-Object -Last 1) -join ' '))
        Start-Sleep -Seconds $Sleep
    }
    return $res
}

# gh 可用性自检（别让后面的失败报成"workflow 不存在"）
$auth = Invoke-Native { gh auth status } $Tries $SleepSeconds 'gh auth'
if (-not $auth.Ok) {
    Write-Host ("gh 不可用（rc=" + $auth.Rc + "）—— 先修网络/凭据，别往下派发")
    exit 2
}

$sha = ""
if (-not $SkipProbeSync) {
    if (-not (Test-Path $ProbeWorktree)) {
        Write-Host ("NO PROBE WORKTREE: " + $ProbeWorktree)
        Write-Host "  探针 worktree 是 .temp\probe_sync（另一个 clone，远端 github = 本 GA 的源）。"
        exit 2
    }
    Write-Host '=== 1) 同步探针分支 ==='
    $r = Invoke-Native { git -C $ProbeWorktree fetch github dev --quiet } $Tries $SleepSeconds 'fetch'
    if (-not $r.Ok) { Write-Host "fetch 失败"; exit 2 }
    $r = Invoke-Native { git -C $ProbeWorktree merge github/dev --no-edit } $Tries $SleepSeconds 'merge'
    if (-not $r.Ok) {
        Write-Host "merge github/dev 失败（多半是有冲突）—— 手工去 .temp\probe_sync 处理后再来"
        $r.Out | Select-Object -Last 15 | ForEach-Object { Write-Host ('  ' + $_) }
        exit 2
    }
    $r = Invoke-Native { git -C $ProbeWorktree push github $Ref } $Tries $SleepSeconds 'push'
    if (-not $r.Ok) { Write-Host "push 失败"; exit 2 }
    $r = Invoke-Native { git -C $ProbeWorktree rev-parse HEAD } $Tries $SleepSeconds 'rev-parse'
    $sha = (($r.Out | Select-Object -Last 1) -join '').Trim()
    Write-Host ("  已推 " + $Ref + " @ " + $sha)
} else {
    $r = Invoke-Native { git -C $ProbeWorktree rev-parse HEAD } $Tries $SleepSeconds 'rev-parse'
    $sha = (($r.Out | Select-Object -Last 1) -join '').Trim()
    Write-Host ("  跳过同步，按当前 HEAD " + $sha + " 派发")
}

function Find-Run([string]$workflow, [string]$branch, [string]$wantSha) {
    $r = Invoke-Native { gh run list --workflow $workflow --branch $branch --limit 15 --json databaseId,status,conclusion,headSha,createdAt } $Tries $SleepSeconds 'run list'
    if (-not $r.Ok) { return $null }
    try {
        $runs = (($r.Out -join "`n") | ConvertFrom-Json)
    } catch { return $null }
    foreach ($run in @($runs)) {
        if ($wantSha -and $run.headSha -eq $wantSha) { return $run }
    }
    return $null
}

function Wait-Run([string]$id, [int]$minutes) {
    $deadline = (Get-Date).AddMinutes($minutes)
    while ((Get-Date) -lt $deadline) {
        $r = Invoke-Native { gh run view $id --json status,conclusion } $Tries $SleepSeconds 'run view'
        if ($r.Ok -and $r.Out) {
            try {
                $o = (($r.Out -join "`n") | ConvertFrom-Json)
                Write-Host ("  [{0}] status={1} conclusion={2}" -f (Get-Date -Format HH:mm:ss), $o.status, $o.conclusion)
                if ($o.status -eq 'completed') { return $o }
            } catch { }
        }
        Start-Sleep -Seconds 20
    }
    return $null
}

Write-Host ''
Write-Host ('=== 2) 派发 "' + $Workflow + '" 到 ' + $Ref + ' ===')
$r = Invoke-Native { gh workflow run $Workflow --ref $Ref } $Tries $SleepSeconds 'workflow run'
if (-not $r.Ok) {
    Write-Host ("派发失败（rc=" + $r.Rc + "）")
    exit 2
}
Start-Sleep -Seconds 5

Write-Host ''
Write-Host '=== 3) 按 headSha 认领 run ==='
$run = $null
for ($i = 1; $i -le 6 -and -not $run; $i++) {
    $run = Find-Run $Workflow $Ref $sha
    if (-not $run) { Start-Sleep -Seconds 10 }
}
if (-not $run) {
    Write-Host ("没认领到 run（workflow=" + $Workflow + " branch=" + $Ref + " sha=" + $sha + "）")
    Write-Host "  手工确认： gh run list --workflow `"$Workflow`" --branch $Ref --limit 5"
    exit 2
}
$runId = "$($run.databaseId)"
Write-Host ("  run = " + $runId + "  headSha = " + $run.headSha + "  status = " + $run.status)
if ($run.headSha -ne $sha) {
    Write-Host ("  ★ headSha 不等于本地 HEAD（" + $run.headSha + " vs " + $sha + "）—— 读数可能串刀")
}

Set-Content -Path (Join-Path $Root ".temp\pd_run.txt") -Value $runId -Encoding ASCII

$conclusion = ''
if (-not $NoWait) {
    Write-Host ''
    Write-Host ('=== 4) 等 run ' + $runId + ' 结束 ===')
    $o = Wait-Run $runId $WaitMinutes
    if ($o) { $conclusion = $o.conclusion } else { $conclusion = 'TIMEOUT' }
} else {
    Write-Host ''
    Write-Host '=== 4) -NoWait：不等 ==='
}

if ($AlsoDevRegression) {
    Write-Host ''
    Write-Host ('=== 5) 派发 "' + $DevWorkflow + '" 到 ' + $DevRef + ' ===')
    $r = Invoke-Native { gh workflow run $DevWorkflow --ref $DevRef } $Tries $SleepSeconds 'workflow run'
    if (-not $r.Ok) {
        Write-Host ("dev 回归派发失败（rc=" + $r.Rc + "）")
    } else {
        Start-Sleep -Seconds 5
        $devSha = ''
        $r = Invoke-Native { git -C $Root ls-remote github ("refs/heads/" + $DevRef) } $Tries $SleepSeconds 'ls-remote'
        if ($r.Ok) {
            $devSha = ((($r.Out | Select-Object -First 1) -join '') -split '\s+')[0]
        }
        $devRun = $null
        for ($i = 1; $i -le 6 -and -not $devRun; $i++) {
            $devRun = Find-Run $DevWorkflow $DevRef $devSha
            if (-not $devRun) { Start-Sleep -Seconds 10 }
        }
        if ($devRun) {
            Write-Host ("  dev run = " + $devRun.databaseId + "  headSha = " + $devRun.headSha)
            Set-Content -Path (Join-Path $Root ".temp\dev_regression_run.txt") -Value "$($devRun.databaseId)" -Encoding ASCII
            if (-not $NoWait) {
                $od = Wait-Run "$($devRun.databaseId)" $WaitMinutes
                if ($od) { Write-Host ("  dev 回归结论 = " + $od.conclusion) }
            }
        } else {
            Write-Host "  没认领到 dev 回归 run（手工查 gh run list）"
        }
    }
}

Write-Host ''
Write-Host '=== 汇总 ==='
Write-Host ("  PD run = " + $runId + "  conclusion = " + $conclusion)
Write-Host ("  读数： pwsh -NoProfile -File scripts\pd_readout.ps1 -Run " + $runId)
exit 0
