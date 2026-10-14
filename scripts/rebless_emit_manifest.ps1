# rebless_emit_manifest.ps1 - 形状门（Emit manifest / shape oracle）红了之后的一条自救命令。
#
# 为什么要有这个脚本（2026-10-09，用户反馈「其他分支推送时太容易挂了，人家不太懂怎么处理」）：
# 这道门的判据是「全语料 398 份输入的 --emit-c 哈希，与仓里 emit-manifest.expected.txt 逐行相同」，
# 红过两次，两次都不是产品缺陷：
#   · 门 #434（head 57b1f64e）= 新夹具没登记哈希（共有 397 行逐行相同）；
#   · head 15e119f3（ai/032 rev2 整型提升）= 7 份发码真的变了，处置本该是「看过、再登记」。
# 而「登记」过去的做法要本机装 MSVC、编一台 C3.exe、跑二十分钟全语料 —— 协作者手上没有这条路的钥匙，
# 于是门红变成卡住别人。但钥匙其实早就在仓库里：CI 每一步都把自己那台二进制交的清单推到
# `ci/emit-manifest` 分支（台账 §B92 那条 git 回读通道），所以**不必本机重跑**就能登记。
#
# 用法:
#   pwsh -File scripts/rebless_emit_manifest.ps1              # 只看：取 CI 那份清单，报三格分类
#   pwsh -File scripts/rebless_emit_manifest.ps1 -Bless       # 看过没问题就登记，登记完自己复算
#   pwsh -File scripts/rebless_emit_manifest.ps1 -Manifest <本地清单> [-Bless]
#   pwsh -File scripts/rebless_emit_manifest.ps1 -Bless -ForceHead   # 明知清单不是当前 HEAD 跑的，仍要登记
# 退出码: 0 已一致（或登记后复算绿）/ 1 需要登记（不带 -Bless 时）/ 2 前置条件坏或来源对不上
param(
    [string]$Root = "",
    [string]$Remote = "",
    [string]$CiBranch = "ci/emit-manifest",
    [string]$Manifest = "",
    [switch]$Bless,
    [switch]$ForceHead
)
$ErrorActionPreference = 'Continue'
if (-not $Root) { $Root = (Split-Path -Parent $PSScriptRoot) }
$Root = (Resolve-Path $Root).Path
$expect = Join-Path $Root 'emit-manifest.expected.txt'
$cmp = Join-Path $PSScriptRoot 'compare_emit_manifest.ps1'
$cov = Join-Path $PSScriptRoot 'check_manifest_coverage.ps1'
foreach ($need in @($expect, $cmp)) {
    if (-not (Test-Path $need)) { Write-Host ('缺文件: ' + $need); exit 2 }
}

function RunCompare([string]$m, [switch]$asBless, [switch]$allowVanish) {
    $argv = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $cmp, '-Manifest', $m, '-Expect', $expect)
    if ($asBless) { $argv += '-Bless' }
    if ($allowVanish) { $argv += '-AllowVanish' }
    # 子进程的输出**必须马上转发给宿主**：函数里不转发就会被当成函数的返回值一起塞进 $rc，
    # 于是 $rc 变成「一整串分类读数 + 末尾一个退出码」的数组 —— `if ($rc -ne 0)` 对数组恒为真，
    # 登记明明成功却被报成「被挡住」，而那些处置文字本来正是给人看的（实测栽过一次）。
    & powershell @argv | ForEach-Object { Write-Host $_ }
    $rc = $LASTEXITCODE
    return $rc
}

# ---- 1) 拿到一份「判据用的清单」 ----
$srcLabel = ''
$runHead = ''
if ($Manifest) {
    if (-not (Test-Path $Manifest)) { Write-Host ('清单不存在: ' + $Manifest); exit 2 }
    $work = (Resolve-Path $Manifest).Path
    $srcLabel = '本地清单 ' + $work
} else {
    if (-not $Remote) {
        $remotes = @(git -C $Root remote)
        $Remote = @('github', 'origin') | Where-Object { $remotes -contains $_ } | Select-Object -First 1
        if (-not $Remote) { Write-Host '找不到 github/origin 这枚 remote，用 -Remote 指定'; exit 2 }
    }
    Write-Host ('取 ' + $Remote + '/' + $CiBranch + '（CI 每次把自己那台的清单提交到这里）...')
    git -C $Root fetch $Remote $CiBranch 2>&1 | ForEach-Object { Write-Host ('  ' + $_) }
    if ($LASTEXITCODE -ne 0) {
        Write-Host ('  fetch 失败：这一格的门还没跑过、或者分支名不对（' + $CiBranch + '）。')
        Write-Host '  那就本机出一份：先 scripts/build.bat，再 scripts/emit_manifest.ps1，最后带 -Manifest emit-manifest.txt 回到这里。'
        exit 2
    }
    $runHead = (git -C $Root log -1 --format='%s' FETCH_HEAD)
    $tmpDir = if (Test-Path (Join-Path $Root '.build')) { Join-Path $Root '.build' } else { $env:TEMP }
    $work = Join-Path $tmpDir 'rebless-from-ci.emit-manifest.txt'
    # 字节必须原样落盘：清单里的路径/注脚是 git blob 的原文，PowerShell 的管道会按控制台代码页重编码，
    # 那会把「登记用的数」变成「我这台的编码」—— 所以这一步借 cmd 的重定向（OS 级，字节不动）。
    $q = '"'
    $cmdLine = 'git -C ' + $q + $Root + $q + ' show FETCH_HEAD:emit-manifest.txt > ' + $q + $work + $q
    & cmd /c $cmdLine
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path $work)) { Write-Host '取回清单失败'; exit 2 }
    $srcLabel = $Remote + '/' + $CiBranch + ' （提交注记: ' + $runHead + '）'
    # 来源核对：这份清单必须是**当前检出这一笔**（或其祖先）跑出来的，否则登记进去的是别人的形状。
    $hm = [regex]::Match($runHead, 'head=([0-9a-f]{6,40})')
    $localHead = (git -C $Root rev-parse HEAD)
    if ($hm.Success) {
        $cid = $hm.Groups[1].Value
        git -C $Root merge-base --is-ancestor $cid $localHead 2>$null
        if ($LASTEXITCODE -ne 0) {
            Write-Host ('注意：这份清单出自 head=' + $cid + '，而它不在你当前 HEAD（' + $localHead.Substring(0, 8) + '）的历史里。')
            Write-Host '      可能别人又推了一笔，或者门跑的不是你这笔改动 —— 登记它会把你没验过的形状钉进期望。'
            Write-Host '      确认没问题再补 -ForceHead；否则先把手上的改动推上去，等门重跑一次再回来。'
            if (-not $ForceHead) { exit 2 }
            Write-Host '      （-ForceHead 已给，继续）'
        }
    }
}

# ---- 2) 分类读数（判据逻辑只有 compare 那一份，这里绝不重抄一遍） ----
Write-Host ('来源: ' + $srcLabel)
$rows = @(Get-Content -LiteralPath $work -Encoding UTF8 | Where-Object { $_ -match '\S' -and $_ -notmatch '^#' })
Write-Host ('清单数据行 = ' + $rows.Count)
$rc = RunCompare $work
if ($rc -eq 0) {
    Write-Host '期望清单与这份清单已经逐行相同 —— 不用登记。'
    exit 0
}
if ($rc -eq 2) { exit 2 }

# ---- 3) 登记 + 自证 ----
if (-not $Bless) {
    Write-Host '上面只是「看」。确认这确实是你这次改动想要的形状之后：'
    Write-Host ('  pwsh -File scripts/rebless_emit_manifest.ps1' + $(if ($Manifest) { ' -Manifest ' + $work } else { ''}) + ' -Bless')
    exit 1
}
$rc = RunCompare $work -asBless
if ($rc -ne 0) { Write-Host '登记被挡住（见上面那句「拒绝登记」），先把成因问清楚。'; exit 2 }
$rc = RunCompare $work
if ($rc -ne 0) { Write-Host '登记后复算仍然不一致 —— 不要提交，直接来找本线。'; exit 2 }
if (Test-Path $cov) {
    & powershell -NoProfile -ExecutionPolicy Bypass -File $cov
    if ($LASTEXITCODE -ne 0) { Write-Host 'check_manifest_coverage 红了：登记面出了问题，别提交。'; exit 2 }
}
Write-Host '复算绿。下一步（只动这一份文件）：'
Write-Host '  git add emit-manifest.expected.txt'
Write-Host '  git commit -m "登记：发码形状随本次改动更新（N 行）"'
exit 0
