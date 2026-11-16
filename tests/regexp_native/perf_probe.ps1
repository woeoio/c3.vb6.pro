param(
    [string]$C3 = ".build\C3.exe",
    [ValidateSet('x86','x64')][string]$Arch = 'x86',
    [string]$OutDir = ""
)
# C30-10 正则引擎的性能读数（**不是门禁**，只出数）。
#
# 为什么单列一份：性能数字会随机器、架构、编译选项变，把它塞进 CI 判据只会制造假红；
# 但"引擎到底多快、灾难回溯会不会跑满"必须有实测，不能靠嘴说。
#
# 跑法：pwsh -File tests\regexp_native\perf_probe.ps1
# 退出码：0 = 6 项都出了数；1 = 编译失败/没出文件。
#
# 口径说明（读数字时必须知道）：
#   * 计时用 VB 的 Timer（秒，Windows 下步进 ~15.6ms），所以每项故意重复到 0.3s 以上，
#     让步进误差 <5%；这不是微基准，是"量级"读数。
#   * 同时打印做了多少事（matches/call 数），避免"快"来自根本没匹配。
#   * 灾难回溯两项特别说明：真组件会把线程跑满；本引擎有 RE_STEP_LIMIT=2000000，
#     超限按"该起点不匹配"收手，所以这两项的数字是**上限行为**，不是"算得快"。

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent        # tests/
$repo = Split-Path $root -Parent
if (-not $OutDir) { $OutDir = Join-Path $repo (".build\re_perf_" + $Arch) }
$c3exe = if ([IO.Path]::IsPathRooted($C3)) { $C3 } else { (Resolve-Path (Join-Path $repo $C3)).Path }
$vbp = Join-Path $PSScriptRoot 'ReApp.vbp'

New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$env:C3_PROJECT_DIR = $repo
$log = & cmd.exe /c "`"$c3exe`" `"$vbp`" --arch $Arch --output-dir `"$OutDir`" 2>&1"
$exe = Join-Path $OutDir 'ReApp.exe'
if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $exe)) {
    Write-Output "FAIL 编译失败（没有 $exe）"
    $log | Select-Object -Last 12 | ForEach-Object { "  " + $_ }
    exit 1
}

$perf = Join-Path $OutDir 're_perf.txt'
if (Test-Path -LiteralPath $perf) { Remove-Item -LiteralPath $perf -Force }
$p = Start-Process -FilePath $exe -ArgumentList '/perf' -WorkingDirectory $OutDir -PassThru
if (-not $p.WaitForExit(600000)) { try { $p.Kill() } catch {} ; Write-Output "FAIL 性能探针超时（600s）"; exit 1 }
if (-not (Test-Path -LiteralPath $perf)) { Write-Output "FAIL 没有产出 re_perf.txt（code=" + $p.ExitCode + "）"; exit 1 }

$lines = @(Get-Content -LiteralPath $perf -Encoding Default | ForEach-Object { $_.TrimEnd() } | Where-Object { $_ -ne "" })
$lines | ForEach-Object { Write-Output ("  " + $_) }
if ($lines.Count -lt 6) { Write-Output ("FAIL 只出了 {0} 项（期望 6）" -f $lines.Count); exit 1 }
Write-Output ("PASS perf 6 项读数（arch={0}）" -f $Arch)
exit 0
