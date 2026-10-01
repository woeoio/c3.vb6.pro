# tests_github/run_t0.ps1 - T0 最小编译门禁 (GitHub CI 专用)
#
# 触发: .github/workflows/ci_t0.yml, 打 tag github-test-* 推送到 GitHub 即运行
#   git tag github-test-001
#   git push <github 远端> github-test-001
# 内容三段:
#   1. C3.exe --version 冒烟 (产物可用性)
#   2. t0_cases\*.bas 全部 --syntax-only (语法门禁, 秒级, 不需要 MSVC 环境)
#   3. smoke.bas 编译 + 运行 + 输出断言 (end-to-end 最小闭环, 需要 MSVC 环境)
# 与 tests\run_tests.ps1 的关系: T0 用例原位于 tests\ (2026-09-20 git mv 到此),
#   run_tests.ps1 跨目录引用同一份文件, 两边不产生副本。
#
# 环境: 环境变量 C3_VCVARSALL 优先 (vcvarsall.bat 完整路径), 缺省 vswhere 自动发现,
#       用法见 scripts\README.md。仅限 Windows 运行。
# 用法: pwsh -File tests_github\run_t0.ps1 [-C3Path .build\C3.exe] [-Verbose] [-Jobs N]
#       [-Shard K -ShardTotal M]
#   -Jobs 1 (默认) = 今天的串行行为, 逐字输出不变; >1 时语法段并行 (每 worker 独立输出目录)。
#   -Shard/-ShardTotal (默认 0/1 = 不分片) 供 CI 多 runner 把语法用例切片并行跑。

param(
    [string]$C3Path = "",
    [switch]$Verbose,
    [int]$Jobs = 1,        # >1 时并行跑语法段用例 (每 worker 独立输出目录); 默认 1 = 今天的串行行为
    [int]$Shard = 0,       # 分片当前编号 (1..ShardTotal); 0 = 不分片整队跑 (供 CI 多 runner 并行)
    [int]$ShardTotal = 1   # 分片总数
)

$ErrorActionPreference = "SilentlyContinue"
# <shared-shard>: 分片算法在 shard.ps1 (三个门禁脚本共用, 不要在这里再抄一遍)
. (Join-Path $PSScriptRoot "shard.ps1")

$Root = Split-Path -Parent $PSScriptRoot
if (-not $C3Path) { $C3Path = Join-Path $Root ".build\C3.exe" }
$CasesDir = Join-Path $PSScriptRoot "t0_cases"
$OutDir = Join-Path $Root "output"

if (-not (Test-Path $C3Path)) {
    Write-Host "[ERROR] C3.exe 不存在: $C3Path (先构建或用 -C3Path 指定)" -ForegroundColor Red
    exit 1
}
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }

# === MSVC 环境 (smoke 编译需要 cl/link; --syntax-only 不需要) ===
# 手法与 tests\run_tests.ps1 一致: cmd 导出 vcvarsall 环境后注入当前进程
$vswhere = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
$VcVars = $env:C3_VCVARSALL
if (-not $VcVars -and (Test-Path $vswhere)) {
    $vsPath = & $vswhere -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath 2>$null
    if ($vsPath) {
        $candidate = Join-Path $vsPath "VC\Auxiliary\Build\vcvarsall.bat"
        if (Test-Path $candidate) { $VcVars = $candidate }
    }
}
if ($VcVars) {
    $envLines = cmd /c "call `"$VcVars`" x64 >nul 2>&1 && set" 2>$null
    foreach ($line in $envLines) {
        if ($line -match '^([^=]+)=(.*)$') {
            [Environment]::SetEnvironmentVariable($matches[1], $matches[2], "Process")
        }
    }
    Write-Host "MSVC env: $VcVars"
} else {
    Write-Host "[WARN] 未找到 vcvarsall.bat, smoke 编译段将跳过 (语法段不受影响)" -ForegroundColor Yellow
    $VcVars = ""
}

$script:pass = 0
$script:fail = 0
$script:skip = 0

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  T0 Minimal Gate (tests_github)" -ForegroundColor Cyan
Write-Host "  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# === 段 1: 版本冒烟 ===
Write-Host "--- 1/3 C3.exe version smoke ---" -ForegroundColor Yellow
& $C3Path --version
if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERROR] C3.exe --version 失败 (退出码 $LASTEXITCODE)" -ForegroundColor Red
    exit 1
}
Write-Host ""

# === 段 2: 语法门禁 (t0_cases\*.bas 全量 --syntax-only) ===
Write-Host "--- 2/3 Syntax gate (t0_cases\*.bas) ---" -ForegroundColor Yellow
$allCases = @(Get-ChildItem -Path $CasesDir -Filter "*.bas" | Sort-Object Name)
if ($allCases.Count -eq 0) {
    Write-Host "[ERROR] t0_cases\ 下无 .bas 用例" -ForegroundColor Red
    exit 1
}
# 分片 (CI 多 runner 并行): 按排序次序连续切片, 各 runner 只取第 Shard 片; 加用例不必改这里
# (手法对齐 tests\run_tests.ps1 的 bas 分片)
# 分片算法见 shard.ps1 (三个脚本共用): ShardTotal<=1 原样返回, >1 按清单次序连续切片
$cases = @(Select-ShardSlice -Items $allCases -Shard $Shard -ShardTotal $ShardTotal -Label "syntax")
# 并行 vs 串行: 默认 Jobs=1 = 今天的串行行为 (逐字输出不变); Jobs>1 每 worker 独立输出目录
if ($Jobs -gt 1 -and $PSVersionTable.PSVersion.Major -ge 7 -and $cases.Count -gt 0) {
    $per = [int][Math]::Ceiling($cases.Count / [double]$Jobs)
    $shards = @()
    for ($i = 0; $i -lt $cases.Count; $i += $per) {
        $end = [Math]::Min($i + $per - 1, $cases.Count - 1)
        $shards += ,@(@( $cases[$i..$end] ), (Join-Path $OutDir ("job" + $shards.Count)))
    }
    # ⚠ -Parallel 的 runspace 调不到脚本函数 ⇒ 判定必须内联 (见 tests\run_tests.ps1 同规则)
    $results = $shards | ForEach-Object -Parallel {
        $items   = $_[0]
        $workDir = $_[1]
        New-Item -ItemType Directory -Path $workDir -Force | Out-Null
        $c3 = $using:C3Path
        $p = 0; $f = 0; $details = @()
        foreach ($c in $items) {
            $r = & $c3 $c.FullName --syntax-only 2>&1
            if ($LASTEXITCODE -eq 0) { $p++ } else {
                # 可观测性约定: FAIL 必须带错误输出, 不依赖 -Verbose
                $f++; $details += "[SYNTAX] $($c.BaseName): FAIL`n" + ((($r | Select-Object -Last 30) | ForEach-Object { "  $_" }) -join "`n")
            }
        }
        [pscustomobject]@{ Pass = $p; Fail = $f; Details = $details }
    } -ThrottleLimit $Jobs
    foreach ($r in $results) {
        $script:pass += $r.Pass
        $script:fail += $r.Fail
        foreach ($d in $r.Details) { Write-Host "  $d" -ForegroundColor Red }
    }
    $sumPass = ($results | Measure-Object -Property Pass -Sum).Sum
    $sumFail = ($results | Measure-Object -Property Fail -Sum).Sum
    Assert-ParallelRan -Items $cases -Results $results -Label "syntax"
    Write-Host "  (parallel: $($results.Count) worker(s), syntax pass=$sumPass fail=$sumFail)"
} else {
    foreach ($c in $cases) {
        Write-Host -NoNewline "  [SYNTAX] $($c.BaseName) ... "
        $result = & $C3Path $c.FullName --syntax-only 2>&1
        if ($LASTEXITCODE -eq 0) {
            $script:pass++
            Write-Host "PASS" -ForegroundColor Green
        } else {
            $script:fail++
            Write-Host "FAIL" -ForegroundColor Red
            # 可观测性约定: FAIL 必须带错误输出, 不依赖 -Verbose
            $result | Select-Object -Last 30 | ForEach-Object { Write-Host "  $_" }
        }
    }
}
Write-Host ""

# === 段 3: smoke end-to-end (编译+运行+输出断言) ===
Write-Host "--- 3/3 Smoke end-to-end ---" -ForegroundColor Yellow
$smokeSrc = Join-Path $CasesDir "smoke.bas"
if (($ShardTotal -gt 1) -and ($Shard -ne 1)) {
    # 端到端冒烟是单项检查: 分片时只由 shard 1 负责, 其余片跳过 (不减检查、不重复计时)
    Write-Host "  (smoke e2e 由 shard 1 负责, 本片跳过)" -ForegroundColor DarkGray
} elseif (-not $VcVars) {
    $script:skip++
    Write-Host "  [SKIP] smoke (无 MSVC 环境)" -ForegroundColor Yellow
} else {
    Write-Host -NoNewline "  [SMOKE] compile ... "
    $smokeOut = & $C3Path $smokeSrc --output-dir $OutDir 2>&1
    if ($LASTEXITCODE -ne 0) {
        $script:fail++
        Write-Host "FAIL (compile)" -ForegroundColor Red
        # 失败时回显 C3 输出尾部与 c3-error.log (cl.exe 错误正文), 否则是黑盒
        Write-Host "  === C3 output (tail 40) ==="
        $smokeOut | Select-Object -Last 40 | ForEach-Object { Write-Host "  $_" }
        $c3Err = Join-Path $OutDir "c3-error.log"
        if (Test-Path $c3Err) {
            Write-Host "  === c3-error.log (tail 30) ==="
            Get-Content $c3Err | Select-Object -Last 30 | ForEach-Object { Write-Host "  $_" }
        }
    } else {
        $smokeExe = Join-Path $OutDir "smoke.exe"
        if (-not (Test-Path $smokeExe)) {
            $script:fail++
            Write-Host "FAIL (no exe)" -ForegroundColor Red
        } else {
            Write-Host "PASS" -ForegroundColor Green
            $expected = @("SMOKE-1:OK", "SMOKE-2:OK", "SMOKE-3:OK", "SMOKE PASS")
            # 运行: .NET Process + 5s 超时 + stdout 重定向 (手法对齐 tests\run_tests.ps1 Invoke-TestExe)
            $stdoutFile = Join-Path $OutDir "t0_smoke.out"
            $proc = $null
            try {
                $psi = New-Object System.Diagnostics.ProcessStartInfo
                $psi.FileName = $smokeExe
                $psi.WorkingDirectory = $OutDir
                $psi.UseShellExecute = $false
                $psi.CreateNoWindow = $false
                $psi.RedirectStandardOutput = $true
                $psi.RedirectStandardError = $true
                $proc = [System.Diagnostics.Process]::Start($psi)
                $soTask = $proc.StandardOutput.ReadToEndAsync()
                $seTask = $proc.StandardError.ReadToEndAsync()
                if (-not $proc.WaitForExit(5000)) {
                    # 读数 (与 tests\run_tests.ps1 同形): 杀之前量 CPU/存活/顶层窗口; 杀掉、
                    # 管道关闭之后才读得到"被杀前已经写出的那部分输出" —— 顺序倒过来则恒空。
                    $cpu = -1; $st = '?'; $win = ''
                    try { $cpu = [int]$proc.TotalProcessorTime.TotalMilliseconds } catch { }
                    try { if ($proc.HasExited) { $st = "exited=$($proc.ExitCode)" } else { $st = 'alive' } } catch { }
                    try { $proc.Refresh(); $win = [string]$proc.MainWindowTitle } catch { }
                    if ($win) { $win = ", win='$win'" }
                    try { $proc.Kill() } catch { }
                    $proc.WaitForExit()
                    $last = ''
                    foreach ($t in @($soTask, $seTask)) {
                        if ($t -and $t.Wait(3000)) {
                            $txt = ''
                            try { $txt = [string]$t.Result } catch { }
                            if ($txt) { $line = @($txt.TrimEnd() -split "`r?`n" | Where-Object { $_ }); if ($line.Count -gt 0) { $last = $line[-1] } }
                        }
                    }
                    $script:fail++
                    Write-Host "  [SMOKE] run FAIL (timeout 5s) (cpu=${cpu}ms, ${st}$win, last='$last')" -ForegroundColor Red
                    $proc = $null
                }
            } catch {
                $script:fail++
                Write-Host "  [SMOKE] run FAIL ($($_.Exception.Message))" -ForegroundColor Red
                $proc = $null
            }
            if ($proc) {
                $stdout = $soTask.Result
                $stderr = $seTask.Result
                [System.IO.File]::WriteAllText($stdoutFile, [string]$stdout, [System.Text.Encoding]::Default)
                [System.IO.File]::WriteAllText(($stdoutFile -replace '\.out$', '.err'), [string]$stderr, [System.Text.Encoding]::Default)
                $runOutput = @(Get-Content $stdoutFile -ErrorAction SilentlyContinue)
                $allMatch = $true
                foreach ($e in $expected) {
                    # 字面子串匹配 (忽略大小写): -like 会把 needle 里的 [ ] * ? 当通配符
                    $found = $false
                    foreach ($l in @($runOutput)) {
                        if ($null -ne $l -and ([string]$l).IndexOf($e, [StringComparison]::OrdinalIgnoreCase) -ge 0) { $found = $true; break }
                    }
                    if (-not $found) { $allMatch = $false; break }
                }
                if ($allMatch) {
                    $script:pass++
                    Write-Host "  [SMOKE] run + output PASS" -ForegroundColor Green
                } else {
                    $script:fail++
                    Write-Host "  [SMOKE] run FAIL (output mismatch)" -ForegroundColor Red
                    if ($Verbose) {
                        Write-Host "  Expected: $($expected -join ', ')"
                        Write-Host "  Got: $($runOutput -join ' | ')"
                    }
                }
            }
        }
    }
}
Write-Host ""

# === 汇总 (语法用例 + smoke 1 项) ===
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  T0 Results: PASS=$($script:pass) FAIL=$($script:fail) SKIP=$($script:skip) TOTAL=$($script:pass + $script:fail + $script:skip)" -ForegroundColor $(if ($script:fail -gt 0) { "Red" } else { "Green" })
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

if ($script:fail -gt 0) { exit 1 } else { exit 0 }
