# tests_github/run_t1.ps1 - T1 最小回归 (GitHub CI 专用, 自包含)
#
# 触发: .github/workflows/ci_t0.yml 的 t1 job (needs: t0, 复用 T0 构建的 C3.exe artifact)
# 内容两段:
#   1. bas 类 33 个: 编译 + 运行 + 输出断言; x64/x86 双架构 (31 个双跑 + 2 个原生 x86
#      专用项 = 64 任务; -Jobs 并行, PS7 ForEach-Object -Parallel, -RunTimeoutSec 超时)
#   2. compile 类 10 个: 只编译不运行 (comprehensive x2 + 8 个窗体 .frm)
# 与 tests\run_tests.ps1 的关系: 用例 2026-09-20 cp 自 tests\ (复制而非引用) ——
#   方向是 tests_github 自包含、后续废弃 tests\; 因此本脚本自带清单与引擎,
#   tests\ 侧后续的用例增删不会自动进入 T1 (有意为之, 勿改回引用)。
# 环境: 环境变量 C3_VCVARSALL 优先 (vcvarsall.bat 完整路径), 缺省 vswhere 自动发现,
#       用法见 scripts\README.md。仅限 Windows + PowerShell 7 (并行需要)。
# 用法: pwsh -File tests_github\run_t1.ps1 [-C3Path .build\C3.exe] [-Jobs 20] [-Verbose]
#       [-Shard K -ShardTotal M]
#   -Jobs 1 (默认) = 今天的执行路径 (输出不变); >1 时 bas 用例在**本 runner 内**并行 (每 worker
#   独立输出目录)。-Shard/-ShardTotal (默认 0/1 = 不分片) 供 CI 把 bas 切给**多 runner** 并行,
#   compile 段与冒烟一样只由 shard 1 跑。清单守卫始终对完整 33/64 校验, 再切执行集。

param(
    [string]$C3Path = "",
    [switch]$Verbose,
    [int]$Jobs = 1,        # >1 时 bas 用例并行 (每 worker 独立输出目录); 默认 1 走今天的路径
    [int]$Shard = 0,       # 分片当前编号 (1..ShardTotal); 0 = 不分片整队跑 (供 CI 多 runner 并行)
    [int]$ShardTotal = 1,  # 分片总数
    [int]$RunTimeoutSec = 60   # 单条用例跑 exe 的墙钟预算。默认 60 与 tests\run_tests.ps1
                               # 的 -RunTimeoutSec 同值 (那里从 5s 提到 60s 就是为这个原因)。
                               # 实测: -Jobs 8 下 test_rtl_x86 空闲 0.03s 跑完, 却在 5s 预算
                               # 下被判 run timeout —— 预算是给并行负载留余量的, 不是给空闲机
                               # 定的; 5s 会把**负载**造成的慢误判成**代码**造成的挂。
)

$ErrorActionPreference = "SilentlyContinue"
# <shared-shard>: 分片算法在 shard.ps1 (三个门禁脚本共用, 不要在这里再抄一遍)
. (Join-Path $PSScriptRoot "shard.ps1")

$Root = Split-Path -Parent $PSScriptRoot
if (-not $C3Path) { $C3Path = Join-Path $Root ".build\C3.exe" }
$CasesDir = Join-Path $PSScriptRoot "t1_cases"
$OutDir = Join-Path $Root "output"

if (-not (Test-Path $C3Path)) {
    Write-Host "[ERROR] C3.exe 不存在: $C3Path (先构建或用 -C3Path 指定)" -ForegroundColor Red
    exit 1
}
if (-not (Test-Path $CasesDir)) {
    Write-Host "[ERROR] 用例目录不存在: $CasesDir" -ForegroundColor Red
    exit 1
}
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }

# === MSVC 环境 (编译需要 cl/link; 手法与 tests\run_tests.ps1 一致: cmd 导出后注入当前进程) ===
$vswhere = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
$VcVars = $env:C3_VCVARSALL
if (-not $VcVars -and (Test-Path $vswhere)) {
    $vsPath = & $vswhere -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath 2>$null
    if ($vsPath) {
        $candidate = Join-Path $vsPath "VC\Auxiliary\Build\vcvarsall.bat"
        if (Test-Path $candidate) { $VcVars = $candidate }
    }
}
if (-not $VcVars) {
    Write-Host "[ERROR] 未找到 vcvarsall.bat (T1 全部用例都需要编译)" -ForegroundColor Red
    exit 1
}
$envLines = cmd /c "call `"$VcVars`" x64 >nul 2>&1 && set" 2>$null
foreach ($line in $envLines) {
    if ($line -match '^([^=]+)=(.*)$') {
        [Environment]::SetEnvironmentVariable($matches[1], $matches[2], "Process")
    }
}
Write-Host "MSVC env: $VcVars"

$script:pass = 0
$script:fail = 0

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  T1 Minimal Regression (tests_github)" -ForegroundColor Cyan
Write-Host "  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# ============================================================
# 段 1: bas 类 33 个 (编译+运行+输出断言, -Jobs 并行)
# ============================================================
Write-Host "--- Bas Tests x64+x86 (parallel, jobs=$Jobs) ---" -ForegroundColor Yellow

$script:basQueue = @()
function Add-BasTest {
    param([string]$Name, [string]$Source, $Expected = @(), [string]$Arch = "")
    $script:basQueue += @{
        Name     = $Name
        Source   = (Join-Path $CasesDir $Source)
        Expected = @($Expected)
        Arch     = $Arch
    }
}

Add-BasTest "hello" "hello.bas"
Add-BasTest "test_m5" "test_m5.bas"
Add-BasTest "test_rtl" "test_rtl.bas"
Add-BasTest "test_array" "test_array.bas"
Add-BasTest "test_fileio" "test_fileio.bas"
Add-BasTest "test_error" "test_error.bas"
Add-BasTest "test_now" "test_now.bas"
Add-BasTest "test_getput" "test_getput.bas" @("PASS1a", "PASS1b", "PASS1c", "PASS2", "PASS3", "PASS4")
Add-BasTest "test_onerror" "test_onerror.bas" @("PASS1", "PASS2", "PASS3a", "PASS3b", "Done")
Add-BasTest "test_ndarray" "test_ndarray.bas" @("2D sum=270", "P8.1 ALL TESTS DONE")
Add-BasTest "test_foreach" "test_foreach4.bas" @("Long For Each: 150", "String For Each: Hello World")
Add-BasTest "test_softkeyword" "test_softkeyword.bas" @("Get=42", "Step=5", "Name=test")
Add-BasTest "test_date" "test_date.bas" @("PASS_Year", "PASS_Month", "PASS_Day", "PASS_NowYear", "Done")
Add-BasTest "test_colon" "test_colon.bas" @("PASS1", "PASS2", "PASS3", "Done")
Add-BasTest "test_variant" "test_variant.bas" @("PASS1a", "PASS1c", "PASS5", "Done")
Add-BasTest "test_compat" "test_compat.bas"
Add-BasTest "test_types" "test_types.bas"
Add-BasTest "test_control" "test_control.bas"
Add-BasTest "test_declare" "test_declare.bas"
Add-BasTest "test_fixes" "test_fixes.bas" @("FIX1:OK", "FIX2:OK", "FIX3:OK", "All fixes passed!")
Add-BasTest "test_com" "test_com.bas" @("COM-1:OK", "COM-2:OK", "COM-3:OK", "COM:3/3")
Add-BasTest "test_com2" "test_com2.bas" @("Users")
Add-BasTest "test_com3" "test_com3.bas" @("COM3-1:OK", "COM3-2:OK", "COM3-3:OK", "COM3-4:OK", "COM3:4/4")
Add-BasTest "test_earlybound" "test_earlybound.bas" @("EB-1:OK", "EB-2:OK", "EB:2/2")
Add-BasTest "test_p1324" "test_p1324.bas" @("P13-1:OK", "P13-3:OK", "P13-5:OK", "P13:8/8")
Add-BasTest "test_p24" "test_p24.bas" @("P24-01a:OK", "P24-01b:OK", "P24-01c:OK", "P24-03a:OK", "P24-03b:OK", "P24:5/5")
Add-BasTest "test_earlybound2" "test_earlybound2.bas" @("EB2-1:OK", "EB2-7:DriveType=2", "EB2-8:OK", "EB2-10:OK", "EB2:10/10") -Arch "x86"
Add-BasTest "test_not_com" "test_not_com.bas" @("NOT-COM:OK", "NOT-COM2:OK", "NOT-COM:PASS") -Arch "x86"
Add-BasTest "test_err_obj" "test_err_obj.bas" @("ERR-1:OK", "ERR-6:OK", "ERR:6/6")
Add-BasTest "test_variant_cmp" "test_variant_cmp.bas" @("VC-1:OK", "VC-4:OK", "VC:4/4")
Add-BasTest "test_com_default_prop" "test_com_default_prop.bas" @("DP-1:OK", "DP-4:OK", "P24-10: 4/4")
Add-BasTest "test_com_optional" "test_com_optional.bas" @("OP-1:OK", "OP-4:OK", "P24-11: 4/4")
Add-BasTest "test_bstr_concat_scalar" "test_bstr_concat_scalar.bas" @("BCS:16/16")

# 清单守卫: 队列数量必须等于 33, 防止清单被误改后静默丢用例
if ($script:basQueue.Count -ne 33) {
    Write-Host "[ERROR] bas 清单数量异常: $($script:basQueue.Count) (应为 33)" -ForegroundColor Red
    exit 1
}

# === 双架构展开 (2026-09-20 用户决策: VB6 生态以 32 位为主, 33 个用例全量双跑) ===
# 31 个用例 x64+x86 各一遍 (x86 任务 Name 加 _x86 后缀); test_earlybound2 / test_not_com
# 原生只登记 x86, 保持不跑 x64 (该组合从未验证过, 不贸然进门禁) => 31*2 + 2 = 64 任务
$fullQueue = @()
foreach ($it in $script:basQueue) {
    if ($it.Arch -eq "x86") {
        $fullQueue += $it
    } else {
        $fullQueue += @{
            Name     = $it.Name
            Source   = $it.Source
            Expected = $it.Expected
            Arch     = ""
        }
        $fullQueue += @{
            Name     = "$($it.Name)_x86"
            Source   = $it.Source
            Expected = $it.Expected
            Arch     = "x86"
        }
    }
}
if ($fullQueue.Count -ne 64) {
    Write-Host "[ERROR] 双架构任务数异常: $($fullQueue.Count) (应为 64)" -ForegroundColor Red
    exit 1
}

# 缺文件守卫: 队列引用的源文件必须都存在
$missing = @($fullQueue | Where-Object { -not (Test-Path $_.Source) })
if ($missing.Count -gt 0) {
    Write-Host "[ERROR] 队列引用了不存在的用例文件:" -ForegroundColor Red
    foreach ($m in $missing) { Write-Host "    $($m.Source)" -ForegroundColor Red }
    exit 1
}

# === 并行执行 (多个 C3 实例同时编译+运行; 每个 worker 独立输出目录避免互相覆盖) ===
# 手法与 tests\run_tests.ps1 Invoke-BasSetParallel 一致: 均分切片 + ForEach-Object -Parallel
function Invoke-BasSetParallel {
    param([object[]]$Items, [int]$Jobs)

    $per = [int][Math]::Ceiling($Items.Count / [double]$Jobs)
    $shards = @()
    for ($i = 0; $i -lt $Items.Count; $i += $per) {
        $end = [Math]::Min($i + $per - 1, $Items.Count - 1)
        $shards += ,@(@( $Items[$i..$end] ), (Join-Path $OutDir ("job" + $shards.Count)))
    }
    if ($shards.Count -eq 0) { return }

    # ⚠ -Parallel 的 runspace 够不到脚本变量 ⇒ 超时预算用 $using: 传进去 (同 run_tests.ps1)
    $runTimeoutMs = $RunTimeoutSec * 1000

    $results = $shards | ForEach-Object -Parallel {
        $shardItems = $_[0]
        $workDir    = $_[1]
        New-Item -ItemType Directory -Path $workDir -Force | Out-Null
        $c3 = $using:C3Path
        $runTimeoutMs = $using:runTimeoutMs
        $p = 0; $f = 0; $details = @()
        foreach ($it in $shardItems) {
            if ($it.Arch) {
                $cr = & $c3 $it.Source --arch $it.Arch --output-dir $workDir 2>&1
            } else {
                $cr = & $c3 $it.Source --output-dir $workDir 2>&1
            }
            $ec = $LASTEXITCODE
            if ($ec -ne 0) {
                # 可观测性约定: FAIL 必须带错误输出 (C3 输出尾部 + c3-error.log 尾部)
                $tail = ($cr | Select-Object -Last 20) -join "`n"
                $c3err = Join-Path $workDir "c3-error.log"
                if (Test-Path $c3err) {
                    $tail += "`n--- c3-error.log (tail 25) ---`n" + ((Get-Content $c3err -Tail 25) -join "`n")
                }
                $f++; $details += "$($it.Name): compile FAIL`n$tail"; continue
            }

            $baseName = [IO.Path]::GetFileNameWithoutExtension($it.Source)
            $exePath = Join-Path $workDir "$baseName.exe"
            if (-not (Test-Path $exePath)) {
                $f++; $details += "$($it.Name): no exe`nC3 tail: " + (($cr | Select-Object -Last 15) -join "`n"); continue
            }

            # --- 运行 (.NET Process + 5s 超时; 语义对齐 run_tests.ps1 Invoke-TestExe) ---
            $stdoutFile = Join-Path $workDir "$($it.Name).out"
            $stderrFile = Join-Path $workDir "$($it.Name).err"
            $runOk = $false
            try {
                $psi = New-Object System.Diagnostics.ProcessStartInfo
                $psi.FileName = $exePath
                $psi.WorkingDirectory = $workDir
                $psi.UseShellExecute = $false
                $psi.CreateNoWindow = $false
                $psi.RedirectStandardOutput = $true
                $psi.RedirectStandardError = $true
                $proc = [System.Diagnostics.Process]::Start($psi)
                $soTask = $proc.StandardOutput.ReadToEndAsync()
                $seTask = $proc.StandardError.ReadToEndAsync()
                if (-not $proc.WaitForExit($runTimeoutMs)) {
                    # 读数: CPU 时间 + 进程状态 + 顶层窗口 + 最后一行输出 (内联 —— -Parallel 的
                    # runspace 够不到脚本函数)。顺序不能倒: 必须 Kill + WaitForExit **之后**
                    # 管道才关闭 ⇒ ReadToEndAsync 才完成 ⇒ 才拿得到被杀前已写出的输出。
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
                    $f++; $details += "$($it.Name): run timeout $($runTimeoutMs)ms (cpu=${cpu}ms, ${st}$win, last='$last')"; continue
                }
                [IO.File]::WriteAllText($stdoutFile, [string]$soTask.Result, [Text.Encoding]::Default)
                [IO.File]::WriteAllText($stderrFile, [string]$seTask.Result, [Text.Encoding]::Default)
                $runOk = $true
            } catch {
                $f++; $details += "$($it.Name): run error ($($_.Exception.Message))"; continue
            }

            if ($it.Expected -and $it.Expected.Count -gt 0 -and $runOk) {
                $runOut = @(Get-Content $stdoutFile -ErrorAction SilentlyContinue)
                $allMatch = $true
                foreach ($exp in $it.Expected) {
                    # ⚠ 这段跑在 -Parallel 的 runspace 里, 调不到脚本函数 ⇒ 字面判定必须内联。
                    # 用字面子串 (忽略大小写) 而非 -like: needle 里的 [ ] * ? 会被当通配符 (见 run_tests.ps1 Test-NeedleHit)
                    $found = $false
                    foreach ($l in @($runOut)) {
                        if ($null -ne $l -and ([string]$l).IndexOf($exp, [StringComparison]::OrdinalIgnoreCase) -ge 0) { $found = $true; break }
                    }
                    if (-not $found) { $allMatch = $false; break }
                }
                if ($allMatch) { $p++ } else {
                    # 可观测性约定: 输出断言失败必须带实际输出
                    $actual = (Get-Content $stdoutFile -ErrorAction SilentlyContinue | Select-Object -First 20) -join "`n"
                    $f++; $details += "$($it.Name): output mismatch`nexpected: $($it.Expected -join ', ')`nactual:`n$actual"
                }
            } else {
                $p++
            }
        }
        [pscustomobject]@{ Pass = $p; Fail = $f; Details = $details }
    } -ThrottleLimit $Jobs

    foreach ($r in $results) {
        $script:pass += $r.Pass
        $script:fail += $r.Fail
        foreach ($d in $r.Details) {
            Write-Host "  [RUN] $d" -ForegroundColor Red
        }
    }
    $sumPass = ($results | Measure-Object -Property Pass -Sum).Sum
    $sumFail = ($results | Measure-Object -Property Fail -Sum).Sum
    Assert-ParallelRan -Items $Items -Results $results -Label "bas"
    Write-Host "  (parallel: $($results.Count) worker(s), pass=$sumPass fail=$sumFail)"
}

# === 分片 (CI 多 runner 并行): 各 runner 只执行 fullQueue 的第 Shard 片 (按调用次序连续切片) ===
# 33/64 清单守卫在上面已对**完整**清单跑过, 这里只切"执行集", 每个 runner 仍会先校验清单完整。
# 分片算法见 shard.ps1 (三个脚本共用): ShardTotal<=1 原样返回, >1 按清单次序连续切片
$fullQueue = @(Select-ShardSlice -Items $fullQueue -Shard $Shard -ShardTotal $ShardTotal -Label "bas")

Invoke-BasSetParallel -Items $fullQueue -Jobs $Jobs
Write-Host ""

# ============================================================
# 段 2: compile 类 10 个 (只编译不运行: comprehensive x2 + 窗体 .frm x8)
# ============================================================
function Test-Compile {
    param([string]$Name, [string]$Source)
    Write-Host -NoNewline "  [COMPILE] $Name ... "
    $result = & $C3Path $Source --output-dir $OutDir 2>&1
    if ($LASTEXITCODE -eq 0) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        # 可观测性约定: FAIL 必须带错误输出, 不依赖 -Verbose
        $result | Select-Object -Last 25 | ForEach-Object { Write-Host "  $_" }
        $c3err = Join-Path $OutDir "c3-error.log"
        if (Test-Path $c3err) { Get-Content $c3err -Tail 25 | ForEach-Object { Write-Host "  $_" } }
    }
}

# compile 段 (10 个, 只编译) 是固定小项: 分片时只由 shard 1 负责, 其余片跳过 (不减检查)
if (($ShardTotal -gt 1) -and ($Shard -ne 1)) {
    Write-Host "--- Compile Tests: 由 shard 1 负责, 本片跳过 ---" -ForegroundColor DarkGray
} else {
    Write-Host "--- Compile Tests (bas x2 + form x8) ---" -ForegroundColor Yellow

    Test-Compile "test_comprehensive" (Join-Path $CasesDir "test_comprehensive.bas")
    Test-Compile "test_comprehensive2" (Join-Path $CasesDir "test_comprehensive2.bas")

    $formTests = @(
        "empty_form.frm",
        "form_test_p74.frm",
        "form_test_p75.frm",
        "form_test_p76.frm",
        "form_test_p78.frm",
        "form_test_p79.frm",
        "form_test_m8.frm",
        "form_mdi_parent.frm"
    )
    foreach ($t in $formTests) {
        Test-Compile ([IO.Path]::GetFileNameWithoutExtension($t)) (Join-Path $CasesDir $t)
    }
}
Write-Host ""

# === 汇总 (bas 64 任务 = 33 用例双架构 + compile 10 = 74) ===
$total = $script:pass + $script:fail
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  T1 Results: PASS=$($script:pass) FAIL=$($script:fail) TOTAL=$total" -ForegroundColor $(if ($script:fail -gt 0) { "Red" } else { "Green" })
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

if ($script:fail -gt 0) { exit 1 } else { exit 0 }
