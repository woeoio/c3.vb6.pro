# tests_github/run_t2.ps1 - T2 全量回归 (GitHub CI 专用, 自包含, 夜间层)
#
# 触发: .github/workflows/ci_t0.yml 的 t2 job (tag github-test-* / 夜间 cron / 手动 dispatch)
# 内容两类 (口径 2026-09-20 与用户对齐: 有交互的 GUI 不做自动化, GUI 用例只编译):
#   1. vbp 控制台工程 8 个: 编译 + 运行 + 输出断言 (无 GUI, 纯 stdout; 串行)
#      test_vbman (VBMANLIB) / test_exe_com_bridge (Scripting.Dictionary) 依赖外部 COM, 未注册自动 SKIP
#   2. GUI vbp 工程 4 个: 只编译不运行 (窗体验证留给本机 T3/L1-L3 体系)
# 与 tests\run_tests.ps1 的关系: 用例 2026-09-20 cp 自 tests\ (复制而非引用) ——
#   方向是 tests_github 自包含、后续废弃 tests\; 本脚本自带清单与引擎。
# 环境: 环境变量 C3_VCVARSALL 优先 (vcvarsall.bat 完整路径), 缺省 vswhere 自动发现,
#       用法见 scripts\README.md。仅限 Windows + PowerShell 7。
# 用法: pwsh -File tests_github\run_t2.ps1 [-C3Path .build\C3.exe] [-Verbose] [-Jobs N]
#       [-Shard K -ShardTotal M]
#   -Jobs 1 (默认) = 今天的串行行为, 逐字输出不变; >1 时 vbp/GUI 在本 runner 内并行 (每 worker
#   独立输出目录, 判定全部内联)。-Shard/-ShardTotal (默认 0/1 = 不分片) 把 vbp+GUI 切给多 runner。
#   8/5 清单守卫始终对完整清单校验, 之后只切"执行集"。

param(
    [string]$C3Path = "",
    [switch]$Verbose,
    [int]$Jobs = 1,        # >1 时 vbp/GUI 用例并行 (每 worker 独立输出目录); 默认 1 = 今天的串行行为
    [int]$Shard = 0,       # 分片当前编号 (1..ShardTotal); 0 = 不分片整队跑 (供 CI 多 runner 并行)
    [int]$ShardTotal = 1,  # 分片总数
    [int]$RunTimeoutSec = 60   # 单条用例跑 exe 的墙钟预算。默认 60 与 tests\run_tests.ps1
                               # 的 -RunTimeoutSec 同值 (那里从 5s 提到 60s 就是为这个原因)。
                               # 实测 run_t1 在 -Jobs 8 下把 test_rtl_x86 (空闲 0.03s 跑完)
                               # 判成 run timeout —— 预算是给并行负载留余量的, 不是给空闲机
                               # 定的; 5s 会把**负载**造成的慢误判成**代码**造成的挂。
                               # (GUI 的 Run3s 那条 3 秒存活自检不走这里, 那是另一回事。)
)

$ErrorActionPreference = "SilentlyContinue"
# <shared-shard>: 分片算法在 shard.ps1 (三个门禁脚本共用, 不要在这里再抄一遍)
. (Join-Path $PSScriptRoot "shard.ps1")

$Root = Split-Path -Parent $PSScriptRoot
if (-not $C3Path) { $C3Path = Join-Path $Root ".build\C3.exe" }
$CasesDir = Join-Path $PSScriptRoot "t2_cases"
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

# === MSVC 环境 (手法与 tests\run_tests.ps1 一致: cmd 导出 vcvarsall 后注入当前进程) ===
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
    Write-Host "[ERROR] 未找到 vcvarsall.bat (T2 全部用例都需要编译)" -ForegroundColor Red
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
$script:skip = 0

# 分片参数一次校验 (vbp 与 GUI 两处切片共用; 默认 0/1 = 不分片)
if ($ShardTotal -gt 1 -and ($Shard -lt 1 -or $Shard -gt $ShardTotal)) {
    Write-Host "[ERROR] Shard 需在 1..ShardTotal" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  T2 Full Regression (tests_github)" -ForegroundColor Cyan
Write-Host "  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# === COM 注册检查 (32 位程序读 WOW6432Node 视图; 未注册的依赖组件走 SKIP 不算 FAIL) ===
function Test-ComRegistered {
    param([string]$ProgId, [string]$Arch)
    if ($Arch -eq "x86") {
        $paths = @("HKLM:\SOFTWARE\WOW6432Node\Classes\$ProgId")
    } else {
        $paths = @("HKLM:\SOFTWARE\Classes\$ProgId")
    }
    foreach ($p in $paths) {
        if (Test-Path -Path $p) { return $true }
    }
    return $false
}

# === 超时读数 (与 tests\run_tests.ps1 同形, 两边改就一起改) ===
# 只量"杀之前"拿得到的东西: CPU 时间 / 进程状态 / **顶层窗口标题**。
# GUI 卡死最有用的一条是"它还带着哪个窗": 有窗 = 卡在消息循环里 (窗体没被 Unload),
# 无窗 = 根本没起来、或者窗已销毁却没退干净。
# ⚠ 这里**不**碰 $OutTask/$ErrTask: 进程活着 ⇒ 重定向管道不关 ⇒ ReadToEndAsync 不可能
#   完成 ⇒ 任何 Wait 都必超时 ⇒ 输出恒空。要拿到"被杀前已写出的那部分输出", 必须排在
#   Kill + WaitForExit **之后** (见 Read-KilledProcOutput)。
function Get-RunTimeoutDiag {
    param($Proc)
    $cpu = -1; $st = '?'; $win = ''
    try { $cpu = [int]$Proc.TotalProcessorTime.TotalMilliseconds } catch { }
    try { if ($Proc.HasExited) { $st = "exited=$($Proc.ExitCode)" } else { $st = 'alive' } } catch { }
    try { $Proc.Refresh(); $win = [string]$Proc.MainWindowTitle } catch { }
    $parts = @("cpu=${cpu}ms", $st)
    if ($win) { $parts += "win='$win'" }
    return " (" + ($parts -join ', ') + ")"
}

# 超时被杀**之后**才读得到的东西: Kill 掉进程 ⇒ 管道关闭 ⇒ ReadToEndAsync 这时才完成,
# 于是拿得到它在被杀前已经写出的那部分输出 (返回 @{ Out; Err; Last; Lines })。
function Read-KilledProcOutput {
    param($OutTask, $ErrTask)
    $out = ''; $err = ''
    try { if ($OutTask -and $OutTask.Wait(3000)) { $out = [string]$OutTask.Result } } catch { }
    try { if ($ErrTask -and $ErrTask.Wait(3000)) { $err = [string]$ErrTask.Result } } catch { }
    $last = ''
    foreach ($txt in @($out, $err)) {
        if ($txt) { $line = @($txt.TrimEnd() -split "`r?`n" | Where-Object { $_ }); if ($line.Count -gt 0) { $last = $line[-1] } }
    }
    $lines = 0
    if ($out) { $lines += @($out -split "`r?`n" | Where-Object { $_ }).Count }
    return @{ Out = $out; Err = $err; Last = $last; Lines = $lines }
}

# === 运行产物并捕获输出 (.NET Process + 5s 超时; 手法对齐 run_tests.ps1 Invoke-TestExe) ===
function Invoke-TestExe {
    param(
        [string]$ExePath,
        [string]$WorkDir,
        [string]$Name
    )
    $stdoutFile = Join-Path $WorkDir "$Name.out"
    $stderrFile = Join-Path $WorkDir "$Name.err"
    # 逐条判存在再删, 不依赖 Remove-Item 对"路径不存在"的宽容度:
    # 某些宿主 (带 safe-delete 钩子的沙箱) 把 Remove-Item 换成 fail-closed 版本, 目标不存在
    # 时抛**终止**异常 ⇒ 客户进程一次都没跑, 表现为"零输出、针全缺"且看不出异常痕迹。
    # (tests\run_tests.ps1 同款修法, 两边保持一致。)
    foreach ($stale in @($stdoutFile, $stderrFile)) {
        if (Test-Path $stale) { Remove-Item $stale -ErrorAction SilentlyContinue }
    }

    $errors = @()
    # --- 方案 A: .NET Process + 重定向 ---
    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $ExePath
        $psi.WorkingDirectory = $WorkDir
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $proc = [System.Diagnostics.Process]::Start($psi)
        $soTask = $proc.StandardOutput.ReadToEndAsync()
        $seTask = $proc.StandardError.ReadToEndAsync()
        if (-not $proc.WaitForExit($RunTimeoutSec * 1000)) {
            # 两步读数: 杀之前只能量 CPU/存活/窗口 (管道还开着, 读不到输出); 杀掉、管道关闭
            # 之后才拿得到"被杀前已经写出的那部分输出"。少了后半步就永远只能报一个裸超时。
            $diag = Get-RunTimeoutDiag -Proc $proc
            try { $proc.Kill() } catch { }
            $proc.WaitForExit()
            $partial = Read-KilledProcOutput -OutTask $soTask -ErrTask $seTask
            # 落盘: 下次超时能直接翻现场, 不用再猜卡在第几行。
            if ($partial.Out) { [System.IO.File]::WriteAllText($stdoutFile, [string]$partial.Out, [System.Text.Encoding]::Default) }
            if ($partial.Err) { [System.IO.File]::WriteAllText($stderrFile, [string]$partial.Err, [System.Text.Encoding]::Default) }
            $tail = " last='$($partial.Last)' lines=$($partial.Lines)"
            return @{ Ok = $false; ExitCode = $null; Output = @(); Detail = "run timeout: ${RunTimeoutSec}s$diag$tail" }
        }
        [System.IO.File]::WriteAllText($stdoutFile, [string]$soTask.Result, [System.Text.Encoding]::Default)
        [System.IO.File]::WriteAllText($stderrFile, [string]$seTask.Result, [System.Text.Encoding]::Default)
        return @{
            Ok       = $true
            ExitCode = $proc.ExitCode
            Output   = (Get-Content $stdoutFile -ErrorAction SilentlyContinue)
            Detail   = "dotnet"
        }
    } catch {
        $errors += ("[dotnet] " + $_.Exception.Message)
    }
    # --- 方案 B: Start-Process 回退 ---
    try {
        $proc = Start-Process -FilePath $ExePath -NoNewWindow -Wait -PassThru `
            -WorkingDirectory $WorkDir `
            -RedirectStandardOutput $stdoutFile `
            -RedirectStandardError $stderrFile `
            -ErrorAction Stop
        return @{
            Ok       = $true
            ExitCode = $proc.ExitCode
            Output   = (Get-Content $stdoutFile -ErrorAction SilentlyContinue)
            Detail   = "start-process"
        }
    } catch {
        $errors += ("[start-process] " + $_.Exception.Message)
    }
    return @{ Ok = $false; ExitCode = $null; Output = @(); Detail = ($errors -join " | ") }
}

# === 断言辅助: 输出行逐一匹配期望子串 (字面匹配, 与 run_tests.ps1 Test-NeedleHit 统一口径) ===
function Assert-Output {
    param([string[]]$ExpectedOutputs, $RunOutput)
    foreach ($expected in $ExpectedOutputs) {
        # 字面子串 (忽略大小写) 而非 -like: -like 会把 needle 里的 [ ] * ? 当通配符 (见 lesson 2)
        $found = $false
        foreach ($l in @($RunOutput)) {
            if ($null -ne $l -and ([string]$l).IndexOf($expected, [StringComparison]::OrdinalIgnoreCase) -ge 0) { $found = $true; break }
        }
        if (-not $found) { return $false }
    }
    return $true
}

# ============================================================
# 类 1: vbp 控制台工程 7 个 (编译 + 运行 + 输出断言, 串行)
# ============================================================
Write-Host "--- VBP Project Tests (run + assert) ---" -ForegroundColor Yellow

$script:vbpQueue = @()
function Add-VbpTest {
    param([string]$Name, [string]$VbpFile, [string[]]$Expected = @(), [string]$Arch = "", [string]$RequiresCom = "")
    $script:vbpQueue += @{
        Name        = $Name
        VbpFile     = (Join-Path $CasesDir $VbpFile)
        Expected    = @($Expected)
        Arch        = $Arch
        RequiresCom = $RequiresCom
    }
}

Add-VbpTest "test_class" "test_class.vbp" @("3", "0")
Add-VbpTest "M6Test" "M6Test.vbp" @("M6A:OK", "M6B:OK", "M6C:OK", "M6D:OK", "M6 PASSED")
Add-VbpTest "modulemethod" "test_modulemethod.vbp" @("30", "21")
Add-VbpTest "test_implements" "test_implements.vbp" @("IMPL1:OK", "IMPL2:OK", "Implements test PASSED")
Add-VbpTest "test_events" "test_events\test_events.vbp" @("Events test PASSED")
Add-VbpTest "M7Test" "m7_test\M7Test.vbp" @("4/4 PASSED")
Add-VbpTest "test_vbman" "test_vbman\test_vbman.vbp" @("P24-04a:OK", "P24-04b:OK", "P24-04:2/2") -Arch "x86" -RequiresCom "VBMANLIB.cVBMAN"
# ExeComBridge 03 回归: EXE 工程类实例过 COM 边界 (New bHello 作 Add 的 Variant 实参,
#   对端 CallByName 晚绑定调用). 修复前在 Add 处 0xC0000005, 修复后 PING-OK + Echo 往返.
Add-VbpTest "test_exe_com_bridge" "test_exe_com_bridge\test_exe_com_bridge.vbp" @("PING-OK", "EXEB:2/2", "EXE-COM-BRIDGE PASSED") -RequiresCom "Scripting.Dictionary"

if ($script:vbpQueue.Count -ne 8) {
    Write-Host "[ERROR] vbp 清单数量异常: $($script:vbpQueue.Count) (应为 8)" -ForegroundColor Red
    exit 1
}
$missing = @($script:vbpQueue | Where-Object { -not (Test-Path $_.VbpFile) })
if ($missing.Count -gt 0) {
    Write-Host "[ERROR] 清单引用了不存在的工程文件:" -ForegroundColor Red
    foreach ($m in $missing) { Write-Host "    $($m.VbpFile)" -ForegroundColor Red }
    exit 1
}

# === 分片 (CI 多 runner): vbp 队列按调用次序连续切片, 每 runner 只取第 Shard 片 ===
# 上面的 8 项清单守卫已对完整清单跑过, 这里只切"执行集"
# 分片算法见 shard.ps1: 本脚本的 vbp 与 GUI 两处队列共用同一对 Shard/ShardTotal
$script:vbpQueue = @(Select-ShardSlice -Items $script:vbpQueue -Shard $Shard -ShardTotal $ShardTotal -Label "vbp")

function Test-Vbp {
    param([object]$It)
    Write-Host -NoNewline "  [VBP] $($It.Name) ... "

    # 依赖的外部 COM 组件未注册: SKIP (环境原因, 不计 FAIL)
    if ($It.RequiresCom -and -not (Test-ComRegistered $It.RequiresCom $It.Arch)) {
        $script:skip++
        Write-Host "SKIP (COM '$($It.RequiresCom)' 未注册)" -ForegroundColor Yellow
        return
    }

    if ($It.Arch) {
        $compileResult = & $C3Path $It.VbpFile --arch $It.Arch --output-dir $OutDir 2>&1
    } else {
        $compileResult = & $C3Path $It.VbpFile --output-dir $OutDir 2>&1
    }
    if ($LASTEXITCODE -ne 0) {
        $script:fail++
        Write-Host "FAIL (compile)" -ForegroundColor Red
        # 可观测性约定: FAIL 必须带错误输出, 不依赖 -Verbose
        $compileResult | Select-Object -Last 25 | ForEach-Object { Write-Host "  $_" }
        $c3err = Join-Path $OutDir "c3-error.log"
        if (Test-Path $c3err) { Get-Content $c3err -Tail 25 | ForEach-Object { Write-Host "  $_" } }
        return
    }

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($It.VbpFile)
    $exePath = Join-Path $OutDir "$baseName.exe"
    if (-not (Test-Path $exePath)) {
        $script:fail++
        Write-Host "FAIL (no exe)" -ForegroundColor Red
        $compileResult | Select-Object -Last 15 | ForEach-Object { Write-Host "  $_" }
        return
    }

    $run = Invoke-TestExe -ExePath $exePath -WorkDir $OutDir -Name $baseName
    if (-not $run.Ok) {
        $script:fail++
        Write-Host "FAIL (run error)" -ForegroundColor Red
        Write-Host ("    " + $run.Detail) -ForegroundColor Red
        return
    }

    if ($It.Expected.Count -gt 0) {
        if (Assert-Output $It.Expected $run.Output) {
            $script:pass++
            Write-Host "PASS" -ForegroundColor Green
        } else {
            $script:fail++
            Write-Host "FAIL (output mismatch)" -ForegroundColor Red
            # 可观测性约定: 断言失败必须带期望与实际输出
            Write-Host "  Expected: $($It.Expected -join ', ')"
            Write-Host "  Got: $($run.Output -join ' | ')"
        }
    } else {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    }
}

# === vbp 并行执行 (多 runner/多 worker; 每个 worker 独立输出目录避免互相覆盖) ===
# ⚠ -Parallel 的 runspace 里调不到脚本函数 (Test-ComRegistered / Invoke-TestExe / Assert-Output),
#   所以 COM 注册判定、运行捕获、输出断言全部**内联**写在这里 (对齐 tests\run_tests.ps1 的规则)。
function Invoke-VbpSetParallel {
    param([object[]]$Items, [int]$Jobs, [string]$DirTag)
    if ($Items.Count -eq 0) { return }
    $per = [int][Math]::Ceiling($Items.Count / [double]$Jobs)
    $shards = @()
    for ($i = 0; $i -lt $Items.Count; $i += $per) {
        $end = [Math]::Min($i + $per - 1, $Items.Count - 1)
        $shards += ,@(@( $Items[$i..$end] ), (Join-Path $OutDir ($DirTag + $shards.Count)))
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
        $p = 0; $f = 0; $sk = 0; $details = @()
        foreach ($it in $shardItems) {
            # --- COM 未注册 => SKIP (与 Test-ComRegistered 语义一致, 内联) ---
            if ($it.RequiresCom) {
                $regPath = if ($it.Arch -eq "x86") { "HKLM:\SOFTWARE\WOW6432Node\Classes\$($it.RequiresCom)" } else { "HKLM:\SOFTWARE\Classes\$($it.RequiresCom)" }
                if (-not (Test-Path -Path $regPath)) { $sk++; $details += "$($it.Name): SKIP (COM '$($it.RequiresCom)' 未注册)"; continue }
            }
            # --- 编译 ---
            if ($it.Arch) { $cr = & $c3 $it.VbpFile --arch $it.Arch --output-dir $workDir 2>&1 } else { $cr = & $c3 $it.VbpFile --output-dir $workDir 2>&1 }
            if ($LASTEXITCODE -ne 0) {
                $tail = ($cr | Select-Object -Last 25) -join "`n"
                $c3err = Join-Path $workDir "c3-error.log"
                if (Test-Path $c3err) { $tail += "`n--- c3-error.log (tail 25) ---`n" + ((Get-Content $c3err -Tail 25) -join "`n") }
                $f++; $details += "$($it.Name): compile FAIL`n$tail"; continue
            }
            $baseName = [IO.Path]::GetFileNameWithoutExtension($it.VbpFile)
            $exePath = Join-Path $workDir "$baseName.exe"
            if (-not (Test-Path $exePath)) { $f++; $details += "$($it.Name): no exe`nC3 tail: " + (($cr | Select-Object -Last 15) -join "`n"); continue }

            # --- 运行 (.NET Process + 5s 超时; 语义对齐 Invoke-TestExe, 内联) ---
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
            } catch { $f++; $details += "$($it.Name): run error ($($_.Exception.Message))"; continue }

            # --- 输出断言 (字面子串, 忽略大小写; 内联, 不用 Assert-Output) ---
            if ($it.Expected.Count -gt 0 -and $runOk) {
                $runOut = @(Get-Content $stdoutFile -ErrorAction SilentlyContinue)
                $allMatch = $true
                foreach ($exp in $it.Expected) {
                    $found = $false
                    foreach ($l in @($runOut)) {
                        if ($null -ne $l -and ([string]$l).IndexOf($exp, [StringComparison]::OrdinalIgnoreCase) -ge 0) { $found = $true; break }
                    }
                    if (-not $found) { $allMatch = $false; break }
                }
                if ($allMatch) { $p++ } else {
                    $actual = ((Get-Content $stdoutFile -ErrorAction SilentlyContinue | Select-Object -First 20) -join "`n")
                    $f++; $details += "$($it.Name): output mismatch`nexpected: $($it.Expected -join ', ')`nactual:`n$actual"
                }
            } else {
                $p++
            }
        }
        [pscustomobject]@{ Pass = $p; Fail = $f; Skip = $sk; Details = $details }
    } -ThrottleLimit $Jobs

    foreach ($r in $results) {
        $script:pass += $r.Pass
        $script:fail += $r.Fail
        $script:skip += $r.Skip
        foreach ($d in $r.Details) { Write-Host "  [VBP] $d" -ForegroundColor Red }
    }
    $sumPass = ($results | Measure-Object -Property Pass -Sum).Sum
    $sumFail = ($results | Measure-Object -Property Fail -Sum).Sum
    Assert-ParallelRan -Items $script:vbpQueue -Results $results -Label "vbp"
    Write-Host "  (parallel vbp: $($results.Count) worker(s), pass=$sumPass fail=$sumFail)"
}

$vbpSw = [Diagnostics.Stopwatch]::StartNew()
if ($Jobs -gt 1 -and $PSVersionTable.PSVersion.Major -ge 7 -and $script:vbpQueue.Count -gt 0) {
    Invoke-VbpSetParallel -Items $script:vbpQueue -Jobs $Jobs -DirTag "vbprun"
} else {
    foreach ($it in $script:vbpQueue) { Test-Vbp $it }
}
$vbpSw.Stop()
Write-Host "  (vbp tests took $([Math]::Round($vbpSw.Elapsed.TotalSeconds))s)"
Write-Host ""

# ============================================================
# 类 2: GUI vbp 工程 3 个 (只编译不运行; 交互式 GUI 不做自动化)
# ============================================================
Write-Host "--- GUI VBP Compile-Only Tests ---" -ForegroundColor Yellow

$script:guiQueue = @()
function Add-GuiCompileTest {
    param([string]$Name, [string]$VbpFile, [string]$Arch = "", [switch]$Run3s)
    $resolved = $VbpFile
    if (-not ([IO.Path]::IsPathRooted($VbpFile))) {
        $resolved = (Join-Path $CasesDir $VbpFile)
    }
    $script:guiQueue += @{
        Name    = $Name
        VbpFile = $resolved
        Arch    = $Arch
        Run3s   = $Run3s
    }
}

Add-GuiCompileTest "VbQRCodegen" "VbQRCodegen-master\test\Project1.vbp"
Add-GuiCompileTest "BalloonTooltips" "BalloonTooltips\prjBalloonTooltips.vbp"
Add-GuiCompileTest "Charts2020" "Charts 2020\Proyecto1.vbp" -Arch "x86"
# ExtShow: 跨模块窗体默认实例"无参" Show (Fix 146 回归靶, 2026-09-20 vbman C2198):
# .bas caller 调 Form2.Show, 定义侧签名 (hMDIClient, modal) 后调用侧须补 modal=0
Add-GuiCompileTest "ExtShow" "ext_show_test\test_ext_show.vbp"
# VBFlexGridDemo: 仓库内部 demo (tests\VBFlexGridDemo), 编译后 3 秒存活自检
# (GUI 自动化口径 2026-09-22: 交互式控件 demo 不做交互断言, 启动 3 秒不崩即 PASS)
Add-GuiCompileTest "VBFlexGridDemo" (Join-Path $Root "tests\VBFlexGridDemo\VBFlexGridDemo.vbp") -Run3s

if ($script:guiQueue.Count -ne 5) {
    Write-Host "[ERROR] GUI 编译清单数量异常: $($script:guiQueue.Count) (应为 5)" -ForegroundColor Red
    exit 1
}
$missingGui = @($script:guiQueue | Where-Object { -not (Test-Path $_.VbpFile) })
if ($missingGui.Count -gt 0) {
    Write-Host "[ERROR] GUI 清单引用了不存在的工程文件:" -ForegroundColor Red
    foreach ($m in $missingGui) { Write-Host "    $($m.VbpFile)" -ForegroundColor Red }
    exit 1
}

# === 分片 (CI 多 runner): GUI 队列同样按调用次序连续切片 (与 vbp 共用 Shard/ShardTotal) ===
# 5 项清单守卫已对完整清单跑过, 这里只切"执行集"
$script:guiQueue = @(Select-ShardSlice -Items $script:guiQueue -Shard $Shard -ShardTotal $ShardTotal -Label "gui")

function Test-GuiCompileOnly {
    param([object]$It)
    Write-Host -NoNewline "  [GUI-COMPILE] $($It.Name) ... "
    if ($It.Arch) {
        $result = & $C3Path $It.VbpFile --arch $It.Arch --output-dir $OutDir 2>&1
    } else {
        $result = & $C3Path $It.VbpFile --output-dir $OutDir 2>&1
    }
    if ($LASTEXITCODE -eq 0) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        # 可观测性约定: FAIL 必须带**真正的错** (不是 tail 的 VB3001 警告)。
        # C3.exe 前端 warning 会灌满 tail-25, 把 cl/link 的 error C/LNK/D 挤出可视区。
        # 三条独立通道一起打:
        #   (a) 按错误模式过滤 $result (stdout+stderr), 前 15 条 error|fatal;
        #   (b) c3-error.log 头 (=== Command === / === MSVC Output === 段就在文件头);
        #   (c) c3-error.log 尾 40 行 (兜底, 覆盖没走 MSVC Output 段的情况)。
        $errPat = 'error\s+[CDL][0-9]+|fatal\s+error|unresolved\s+external|Permission\s+denied|cannot\s+open|C[0-9]{4}\s*:'
        $errLines = @($result | ForEach-Object { "$_" } | Where-Object { $_ -match $errPat } | Select-Object -First 15)
        if ($errLines.Count -gt 0) {
            Write-Host "  --- error-pattern matches (stdout+stderr) ---"
            $errLines | ForEach-Object { Write-Host "  $_" }
        }
        $result | Select-Object -Last 25 | ForEach-Object { Write-Host "  $_" }
        $c3err = Join-Path $OutDir "c3-error.log"
        # (b) 若 C3.exe 报了 "intermediates kept at: <path>" (msvc_driver.cpp 的
        # c3-error.log 会落在 objDir 而不是 outputDir), 也从那里读一次。
        $keptLine = ($result | ForEach-Object { "$_" } | Where-Object { $_ -match 'intermediates kept at: (.+)$' } | Select-Object -First 1)
        if ($keptLine -and $keptLine -match 'intermediates kept at: (.+)$') {
            $keptErr = Join-Path $Matches[1].Trim() "c3-error.log"
            if ((Test-Path $keptErr) -and ($keptErr -ne $c3err)) { $c3err = $keptErr }
        }
        if (Test-Path $c3err) {
            $c3Lines = @(Get-Content $c3err)
            $msvcIdx = [Array]::IndexOf($c3Lines, "=== MSVC Output ===")
            Write-Host "  --- c3-error.log head (Command / Response File) ---"
            if ($msvcIdx -gt 0) {
                $c3Lines[0..([Math]::Min($msvcIdx, $c3Lines.Count)-1)] | Select-Object -First 30 | ForEach-Object { Write-Host "  $_" }
            } else {
                $c3Lines | Select-Object -First 30 | ForEach-Object { Write-Host "  $_" }
            }
            # MSVC Output 段整块打 (cl/link 真错就在这里; 前端 VB3001 warning 是
            # driver_args.cpp writeErrorLog 之后追加的 === C3 Diagnostics === 段,
            # 会挤爆 tail-40, 掩盖 cl/link 侧的信息)。
            if ($msvcIdx -ge 0) {
                $diagIdx = -1
                for ($k = $msvcIdx + 1; $k -lt $c3Lines.Count; $k++) {
                    if ($c3Lines[$k] -like "=== C3 Diagnostics*") { $diagIdx = $k; break }
                }
                $upper = if ($diagIdx -ge 0) { $diagIdx - 1 } else { $c3Lines.Count - 1 }
                $msvcBody = @()
                if ($upper -ge $msvcIdx) { $msvcBody = $c3Lines[$msvcIdx..$upper] }
                Write-Host "  --- c3-error.log MSVC section ($($msvcBody.Count) lines) ---"
                if ($msvcBody.Count -eq 0) {
                    Write-Host "  (cl/link returned non-zero without any stdout/stderr output)"
                } else {
                    $msvcBody | ForEach-Object { Write-Host "  $_" }
                }
            }
            Write-Host "  --- c3-error.log tail 40 ---"
            $c3Lines | Select-Object -Last 40 | ForEach-Object { Write-Host "  $_" }
        }
    }
    if ($LASTEXITCODE -eq 0 -and $It.Run3s) {
        # 3 秒存活自检: 启动后若 3 秒内自行退出 -> 视为启动崩溃 (FAIL); 存活则强杀后 PASS
        $exe = Join-Path $OutDir "$([IO.Path]::GetFileNameWithoutExtension($It.VbpFile)).exe"
        Write-Host -NoNewline "  [GUI-RUN3S] $($It.Name) ... "
        if (-not (Test-Path $exe)) {
            $script:fail++
            Write-Host "FAIL (exe not found: $exe)" -ForegroundColor Red
            return
        }
        $proc = Start-Process -FilePath $exe -PassThru
        $crashed = $proc.WaitForExit(3000)
        if ($crashed) {
            $script:fail++
            Write-Host "FAIL (exited early, code=$($proc.ExitCode))" -ForegroundColor Red
        } else {
            Stop-Process -Id $proc.Id -Force
            $script:pass++
            Write-Host "PASS (alive 3s, killed)" -ForegroundColor Green
        }
    }
}

# === GUI 并行执行 (只编译 + 可选 Run3s 存活自检; 每个 worker 独立输出目录) ===
# ⚠ 判定内联 (同 vbp 并行规则): runspace 里调不到 Test-GuiCompileOnly。Run3s 只做"3 秒不崩"
#   存活自检 (Start-Process + WaitForExit(3000)), 不做窗口效果断言, 因此并行不违反 GUI 口径。
function Invoke-GuiSetParallel {
    param([object[]]$Items, [int]$Jobs, [string]$DirTag)
    if ($Items.Count -eq 0) { return }
    $per = [int][Math]::Ceiling($Items.Count / [double]$Jobs)
    $shards = @()
    for ($i = 0; $i -lt $Items.Count; $i += $per) {
        $end = [Math]::Min($i + $per - 1, $Items.Count - 1)
        $shards += ,@(@( $Items[$i..$end] ), (Join-Path $OutDir ($DirTag + $shards.Count)))
    }
    if ($shards.Count -eq 0) { return }
    $results = $shards | ForEach-Object -Parallel {
        $shardItems = $_[0]
        $workDir    = $_[1]
        New-Item -ItemType Directory -Path $workDir -Force | Out-Null
        $c3 = $using:C3Path
        $p = 0; $f = 0; $details = @()
        foreach ($it in $shardItems) {
            if ($it.Arch) { $r = & $c3 $it.VbpFile --arch $it.Arch --output-dir $workDir 2>&1 } else { $r = & $c3 $it.VbpFile --output-dir $workDir 2>&1 }
            if ($LASTEXITCODE -eq 0) {
                $p++
            } else {
                $f++; $details += "[GUI-COMPILE] $($it.Name): FAIL`n" + ((($r | Select-Object -Last 25) | ForEach-Object { "  $_" }) -join "`n")
                $c3err = Join-Path $workDir "c3-error.log"
                if (Test-Path $c3err) { $details += ((Get-Content $c3err -Tail 25 | ForEach-Object { "  $_" }) -join "`n") }
            }
            if ($LASTEXITCODE -eq 0 -and $it.Run3s) {
                $exe = Join-Path $workDir "$([IO.Path]::GetFileNameWithoutExtension($it.VbpFile)).exe"
                if (-not (Test-Path $exe)) {
                    $f++; $details += "[GUI-RUN3S] $($it.Name): FAIL (exe not found: $exe)"; continue
                }
                $proc = Start-Process -FilePath $exe -PassThru
                $crashed = $proc.WaitForExit(3000)
                if ($crashed) {
                    $f++; $details += "[GUI-RUN3S] $($it.Name): FAIL (exited early, code=$($proc.ExitCode))"
                } else {
                    Stop-Process -Id $proc.Id -Force
                    $p++
                }
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
    Assert-ParallelRan -Items $script:guiQueue -Results $results -Label "gui"
    Write-Host "  (parallel gui: $($results.Count) worker(s), pass=$sumPass fail=$sumFail)"
}

# GUI 段一律串行, 不走 ForEach-Object -Parallel。
#   1) VBFlexGridDemo 一份 63k 行 VBFlexGrid.c, cl 单进程 /MP 就能吃满 runner 全部核;
#      同 runner 上再叠一枚 GUI (Charts2020 x86 / ExtShow 之类) 并行, 两批 cl 的
#      c1.exe 会共享 %TEMP% 落 _CL_*.tmp 互相踩 (D8050 / C1083 Permission denied)。
#      C3.exe 侧已经按 session objDir 隔离了 TMP/TEMP (Fix <vbeclipse> D8050),
#      但同 runner 上多枚 GUI 一起编时 wall-clock 与内存都吃紧, 收益也有限。
#   2) 串行下 tail-25 的可观测性也够 —— 单条用例失败, 不用去区分是哪个并行 worker
#      的 stdout 交错。
# 恢复并行前先给 Test-GuiCompileOnly 独立 -OutDir, 并把 Jobs 上限压到 2。
foreach ($it in $script:guiQueue) { Test-GuiCompileOnly $it }
Write-Host ""

# === 汇总 (vbp 7 + gui 只编译 4 = 11) ===
$total = $script:pass + $script:fail + $script:skip
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  T2 Results: PASS=$($script:pass) FAIL=$($script:fail) SKIP=$($script:skip) TOTAL=$total" -ForegroundColor $(if ($script:fail -gt 0) { "Red" } else { "Green" })
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

if ($script:fail -gt 0) { exit 1 } else { exit 0 }
