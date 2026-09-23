# c3 编译器自动化测试框架
# 用途: .\run_tests.ps1 [-Category <all|compile|run|syntax>] [-Verbose]
#
# 参数说明:
#   smoke   - 冒烟测试 (编译+链接+运行; 主要验证 .build\C3.exe 能被正常生成并运行)
#   compile - 编译测试 (c3 .bas -> .exe, 只编译不运行)
#   run     - 运行测试 (编译+链接后运行, 校验输出是否符合预期)
#   syntax  - 语法测试 (仅 --syntax-only, 不生成可执行文件)
#   all     - 全部测试 (编译+链接+运行所有分类)

param(
    [string]$Category = "all",
    [switch]$Verbose,
    [string]$OutputDirectory = "",
    [int]$Jobs = 1,          # >1 时并行运行纯 .bas 用例 (每 worker 独立输出目录); GUI/VBP 始终串行
    [int]$BasShard = 0,      # bas 分片当前编号 (1..BasShardTotal); 0=不分片整队跑
    [int]$BasShardTotal = 1  # bas 分片总数 (供 CI 用多 runner 并行跑 bas 用例)
)

$ErrorActionPreference = "SilentlyContinue"

# 输出编码说明: PS5.1 终端按 Windows 控制台代码页解释输出; 本脚本统一以 UTF-8 写入
# 兼容 VS2022 的 Community/Professional/Enterprise/BuildTools 任一版本 (供 CI 使用)。详见 scripts\README.md
$Root = Split-Path -Parent $PSScriptRoot
$C3 = Join-Path $Root ".build\C3.exe"
$Tests = $PSScriptRoot
# T0 拆分 (2026-09-20): 语法/冒烟用例 git mv 至 tests_github\t0_cases\, 跨目录引用同一份文件 (无副本)
$GHTests = Join-Path $Tests "..\tests_github\t0_cases"
$OutDir = if ($OutputDirectory) { $OutputDirectory } else { Join-Path $Root "output" }
# === 获取 MSVC 编译环境 (通用) ===
# 设计原则: 不依赖 cmd.exe / vcvarsall 解析 (易因安全策略/编码/PATH 大小写失效),
# 不硬编码 VS 版本 (2019/2022/2026 均可) 与 Windows SDK 版本号 (动态发现)。
# 优先用环境变量 C3_VCVARSALL 指向的 vcvarsall.bat 推导 VS 根; 否则用 vswhere 取最新已装 VS。
function Get-MsvcToolset {
    # 返回 @{ VsRoot; ToolVer; SdkVer; BinX64; BinX86; Include; LibX64; LibX86 }
    # 1) C3_VCVARSALL -> 上溯 4 级 (...\<Edition>\VC\Auxiliary\Build\vcvarsall.bat) 得 VS 根目录
    $vsRoot = $null
    if ($env:C3_VCVARSALL -and (Test-Path $env:C3_VCVARSALL)) {
        $p = $env:C3_VCVARSALL
        for ($i = 0; $i -lt 4; $i++) { $p = Split-Path -Parent $p }
        $vsRoot = $p
    }
    # 2) 否则 vswhere 探测已安装的最新 VS (Community/Pro/Enterprise/BuildTools 任一版本)
    if (-not $vsRoot) {
        $vswhere = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
        if (Test-Path $vswhere) {
            $vsPath = & $vswhere -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath 2>$null
            if ($vsPath) { $vsRoot = $vsPath }
        }
    }
    if (-not $vsRoot -or -not (Test-Path $vsRoot)) {
        Write-Host "[ERROR] 未找到 Visual Studio (含 VC.Tools 工作负载)" -ForegroundColor Red
        Write-Host "        请安装任意版本 VS 并勾选 [使用 C++ 的桌面开发], 或设置环境变量 C3_VCVARSALL 指向 vcvarsall.bat" -ForegroundColor Red
        exit 1
    }
    # 3) 动态发现 MSVC toolset 版本 (VC\Tools\MSVC\<ver>) -- 不硬编码
    $msvcRoot = Join-Path $vsRoot "VC\Tools\MSVC"
    $toolVer = Get-ChildItem $msvcRoot -Directory -ErrorAction SilentlyContinue |
        Sort-Object Name -Descending | Select-Object -First 1 -ExpandProperty Name
    if (-not $toolVer) { Write-Host "[ERROR] 未找到 MSVC toolset ($msvcRoot)" -ForegroundColor Red; exit 1 }
    # 4) 动态发现 Windows SDK 根 (Windows Kits\10) -- 不硬编码盘符/版本
    #    优先读注册表 KitsRoot10 (vcvarsall 自身定位方式), 其次标准 Program Files 位置,
    #    再回退扫描常见盘符。
    $kitRoot = $null
    foreach ($base in @(${env:ProgramFiles(x86)}, $env:ProgramFiles)) {
        if ($base -and (Test-Path (Join-Path $base "Windows Kits\10\Include"))) {
            $kitRoot = Join-Path $base "Windows Kits\10"; break
        }
    }
    if (-not $kitRoot) {
        foreach ($rp in @("HKLM:\SOFTWARE\Microsoft\Windows Kits\Installed Roots",
                          "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows Kits\Installed Roots")) {
            if (Test-Path $rp) {
                $kr = (Get-ItemProperty -Path $rp -ErrorAction SilentlyContinue).KitsRoot10
                if ($kr) {
                    $kr = $kr.TrimEnd('\')
                    if (Test-Path (Join-Path $kr "Include")) { $kitRoot = $kr; break }
                }
            }
        }
    }
    if (-not $kitRoot) {
        foreach ($d in @("D:","E:","F:")) {
            $cand = Join-Path $d "Windows Kits\10"
            if (Test-Path (Join-Path $cand "Include")) { $kitRoot = $cand; break }
        }
    }
    if (-not $kitRoot) { Write-Host "[ERROR] 未找到 Windows SDK (Windows Kits\10\Include)" -ForegroundColor Red; exit 1 }
    $sdkVer = Get-ChildItem (Join-Path $kitRoot "Include") -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^\d+\.\d+\.\d+\.\d+$' } |
        Sort-Object Name -Descending | Select-Object -First 1 -ExpandProperty Name
    if (-not $sdkVer) { Write-Host "[ERROR] 未找到 Windows SDK Include 版本 ($kitRoot\Include)" -ForegroundColor Red; exit 1 }

    $tool   = Join-Path $msvcRoot $toolVer
    $binX64 = Join-Path $tool "bin\Hostx64\x64"
    $binX86 = Join-Path $tool "bin\Hostx64\x86"
    $incDir = Join-Path $kitRoot "Include"
    $libDir = Join-Path $kitRoot "Lib"
    # Join-Path 三位置参数写法在 Windows PowerShell 5.1 下静默丢掉第三段, 故显式拼接路径
    $Include = "$tool\include;$incDir\$sdkVer\um;$incDir\$sdkVer\ucrt;$incDir\$sdkVer\shared;$incDir\$sdkVer\winrt;$incDir\$sdkVer\cppwinrt"
    $LibX64  = "$tool\lib\x64;$libDir\$sdkVer\um\x64;$libDir\$sdkVer\ucrt\x64"
    $LibX86  = "$tool\lib\x86;$libDir\$sdkVer\um\x86;$libDir\$sdkVer\ucrt\x86"
    return [pscustomobject]@{ VsRoot=$vsRoot; ToolVer=$toolVer; SdkVer=$sdkVer;
        BinX64=$binX64; BinX86=$binX86; Include=$Include; LibX64=$LibX64; LibX86=$LibX86 }
}

$msvc = Get-MsvcToolset
# 原生 $env: 赋值: 子进程(含 C3 启动的 cl/link)可靠继承; 同时含 x64 与 x86 交叉工具链
$env:PATH    = "$($msvc.BinX64);$($msvc.BinX86);$env:PATH"
$env:INCLUDE = $msvc.Include
$env:LIB     = "$($msvc.LibX64);$($msvc.LibX86)"
$clCmd = Get-Command cl.exe -ErrorAction SilentlyContinue
if (-not $clCmd) { Write-Host "[ERROR] 设置 MSVC 环境后仍找不到 cl.exe (PATH 前段=$($msvc.BinX64))" -ForegroundColor Red; exit 1 }
Write-Host ("  MSVC 环境: VS=$($msvc.VsRoot)  toolset=$($msvc.ToolVer)  SDK=$($msvc.SdkVer)") -ForegroundColor Gray
Write-Host ("  cl.exe: $($clCmd.Source)") -ForegroundColor Gray

if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }

# === 测试基础函数 ===
$script:pass = 0
$script:fail = 0
$script:skip = 0
$script:total = 0

# === COM 测试前: 检查相关 COM 组件是否已注册 ===
# 仅当所需的 COM 组件已注册时, 才执行对应的 COM 测试 (例如 VBMANLIB)
# 若所需 COM 组件缺失: 相关用例 SKIP 而非 FAIL
# 32 位程序 (Arch=x86) 需读取 32 位注册表视图 (WOW6432Node), 本脚本统一处理
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

# === 冒烟测试 (编译+链接+运行) ===
function Test-Compile {
    param([string]$Name, [string]$Source)
    $script:total++
    Write-Host -NoNewline "  [COMPILE] $Name ... "
    
    $result = & $C3 $Source --output-dir $OutDir 2>&1
    $exitCode = $LASTEXITCODE
    
    if ($exitCode -eq 0) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        if ($Verbose) { Write-Host ($result | Out-String) }
    }
}

# === 编译测试 (能生成 exe 但没有 Main, 仅验证编译通过) ===
# 返回 @{ Ok; ExitCode; Output; Detail } 供调用方判定编译/运行结果
#
# 输出日志说明: 不同终端编码下, 输出内容可能略有差异,
# 使用方法: 参考 codes\smoke.bas 中带主入口的示例
#   start ".\build\C3.exe" 由调用方传入, 这里统一封装运行逻辑
#   启动并等待, 捕获标准输出/错误, 汇总为 @{ Ok; ExitCode; Output; Detail }
#   方案 A: .NET Process + Diagnostic (ReadToEndAsync, 不阻塞)
#   方案 B: Start-Process -RedirectStandardOutput (简单但不支持实时读取)
# 若启用了 VCVARSALL, 会自动按它配置编译环境, 生成的目标写于 "output\" 目录
function Invoke-TestExe {
    param(
        [string]$ExePath,
        [string]$WorkDir,
        [string]$Name
    )

    # 日志统一写入 output 目录: 使用 Open ... For Output 追加写入同一日志文件
    # (scores.txt / test_output.txt / *.dat 等), 不同进程写入不同实时文件
    $stdoutFile = Join-Path $WorkDir "$Name.out"
    $stderrFile = Join-Path $WorkDir "$Name.err"
    Remove-Item $stdoutFile, $stderrFile -ErrorAction SilentlyContinue

    $errors = @()

    # --- 方案 A: .NET Process + 重定向 (实时读取, 推荐) ---
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
        if (-not $proc.WaitForExit(5000)) {
            try { $proc.Kill() } catch { }
            $proc.WaitForExit()
            return @{
                Ok       = $false
                ExitCode = $null
                Output   = @()
                Detail   = "run timeout: 5s"
            }
        }
        $stdout = $soTask.Result
        $stderr = $seTask.Result

        # 输出可能包含 ANSI/UTF-8 混合编码, 按标准输出逐行读取以降低乱码 (当前按 ANSI 处理)
        [System.IO.File]::WriteAllText($stdoutFile, [string]$stdout, [System.Text.Encoding]::Default)
        [System.IO.File]::WriteAllText($stderrFile, [string]$stderr, [System.Text.Encoding]::Default)

        return @{
            Ok       = $true
            ExitCode = $proc.ExitCode
            Output   = (Get-Content $stdoutFile -ErrorAction SilentlyContinue)
            Detail   = "dotnet"
        }
    } catch {
        $errors += ("[dotnet] " + $_.Exception.Message)
    }

    # --- 方案 B: Start-Process -RedirectStandard* (简单, 但编码不可控) ---
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

    return @{
        Ok       = $false
        ExitCode = $null
        Output   = @()
        Detail   = ($errors -join " | ")
    }
}

# GUI smoke: require a visible main window, then close only the process we launch.
# This checks startup, not screenshot correctness or QR decoding.
function Test-GuiVbp {
    param([string]$Name, [string]$VbpFile, [string]$ExeName = "", [string]$Arch = "", [int]$AutoExitSec = 0)
    $script:total++
    Write-Host -NoNewline "  [GUI] $Name ... "
    $guiOut = Join-Path $OutDir $Name
    New-Item -ItemType Directory -Path $guiOut -Force | Out-Null
    if ($Arch) {
        $compileResult = & $C3 $VbpFile --arch $Arch --output-dir $guiOut 2>&1
    } else {
        $compileResult = & $C3 $VbpFile --output-dir $guiOut 2>&1
    }
    if ($LASTEXITCODE -ne 0) {
        $script:fail++
        Write-Host "FAIL (compile)" -ForegroundColor Red
        if ($Verbose) { Write-Host ($compileResult | Out-String) }
        return
    }
    # 自动把工程目录下的原生依赖 (OCX/DLL/TLB) 复制到 exe 所在目录 (与产物同目录),
    # 以支持免注册便携部署: 如第三方 OCX 控件无需本机注册即可 LoadLibrary 加载
    $srcDir = Split-Path $VbpFile -Parent
    if (Test-Path $srcDir) {
        Get-ChildItem $srcDir -File | Where-Object { $_.Extension -match '\.(ocx|dll|tlb)$' } | ForEach-Object {
            Copy-Item -Force $_.FullName $guiOut | Out-Null
        }
    }
    # VBP ExeName32 may differ from the vbp file name; pass -ExeName to override.
    $exeBase = if ($ExeName) { $ExeName } else { [IO.Path]::GetFileNameWithoutExtension($VbpFile) }
    $exe = Join-Path $guiOut ($exeBase + ".exe")
    $proc = $null
    try {
        $proc = Start-Process -FilePath $exe -WorkingDirectory $guiOut -PassThru -ErrorAction Stop
        $watch = [Diagnostics.Stopwatch]::StartNew()
        $windowSeen = $false
        while ($watch.ElapsedMilliseconds -lt 5000) {
            $proc.Refresh()
            if ($proc.HasExited) { break }
            if ($proc.MainWindowHandle -ne [IntPtr]::Zero) { $windowSeen = $true; break }
            Start-Sleep -Milliseconds 100
        }
        if (-not $windowSeen) { throw "Main window not available within 5s" }
        if ($AutoExitSec -gt 0) {
            # GUI demo with no clean-exit contract: window shown is enough;
            # auto-kill after the timeout so the suite never hangs.
            # Fix 189: 但"它自己先退了"不是我们让它退的。旧代码在这里无条件
            # pass++ 且不读退出码, 于是运行期崩溃 (实测窗口出现后 18s 的
            # 0xC0000005) 也算 PASS —— 只有超时兜底那条路径才该免检。
            $watch = [Diagnostics.Stopwatch]::StartNew()
            while ($watch.ElapsedMilliseconds -lt $AutoExitSec * 1000) {
                $proc.Refresh()
                if ($proc.HasExited) { break }
                Start-Sleep -Milliseconds 100
            }
            $proc.Refresh()
            if ($proc.HasExited) {
                $selfSec = [int]($watch.ElapsedMilliseconds / 1000)
                $code = $proc.ExitCode
                if ($code -ne 0) {
                    throw ("exited on its own at ~{0}s with code 0x{1:X8}" -f $selfSec, $code)
                }
                Write-Host "PASS (compile, window, self-exit at ~${selfSec}s, code 0)" -ForegroundColor Green
            } else {
                $proc.Kill(); $proc.WaitForExit(5000) | Out-Null
                Write-Host "PASS (compile, window, auto-exit after ${AutoExitSec}s)" -ForegroundColor Green
            }
            $script:pass++
        } else {
            if (-not $proc.CloseMainWindow()) { throw "Main window refused close" }
            if (-not $proc.WaitForExit(2000)) { throw "Application did not exit after close" }
            $proc.Refresh()
            if ($proc.ExitCode -ne 0) { throw "Exit code $($proc.ExitCode)" }
            $script:pass++
            Write-Host "PASS (compile, window, clean exit)" -ForegroundColor Green
        }
    } catch {
        $script:fail++
        Write-Host "FAIL ($($_.Exception.Message))" -ForegroundColor Red
    } finally {
        if ($proc -and -not $proc.HasExited) { $proc.Kill(); $proc.WaitForExit() }
    }
}

# === 编译冒烟测试 (编译+链接, 不运行生成物) ===
function Test-Run {
    param(
        [string]$Name, 
        [string]$Source,
        [string[]]$ExpectedOutputs,  # 预期输出 (可含多个子串, 逐一匹配)
        [string]$Arch = ""            # 可选架构参数 (x86/x64)
    )
    $script:total++
    Write-Host -NoNewline "  [RUN] $Name ... "
    
    # 编译: 输入源文件, 经中间C代码 -> cl/link -> 生成目标 (默认输出到 output 目录)
    if ($Arch) {
        $compileResult = & $C3 $Source --arch $Arch --output-dir $OutDir 2>&1
    } else {
        $compileResult = & $C3 $Source --output-dir $OutDir 2>&1
    }
    if ($LASTEXITCODE -ne 0) {
        $script:fail++
        Write-Host "FAIL (compile)" -ForegroundColor Red
        if ($Verbose) { Write-Host ($compileResult | Out-String) }
        return
    }
    
    # 验证编译产物 exe 是否存在: 若缺失则标记 FAIL (no exe) 并中断本次用例
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($Source)
    $exePath = Join-Path $OutDir "$baseName.exe"
    if (-not (Test-Path $exePath)) {
        $script:fail++
        Write-Host "FAIL (no exe)" -ForegroundColor Red
        return
    }
    
    # 运行冒烟测试: 校验输出 (详见 smoke 用例; 若超时则按 SKIP 处理)
    $run = Invoke-TestExe -ExePath $exePath -WorkDir $OutDir -Name $baseName
    if (-not $run.Ok) {
        $script:fail++
        Write-Host "FAIL (run error)" -ForegroundColor Red
        Write-Host ("    " + $run.Detail) -ForegroundColor Red
        return
    }
    $runOutput = $run.Output
    
    # 编译通过 + 冒烟运行通过
    if ($ExpectedOutputs -and $ExpectedOutputs.Count -gt 0) {
        $allMatch = $true
        foreach ($expected in $ExpectedOutputs) {
            $found = $runOutput | Where-Object { $_ -like "*$expected*" }
            if (-not $found) {
                $allMatch = $false
                break
            }
        }
        if ($allMatch) {
            $script:pass++
            Write-Host "PASS" -ForegroundColor Green
        } else {
            $script:fail++
            Write-Host "FAIL (output mismatch)" -ForegroundColor Red
            if ($Verbose) {
                Write-Host "  Expected: $($ExpectedOutputs -join ', ')"
                Write-Host "  Got: $($runOutput -join '`n')"
            }
        }
    } else {
        # 该测试用例预期会编译失败 (负向用例)
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    }
}

# === 纯 .bas 用例串行执行 (默认路径; 与 Tests 逐个调用 Test-Run 等价) ===
function Invoke-BasSetSerial {
    param([object[]]$Items)
    foreach ($it in $Items) {
        Test-Run $it.Name $it.Source $it.Expected $it.Arch
    }
}

# === 纯 .bas 用例并行执行 (多个 C3 实例同时编译+运行) ===
# 每个 worker 独立输出目录 (避免 exe/日志文件互相覆盖); 仅限 PowerShell 7+
# (ForEach-Object -Parallel); 每个 worker 返回汇总对象, 由调用方合并计数.
function Invoke-BasSetParallel {
    param([object[]]$Items, [int]$Jobs)
    if ($Items.Count -eq 0) { return }

    # 均分 (按遍历顺序切片, 每片尽可能均匀)
    $per = [int][Math]::Ceiling($Items.Count / [double]$Jobs)
    $shards = @()
    for ($i = 0; $i -lt $Items.Count; $i += $per) {
        $end = [Math]::Min($i + $per - 1, $Items.Count - 1)
        $shards += ,@(@( $Items[$i..$end] ), (Join-Path $OutDir ("job" + $shards.Count)))
    }
    if ($shards.Count -eq 0) { return }

    $results = $shards | ForEach-Object -Parallel {
        $shardItems = $_[0]
        $workDir    = $_[1]
        New-Item -ItemType Directory -Path $workDir -Force | Out-Null
        $c3 = $using:C3
        $p = 0; $f = 0; $details = @()
        foreach ($it in $shardItems) {
            if ($it.Arch) {
                $cr = & $c3 $it.Source --arch $it.Arch --output-dir $workDir 2>&1
            } else {
                $cr = & $c3 $it.Source --output-dir $workDir 2>&1
            }
            $ec = $LASTEXITCODE
            if ($ec -ne 0) { $f++; $details += "$($it.Name): compile FAIL"; continue }

            $baseName = [IO.Path]::GetFileNameWithoutExtension($it.Source)
            $exePath = Join-Path $workDir "$baseName.exe"
            if (-not (Test-Path $exePath)) { $f++; $details += "$($it.Name): no exe"; continue }

            # --- 运行 (语义与 Invoke-TestExe 一致; 5s 超时) ---
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
                if (-not $proc.WaitForExit(5000)) {
                    try { $proc.Kill() } catch { }
                    $proc.WaitForExit()
                    $f++; $details += "$($it.Name): run timeout 5s"; continue
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
                    $found = $runOut | Where-Object { $_ -like "*$exp*" }
                    if (-not $found) { $allMatch = $false; break }
                }
                if ($allMatch) { $p++ } else { $f++; $details += "$($it.Name): output mismatch" }
            } else {
                $p++
            }
        }
        [pscustomobject]@{ Pass = $p; Fail = $f; Details = $details }
    } -ThrottleLimit $Jobs

    foreach ($r in $results) {
        $script:pass += $r.Pass
        $script:fail += $r.Fail
        $script:total += ($r.Pass + $r.Fail)
        foreach ($d in $r.Details) {
            Write-Host "  [RUN] $d" -ForegroundColor Red
        }
    }
    $sumPass = ($results | Measure-Object -Property Pass -Sum).Sum
    $sumFail = ($results | Measure-Object -Property Fail -Sum).Sum
    Write-Host "  (parallel: $($results.Count) worker(s), pass=$sumPass fail=$sumFail)"
}

# === VBP 工程测试 (编译+链接+运行) ===
function Test-Vbp {
    param(
        [string]$Name,
        [string]$VbpFile,
        [string[]]$ExpectedOutputs,
        [string]$Arch = "",           # 可选架构参数 (x86/x64)
        [string]$RequiresCom = ""     # 依赖的 COM ProgId (未注册则 SKIP, 否则 FAIL)
    )
    $script:total++
    Write-Host -NoNewline "  [VBP] $Name ... "

    # 若依赖的 COM 组件未注册 (即缺失): 标记 SKIP (不影响通过率, 不计 FAIL)
    if ($RequiresCom -and -not (Test-ComRegistered $RequiresCom $Arch)) {
        $script:skip++
        $view = if ($Arch -eq "x86") { "WOW6432Node (32-bit)" } else { "64-bit" }
        Write-Host "SKIP (COM '$RequiresCom' 未注册 $view 视图)" -ForegroundColor Yellow
        return
    }

    # 编译 VBP 工程: 输入 VBP 文件, 经 cl/link 生成可执行文件
    if ($Arch) {
        $compileResult = & $C3 $VbpFile --arch $Arch --output-dir $OutDir 2>&1
    } else {
        $compileResult = & $C3 $VbpFile --output-dir $OutDir 2>&1
    }
    if ($LASTEXITCODE -ne 0) {
        $script:fail++
        Write-Host "FAIL (compile)" -ForegroundColor Red
        if ($Verbose) { Write-Host ($compileResult | Out-String) }
        return
    }

    # 处理 VBP 编译输出: 校验是否包含 expected 输出 (可含多个子串)
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($VbpFile)
    $exePath = Join-Path $OutDir "$baseName.exe"
    if (-not (Test-Path $exePath)) {
        $script:fail++
        Write-Host "FAIL (no exe)" -ForegroundColor Red
        return
    }

    # 编译失败则标记 FAIL (compile); 生成物缺失标记 FAIL (no exe)
    $run = Invoke-TestExe -ExePath $exePath -WorkDir $OutDir -Name $baseName
    if (-not $run.Ok) {
        $script:fail++
        Write-Host "FAIL (run error)" -ForegroundColor Red
        Write-Host ("    " + $run.Detail) -ForegroundColor Red
        return
    }
    $runOutput = $run.Output

    # 编译失败提示
    if ($ExpectedOutputs -and $ExpectedOutputs.Count -gt 0) {
        $allMatch = $true
        foreach ($expected in $ExpectedOutputs) {
            $found = $runOutput | Where-Object { $_ -like "*$expected*" }
            if (-not $found) {
                $allMatch = $false
                break
            }
        }
        if ($allMatch) {
            $script:pass++
            Write-Host "PASS" -ForegroundColor Green
        } else {
            $script:fail++
            Write-Host "FAIL (output mismatch)" -ForegroundColor Red
            if ($Verbose) {
                Write-Host "  Expected: $($ExpectedOutputs -join ', ')"
                Write-Host "  Got: $($runOutput -join '`n')"
            }
        }
    } else {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    }
}

# === 语法检查测试 ===
function Test-Syntax {
    param([string]$Name, [string]$Source)
    $script:total++
    Write-Host -NoNewline "  [SYNTAX] $Name ... "
    
    $result = & $C3 $Source --syntax-only 2>&1
    if ($LASTEXITCODE -eq 0) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        if ($Verbose) { Write-Host ($result | Out-String) }
    }
}

# =============================================
# 语法检查: 用 -syntax-only 验证源码合法性; 未生成目标时仍计为通过
# =============================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  C3 Compiler Test Suite" -ForegroundColor Cyan
Write-Host "  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# --- 语法测试 (--syntax-only) ---
# 说明 C3.exe 的构建链路: "源文件.bas -> 中间C -> cl/link -> 可执行文件"
# 参考 tests_github\t0_cases\smoke.bas (内含防呆提醒: 运行中弹 MsgBox 的用例需获得前台焦点)
if ($Category -in @("all", "smoke")) {
    Write-Host "--- Smoke Test (C3.exe end-to-end) ---" -ForegroundColor Yellow

    Test-Run "smoke" "$GHTests\smoke.bas" @("SMOKE-1:OK", "SMOKE-2:OK", "SMOKE-3:OK", "SMOKE PASS")
    Write-Host ""
}

# --- 纯 .bas 用例 (编译+运行; 支持 -Jobs 并行) ---
if ($Category -in @("all", "run", "bas")) {
    # 纯 .bas 用例统一入队; VBP/GUI 用例在独立分类 (vbp) 中保持串行
    $basQueue = @()
    function Add-BasTest {
        param([string]$Name, [string]$Source, $Expected = @(), [string]$Arch = "")
        $script:basQueue += @{
            Name     = $Name
            Source   = $Source
            Expected = @($Expected)
            Arch     = $Arch
        }
    }

    Write-Host "--- Regression (Compile+Run) ---" -ForegroundColor Yellow

    Add-BasTest "hello" "$Tests\hello.bas"
    Add-BasTest "test_m5" "$Tests\test_m5.bas"
    Add-BasTest "test_rtl" "$Tests\test_rtl.bas"
    Add-BasTest "test_array" "$Tests\test_array.bas" @("wa-clone=22", "wa-ub=3", "wa-str=65", "wa-rt=65", "=== Array Tests PASSED ===")
    Add-BasTest "test_fileio" "$Tests\test_fileio.bas"
    Add-BasTest "test_error" "$Tests\test_error.bas"
    Add-BasTest "test_now" "$Tests\test_now.bas"
    Add-BasTest "test_getput" "$Tests\test_getput.bas" @("PASS1a", "PASS1b", "PASS1c", "PASS2", "PASS3", "PASS4")
    Add-BasTest "test_onerror" "$Tests\test_onerror.bas" @("PASS1", "PASS2", "PASS3a", "PASS3b", "Done")
    Add-BasTest "test_ndarray" "$Tests\test_ndarray.bas" @("2D sum=270", "P8.1 ALL TESTS DONE")
    Add-BasTest "test_foreach" "$Tests\test_foreach4.bas" @("Long For Each: 150", "String For Each: Hello World")
    Add-BasTest "test_softkeyword" "$Tests\test_softkeyword.bas" @("Get=42", "Step=5", "Name=test")
    Add-BasTest "test_date" "$Tests\test_date.bas" @("PASS_Year", "PASS_Month", "PASS_Day", "PASS_NowYear", "Done")
    Add-BasTest "test_colon" "$Tests\test_colon.bas" @("PASS1", "PASS2", "PASS3", "Done")
    Add-BasTest "test_goto" "$GHTests\test_goto.bas" @("GOTO-PASS1", "GOTO-PASS2", "GOTO-PASS3", "GOTO-PASS4", "GOTO-PASS5", "GOTO-PASS6", "GOTO-PASS7", "GOTO DONE")
    Add-BasTest "test_gosub" "$GHTests\test_gosub.bas" @("GOSUB-PASS1", "GOSUB-PASS2", "GOSUB-PASS3", "GOSUB DONE")
    Add-BasTest "test_nested_udt_array" "$Tests\test_nested_udt_array.bas" @("NA1=1;NB1=D0;NC1=1.2", "NA5=5;NB5=D4;NC5=5.2", "HDR@ABC", "NESTED-DONE")
    Add-BasTest "test_udt_assign" "$Tests\test_udt_assign.bas" @("A1=1;S1=hello;N1=42", "A2=99;S2=world;N2=7", "H1=11;HS1=alpha", "E1=5;ES1=five", "L1=2;LS1=hello", "UDT-ASSIGN-DONE")
    Add-BasTest "test_date_display" "$Tests\test_date_display.bas" @("D1-noserial=Y", "D2-year=Y", "D2b-notime=Y", "D3-nextday=Y", "D4-nextyear=Y", "D5-diff0=Y", "D6-param=Y", "D7-longparam=Y", "D8-cstr=Y", "D9-format=Y", "DATE-DONE")
    Add-BasTest "test_variant" "$Tests\test_variant.bas" @("PASS1a", "PASS1c", "PASS5", "Done")
    Write-Host ""

        # --- P5.5 数据类型兼容性测试 ---
    Write-Host "--- Compat Tests (P5.5) ---" -ForegroundColor Yellow

    Add-BasTest "test_compat" "$Tests\test_compat.bas"
    Add-BasTest "test_types" "$Tests\test_types.bas"
    Add-BasTest "test_control" "$Tests\test_control.bas"
    Add-BasTest "test_declare" "$Tests\test_declare.bas"
    Write-Host ""

    # --- P5.7 语法/语义检查用例组 ---
    Write-Host "--- Bugfix Tests (P5.7) ---" -ForegroundColor Yellow

    Add-BasTest "test_fixes" "$Tests\test_fixes.bas" @("FIX1:OK", "FIX2:OK", "FIX3:OK", "All fixes passed!")
    Write-Host ""

    # --- P6 预处理器/冒烟测试用例组 ---
    Write-Host "--- COM Tests (P6) ---" -ForegroundColor Yellow

    Add-BasTest "test_com" "$Tests\test_com.bas" @("COM-1:OK", "COM-2:OK", "COM-3:OK", "COM:3/3")
    Add-BasTest "test_com2" "$Tests\test_com2.bas" @("Users")
    Add-BasTest "test_com3" "$Tests\test_com3.bas" @("COM3-1:OK", "COM3-2:OK", "COM3-3:OK", "COM3-4:OK", "COM3:4/4")
    Add-BasTest "test_earlybound" "$Tests\test_earlybound.bas" @("EB-1:OK", "EB-2:OK", "EB:2/2")
    Add-BasTest "test_p1324" "$Tests\test_p1324.bas" @("P13-1:OK", "P13-3:OK", "P13-5:OK", "P13:8/8")

    # --- P24 COM 测试用例 ---
    Write-Host "--- P24 COM Optimization Tests ---" -ForegroundColor Yellow

    Add-BasTest "test_p24" "$Tests\test_p24.bas" @("P24-01a:OK", "P24-01b:OK", "P24-01c:OK", "P24-03a:OK", "P24-03b:OK", "P24:5/5")
    Add-BasTest "test_earlybound2" "$Tests\test_earlybound2.bas" @("EB2-1:OK", "EB2-7:DriveType=2", "EB2-8:OK", "EB2-10:OK", "EB2:10/10") -Arch "x86"
    Add-BasTest "test_not_com" "$Tests\test_not_com.bas" @("NOT-COM:OK", "NOT-COM2:OK", "NOT-COM:PASS") -Arch "x86"
    Add-BasTest "test_err_obj" "$Tests\test_err_obj.bas" @("ERR-1:OK", "ERR-6:OK", "ERR:6/6")
    Add-BasTest "test_variant_cmp" "$Tests\test_variant_cmp.bas" @("VC-1:OK", "VC-4:OK", "VC:4/4")
    Add-BasTest "test_com_default_prop" "$Tests\test_com_default_prop.bas" @("DP-1:OK", "DP-4:OK", "P24-10: 4/4")
    Add-BasTest "test_com_optional" "$Tests\test_com_optional.bas" @("OP-1:OK", "OP-4:OK", "P24-11: 4/4")
    Add-BasTest "test_bstr_concat_scalar" "$Tests\test_bstr_concat_scalar.bas" @("BCS:16/16")
    # Fix 190: Declare "As Any" ByRef 的下标链实参必须取地址, 不能把元素值当指针
    Add-BasTest "test_asany_subscript" "$Tests\test_asany_subscript.bas" @("WITH-SUB=Y", "EXPR-SUB=Y", "SCALAR=Y", "CHAIN=Y", "ASANY-DONE")
    # Delegate (tB extension): typed function pointers, stdcall/cdecl thunks, both arches
    Add-BasTest "test_delegate" "$Tests\test_delegate.bas" @("CALL=Y", "INIT=Y", "REASSIGN=Y", "BITCMP=Y", "APICB=Y", "QSORT=Y", "MODVAR=Y", "DELEGATE-DONE")
    Add-BasTest "test_delegate_x86" "$Tests\test_delegate.bas" @("CALL=Y", "INIT=Y", "REASSIGN=Y", "BITCMP=Y", "APICB=Y", "QSORT=Y", "MODVAR=Y", "DELEGATE-DONE") -Arch "x86"
    # Overloading (tB extension): same-name by-type/arity resolution, Optional span, Variant tier
    Add-BasTest "test_overload" "$Tests\test_overload.bas" @("F-L5", "F-Shi", "F-L6", "3", "5", "O0", "O17", "21", "OVERLOAD-DONE")
    Add-BasTest "test_overload_x86" "$Tests\test_overload.bas" @("F-L5", "F-Shi", "F-L6", "3", "5", "O0", "O17", "21", "OVERLOAD-DONE") -Arch "x86"
    # Generics (tB extension): monomorphizing prepass — generic UDT (nested/multi-arity),
    # generic Function/Sub with explicit instantiation + call-site type inference.
    Add-BasTest "test_generics" "$Tests\test_generics.bas" @("G-A=12", "G-B=hi", "G-C=21 abc!", "G-D=10", "G-E=10", "G-F=ab", "G-G=20", "LEN=2", "G-H=0", "G-I=20", "GENERICS-DONE")
    Add-BasTest "test_generics_x86" "$Tests\test_generics.bas" @("G-A=12", "G-B=hi", "G-C=21 abc!", "G-D=10", "G-E=10", "G-F=ab", "G-G=20", "LEN=2", "G-H=0", "G-I=20", "GENERICS-DONE") -Arch "x86"

    # 分片: CI 用多 runner 并行跑 bas 用例时, 各 runner 只取第 BasShard 片
    if ($BasShardTotal -gt 1) {
        if ($BasShard -lt 1 -or $BasShard -gt $BasShardTotal) {
            Write-Host "[ERROR] BasShard 需在 1..BasShardTotal" -ForegroundColor Red
            exit 1
        }
        $shardLen = [Math]::Ceiling($basQueue.Count / $BasShardTotal)
        $shardStart = ($BasShard - 1) * $shardLen
        $basQueue = @($basQueue[$shardStart..([Math]::Min($shardStart + $shardLen - 1, $basQueue.Count - 1))])
        Write-Host "  (bas shard ${BasShard}/${BasShardTotal}: $($basQueue.Count) tests)" -ForegroundColor Cyan
    }

    # 执行纯 .bas 用例 (串行或并行)
    $basSw = [Diagnostics.Stopwatch]::StartNew()
    if ($Jobs -gt 1 -and $PSVersionTable.PSVersion.Major -ge 7) {
        Write-Host "--- Bas Tests (parallel, jobs=$Jobs) ---" -ForegroundColor Yellow
        Invoke-BasSetParallel -Items $basQueue -Jobs $Jobs
    } else {
        if ($Jobs -gt 1) {
            Write-Host "[WARN] -Jobs>1 需要 PowerShell 7+, 当前 $($PSVersionTable.PSVersion) 回退串行" -ForegroundColor Yellow
        }
        Invoke-BasSetSerial -Items $basQueue
    }
    $basSw.Stop()
    Write-Host "  (bas tests took $([Math]::Round($basSw.Elapsed.TotalSeconds))s)"
}

# --- VBP 工程测试 (串行; GUI 窗口效果无法通过自动校验并行确认) ---
if ($Category -in @("all", "run", "vbp")) {
    # --- VBP 工程测试 (P5) --- (串行)
    Write-Host "--- VBP Project Tests (P5) ---" -ForegroundColor Yellow
    $vbpSw = [Diagnostics.Stopwatch]::StartNew()

    Test-Vbp "test_class" "$Tests\test_class.vbp" @("3", "0")

    # Cross-module overload resolution (O3): modA exports Sub/Function overload
    # groups; modB Main calls them. Deferred re-resolution must pick variants.
    Test-Vbp "ovl_xmod" "$Tests\ovl_xmod\test_ovl_xmod.vbp" @("XM1=105", "XM2=203", "XM3=SUB7", "XM3=STRhi", "XMOD-DONE")

    # Cross-module generics (G3): generic UDT + Function/Sub exported from mod A,
    # consumed in mod B via explicit instantiation and bare-call inference.
    Test-Vbp "gen_xmod" "$Tests\gen_xmod\gen_xmod.vbp" @("XM1=35", "XM2=ab", "XM3=12", "XM4=hi", "XM5=7", "XM6=99", "XLEN=2", "XM7=34", "XMOD-DONE")
    # Generic classes (G4): .cls `Class Name(Of T)` header template + whole-module
    # specialization clone (two specializations coexist; self-referential LNode).
    Test-Vbp "gen_cls" "$Tests\gen_cls\gen_cls.vbp" @("GC1=42", "GC2=box:42", "GC3=hey", "GC4=box:hey", "GC5=33", "GC6=y", "GCLS-DONE")

    Test-Vbp "M6Test" "$Tests\M6Test.vbp" @("M6A:OK", "M6B:OK", "M6C:OK", "M6D:OK", "M6 PASSED")
    Test-Vbp "modulemethod" "$Tests\test_modulemethod.vbp" @("30", "21")

    # QR code project (tests\VbQRCodegen-master): form loads, sets Image1.Picture via Stretch
    Test-GuiVbp "VbQRCodegen" "$Tests\VbQRCodegen-master\test\Project1.vbp"
    # BalloonTooltips: form loads with controls + creates its common-controls tooltip windows (x64).
    Test-GuiVbp "BalloonTooltips" "$Tests\BalloonTooltips\prjBalloonTooltips.vbp" -ExeName "BalloonTooltips"
    # Charts 2020 demo (3rd-party UserControl charts): windowless chart controls (x86 first;
    # x64 after LongPtr port of API pointers/handles in the .ctl/.cls sources).
    Test-GuiVbp "Charts2020" "$Tests\Charts 2020\Proyecto1.vbp" -Arch "x86" -AutoExitSec 3
    # czUI (czForm): 自定义 GDI+ UserControl (.ctl) 无边框窗体 demo, 需 -Arch x86 (32 位)
    Test-GuiVbp "czUI" "$Tests\czUI-main\czFormDemo.vbp" -Arch "x86" -AutoExitSec 3
    # NewTab: 第三方 OCX 控件 (NewTab01.ocx, 32 位) 真宿主验证. 免注册便携部署 (OCX 在工程目录, 由 harness 复制到 exe 旁, 不依赖本机注册);
    # 无边框窗体无关闭按钮/无自动退出逻辑, 用 -AutoExitSec 3 收尾避免阻塞后续测试
    Test-GuiVbp "NewTab" "$Tests\NewTab-test\Test.vbp" -Arch "x86" -AutoExitSec 3
    # ExtShow: 跨模块窗体默认实例"无参" Show (Fix 146 回归靶, 2026-09-20 vbman C2198):
    # .bas caller 调 Form2.Show, 定义侧签名 (hMDIClient, modal) 后调用侧须补 modal=0
    Test-GuiVbp "ExtShow" "$Tests\ext_show_test\test_ext_show.vbp" -AutoExitSec 3
    Write-Host ""

    Test-Vbp "test_implements" "$Tests\test_implements.vbp" @("IMPL1:OK", "IMPL2:OK", "Implements test PASSED")
    Test-Vbp "test_events" "$Tests\test_events\test_events.vbp" @("Events test PASSED")
    Test-Vbp "M7Test" "$Tests\m7_test\M7Test.vbp" @("4/4 PASSED")

    # test_vbman 用于验证外部 COM 组件 VBMANLIB (x86 DLL, 供 32 位程序调用)
    Test-Vbp "test_vbman" "$Tests\test_vbman\test_vbman.vbp" @("P24-04a:OK", "P24-04b:OK", "P24-04:2/2") -Arch "x86" -RequiresCom "VBMANLIB.cVBMAN"
    $vbpSw.Stop()
    Write-Host "  (vbp/gui tests took $([Math]::Round($vbpSw.Elapsed.TotalSeconds))s)"
    Write-Host ""
}

if ($Category -in @("all", "compile")) {
    # --- 综合测试 (编译+运行, 以 Main 为程序入口) ---
    Write-Host "--- Compile Tests ---" -ForegroundColor Yellow
    
    Test-Compile "test_comprehensive" "$Tests\test_comprehensive.bas"
    Test-Compile "test_comprehensive2" "$Tests\test_comprehensive2.bas"
    Write-Host ""
    
    # --- P7 窗体测试 (GUI 验证: 窗体正常加载即可) ---
    Write-Host "--- Form Compile Tests (P7) ---" -ForegroundColor Yellow
    
    $formTests = @(
        "test_form\empty_form.frm",
        "test_form\form_test_p74.frm",
        "form_test_p75.frm",
        "form_test_p76.frm",
        "form_test_p78.frm",
        "form_test_p79.frm",
        "form_test_m8.frm",
        "form_mdi_parent.frm"
    )
    
    foreach ($t in $formTests) {
        $path = Join-Path $Tests $t
        if (Test-Path $path) {
            $name = [System.IO.Path]::GetFileNameWithoutExtension($t)
            Test-Compile $name $path
        }
    }
    Write-Host ""
}

if ($Category -in @("all", "syntax")) {
    # --- 生成物 / 输出目录说明 ---
    Write-Host "--- Syntax/Semantic Tests ---" -ForegroundColor Yellow
    
    # T0 拆分: 11 个用例移入 tests_github\t0_cases\; test_onerror 留 tests\ (run 类双登记, 不可挪)
    # test_goto / test_gosub 已升级为「编译+运行+输出比对」用例, 移入 bas 队列
    $syntaxTests = @(
        "$GHTests\test_basic.bas",
        "$GHTests\test_for.bas", "$GHTests\test_for2.bas",
        "$GHTests\test_select.bas",
        "$Tests\test_onerror.bas",
        "$GHTests\test_redim.bas",
        "$GHTests\test_setlet.bas", "$GHTests\test_setonly.bas",
        "$GHTests\test_assign.bas",
        "$GHTests\test_sem_minimal.bas", "$GHTests\test_sem_proc.bas", "$GHTests\test_semantic.bas"
    )

    foreach ($t in $syntaxTests) {
        $path = $t
        if (Test-Path $path) {
            $name = [System.IO.Path]::GetFileNameWithoutExtension($t)
            Test-Syntax $name $path
        }
    }
    Write-Host ""
    
    # --- 生成环境检查与汇总 (冒烟+语法+VBP+run 计数) ---
    Write-Host "--- Preprocessor Tests ---" -ForegroundColor Yellow
    
    # T0 拆分: pp 系列全部移入 tests_github\t0_cases\, 跨目录引用
    $ppTests = @(
        "$GHTests\test_preprocess.bas",
        "$GHTests\test_pp_minimal.bas",
        "$GHTests\test_pp2.bas", "$GHTests\test_pp3.bas", "$GHTests\test_pp4.bas", "$GHTests\test_pp5.bas",
        "$GHTests\test_pp6.bas", "$GHTests\test_pp7.bas", "$GHTests\test_pp8.bas", "$GHTests\test_pp9.bas",
        "$GHTests\test_pp10.bas", "$GHTests\test_pp11.bas", "$GHTests\test_pp12.bas", "$GHTests\test_pp13.bas",
        "$GHTests\test_pp14.bas", "$GHTests\test_pp15.bas", "$GHTests\test_pp16.bas", "$GHTests\test_pp17.bas",
        "$GHTests\test_pp18.bas"
    )

    foreach ($t in $ppTests) {
        $path = $t
        if (Test-Path $path) {
            $name = [System.IO.Path]::GetFileNameWithoutExtension($t)
            Test-Syntax $name $path
        }
    }
    Write-Host ""
}

# =============================================
# 若需调试单用例, 可设置 \$Verbose 后调用 Test-Run/Test-Compile/Test-Gui 子函数
# =============================================
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Results: PASS=$script:pass FAIL=$script:fail SKIP=$script:skip TOTAL=$script:total" -ForegroundColor $(if ($script:fail -gt 0) { "Red" } else { "Green" })
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

if ($script:fail -gt 0) { exit 1 } else { exit 0 }
