# baseline.ps1 - M1 基线固化：一条命令跑通三条最小路径
#
#   1) tests\hello.bas                            控制台：编译 + 运行 + 输出断言 + 产物大小
#   2) tests\btnfocus\BfApp.vbp                   .frm/控件：编译 + 运行(Timer 自关) + 日志检查点
#   3) tests\test_activex_dll\test_activex_dll.vbp  ActiveX DLL：--dll 编译 + 产物 + COM 导出
#
# 本地与 CI 用的是**同一份脚本**；CI 里由 .github/workflows/ci.yml 的 baseline job 调用。
# 只做构建/运行/断言，不改编译器语义。
#
# TEMP/TMP 重定向到 <repo>\.temp\baseline：C3 会把内嵌 RTL 解包到 %TEMP%\C3C\<会话>\rtl，
# 而受限环境（沙箱/CI 外层策略）下直接写系统 %TEMP% 会被拒（实测 "cannot create session
# directory ... rtl (拒绝访问)"）。重定向后这条路径在仓库内，本进程与其子进程都继承。
#
# 用法:
#   powershell -ExecutionPolicy Bypass -File tests\baseline.ps1
#   powershell -ExecutionPolicy Bypass -File tests\baseline.ps1 -Verbose
#   powershell -ExecutionPolicy Bypass -File tests\baseline.ps1 -C3Path D:\path\C3.exe
#
# 退出码: 0 = 三条全过; 1 = 至少一条 FAIL; 2 = 环境缺失（SKIP，不是产品红）

param(
    [string]$C3Path = "",
    [string]$OutputDirectory = "",
    [switch]$Verbose
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
$Tests = $PSScriptRoot

# ---------------------------------------------------------------------------
# 0. 环境默认值：TEMP 重定向
# ---------------------------------------------------------------------------
$TempRoot = Join-Path $Root ".temp\baseline"
New-Item -ItemType Directory -Force -Path $TempRoot | Out-Null
$env:TEMP = $TempRoot
$env:TMP = $TempRoot

if (-not $C3Path) { $C3Path = Join-Path $Root ".build\C3.exe" }
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $Root "output\baseline" }
$C3 = $C3Path
$script:Pass = 0
$script:Fail = 0
$script:Skip = 0

# ---------------------------------------------------------------------------
# 1. 断言/进程助手（必须先于任何调用定义：PS 是顺序执行）
# ---------------------------------------------------------------------------
function Test-NeedleHit {
    param($Lines, [string]$Needle)
    foreach ($l in @($Lines)) {
        if ($null -eq $l) { continue }
        if (([string]$l).IndexOf($Needle, [StringComparison]::OrdinalIgnoreCase) -ge 0) { return $true }
    }
    return $false
}

function Assert-Artifact {
    param([string]$Path, [int]$MinBytes, [string]$Label)
    if (-not (Test-Path -LiteralPath $Path)) {
        throw ("$Label 产物缺失: $Path")
    }
    $len = (Get-Item -LiteralPath $Path).Length
    if ($len -lt $MinBytes) {
        throw ("$Label 产物过小: $Path ($len B < $MinBytes B)")
    }
    return $len
}

# 统一用 .NET Process 捕获 stdout/stderr：既避开 PS5.1 下 `native 2>&1` +
# $ErrorActionPreference=Stop 会抛 NativeCommandError 的坑，也保证超时可控。
function Invoke-Capture {
    param(
        [string]$ExePath,
        [string[]]$Arguments = @(),
        [string]$WorkDir = "",
        [int]$TimeoutSec = 60,
        [hashtable]$EnvVars = @{}
    )
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $ExePath
    if ($Arguments -and $Arguments.Count -gt 0) {
        $quoted = @()
        foreach ($a in $Arguments) {
            if (([string]$a) -match '\s') { $quoted += ('"' + [string]$a + '"') } else { $quoted += [string]$a }
        }
        $psi.Arguments = $quoted -join ' '
    }
    if ($WorkDir) { $psi.WorkingDirectory = $WorkDir }
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    foreach ($k in $EnvVars.Keys) { $psi.EnvironmentVariables[$k] = [string]$EnvVars[$k] }

    $proc = [System.Diagnostics.Process]::Start($psi)
    $so = $proc.StandardOutput.ReadToEndAsync()
    $se = $proc.StandardError.ReadToEndAsync()
    if (-not $proc.WaitForExit($TimeoutSec * 1000)) {
        try { $proc.Kill() } catch { }
        try { $proc.WaitForExit(5000) | Out-Null } catch { }
        return @{
            Ok = $false; Timeout = $true; ExitCode = $null
            Out = [string]$so.Result; Err = [string]$se.Result
        }
    }
    return @{
        Ok = $true; Timeout = $false; ExitCode = $proc.ExitCode
        Out = [string]$so.Result; Err = [string]$se.Result
    }
}

function Write-CompileLog {
    param([string]$Path, $Result)
    try {
        $text = "exit=$($Result.ExitCode)`r`n--- stdout ---`r`n$($Result.Out)`r`n--- stderr ---`r`n$($Result.Err)`r`n"
        [System.IO.File]::WriteAllText($Path, $text, (New-Object System.Text.UTF8Encoding($true)))
    } catch { }
}

function Get-CompileTail {
    param($Result)
    $all = @(([string]$Result.Err -split "`r?`n") + ([string]$Result.Out -split "`r?`n")) |
        Where-Object { $_ -and $_.Trim().Length -gt 0 }
    return (($all | Select-Object -Last 8) -join " | ")
}

function Invoke-Compile {
    param([string]$Dir, [string]$LogName, [string[]]$Arguments)
    if (Test-Path -LiteralPath $Dir) { Remove-Item -LiteralPath $Dir -Recurse -Force }
    New-Item -ItemType Directory -Force -Path $Dir | Out-Null
    $result = Invoke-Capture -ExePath $C3 -Arguments $Arguments -WorkDir $Dir -TimeoutSec 600
    $log = Join-Path $Dir $LogName
    Write-CompileLog $log $result
    if ($result.Timeout) {
        throw ("编译超时 (600s); 日志: $log; 尾部: " + (Get-CompileTail $result))
    }
    if ($result.ExitCode -ne 0) {
        throw ("编译失败 rc=$($result.ExitCode); 日志: $log; 尾部: " + (Get-CompileTail $result))
    }
    return $result
}

function Assert-RunOk {
    param($Run, [string]$Dir, [int]$TimeoutSec, [string]$What)
    [System.IO.File]::WriteAllText((Join-Path $Dir "run.out"), $Run.Out, [System.Text.Encoding]::Default)
    [System.IO.File]::WriteAllText((Join-Path $Dir "run.err"), $Run.Err, [System.Text.Encoding]::Default)
    if ($Run.Timeout) { throw ("$What 运行超时 (${TimeoutSec}s); 日志: $Dir") }
    if ($Run.ExitCode -ne 0) {
        throw ("$What 运行退出码 $($Run.ExitCode) (期望 0); 日志: " + (Join-Path $Dir "run.err"))
    }
}

# 加载 MSVC 环境。做法与 scripts\env.ps1 相同（vcvarsall -> cmd set -> 导入当前进程），
# 但不能直接 dot-source 它：该文件首行有一个历史遗留的 `?` 前缀，在 try/catch 里会变成
# 终止错误并中断整个脚本（实测 cl.exe 因此找不到）。这里复用同一套 vcvarsall/vswhere
# 发现逻辑，避免动那个与本次任务无关的文件。
function Import-MsvcEnv {
    param([int]$TimeoutSec = 180)
    $vcvars = $env:C3_VCVARSALL
    if (-not $vcvars -or -not (Test-Path -LiteralPath $vcvars)) {
        $vswhere = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
        if (Test-Path -LiteralPath $vswhere) {
            $vsPath = & $vswhere -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath 2>$null
            if ($vsPath) {
                $candidate = Join-Path $vsPath "VC\Auxiliary\Build\vcvarsall.bat"
                if (Test-Path -LiteralPath $candidate) { $vcvars = $candidate }
            }
        }
    }
    if (-not $vcvars -or -not (Test-Path -LiteralPath $vcvars)) { return 0 }
    $bat = Join-Path $TempRoot ("vcvars_" + [Guid]::NewGuid().ToString("N") + ".bat")
    [System.IO.File]::WriteAllText($bat, ('call "' + $vcvars + '" x64 >nul 2>&1' + "`r`n" + "set`r`n"), [System.Text.Encoding]::ASCII)
    $r = Invoke-Capture -ExePath $env:ComSpec -Arguments @("/c", $bat) -TimeoutSec $TimeoutSec
    $n = 0
    foreach ($line in (([string]$r.Out) -split "`r?`n")) {
        if ($line -match '^([^=]+)=(.*)$') {
            [Environment]::SetEnvironmentVariable($matches[1], $matches[2], "Process")
            $n++
        }
    }
    return $n
}

# ---------------------------------------------------------------------------
# 2. 头信息 + 工具链
# ---------------------------------------------------------------------------
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  C3 baseline (M1)" -ForegroundColor Cyan
Write-Host "  C3      : $C3" -ForegroundColor Gray
Write-Host "  output  : $OutputDirectory" -ForegroundColor Gray
Write-Host "  TEMP    : $env:TEMP" -ForegroundColor Gray
Write-Host "========================================" -ForegroundColor Cyan

if (-not (Test-Path -LiteralPath $C3)) {
    Write-Host "FAIL: C3.exe 不存在: $C3" -ForegroundColor Red
    Write-Host "      先构建: scripts\build.bat  (或 cmake --build .build)" -ForegroundColor Red
    exit 1
}
$verRun = Invoke-Capture -ExePath $C3 -Arguments @("--version") -WorkDir $Root -TimeoutSec 60
$ver = (([string]$verRun.Out) + ([string]$verRun.Err)).Trim()
Write-Host "  version : $ver" -ForegroundColor Gray
if (-not $verRun.Ok -or $verRun.ExitCode -ne 0) {
    Write-Host "FAIL: C3 --version 退出码 $($verRun.ExitCode)" -ForegroundColor Red
    exit 1
}

# 复用仓库同一套 vcvarsall/vswhere 发现逻辑。失败时 SKIP 而非产品红。
$envCount = Import-MsvcEnv
$msvcOk = $null -ne (Get-Command cl.exe -ErrorAction SilentlyContinue)
if (-not $msvcOk) {
    Write-Host "SKIP: 找不到 MSVC (cl.exe)；baseline 需要 cl/link 才能产出可执行文件" -ForegroundColor Yellow
    Write-Host "      缺什么: Visual Studio 的 [使用 C++ 的桌面开发] 工作负载 / vcvarsall.bat" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  Baseline: PASS=0 FAIL=0 SKIP=3 TOTAL=3" -ForegroundColor Yellow
    exit 2
}
Write-Host ("  cl.exe  : " + (Get-Command cl.exe).Source + "  (env vars imported: $envCount)") -ForegroundColor Gray

if (-not (Test-Path -LiteralPath $OutputDirectory)) {
    New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
}

# ---------------------------------------------------------------------------
# 3. tests\hello.bas —— 控制台输出
# ---------------------------------------------------------------------------
function Test-BaselineHello {
    $name = "hello"
    $dir = Join-Path $OutputDirectory $name
    $src = Join-Path $Tests "hello.bas"
    Write-Host ""
    Write-Host "[1/3] hello.bas" -ForegroundColor Yellow

    Invoke-Compile -Dir $dir -LogName "compile.log" -Arguments @($src, "--output-dir", $dir) | Out-Null
    $exe = Join-Path $dir "hello.exe"
    $size = Assert-Artifact -Path $exe -MinBytes 100000 -Label "hello"
    Write-Host ("  compile: rc=0  exe={0:N0} B" -f $size) -ForegroundColor Gray

    $run = Invoke-Capture -ExePath $exe -WorkDir $dir -TimeoutSec 30
    Assert-RunOk -Run $run -Dir $dir -TimeoutSec 30 -What "hello"

    $needles = @("Count: ", "Sum = ", "5050")
    $lines = @(([string]$run.Out) -split "`r?`n")
    $missing = @()
    foreach ($n in $needles) { if (-not (Test-NeedleHit -Lines $lines -Needle $n)) { $missing += $n } }
    if ($missing.Count -gt 0) {
        throw ("输出缺少断言串: " + ($missing -join ", ") + "; 实到: " + (($lines | Where-Object { $_ }) -join " | ") + "; 日志: " + (Join-Path $dir "run.out"))
    }
    Write-Host "  run    : rc=0  checkpoints=3/3 (Count: / Sum = / 5050)" -ForegroundColor Gray
    $script:Pass++
    Write-Host "  PASS" -ForegroundColor Green
}

# ---------------------------------------------------------------------------
# 4. tests\btnfocus\BfApp.vbp —— .frm + 嵌套控件 + Timer 自关
# ---------------------------------------------------------------------------
function Test-BaselineForm {
    $name = "bfapp"
    $dir = Join-Path $OutputDirectory $name
    $vbp = Join-Path $Tests "btnfocus\BfApp.vbp"
    Write-Host ""
    Write-Host "[2/3] btnfocus\BfApp.vbp  (.frm)" -ForegroundColor Yellow

    Invoke-Compile -Dir $dir -LogName "compile.log" -Arguments @($vbp, "--output-dir", $dir) | Out-Null
    $exe = Join-Path $dir "BfApp.exe"
    $size = Assert-Artifact -Path $exe -MinBytes 100000 -Label "BfApp"
    Write-Host ("  compile: rc=0  exe={0:N0} B" -f $size) -ForegroundColor Gray

    $log = Join-Path $dir "btnfocus.log"
    if (Test-Path -LiteralPath $log) { Remove-Item -LiteralPath $log -Force }
    $run = Invoke-Capture -ExePath $exe -WorkDir $dir -TimeoutSec 60
    [System.IO.File]::WriteAllText((Join-Path $dir "run.out"), $run.Out, [System.Text.Encoding]::Default)
    [System.IO.File]::WriteAllText((Join-Path $dir "run.err"), $run.Err, [System.Text.Encoding]::Default)
    if ($run.Timeout) { throw ("BfApp 运行超时 (60s); 未形成自关 (t1_Timer -> Unload Me); 日志: $dir") }
    if ($run.ExitCode -ne 0) { throw ("BfApp 运行退出码 $($run.ExitCode) (期望 0); 日志: " + (Join-Path $dir "run.err")) }
    if (-not (Test-Path -LiteralPath $log)) { throw ("缺少窗体运行日志: $log (期望 t1_Timer 写出 BF-* 检查点)") }

    $text = [System.IO.File]::ReadAllText($log, [System.Text.Encoding]::Default)
    $lines = @($text -split "`r?`n")
    $needles = @("BF-both=Y", "BF-noclick=Y", "BF-each=Y", "BTNFOCUS-DONE")
    $missing = @()
    foreach ($n in $needles) { if (-not (Test-NeedleHit -Lines $lines -Needle $n)) { $missing += $n } }
    if ($missing.Count -gt 0) {
        throw ("窗体检查点缺失: " + ($missing -join ", ") + "; 实到: " + (($lines | Where-Object { $_ }) -join " | ") + "; 日志: $log")
    }
    if ($text -match "BF-each=N") { throw ("窗体检查点 BF-each=N（控件事件计数不齐）; 日志: $log") }
    Write-Host ("  run    : rc=0  checkpoints=4/4  (窗口存活到 Timer 自关)") -ForegroundColor Gray
    $script:Pass++
    Write-Host "  PASS" -ForegroundColor Green
}

# ---------------------------------------------------------------------------
# 5. tests\test_activex_dll\test_activex_dll.vbp —— ActiveX DLL
# ---------------------------------------------------------------------------
function Test-BaselineActiveXDll {
    $name = "axdll"
    $dir = Join-Path $OutputDirectory $name
    $vbp = Join-Path $Tests "test_activex_dll\test_activex_dll.vbp"
    Write-Host ""
    Write-Host "[3/3] test_activex_dll.vbp  (ActiveX DLL)" -ForegroundColor Yellow

    Invoke-Compile -Dir $dir -LogName "compile.log" -Arguments @($vbp, "--dll", "--output-dir", $dir) | Out-Null
    $dll = Join-Path $dir "test_activex_dll.dll"
    $size = Assert-Artifact -Path $dll -MinBytes 100000 -Label "ActiveX DLL"
    Write-Host ("  compile: rc=0  dll={0:N0} B" -f $size) -ForegroundColor Gray

    # COM 服务器必须导出这三枚（缺一则不是可注册的 in-proc server）。
    $dumpbin = Get-Command dumpbin.exe -ErrorAction SilentlyContinue
    if ($dumpbin) {
        $dump = Invoke-Capture -ExePath $dumpbin.Source -Arguments @("/exports", $dll) -WorkDir $dir -TimeoutSec 120
        $exports = ([string]$dump.Out) + ([string]$dump.Err)
        $need = @("DllGetClassObject", "DllRegisterServer", "DllUnregisterServer")
        $missing = @($need | Where-Object { $exports.IndexOf($_, [StringComparison]::OrdinalIgnoreCase) -lt 0 })
        if ($missing.Count -gt 0) {
            throw ("COM 导出缺失: " + ($missing -join ", ") + "; dumpbin rc=$($dump.ExitCode); 输出: " + $exports.Substring(0, [Math]::Min(400, $exports.Length)))
        }
        Write-Host "  exports: DllGetClassObject / DllRegisterServer / DllUnregisterServer" -ForegroundColor Gray
    } else {
        Write-Host "  exports: dumpbin.exe 不在 PATH，跳过导出断言（产物存在/大小已断言）" -ForegroundColor DarkYellow
    }
    $script:Pass++
    Write-Host "  PASS" -ForegroundColor Green
}

# ---------------------------------------------------------------------------
# 6. 执行三条；任何一条失败不影响其余两条（便于一次看到全部现场）
# ---------------------------------------------------------------------------
$steps = @(
    @{ Name = "hello"; Fn = { Test-BaselineHello } },
    @{ Name = "bfapp"; Fn = { Test-BaselineForm } },
    @{ Name = "axdll"; Fn = { Test-BaselineActiveXDll } }
)
foreach ($s in $steps) {
    try {
        & $s.Fn
    } catch {
        $script:Fail++
        Write-Host ("  FAIL: " + $_.Exception.Message) -ForegroundColor Red
        Write-Host ("        现场目录: " + (Join-Path $OutputDirectory $s.Name)) -ForegroundColor DarkGray
        if ($Verbose) { Write-Host ($_.ScriptStackTrace) -ForegroundColor DarkGray }
    }
}

Write-Host ""
$color = if ($script:Fail -gt 0) { "Red" } elseif ($script:Skip -gt 0) { "Yellow" } else { "Green" }
Write-Host ("  Baseline: PASS={0} FAIL={1} SKIP={2} TOTAL=3" -f $script:Pass, $script:Fail, $script:Skip) -ForegroundColor $color
if ($script:Fail -gt 0) { exit 1 }
if ($script:Skip -gt 0) { exit 2 }
exit 0
