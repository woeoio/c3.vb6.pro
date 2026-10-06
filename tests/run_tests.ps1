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
    [int]$BasShardTotal = 1, # bas 分片总数 (供 CI 用多 runner 并行跑 bas 用例)
    [int]$VbpShard = 0,      # vbp 分片当前编号 (1..VbpShardTotal); 0=不分片整队跑
    [int]$VbpShardTotal = 1, # vbp 分片总数 (供 CI 用多 runner 并行跑 vbp 用例)
    # 单条用例「运行」阶段的墙钟预算 (秒)。5s 在 -Jobs 20 的并行编译下会误杀: 4 vCPU runner 上
    # 被 cl/link 压住时, 健康的小 exe 也可能 >5s 才跑完 (实测卡住的进程只有 15ms CPU、线程 Ready)。
    # 真挂 (模态框/死锁) 靠这个上限兜底; 超时时会把 CPU 时间/状态/最后一行输出写进日志, 见下两处。
    [int]$RunTimeoutSec = 60,
    # ai/030 T30-B(本地先行): 传了这个开关, 套件里每次**真实构建**才带 --incremental。
    # 不传时 @IncArg 是空数组, 命令行逐字不变 (空数组 splat 的透明性单独验过: 输出与退出码都不动)。
    # 默认路径 = 今天的默认路径。
    [switch]$Incremental
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

# ai/030 T30-B: 只有 -Incremental 时才有的实参; 见 param 处的说明。
$IncArg = @()
if ($Incremental) { $IncArg = @('--incremental') }
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
    $Include = "$tool\include;$(Join-Path $incDir $sdkVer um);$(Join-Path $incDir $sdkVer ucrt);$(Join-Path $incDir $sdkVer shared);$(Join-Path $incDir $sdkVer winrt);$(Join-Path $incDir $sdkVer cppwinrt)"
    $LibX64  = "$(Join-Path $tool lib x64);$(Join-Path $libDir $sdkVer um x64);$(Join-Path $libDir $sdkVer ucrt x64)"
    $LibX86  = "$(Join-Path $tool lib x86);$(Join-Path $libDir $sdkVer um x86);$(Join-Path $libDir $sdkVer ucrt x86)"
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
# 运行超时时的读数: 报出已烧掉的 CPU 时间、进程状态与最后一行输出。
# 判据: CPU≈0 且没有任何输出 = 进程根本没被调度 (机器被并行编译压满, 不是用例的错);
#       CPU 不小却一直不退出 = 它自己在转/在等, 那是真问题。
function Get-RunTimeoutDiag {
    param($Proc)
    $cpu = -1; $st = '?'; $win = ''
    try { $cpu = [int]$Proc.TotalProcessorTime.TotalMilliseconds } catch { }
    try { if ($Proc.HasExited) { $st = "exited=$($Proc.ExitCode)" } else { $st = 'alive' } } catch { }
    # GUI 用例最有用的那一条: 它还带着哪个顶层窗。有窗 = 卡在消息循环里 (窗体没被 Unload);
    # 无窗 = 卡在创建/退出之外的什么地方 (或压根没起来)。
    try { $Proc.Refresh(); $win = [string]$Proc.MainWindowTitle } catch { }
    $parts = @("cpu=${cpu}ms", $st)
    if ($win) { $parts += "win='$win'" }
    return " (" + ($parts -join ', ') + ")"
}

# 超时被杀**之后**才读得到的东西: Kill 掉进程 ⇒ 管道关闭 ⇒ ReadToEndAsync 这时才完成,
# 于是拿得到它在被杀前已经写出的那部分输出 (返回 @{ Out; Err; Last; Lines })。
#
# ⚠ 旧版把这段放在 Kill **之前**: 进程还活着 ⇒ 管道绝不关闭 ⇒ ReadToEndAsync 不可能完成,
#   Wait(1500) 必然超时 ⇒ last 恒为 '' —— 每一次卡死都只能读到 `last=''`, 等于没有线索
#   (GA 上 ctrlsstab / ctrldlg_probe 那两例GUI 超时就是这么"看不见"的)。
#   顺序必须是: 量 CPU/存活/窗口 → Kill → WaitForExit → 才读输出。
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
$script:pass = 0
$script:fail = 0
$script:skip = 0
$script:total = 0

# vbp 分片闸门: 每个 CI job 只跑 vbp 队列的 1/N, 把那条 12.8 min 的关键路径摊到 N 个
# runner 上 (实测 123 例 / 714s, 最慢一例 41s, 分布很平 ⇒ 分片收益几乎线性)。
#
# 按**调用次序取模**而不是切段, 两个理由:
#   1. vbp 的用例是就地写在分类块里的一条条 Test-* 调用, 没有 bas 那样的 $basQueue 可切。
#      按次序就不必把上百个调用点重排成数据 —— 改动面小一个数量级, 也就不会碰到那些
#      内联断言块里的细节 (它们只能整体包 if, 不能像函数体那样提前 return)。
#   2. 不需要知道用例总数。加一条 vbp 用例不用同步改任何常数, 也就不会出现
#      "总数写死、加了用例就静默漏跑" —— 分片漏跑 = 覆盖面静默缩小, 是最坏的一类退化。
#
# 可复现性: 一个 job 只跑一个 category, 且 vbp/GUI 始终串行 ⇒ 同一次提交下第 k 片的
# 内容每次都一样, 不会这次归 1 片那次归 2 片。
#
# **分组**: vbp 里有几处用例靠 $OutDir 里的文件互相接续 (先编出 exe, 后一条再去查那个 exe)。
# 纯取模会把这一对拆到两个 runner, 后一条就报 "no exe" —— 头一次分片就是这么红的
# (run 36423080858: vbp #1/#2/#3 挂在 frx_extract_content / ctrlmanifest_builtin /
# balloon_manifest_is_user_supplied 三处)。所以带 Group 的调用按**组首次出现时占到的槽**
# 定位, 同组永远落同一片; 原本的隐式顺序依赖就此变成代码里写明的事实。
$script:VbpCaseNo = 0
$script:VbpSkipped = 0
$script:VbpGroupSlot = @{}
function Enter-VbpShard {
    param([string]$Group = "")
    if ($VbpShardTotal -le 1) { return $true }   # 不分片: 计数器都不必动
    if ($Group -ne "" -and $script:VbpGroupSlot.ContainsKey($Group)) {
        $slot = $script:VbpGroupSlot[$Group]     # 同组复用首次的槽
    } elseif ($Group -ne "") {
        $slot = $script:VbpCaseNo++
        $script:VbpGroupSlot[$Group] = $slot
    } else {
        $slot = $script:VbpCaseNo++
    }
    $mine = $VbpShard - 1
    if (($slot % $VbpShardTotal) -eq $mine) { return $true }
    $script:VbpSkipped++
    return $false
}
if ($VbpShardTotal -gt 1) {
    if ($VbpShard -lt 1 -or $VbpShard -gt $VbpShardTotal) {
        Write-Host "[ERROR] VbpShard 需在 1..VbpShardTotal" -ForegroundColor Red
        exit 1
    }
    Write-Host "  (vbp shard ${VbpShard}/${VbpShardTotal})" -ForegroundColor Cyan
}
# ai/022 B13c: 类型库 / COM 服务器表 / 接口 vtable ä¸条通道是否同值,
# 取读数的助手单独一个文件（tests\tlb_identity.ps1）。它只用 $C3/$Tests/$OutDir,
# 这里都已经就位。
. (Join-Path $PSScriptRoot "tlb_identity.ps1")
# ai/022 B14: 同样单独一个文件 —— 这回是「真客户端」：
# tests\disp_invoke.ps1 用 tests\tools\disp_probe.c（LoadLibrary + DllGetClassObject）
# 把产出的 DLL 真的按 IDispatch 调一遍，不再只比字节。
. (Join-Path $PSScriptRoot "disp_invoke.ps1")
# ai/022 B16: 类型库的「契约面」读数（tests\tools\tlb_slots.cpp）——
# 库里那一档真接口发不发成员、槽偏移/调用约定对不对；x86 与 x64 各一条读数。
. (Join-Path $PSScriptRoot "tlb_contract.ps1")
# ai/022 B17: 外部激活 —— 真注册 (DllRegisterServer) + 走系统那条路的客户：
# tests\tools\com_act_probe.c 按 CLSID/ProgID `CoCreateInstance` 拿 IDispatch（= CreateObject
# 那条路）与接口薄指针（早绑定直调契约槽），并证明反注册后三类键都不留。
. (Join-Path $PSScriptRoot "com_activate.ps1")
# ai/022 B19: 控制台/管道的**编码**用例 —— 程序输出在 cmd 与重定向下都不许乱码
# (中文/日文/韩文/英文四语种; 覆盖 WriteConsoleW 那条与“控制台代码页”那条字节路)。
. (Join-Path $PSScriptRoot "console_enc.ps1")

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

# === 静态对表: DI 桩 census ===
# 判据与基线都在 scripts/check_di_stubs.ps1 + scripts/di_stubs_manifest.txt (只许增不许静默减)。
# 子进程跑 (脚本以 exit 收尾, 直接 & 调用会把本 runner 一起 exit 掉)。没有 pwsh 时 SKIP ——
# 本 runner 本来就可跑在 Windows PowerShell 5.1 下。
function Test-DiStubCensus {
    $script:total++
    Write-Host -NoNewline "  [STATIC] di_stubs_census ... "
    $ps7 = Get-Command pwsh -ErrorAction SilentlyContinue
    if (-not $ps7) {
        $script:skip++
        Write-Host "SKIP (无 pwsh)" -ForegroundColor Yellow
        return
    }
    $out = & $ps7.Source -NoProfile -File "$Root\scripts\check_di_stubs.ps1" 2>&1
    if ($LASTEXITCODE -eq 0) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        $out | Select-Object -Last 14 | ForEach-Object { Write-Host "  $_" }
    }
}

# === 冒烟测试 (编译+链接+运行) ===
function Test-Compile {
    param([string]$Name, [string]$Source)
    $script:total++
    Write-Host -NoNewline "  [COMPILE] $Name ... "
    
    $result = & $C3 $Source --output-dir $OutDir @IncArg 2>&1
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

# 判据匹配器 —— 进程内路径的单一出口, 按**字面**子串比 (忽略大小写, 与 -like 原本的
# 大小写行为一致, 只去掉通配语义)。
# 旧代码在 4 个点各自写 -like 通配匹配, 而 [ ] * ? 在 needle 里会被当成通配符:
# 输出本来带方括号的读数 (例如 Join(Array("a","b"), "|") 打出的 [a|b]) 永远匹配不上,
# 因为 [ab] 被解释成"取一个字符的字符集"。上一批 Fix 161b-decl-out 撞上同一件事后
# 改的是夹具 (另打一行无括号标记) 并留了一条警告注释 —— 那是让测试迁就工具,
# 下一个带方括号的读数还会再红一次。这里把语义钉死: needle 是字面子串。
# ⚠ 本函数**不能**在 -Parallel 的 runspace 里调用 (runspace 看不到脚本函数, 表现为
#   每条 needle 都判不中 => 整批 output mismatch)。那一段的比较必须内联, 见
#   Invoke-BasParallel 里的内联判定。
# 注: 脚本里另有一批 .Contains 比较点, 语义与本函数一致 (同为字面), 不动。
function Test-NeedleHit {
    param($Lines, [string]$Needle)
    foreach ($l in @($Lines)) {
        if ($null -eq $l) { continue }
        if (([string]$l).IndexOf($Needle, [StringComparison]::OrdinalIgnoreCase) -ge 0) { return $true }
    }
    return $false
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
        [string]$Name,
        [hashtable]$EnvVars = @{}   # 显式注入子进程环境 (不依赖宿主进程级变量)
    )

    # 日志统一写入 output 目录: 使用 Open ... For Output 追加写入同一日志文件
    # (scores.txt / test_output.txt / *.dat 等), 不同进程写入不同实时文件
    $stdoutFile = Join-Path $WorkDir "$Name.out"
    $stderrFile = Join-Path $WorkDir "$Name.err"
    # #44: 逐条判存在再删, 不依赖 Remove-Item 对"路径不存在"的宽容度。
    # 实证: 某些宿主(带 safe-delete 钩子的沙箱)把 Remove-Item 换成 fail-closed 版本,
    # 目标不存在时抛**终止**异常 —— 客户进程因此一次都没跑, 表现为"零输出、
    # 全部 needle 缺失", 且 detail 里看不出任何异常痕迹(最难查的一种红)。
    foreach ($stale in @($stdoutFile, $stderrFile)) {
        if (Test-Path $stale) { Remove-Item $stale -ErrorAction SilentlyContinue }
    }

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
        # P20-44: 显式注入子进程环境块 (RTL 的 GetEnvironmentVariableW 读的就是它)。
        # 比只改宿主进程级变量可靠: 子进程的环境块在这里被直接写定。
        foreach ($ek in $EnvVars.Keys) { $psi.EnvironmentVariables[$ek] = [string]$EnvVars[$ek] }

        $proc = [System.Diagnostics.Process]::Start($psi)
        $soTask = $proc.StandardOutput.ReadToEndAsync()
        $seTask = $proc.StandardError.ReadToEndAsync()
        if (-not $proc.WaitForExit($RunTimeoutSec * 1000)) {
            # 两步读数: 杀之前只能量 CPU/存活/窗口 (管道还开着, 读不到输出); 杀掉、管道关闭
            # 之后才拿得到"被杀前已经写出的那部分输出"。缺了后半步就只剩 last='' —— 看不出
            # 卡在哪一步 (是 Form_Load 就没起来, 还是 Timer 没触发, 还是 Unload 没落地)。
            $diag = Get-RunTimeoutDiag -Proc $proc
            try { $proc.Kill() } catch { }
            $proc.WaitForExit()
            $partial = Read-KilledProcOutput -OutTask $soTask -ErrTask $seTask
            # 落盘: CI 的 "Upload test logs" 收的就是这两个文件, 下次超时能直接翻到现场。
            if ($partial.Out) { [System.IO.File]::WriteAllText($stdoutFile, [string]$partial.Out, [System.Text.Encoding]::Default) }
            if ($partial.Err) { [System.IO.File]::WriteAllText($stderrFile, [string]$partial.Err, [System.Text.Encoding]::Default) }
            $tail = " last='$($partial.Last)' lines=$($partial.Lines)"
            return @{
                Ok       = $false
                ExitCode = $null
                Output   = @()
                Detail   = "run timeout: ${RunTimeoutSec}s$diag$tail"
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

# 账 #166: 产物名只有**一个权威** —— `.vbp` 里的 `ExeName32`。驱动落盘用的就是它
# (`src/driver/driver_compile.cpp:88`：优先 `ExeName32` 的 stem，缺键才回落到 vbp 文件名)，
# 而这里以前默认取**文件名** —— 两个权威管同一个名字。不一致时的症状很有欺骗性：本地手动跑
# (跑的是那个不一致的名字) 退出码与读数全对，只有门禁两条架构一起 `FAIL (no exe)`
# (门 #218 的 tabwalk 就是这么红的)。VB6 对这个键可写 "Foo" 也可写 "Foo.exe"，两种都折成 stem。
# `-ExeName` 降级成**显式覆盖**（只剩 BalloonTooltips 一处用它，因为那个工程的名字确实不同）。
function Resolve-VbpExeBase {
    param([string]$VbpFile, [string]$ExeName = "")
    if ($ExeName) { return [IO.Path]::GetFileNameWithoutExtension($ExeName) }
    $stem = [IO.Path]::GetFileNameWithoutExtension($VbpFile)
    if (Test-Path $VbpFile) {
        foreach ($ln in (Get-Content $VbpFile)) {
            if ($ln -match '^\s*ExeName32\s*=\s*(.+?)\s*$') {
                $v = $Matches[1].Trim().Trim('"').Trim()
                if ($v) { $stem = [IO.Path]::GetFileNameWithoutExtension($v) }
                break
            }
        }
    }
    return $stem
}

# GUI smoke: require a visible main window, then close only the process we launch.
# This checks startup, not screenshot correctness or QR decoding.
function Test-GuiVbp {
    param([string]$Name, [string]$VbpFile, [string]$ExeName = "", [string]$Arch = "", [int]$AutoExitSec = 0, [string]$ShardGroup = "", [int]$DumpMinColors = 0)
    # vbp 分片: 本片不跑这例 —— 闸门必须在任何副作用之前 (建目录/起进程)。
    if (-not (Enter-VbpShard -Group $ShardGroup)) { return }
    $script:total++
    Write-Host -NoNewline "  [GUI] $Name ... "
    $guiOut = Join-Path $OutDir $Name
    New-Item -ItemType Directory -Path $guiOut -Force | Out-Null
    if ($Arch) {
        $compileResult = & $C3 $VbpFile --arch $Arch --output-dir $guiOut @IncArg 2>&1
    } else {
        $compileResult = & $C3 $VbpFile --output-dir $guiOut @IncArg 2>&1
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
    # 账 #166: 名字走单一权威 `Resolve-VbpExeBase`（读 .vbp 的 ExeName32），不再各算各的。
    $exeBase = Resolve-VbpExeBase -VbpFile $VbpFile -ExeName $ExeName
    $exe = Join-Path $guiOut ($exeBase + ".exe")
    $proc = $null
    try {
        # 账 #175: 子进程继承本进程环境 ⇒ 用前先设、finally 里清, 不漏给后面的用例
        $dumpDir = ""
        if ($DumpMinColors -gt 0) {
            $dumpDir = Join-Path $guiOut "ucdump"
            New-Item -ItemType Directory -Force -Path $dumpDir | Out-Null
            $env:C3_UC_DUMPDIR = $dumpDir
        }
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
                if ($DumpMinColors -gt 0) { Assert-UcDumpPainted -Dir $dumpDir -MinColors $DumpMinColors }
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
        if ($DumpMinColors -gt 0) { Remove-Item Env:\C3_UC_DUMPDIR -ErrorAction SilentlyContinue }
        if ($proc -and -not $proc.HasExited) { $proc.Kill(); $proc.WaitForExit() }
    }
}

# 账 #175: 读 C3_UC_DUMPDIR 落下的每控件绘制缓冲, 数**不同颜色**个数。
# 为什么用这个读数: 图体空白在「窗口出来了」「退出码 0」两条判据下都是绿的 —— 只有
# 缓冲本身能说话。BASE (改前) 实测个位数, NEW 实测 >=63, 阈值取 40 留两侧余量。
function Assert-UcDumpPainted {
    param([string]$Dir, [int]$MinColors, [int]$MinFiles = 3)
    $files = @(Get-ChildItem $Dir -Filter *.bmp -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike 'FORM_*' })
    if ($files.Count -lt $MinFiles) { throw "ucdump: only $($files.Count) bmp (< $MinFiles) in $Dir" }
    foreach ($f in $files) {
        $b = [IO.File]::ReadAllBytes($f.FullName)
        if ($b.Length -lt 54) { throw ("ucdump: " + $f.Name + " too short") }
        $off = [BitConverter]::ToInt32($b, 10)
        $w = [BitConverter]::ToInt32($b, 18)
        $h = [BitConverter]::ToInt32($b, 22)
        $bpp = [BitConverter]::ToInt16($b, 28)
        $bytes = $bpp / 8
        $stride = ((($w * $bpp) + 31) / 32) * 4
        $set = @{}
        for ($y = 0; $y -lt $h; $y += 2) {
            $row = $off + $y * $stride
            for ($x = 0; $x -lt $w; $x += 2) {
                $k = $row + $x * $bytes
                if ($k + 2 -lt $b.Length) { $set[[int]$b[$k] * 65536 + [int]$b[$k+1] * 256 + [int]$b[$k+2]] = 1 }
            }
        }
        if ($set.Count -lt $MinColors) {
            throw ("ucdump blank: " + $f.Name + " " + $w + "x" + $h + " distinct=" + $set.Count + " (< " + $MinColors + ")")
        }
    }
}

function Test-UcScaleUnitsCensus {
    $script:total++
    Write-Host -NoNewline "  [STATIC] uc_scale_units_census ... "
    $out = & powershell -NoProfile -ExecutionPolicy Bypass -File "$Root\scripts\check_uc_scale_units.ps1" 2>&1
    if ($LASTEXITCODE -eq 0) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        $out | Select-Object -Last 14 | ForEach-Object { Write-Host "  $_" }
    }
}

# 账 #159: 宿主伪成员只有一张表 (名字 / 类型 / 可否裸写 / 赋值拆数值)。哨兵逐行把
# 表与 src/rtl/core/vb6rtl/vb6rtl_userctl.h 的 extern 声明对照, 并禁止那五份旧清单
# 回潮 (成员名字面量再出现在消费点 = 抄了第二份)。
function Test-HostPseudoTableCensus {
    $script:total++
    Write-Host -NoNewline "  [STATIC] host_pseudo_table_census ... "
    $out = & powershell -NoProfile -ExecutionPolicy Bypass -File "$Root\scripts\check_host_pseudo_table.ps1" 2>&1
    if ($LASTEXITCODE -eq 0) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        $out | Select-Object -Last 14 | ForEach-Object { Write-Host "  $_" }
    }
}

# 账 #184: AddressOf 交出去的必须是桩地址。口径收在 CCodeGen::addressOfTargetCName 一处,
# 这条哨兵拦"又加一个把 (void*) 直接拼 cProcName 的取址出口" —— 那种写法 x64 毫无症状,
# 只有 x86 运行期才炸, 所以必须静态钉住 (负控: 同一份脚本在 HEAD~1 的树上四条规则全红)。
# 账 #185: 子类化分层的槽位哨兵 (一个属性名 = 一层, 谁都不许跟谁抢)。
# 这一类缺陷的症状是「装了但没装」—— 编得过、发码对、事件一次也不响, 只有运行期看得见,
# 所以静态钉住。同一份脚本在 HEAD 的树上 S1~S4 一起红 (实测 6 行)。
# 账 #188: 浮点字面量的形状只有一个出口 (src/common/float_literal.hpp)。
# 手拼 "f" 后缀这件事在 x86/x64 都不会有运行期症状 —— 它是**编不过**, 而带小数点的设计值 (8.25)
# 恰好躲过, 所以必须静态钉。同一份脚本在 HEAD 的树上 R1~R3 全红 (R1 当场点出那 6 个手拼点)。
function Test-IntSuffixSites {
    $script:total++
    Write-Host -NoNewline "  [STATIC] int_suffix_sites ... "
    $out = & powershell -NoProfile -ExecutionPolicy Bypass -File "$Root\scripts\check_int_suffix_sites.ps1" 2>&1
    if ($LASTEXITCODE -eq 0) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        $out | Select-Object -First 6 | ForEach-Object { Write-Host "  $_" }
    }
}
function Test-ScaleModeWriters {
    $script:total++
    Write-Host -NoNewline "  [STATIC] scalemode_writers ... "
    $out = & powershell -NoProfile -ExecutionPolicy Bypass -File "$Root\scripts\check_scalemode_writers.ps1" 2>&1
    if ($LASTEXITCODE -eq 0) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        $out | Select-Object -First 6 | ForEach-Object { Write-Host "  $_" }
    }
}
function Test-ControlDC {
    $script:total++
    Write-Host -NoNewline "  [STATIC] control_dc ... "
    $out = & powershell -NoProfile -ExecutionPolicy Bypass -File "$Root\scripts\check_control_dc.ps1" 2>&1
    if ($LASTEXITCODE -eq 0) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        $out | Select-Object -First 6 | ForEach-Object { Write-Host "  $_" }
    }
}
function Test-SaAccess {
    $script:total++
    Write-Host -NoNewline "  [STATIC] sa_access ... "
    $out = & powershell -NoProfile -ExecutionPolicy Bypass -File "$Root\scripts\check_sa_access.ps1" 2>&1
    if ($LASTEXITCODE -eq 0) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        $out | Select-Object -First 6 | ForEach-Object { Write-Host "  $_" }
    }
}
function Test-BitwiseAuthority {
    $script:total++
    Write-Host -NoNewline "  [STATIC] bitwise_authority ... "
    $out = & powershell -NoProfile -ExecutionPolicy Bypass -File "$Root\scripts\check_bitwise_authority.ps1" 2>&1
    if ($LASTEXITCODE -eq 0) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        $out | Select-Object -First 6 | ForEach-Object { Write-Host "  $_" }
    }
}
function Test-BodyDeclShape {
    $script:total++
    Write-Host -NoNewline "  [STATIC] bodydecl_shape ... "
    $out = & powershell -NoProfile -ExecutionPolicy Bypass -File "$Root\scripts\check_bodydecl_shape.ps1" 2>&1
    if ($LASTEXITCODE -eq 0) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        $out | Select-Object -First 6 | ForEach-Object { Write-Host "  $_" }
    }
}
function Test-DocHostAuthority {
    $script:total++
    Write-Host -NoNewline "  [STATIC] dochost_authority ... "
    $out = & powershell -NoProfile -ExecutionPolicy Bypass -File "$Root\scripts\check_dochost_authority.ps1" 2>&1
    if ($LASTEXITCODE -eq 0) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        $out | Select-Object -First 6 | ForEach-Object { Write-Host "  $_" }
    }
}
function Test-CaseIsShape {
    $script:total++
    Write-Host -NoNewline "  [STATIC] caseis_shape ... "
    $out = & powershell -NoProfile -ExecutionPolicy Bypass -File "$Root\scripts\check_caseis_shape.ps1" 2>&1
    if ($LASTEXITCODE -eq 0) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        $out | Select-Object -First 6 | ForEach-Object { Write-Host "  $_" }
    }
}
function Test-RtlNakedNames {
    $script:total++
    Write-Host -NoNewline "  [STATIC] rtl_naked_names ... "
    $out = & powershell -NoProfile -ExecutionPolicy Bypass -File "$Root\scripts\check_rtl_naked_names.ps1" 2>&1
    if ($LASTEXITCODE -eq 0) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        $out | Select-Object -First 6 | ForEach-Object { Write-Host "  $_" }
    }
}
function Test-RtlResourceIds {
    $script:total++
    Write-Host -NoNewline "  [STATIC] rtl_resource_ids ... "
    $out = & powershell -NoProfile -ExecutionPolicy Bypass -File "$Root\scripts\check_rtl_resource_ids.ps1" 2>&1
    if ($LASTEXITCODE -eq 0) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        $out | Select-Object -First 6 | ForEach-Object { Write-Host "  $_" }
    }
}
function Test-UcArrayEventSites {
    $script:total++
    Write-Host -NoNewline "  [STATIC] uc_array_event_sites ... "
    $out = & powershell -NoProfile -ExecutionPolicy Bypass -File "$Root\scripts\check_uc_array_event_sites.ps1" 2>&1
    if ($LASTEXITCODE -eq 0) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        $out | Select-Object -First 8 | ForEach-Object { Write-Host "  $_" }
    }
}
function Test-IntLiteralShape {
    $script:total++
    Write-Host -NoNewline "  [STATIC] int_literal_shape ... "
    $out = & powershell -NoProfile -ExecutionPolicy Bypass -File "$Root\scripts\check_int_literal_shape.ps1" 2>&1
    if ($LASTEXITCODE -eq 0) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        $out | Select-Object -First 6 | ForEach-Object { Write-Host "  $_" }
    }
}
function Test-FloatLiteralShape {
    $script:total++
    Write-Host -NoNewline "  [STATIC] float_literal_shape ... "
    $out = & powershell -NoProfile -ExecutionPolicy Bypass -File "$Root\scripts\check_float_literal_shape.ps1" 2>&1
    if ($LASTEXITCODE -eq 0) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        $out | Select-Object -First 6 | ForEach-Object { Write-Host "  $_" }
    }
}


function Test-SubclassSlotSites {
    $script:total++
    Write-Host -NoNewline "  [STATIC] subclass_slot_sites ... "
    $out = & powershell -NoProfile -ExecutionPolicy Bypass -File "$Root\scripts\check_subclass_slots.ps1" 2>&1
    if ($LASTEXITCODE -eq 0) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        $out | Select-Object -Last 14 | ForEach-Object { Write-Host "  $_" }
    }
}

function Test-CtrlArrayMemberSites {
    $script:total++
    Write-Host -NoNewline "  [STATIC] ctrl_array_members ... "
    $out = & powershell -NoProfile -ExecutionPolicy Bypass -File "$Root\scripts\check_ctrl_array_members.ps1" 2>&1
    if ($LASTEXITCODE -eq 0) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        $out | Select-Object -First 6 | ForEach-Object { Write-Host "  $_" }
    }
}

function Test-CtrlPropTypeAuthority {
    $script:total++
    Write-Host -NoNewline "  [STATIC] ctrl_prop_type_authority ... "
    $out = & powershell -NoProfile -ExecutionPolicy Bypass -File "$Root\scripts\check_ctrl_prop_type_authority.ps1" 2>&1
    if ($LASTEXITCODE -eq 0) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        $out | Select-Object -First 6 | ForEach-Object { Write-Host "  $_" }
    }
}

function Test-FormDrawState {
    $script:total++
    Write-Host -NoNewline "  [STATIC] form_draw_state ... "
    $out = & powershell -NoProfile -ExecutionPolicy Bypass -File "$Root\scripts\check_form_draw_state.ps1" 2>&1
    if ($LASTEXITCODE -eq 0) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        $out | Select-Object -First 6 | ForEach-Object { Write-Host "  $_" }
    }
}
function Test-VariantCmpBoxing {
    $script:total++
    Write-Host -NoNewline "  [STATIC] variant_cmp_boxing ... "
    $out = & powershell -NoProfile -ExecutionPolicy Bypass -File "$Root\scripts\check_variant_cmp_boxing.ps1" 2>&1
    if ($LASTEXITCODE -eq 0) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        $out | Select-Object -First 6 | ForEach-Object { Write-Host "  $_" }
    }
}
function Test-FixtureTimerClose {
    $script:total++
    Write-Host -NoNewline "  [STATIC] fixture_timer_close ... "
    $out = & powershell -NoProfile -ExecutionPolicy Bypass -File "$Root\scripts\check_fixture_timer_close.ps1" 2>&1
    if ($LASTEXITCODE -eq 0) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        $out | Select-Object -First 6 | ForEach-Object { Write-Host "  $_" }
    }
}


function Test-EventHandlerNames {
    $script:total++
    Write-Host -NoNewline "  [STATIC] event_handler_names ... "
    $out = & powershell -NoProfile -ExecutionPolicy Bypass -File "$Root\scripts\check_event_handler_names.ps1" 2>&1
    if ($LASTEXITCODE -eq 0) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        $out | Select-Object -First 6 | ForEach-Object { Write-Host "  $_" }
    }
}

function Test-UcInstanceExit {
    $script:total++
    Write-Host -NoNewline "  [STATIC] uc_instance_exit ... "
    $out = & powershell -NoProfile -ExecutionPolicy Bypass -File "$Root\scripts\check_uc_instance_exit.ps1" 2>&1
    if ($LASTEXITCODE -eq 0) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        $out | Select-Object -First 6 | ForEach-Object { Write-Host "  $_" }
    }
}


function Test-AddressOfThunkSites {
    $script:total++
    Write-Host -NoNewline "  [STATIC] addrof_thunk_sites ... "
    $out = & powershell -NoProfile -ExecutionPolicy Bypass -File "$Root\scripts\check_addressof_thunk_sites.ps1" 2>&1
    if ($LASTEXITCODE -eq 0) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        $out | Select-Object -Last 14 | ForEach-Object { Write-Host "  $_" }
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
        $compileResult = & $C3 $Source --arch $Arch --output-dir $OutDir @IncArg 2>&1
    } else {
        $compileResult = & $C3 $Source --output-dir $OutDir @IncArg 2>&1
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
    # 可选环境变量 (P20-44: frmevents 的 OLE 无头联测要 C3_OLEDDB_TEST=1 才驱动)
    $savedEnv = @{}
    $envMap = @{}
    if ($Env) {
        foreach ($kv in $Env.Split(';')) {
            if (-not $kv) { continue }
            $pp = $kv.Split('=', 2)
            $envMap[$pp[0]] = $pp[1]
            $savedEnv[$pp[0]] = [Environment]::GetEnvironmentVariable($pp[0])
            [Environment]::SetEnvironmentVariable($pp[0], $pp[1])   # 兜底
        }
    }
    $run = Invoke-TestExe -ExePath $exePath -WorkDir $OutDir -Name $baseName -EnvVars $envMap
    foreach ($k in $savedEnv.Keys) {
        [Environment]::SetEnvironmentVariable($k, $savedEnv[$k])
    }
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
            $found = Test-NeedleHit -Lines $runOutput -Needle $expected
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

    # -Parallel 的 runspace 里调不到脚本函数 ⇒ 超时预算用 $using: 传, 读数逻辑内联
    $runTimeoutMs = $RunTimeoutSec * 1000
    $results = $shards | ForEach-Object -Parallel {
        $shardItems = $_[0]
        $workDir    = $_[1]
        New-Item -ItemType Directory -Path $workDir -Force | Out-Null
        $c3 = $using:C3
        $runTimeoutMs = $using:runTimeoutMs
        $incArg = $using:IncArg   # ai/030 T30-B: runspace 里够不到脚本变量
        $p = 0; $f = 0; $details = @()
        foreach ($it in $shardItems) {
            if ($it.Arch) {
                $cr = & $c3 $it.Source --arch $it.Arch --output-dir $workDir @IncArg 2>&1
            } else {
                $cr = & $c3 $it.Source --output-dir $workDir @IncArg 2>&1
            }
            $ec = $LASTEXITCODE
            if ($ec -ne 0) {
                $f++
                # 可观测性约定 (同下面的 output mismatch): 编译失败必须带实际错误行。
                # 以前只报 "compile FAIL" —— 编译输出既不落盘也不在 CI 工件里 (工件只收
                # output/**/*.out 与 *.err, 那两份是**运行**期的 stdout/stderr), 于是
                # GA 上一条编译红根本无从判断是哪一条 C 错误 (实测 test_variant_boxing
                # 在 CI 上 compile FAIL、本地同一条命令干净, 来回三轮都拿不到信息)。
                $ce = @($cr | ForEach-Object { [string]$_ } | Where-Object { $_ -match 'error C\d+' } | Select-Object -First 2)
                if ($ce.Count -gt 0) {
                    $details += "$($it.Name): compile FAIL :: " + ($ce -join " | ")
                } else {
                    $tail = @($cr | ForEach-Object { [string]$_ } | Select-Object -Last 2)
                    $details += "$($it.Name): compile FAIL (exit $ec, 无 error C 行) :: " + ($tail -join " / ")
                }
                continue
            }

            $baseName = [IO.Path]::GetFileNameWithoutExtension($it.Source)
            $exePath = Join-Path $workDir "$baseName.exe"
            if (-not (Test-Path $exePath)) { $f++; $details += "$($it.Name): no exe"; continue }

            # --- 运行 (语义与 Invoke-TestExe 一致; 见 -RunTimeoutSec) ---
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
                    # 读数: CPU 时间 + 进程状态 + 顶层窗口 + 最后一行输出 (CPU≈0 且无输出 = 没被调度, 不是挂)
                    $cpu = -1; $st = '?'; $win = ''
                    try { $cpu = [int]$proc.TotalProcessorTime.TotalMilliseconds } catch { }
                    try { if ($proc.HasExited) { $st = "exited=$($proc.ExitCode)" } else { $st = 'alive' } } catch { }
                    try { $proc.Refresh(); $win = [string]$proc.MainWindowTitle } catch { }
                    if ($win) { $win = ", win='$win'" }
                    # ⚠ 顺序不能倒: 必须 Kill + WaitForExit **之后**管道才关闭, ReadToEndAsync
                    #   才完成 ⇒ 这时才拿得到被杀前已写出的输出 (杀之前读只会拿到空串)。
                    try { $proc.Kill() } catch { }
                    $proc.WaitForExit()
                    $last = ''
                    foreach ($t in @($soTask, $seTask)) {
                        if ($t -and $t.Wait(1500)) {
                            $txt = ''
                            try { $txt = [string]$t.Result } catch { }
                            if ($txt) { $line = @($txt.TrimEnd() -split "`r?`n" | Where-Object { $_ }); if ($line.Count -gt 0) { $last = $line[-1] } }
                        }
                    }
                    $f++; $details += "$($it.Name): run timeout (cpu=${cpu}ms, ${st}$win, last='$last')"; continue
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
                    # ⚠ 这段跑在 -Parallel 的 runspace 里, 调不到脚本函数 (见上面
                    # Invoke-BasParallel 里那条同规则注释) ⇒ 字面判定必须内联写在这里。
                    $found = $false
                    foreach ($l in @($runOut)) {
                        if ($null -ne $l -and ([string]$l).IndexOf($exp, [StringComparison]::OrdinalIgnoreCase) -ge 0) { $found = $true; break }
                    }
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
# ai/029 C29-M: 断**产物里的应用清单**（不重建，拿 Test-Vbp 刚产出的那份 exe）。
# 为什么在产物面上断：清单的作用全在运行期 —— 少了那条 #1 RT_MANIFEST，SxS 就把 comctl32
# 解析成 System32 的 5.82，本线所有原生控件换成 v5 的类表与消息语义，而编译/链接/退出码
# 全都好看（这条洞就是 C29-M 修的那件事）。
# 断两样：`</assembly>` 的**个数**（两份 #1 会让加载器直接报错，所以"恰好一份"是硬要求，
# 也是"内置那份有没有叠到用户那份上"的读数）+ 内容里的特征串（用户那份带 dpiAware，
# 内置那份带 Microsoft.Windows.Common-Controls 但没有 dpiAware）。
function Test-ProductManifest {
    param(
        [string]$Name,
        [string]$ExeFile,
        [int]$ManifestCount,
        [string]$MustContain = "",
        [string]$MustNotContain = "",
        [string]$ShardGroup = ""   # 见 Enter-VbpShard: 与别的用例靠产物接续时同组同片
    )
    # vbp 分片: 本片不跑这例 —— 闸门必须在任何副作用之前 (建目录/起进程)。
    if (-not (Enter-VbpShard -Group $ShardGroup)) { return }
    $script:total++
    Write-Host -NoNewline "  [MANIFEST] $Name ... "
    if (-not (Test-Path $ExeFile)) {
        $script:fail++
        Write-Host "FAIL (no exe: $ExeFile)" -ForegroundColor Red
        return
    }
    $text = [System.Text.Encoding]::ASCII.GetString([System.IO.File]::ReadAllBytes($ExeFile))
    $got = ([regex]::Matches($text, '</assembly>')).Count
    $problems = @()
    if ($got -ne $ManifestCount) { $problems += "清单数=$got 期望=$ManifestCount" }
    if ($MustContain -and $text.IndexOf($MustContain) -lt 0) { $problems += "缺串 '$MustContain'" }
    if ($MustNotContain -and $text.IndexOf($MustNotContain) -ge 0) { $problems += "不该有串 '$MustNotContain'" }
    if ($problems.Count -gt 0) {
        $script:fail++
        $msg = "FAIL (" + ($problems -join '; ') + ")"
        Write-Host $msg -ForegroundColor Red
        return
    }
    $script:pass++
    Write-Host "OK ($ManifestCount 份清单)" -ForegroundColor Green
}

function Test-Vbp {
    param(
        [string]$Name,
        [string]$VbpFile,
        [string[]]$ExpectedOutputs,
        [string]$Arch = "",           # 可选架构参数 (x86/x64)
        [string]$RequiresCom = "",    # 依赖的 COM ProgId (未注册则 SKIP, 否则 FAIL)
        [string]$Env = "",            # 可选环境变量, "K1=V1;K2=V2" (跑 exe 前设, 跑完还原)
        [string]$ShardGroup = ""      # vbp 分片组名 (见 Enter-VbpShard): 与别的用例靠 $OutDir 里的
                                      # 文件接续时, 同组用例必须落同一片, 否则后一条报 "no exe"
    )
    # vbp 分片: 本片不跑这例 —— 闸门必须在任何副作用之前 (建目录/起进程)。
    if (-not (Enter-VbpShard -Group $ShardGroup)) { return }
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
        $compileResult = & $C3 $VbpFile --arch $Arch --output-dir $OutDir @IncArg 2>&1
    } else {
        $compileResult = & $C3 $VbpFile --output-dir $OutDir @IncArg 2>&1
    }
    if ($LASTEXITCODE -ne 0) {
        $script:fail++
        Write-Host "FAIL (compile)" -ForegroundColor Red
        if ($Verbose) { Write-Host ($compileResult | Out-String) }
        return
    }

    # 处理 VBP 编译输出: 校验是否包含 expected 输出 (可含多个子串)
    # 账 #166: 名字走单一权威 `Resolve-VbpExeBase`（读 .vbp 的 ExeName32）。以前这里只认**文件名**，
    # 连 `-ExeName` 的口子都没有 ⇒ 凡 ExeName32 与文件名不同的工程，本地看不出、门禁必红。
    $baseName = Resolve-VbpExeBase -VbpFile $VbpFile
    $exePath = Join-Path $OutDir "$baseName.exe"
    if (-not (Test-Path $exePath)) {
        $script:fail++
        # 哨兵：把"按什么名字找的 / .vbp 里声明的是什么"一起打出来，下次这类红一句话就能定位
        $declared = '(no ExeName32 key)'
        if (Test-Path $VbpFile) {
            foreach ($ln0 in (Get-Content $VbpFile)) {
                if ($ln0 -match '^\s*ExeName32\s*=\s*(.+?)\s*$') { $declared = $Matches[1].Trim(); break }
            }
        }
        Write-Host "FAIL (no exe)" -ForegroundColor Red
        Write-Host ("    找的是 " + $exePath + " ；vbp 文件名=" + [IO.Path]::GetFileNameWithoutExtension($VbpFile) + " ；声明的 ExeName32=" + $declared) -ForegroundColor Red
        return
    }

    # 免注册便携部署: 把工程目录下的原生依赖 (OCX/DLL/TLB) 复制到 exe 同目录 ($OutDir),
    # 与 Test-GuiVbp 行为一致 (否则依赖免注册 OCX 的用例在 CI 上拿不到控件文件 → LoadLibrary 失败)。
    # Fix <vbeclipse> 2026-10-06: axcontrolsadd 的运行时 Me.Controls.Add 走 Object= 免注册表,
    # 需要 NewTab01.ocx 在 exe 旁才能 regfree 加载。
    $srcDir = Split-Path $VbpFile -Parent
    if (Test-Path $srcDir) {
        Get-ChildItem $srcDir -File | Where-Object { $_.Extension -match '\.(ocx|dll|tlb)$' } | ForEach-Object {
            Copy-Item -Force $_.FullName $OutDir | Out-Null
        }
    }

    # 可选环境变量 "K1=V1;K2=V2" → 显式注入子进程环境块 (RTL 用 GetEnvironmentVariableW 读它)。
    # ⚠ Test-Vbp 曾经声明了 $Env 却忘了往下传 ⇒ 调用点写的 -Env 被静默吞掉: frmevents 的
    # OLE 无头联测 (C3_OLEDDB_TEST=1) 在 CI 上从不启用, 症状是 EV24/EV25 缺失 → FAIL
    # (output mismatch)。这是脚本 bug, 不是编译器回归 —— 同一形态的判断先查这里。
    $envMap = @{}
    if ($Env) {
        foreach ($kv in $Env.Split(';')) {
            if (-not $kv) { continue }
            $pp = $kv.Split('=', 2)
            if ($pp.Count -eq 2) { $envMap[$pp[0]] = $pp[1] }
        }
    }
    # 编译失败则标记 FAIL (compile); 生成物缺失标记 FAIL (no exe)
    $run = Invoke-TestExe -ExePath $exePath -WorkDir $OutDir -Name $baseName -EnvVars $envMap
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
            $found = Test-NeedleHit -Lines $runOutput -Needle $expected
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
# Negative syntax case: --syntax-only must FAIL and report the given (ASCII) text.
# Used by the tB-extension contract diagnostics (ai/022 B01+).
# Multi-source --syntax-only probes (ai/022 B07): C3.exe accepts several positional files
# and driver_frontend picks the module kind per extension, so class-Inherits cases that need
# the base class in a *second* module are testable without a full vbp build.
function Invoke-SyntaxProj {
    param([array]$Sources)
    $argList = (($Sources | ForEach-Object { '"' + $_ + '"' }) -join ' ')
    $out = & cmd /c ('"' + $C3 + '" ' + $argList + ' --syntax-only 2>&1')
    $script:syntaxProjExit = $LASTEXITCODE
    return (($out | Out-String) -replace '\s+', ' ')
}

function Test-SyntaxFailMulti {
    param([string]$Name, [array]$Sources, [string]$Needle)
    $script:total++
    Write-Host -NoNewline "  [SYNTAX-FAIL] $Name ... "
    $text = Invoke-SyntaxProj $Sources
    if ($script:syntaxProjExit -ne 0 -and $text.Contains($Needle)) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        Write-Host "  expected failing compile containing: $Needle" -ForegroundColor DarkGray
        if ($Verbose) { Write-Host $text }
    }
}

# ai/022 B11/C04: --syntax-only must SUCCEED and the merged output must carry every needle
# (and none of the -Absent ones). Folding legacy header attributes is a read-only information
# channel, so "what got folded" and "what deliberately did NOT" is asserted here, not by
# an exit code (D52).
function Test-SyntaxNote {
    param([string]$Name, [array]$Sources, [array]$Needles, [array]$Absent = @())
    $script:total++
    Write-Host -NoNewline "  [SYNTAX-NOTE] $Name ... "
    $text = Invoke-SyntaxProj $Sources
    $bad = @($Needles | Where-Object { -not $text.Contains($_) })
    $hit = @($Absent | Where-Object { $text.Contains($_) })
    if ($script:syntaxProjExit -eq 0 -and $bad.Count -eq 0 -and $hit.Count -eq 0) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        Write-Host ("  exit=" + $script:syntaxProjExit + " missing: " + ($bad -join ' | ') +
                    " unexpected: " + ($hit -join ' | ')) -ForegroundColor DarkGray
        if ($Verbose) { Write-Host $text }
    }
}

function Test-SyntaxMulti {
    param([string]$Name, [array]$Sources)
    $script:total++
    Write-Host -NoNewline "  [SYNTAX] $Name ... "
    $text = Invoke-SyntaxProj $Sources
    if ($script:syntaxProjExit -eq 0) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        if ($Verbose) { Write-Host $text }
    }
}

# Codegen-stage probes (ai/022 B08e-6): diagnostics raised while C is being emitted never
# reach --syntax-only (driver_compile.cpp returns before codegen), and a full build would
# pay for cl.exe + link. `--emit-c` runs the whole front end plus codegen and stops there.
function Invoke-CodegenProj {
    param([array]$Sources)
    $argList = (($Sources | ForEach-Object { '"' + $_ + '"' }) -join ' ')
    $out = & cmd /c ('"' + $C3 + '" ' + $argList + ' --emit-c 2>&1')
    $script:codegenProjExit = $LASTEXITCODE
    return (($out | Out-String) -replace '\s+', ' ')
}

function Test-CompileFail {
    param([string]$Name, [array]$Sources, [string]$Needle)
    $script:total++
    Write-Host -NoNewline "  [COMPILE-FAIL] $Name ... "
    $text = Invoke-CodegenProj $Sources
    if ($script:codegenProjExit -ne 0 -and $text.Contains($Needle)) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        Write-Host "  expected failing codegen containing: $Needle" -ForegroundColor DarkGray
        if ($Verbose) { Write-Host $text }
    }
}

# ai/028 V2: --emit-c 通路上断言"必须出现的读数" (退出码 0 + 每条 needle 都在)。
# 语义层的检查 (未声明标识符 VB3001 一族) 在 --syntax-only 上根本看不见 —— 那条通路停在
# parse 之后 —— 所以"孔里就是普通表达式"这条判据只能走 codegen 通路量。
function Test-CodegenNote {
    param([string]$Name, [array]$Sources, [array]$Needles, [array]$Absent = @())
    $script:total++
    Write-Host -NoNewline "  [CODEGEN-NOTE] $Name ... "
    $text = Invoke-CodegenProj $Sources
    $bad = @($Needles | Where-Object { -not $text.Contains($_) })
    $hit = @($Absent | Where-Object { $text.Contains($_) })
    if ($script:codegenProjExit -eq 0 -and $bad.Count -eq 0 -and $hit.Count -eq 0) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host ("  exit=" + $script:codegenProjExit + " missing: " + ($bad -join ' | ') +
                    " unexpected: " + ($hit -join ' | ')) -ForegroundColor DarkGray
        if ($Verbose) { Write-Host $text }
    }
}

function Test-Compile {
    param([string]$Name, [array]$Sources)
    $script:total++
    Write-Host -NoNewline "  [COMPILE] $Name ... "
    $text = Invoke-CodegenProj $Sources
    if ($script:codegenProjExit -eq 0) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        if ($Verbose) { Write-Host $text }
    }
}

# ai/028 V1: --emit-c 的字面量形状断言 (故意不折叠空白 —— 行结构本身就是读数)。
# 多行串必须在词法出口折成「一行 C 字面量 + \r\n 转义」; 真换行若漏进 C 源码就会把一枚
# 字面量劈成两行, 于是「带字面量的行里未转义的双引号必须成对」就是这条判据的不变式。
function Test-EmitcShape {
    param([string]$Name, [array]$Sources, [array]$Needles)
    # vbp 分片: 本片不跑这例 —— 闸门必须在任何副作用之前 (建目录/起进程)。
    if (-not (Enter-VbpShard)) { return }
    $script:total++
    Write-Host -NoNewline "  [EMITC-SHAPE] $Name ... "
    $argList = (($Sources | ForEach-Object { '"' + $_ + '"' }) -join ' ')
    $out = & cmd /c ('"' + $C3 + '" ' + $argList + ' --emit-c 2>&1')
    $code = $LASTEXITCODE
    $raw = ($out | Out-String)
    $bad = @($Needles | Where-Object { -not $raw.Contains($_) })
    $odd = 0
    foreach ($ln in ($raw -split "`r?`n")) {
        if ($ln.Contains('vb6_BSTR_FromStr(')) {
            if (([regex]::Matches($ln, '(?<!\\)"')).Count % 2 -ne 0) { $odd++ }
        }
    }
    if ($code -eq 0 -and $bad.Count -eq 0 -and $odd -eq 0) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host ("  exit=" + $code + " missing: " + ($bad -join ' | ') +
                    " odd-literal-lines=" + $odd) -ForegroundColor DarkGray
        if ($Verbose) { Write-Host $raw }
    }
}

# Test-EmitcShape 的反面：断 --emit-c 的输出里**没有**某些形状。用在 D6 那一类改动上
# （"这一类控件不再走 OCX 晚绑定"）—— 只断"原生入口在"不够，残留的 OCX 形状会让两条路
# 并存，读数目视上全绿、发码却还在 CoCreateInstance。
function Test-EmitcAbsent {
    param([string]$Name, [array]$Sources, [array]$Needles)
    # vbp 分片: 本片不跑这例 —— 闸门必须在任何副作用之前 (建目录/起进程)。
    if (-not (Enter-VbpShard)) { return }
    $script:total++
    Write-Host -NoNewline "  [EMITC-ABSENT] $Name ... "
    $argList = (($Sources | ForEach-Object { '"' + $_ + '"' }) -join ' ')
    $out = & cmd /c ('"' + $C3 + '"' + ' ' + $argList + ' --emit-c 2>&1')
    $code = $LASTEXITCODE
    $raw = ($out | Out-String)
    $hit = @($Needles | Where-Object { $raw.Contains($_) })
    if ($code -eq 0 -and $hit.Count -eq 0) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host ("  exit=" + $code + " 还在场: " + ($hit -join ' | ')) -ForegroundColor DarkGray
        if ($Verbose) { Write-Host $raw }
    }
}

function Test-SyntaxFail {
    param([string]$Name, [string]$Source, [string]$Needle)
    $script:total++
    Write-Host -NoNewline "  [SYNTAX-FAIL] $Name ... "
    # Native stderr under SilentlyContinue is dropped by `& 2>&1`; let cmd.exe merge the
    # streams, and collapse whitespace so long diagnostics cannot be word-wrapped apart.
    $result = & cmd /c ('"' + $C3 + '" "' + $Source + '" --syntax-only 2>&1')
    $text = (($result | Out-String) -replace '\s+', ' ')
    if ($LASTEXITCODE -ne 0 -and $text.Contains($Needle)) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        Write-Host "  expected failing compile containing: $Needle" -ForegroundColor DarkGray
        if ($Verbose) { Write-Host $text }
    }
}

# ai/022 B11/C02: CoClass identity assertions. The resolver prints one info line per block on
# stderr -- note-level *diagnostics* cannot be the surface here, because Diagnostics::toString()
# is only dumped when a stage fails, so a note never reaches a successful compile (ai/022 D46).
function Invoke-IdentityText {
    param([string]$Proj)
    $out = & cmd /c ('"' + $C3 + '" "' + $Proj + '" --syntax-only 2>&1')
    $script:identityExit = $LASTEXITCODE
    return (((@($out | Where-Object { "$_" -match 'identity:' })) | Out-String) -replace '\s+', ' ')
}

function Test-IdentityNote {
    param([string]$Name, [string]$Proj, [array]$Needles)
    $script:total++
    Write-Host -NoNewline "  [IDENTITY] $Name ... "
    $text = Invoke-IdentityText $Proj
    $bad = @($Needles | Where-Object { -not $text.Contains($_) })
    if ($script:identityExit -eq 0 -and $bad.Count -eq 0) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        Write-Host ("  exit=" + $script:identityExit + " missing: " + ($bad -join ' | ')) -ForegroundColor DarkGray
        if ($Verbose) { Write-Host $text }
    }
}

function Test-IdentityStable {
    param([string]$Name, [string]$Proj)
    $script:total++
    Write-Host -NoNewline "  [IDENTITY] $Name ... "
    $a = Invoke-IdentityText $Proj
    $b = Invoke-IdentityText $Proj
    if ($a.Length -gt 0 -and $a -ceq $b) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        Write-Host "  two runs of the same project differ (or printed nothing)" -ForegroundColor DarkGray
    }
}

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

# ai/023 S01: vbp-level negative case. The package hard checks run in driver stage 0
# (before the pipeline), so --syntax-only is enough: compile must FAIL and the output
# must contain the given ASCII needle.
# ai/030 T30-A: 内容寻址 obj store 的判据。用例跑在自己的 store 里 (C3 的缓存根取自
# %LOCALAPPDATA%，拿不到才退回 <outputDir>/.c3obj)，于是这几条能精确断、不受机器上历史
# 缓存摆布：空 store 首编必全 miss、次编必全命中 (缓存真跨构建复用)、换 -O 2 之后 RTL 那一族
# 仍是一格不多 (RTL 的编译档与用户 -O 解耦) 而用户码必须多占一格 (键确实跟着优化档走)、
# 三臂产物跑起来 stdout 逐字相同 (命中不改产物 —— 这条才是"敢默认开"的前提)。
# 本机实测：冷编 24.5 s / 全命中 2.9 s / 换 -O 2 6.5 s / 换架构 x86 24.6 s。
function Test-ObjCache {
    param([string]$Name, [string]$Source)
    # vbp 分片: 本片不跑这例 —— 闸门必须在任何副作用之前 (建目录/起进程)。
    if (-not (Enter-VbpShard)) { return }
    $script:total++
    Write-Host -NoNewline "  [OBJCACHE] $Name ... "
    $reasons = @()
    $outs = @('', '', '')
    $rtl = @('', '', ''); $usr = @('', '', '')
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($Source)

    $sandbox = Join-Path $OutDir ("{0}_store" -f $Name)
    if (Test-Path $sandbox) { Remove-Item -Recurse -Force $sandbox }
    New-Item -ItemType Directory -Force -Path $sandbox | Out-Null
    $ladSaved = $env:LOCALAPPDATA
    $env:LOCALAPPDATA = $sandbox
    try {
        for ($i = 0; $i -lt 3; $i++) {
            # NB: 别拿 @(@(), @(), @('-O','2')) 枚举三臂 —— PS 把里面的空数组压平, 三臂会
            # 悄悄变成 "-O" / "2" / 无 (踩过)。
            $extra = @()
            if ($i -eq 2) { $extra = @('-O', '2') }
            $dir = Join-Path $OutDir ("{0}_{1}" -f $Name, $i)
            if (Test-Path $dir) { Remove-Item -Recurse -Force $dir }
            New-Item -ItemType Directory -Force -Path $dir | Out-Null
            $c3Args = @($Source, '--output-dir', $dir, '--incremental') + $extra
            $log = (& $C3 @c3Args 2>&1) | Out-String
            if ($LASTEXITCODE -ne 0) { $reasons += ("build{0} rc={1}" -f $i, $LASTEXITCODE); continue }
            $m = [regex]::Match($log, 'OBJCACHE rtl=(\d+)/(\d+) user=(\d+)/(\d+)')
            if (-not $m.Success) { $reasons += ("build{0} 没读到 OBJCACHE 读数" -f $i); continue }
            $rtl[$i] = ('{0}/{1}' -f [int]$m.Groups[1].Value, [int]$m.Groups[2].Value)
            $usr[$i] = ('{0}/{1}' -f [int]$m.Groups[3].Value, [int]$m.Groups[4].Value)
            $exe = Join-Path $dir ($stem + '.exe')
            if (-not (Test-Path $exe)) { $reasons += ("build{0} 无 exe" -f $i); continue }
            $run = Invoke-TestExe -ExePath $exe -WorkDir $dir -Name ("{0}run{1}" -f $Name, $i)
            if (-not $run.Ok) { $reasons += ("build{0} 跑失败 {1}" -f $i, $run.Detail); continue }
            $outs[$i] = ($run.Output -join "`n")
        }
    } finally {
        $env:LOCALAPPDATA = $ladSaved
    }

    # 每个源占几格: 键 = 去掉尾部 _<hex> 之后的源名。
    $store = Join-Path (Join-Path $sandbox 'C3') 'objcache'
    $nUser = -1; $nRtl = 0
    if (Test-Path $store) {
        $slots = @{}
        foreach ($f in [System.IO.Directory]::GetFiles($store)) {
            $mm = [regex]::Match([System.IO.Path]::GetFileName($f), '^(.+)_[0-9a-f]{8,}\.obj$')
            if (-not $mm.Success) { continue }
            $k = $mm.Groups[1].Value
            if ($slots.ContainsKey($k)) { $slots[$k] = $slots[$k] + 1 } else { $slots[$k] = 1 }
        }
        foreach ($k in $slots.Keys) {
            if ($k -eq $stem) { $nUser = $slots[$k] }
            elseif ($k -like 'vb6rtl*') { if ($slots[$k] -gt $nRtl) { $nRtl = $slots[$k] } }
        }
    }
    $a = @($rtl[0].Split('/'))

    if ($reasons.Count -eq 0) {
        if ($a.Count -ne 2 -or [int]$a[1] -le 0) { $reasons += 'RTL 计数读不到 (读数形状变了?)' }
        elseif ([int]$a[0] -ne 0) { $reasons += ("空 store 第一次竟命中 {0}" -f $rtl[0]) }
        elseif ($rtl[1] -ne ('{0}/{0}' -f $a[1])) { $reasons += ("第二次没全命中: {0}" -f $rtl[1]) }
        elseif ($rtl[2] -ne $rtl[1]) { $reasons += ("-O 2 之后 RTL 变了: {0} (解耦破了)" -f $rtl[2]) }
        elseif ([int](@($usr[2].Split('/'))[0]) -ne 0) { $reasons += ("-O 2 竟复用了 /O0 的用户码 obj: {0}" -f $usr[2]) }
        elseif ($nRtl -ne 1) { $reasons += ("RTL 占了 {0} 格 (与用户 -O 无关, 该只 1 格)" -f $nRtl) }
        elseif ($nUser -ne 2) { $reasons += ("用户码 {1}.c 占了 {0} 格 (该是 /O0 与 /O2 两格)" -f $nUser, $stem) }
        elseif ($outs[0] -ne $outs[1] -or $outs[1] -ne $outs[2]) { $reasons += '三臂产物输出不一致 (命中改了产物)' }
    }
    if ($reasons.Count -eq 0) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host ("FAIL: " + ($reasons -join ' | ')) -ForegroundColor Red
        if ($Verbose) { Write-Host ("  rtl=" + ($rtl -join ' ') + " user=" + ($usr -join ' ')) }
    }
}

function Test-VbpFail {
    param([string]$Name, [string]$VbpFile, [string]$Needle)
    $script:total++
    Write-Host -NoNewline "  [VBP-FAIL] $Name ... "
    $result = & cmd /c ('"' + $C3 + '" "' + $VbpFile + '" --syntax-only 2>&1')
    $text = (($result | Out-String) -replace '\s+', ' ')
    if ($LASTEXITCODE -ne 0 -and $text.Contains($Needle)) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        Write-Host "  expected failing compile containing: $Needle" -ForegroundColor DarkGray
        if ($Verbose) { Write-Host $text }
    }
}

# ai/023 S03: vbp-level negative case that only surfaces in the FULL pipeline
# (package export boundary fires at visit time / stage 3.5, not syntax-only).
function Test-VbpBuildFail {
    param([string]$Name, [string]$VbpFile, [string]$Needle)
    $script:total++
    Write-Host -NoNewline "  [VBP-BUILD-FAIL] $Name ... "
    $result = & cmd /c ('"' + $C3 + '" "' + $VbpFile + '" --output-dir "' + $OutDir + '" 2>&1')
    $text = (($result | Out-String) -replace '\s+', ' ')
    if ($LASTEXITCODE -ne 0 -and $text.Contains($Needle)) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        Write-Host "  expected failing build containing: $Needle" -ForegroundColor DarkGray
        if ($Verbose) { Write-Host $text }
    }
}

# ai/023 S05: build must SUCCEED but emit the given warning text (D7: 校验不拒收,
# 警告必须可见 —— 成功路径吞警告的坑在 S01 已修, 本用例防回归).
# 账 #187: 真工程「编得过、链得动、出得了 exe」的正面积。三个 helper 各管一头 ——
# Test-Vbp 要跑起来并校验输出（这些第三方工程没有自退出口，CI 里会一直挂着）、
# Test-VbpBuildFail 只钉"必须红"，中间那一格"必须绿"以前没人管：本轮 #188/#190/#191 三刀
# 全落在 Charts 2020 的 UC 子工程上，而它们压根不在任何清单里，每次都靠手工真编译才看得见
# （ucChartBar 的 LNK2019 就是手工编译才抓到的 —— 编不过的东西在门禁里连"红"都算不上，是根本不存在）。
# 输出目录**按用例隔离**：这四份 .vbp 的 ExeName32 全写着 Proyecto1.exe，共用 $OutDir 会互相盖掉
# （#166 那条"exe 名有两个权威"的姊妹坑：名字修对了还会串味）。先删干净再编，
# 免得 Test-Path 命中上一轮的旧 exe —— 记忆里那条教训原话是「Test-Path $exe 不是构建成功的判据」。
function Test-VbpBuild {
    param([string]$Name, [string]$VbpFile, [string]$Arch = "")
    $script:total++
    if (-not (Enter-VbpShard)) { return }
    Write-Host -NoNewline "  [VBP-BUILD] $Name ... "
    $buildDir = Join-Path $OutDir ("build_" + $Name)
    if (Test-Path $buildDir) { Remove-Item -Recurse -Force $buildDir }
    New-Item -ItemType Directory -Force -Path $buildDir | Out-Null
    $argList = '"' + $VbpFile + '" --output-dir "' + $buildDir + '"'
    if ($Arch -ne "") { $argList += " --arch " + $Arch }
    $result = & cmd /c ('"' + $C3 + '" ' + $argList + ' 2>&1')
    $text = (($result | Out-String) -replace '\s+', ' ')
    $rc = $LASTEXITCODE
    $exePath = Join-Path $buildDir ((Resolve-VbpExeBase -VbpFile $VbpFile) + ".exe")
    if ($rc -eq 0 -and (Test-Path $exePath)) {
        $script:pass++
        Write-Host ("PASS (" + (Get-Item $exePath).Length + " bytes)") -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host ("FAIL rc=" + $rc + " exe=" + (Test-Path $exePath)) -ForegroundColor Red
        Write-Host ("    找的是 " + $exePath) -ForegroundColor DarkGray
        if ($Verbose) { Write-Host $text }
    }
}

function Test-VbpWarn {
    param([string]$Name, [string]$VbpFile, [string]$Needle)
    $script:total++
    Write-Host -NoNewline "  [VBP-WARN] $Name ... "
    $result = & cmd /c ('"' + $C3 + '" "' + $VbpFile + '" --output-dir "' + $OutDir + '" 2>&1')
    $text = (($result | Out-String) -replace '\s+', ' ')
    if ($LASTEXITCODE -eq 0 -and $text.Contains($Needle)) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        Write-Host "  expected successful build containing: $Needle" -ForegroundColor DarkGray
        if ($Verbose) { Write-Host $text }
    }
}

# ai/023 S05/S06: generic CLI check — exit 0 and output contains needle.
function Test-CliOk {
    param([string]$Name, [string[]]$C3Args, [string]$Needle, [string]$ShardGroup = "")
    # vbp 分片: 本片不跑这例 —— 闸门必须在任何副作用之前 (建目录/起进程)。
    if (-not (Enter-VbpShard -Group $ShardGroup)) { return }
    $script:total++
    Write-Host -NoNewline "  [CLI] $Name ... "
    $result = & cmd /c (('"' + $C3 + '" ' + ($C3Args -join ' ') + ' 2>&1'))
    $text = (($result | Out-String) -replace '\s+', ' ')
    if ($LASTEXITCODE -eq 0 -and $text.Contains($Needle)) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        Write-Host "  expected exit 0 containing: $Needle" -ForegroundColor DarkGray
        if ($Verbose) { Write-Host $text }
    }
}

# ai/023 S06: pack -> unpack -> every file byte-identical (acceptance: 逐文件 cmp).
function Test-PackRoundtrip {
    param([string]$Name, [string]$PkgDir)
    $script:total++
    Write-Host -NoNewline "  [PACK-RT] $Name ... "
    $tmp = Join-Path $OutDir ("packrt_" + [guid]::NewGuid().ToString("N").Substring(0,8))
    try {
        New-Item -ItemType Directory -Path $tmp | Out-Null
        Copy-Item -Recurse -Path "$PkgDir\*" -Destination $tmp
        & cmd /c (('"' + $C3 + '" --pack "' + $tmp + '" 2>&1')) | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "pack failed" }
        $pkgFile = Get-ChildItem -Path $tmp -Filter "*.c3pkg" | Select-Object -First 1
        if (-not $pkgFile) { throw "no .c3pkg produced" }
        $outDir = Join-Path $tmp "unpacked"
        & cmd /c (('"' + $C3 + '" --unpack "' + $pkgFile.FullName + '" --output-dir "' + $outDir + '" 2>&1')) | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "unpack failed" }
        foreach ($f in (Get-ChildItem -Path $PkgDir -File)) {
            if ($f.Extension -eq ".c3pkg") { continue }
            $srcHash = (Get-FileHash -Algorithm SHA1 $f.FullName).Hash
            $dst = Join-Path $outDir $f.Name
            if (-not (Test-Path $dst)) { throw "missing after unpack: $($f.Name)" }
            $dstHash = (Get-FileHash -Algorithm SHA1 $dst).Hash
            if ($srcHash -ne $dstHash) { throw "content differs: $($f.Name)" }
        }
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } catch {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        Write-Host "  $($_.Exception.Message)" -ForegroundColor DarkGray
    } finally {
        if (Test-Path $tmp) { Remove-Item -Recurse -Force $tmp }
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
    # <vbeclipse>: Array() 空数组那一形 (UBound=-1/LBound=0/For Each 零次) 与"未分配 →
    # 运行时错误 9"两形; 读数全取自真实输出。常量实参那形见 arr_n01_ubound_vbnull。
    Add-BasTest "test_arr_empty" "$Tests\test_arr_empty.bas" @("EA-ub=-1", "EA-lb=0", "EA-n=0", "EA-isarr=True", "EA-join=[", "EA-fe=0", "EA-vub=-1", "EA-3ub=2", "EA-e2=9", "EA-rb-lb=1", "EA-rb-ub=3", "EA-err=9", "EA-ea-nullread=9", "EA-ea-nullwrite=9",
        "EA-ea-in=20/30/50", "EA-ea-above=9", "EA-ea-below=9", "EA-ea-udt-in=11/22/0/0",
        "EA-ea-udt=9", "EA-ea-before=8", "EA-ea-erase=9", "EA-ea-loop=100", "EA-DONE")
    # <vbeclipse> 账 #214: 多维数组的**元素**访问收成一处带检查的出口 —— 越界 / 未分配一律错误 9
    # (以前: 二维越界静默拿到隔壁那格、二维未分配读 NULL->data 当场 AV、4 秩连"在范围内"都崩 ——
    #  发码那一条把实参拼成 (int[]){i0, i1} 只塞了两个下标却按实际秩数交出去)。NA-in / NA-dyn-in 是范围内的负控。
    Add-BasTest "test_arr_nd" "$Tests\test_arr_nd.bas" @("NA-udt=5/6", "NA-in=0/23/211/2112",
        "NA-above=9", "NA-below=9", "NA-null=9", "NA-dyn-in=102/23", "NA-rank=9", "NA-DONE")
    # <vbeclipse> 账 #216: 位运算 (And/Or/Xor/Eqv/Imp) 与一元 Not 的**结果类型**收成一处权威
    # (TypeSystem::bitwiseResult / logicalNotResult，语义层与发码层都调它)。改前发码那份
    # oracle 恒答 Boolean，于是位运算数一进字符串上下文就打 True/False，COM 实参还被装箱成
    # VT_BOOL (真工程里 MainForm 的 Render(hdc Or 0, ...) 就是这一条)。两头钉：数值档按数答
    # (BF-or/xor/and/not/eqv/imp/byte*/int-not/dblnot)，两侧都是 Boolean 时**仍**是 Boolean
    # (BF-bool-or / BF-bool-and / BF-boolbox)，条件位照旧当真假用 (BF-cond) —— 拦反向错。
    $bitsNeedles = @("BF-or-lit=51", "BF-or-var=51", "BF-xor=17", "BF-and=34", "BF-not=-6",
        "BF-int-or=3", "BF-mixed=-1", "BF-cstr=51", "BF-left=51", "BF-byte-and=80",
        "BF-byte-or=95", "BF-byte-not=-86", "BF-int-not=-2", "BF-dblnot=-3", "BF-eqv=-7",
        "BF-imp=-5", "BF-assigned=51", "BF-bool-or=True", "BF-bool-and=False", "BF-boolbox=False",
        "BF-cond=hit", "BF-boolcond=miss", "BF-sum=85", "BF-DONE")
    Add-BasTest "test_bitops" "$Tests\test_bitops.bas" $bitsNeedles
    Add-BasTest "test_bitops_x86" "$Tests\test_bitops.bas" $bitsNeedles -Arch "x86"
    # <vbeclipse> 账 #215: 体级声明收成「一条声明符一条 LocalDeclStmt」。改前四条体级路两种形状 ——
    # Dim 自己手写一遍展开(那份副本漏了 WithEvents 与「后缀即类型」)，Const/Static/体级 Public 把
    # MultiDecl 原样交给语义层(switch 不认) ⇒ 一枚名字都不登记、每条使用一条 VB3001(真工程一片 276 条)，
    # 而 `Dim a&, b&` 第二枚落回 Variant 是**真值差**(BD-empty 改前 3/0)。BD-* 其余六条是护栏(改前改后同)。
    $bdNeedles = @("BD-dim=9", "BD-suffix=Long/Long/15", "BD-empty=3/3", "BD-const=345",
        "BD-static=33", "BD-arr=13/2/3", "BD-variant=Long/String/1", "BD-DONE")
    Add-BasTest "test_bodydecl" "$Tests\test_bodydecl.bas" $bdNeedles
    Add-BasTest "test_bodydecl_x86" "$Tests\test_bodydecl.bas" $bdNeedles -Arch "x86"
    # <vbeclipse> 账 #217 第一刀: `Case Is > 2` 里的 Is 不是标识符。改前 parser 造一枚
    # IdentifierExpr("Is") 留在 AST 里借优先级解析，语义层就把它当未声明的名字 —— 模块写着
    # Option Explicit 时每形一条 VB3001(两份真工程 36 条，发码语料全仓 138 条)，没写的那面
    # 每枚用到的过程发一枚没人引用的 vb6_VARIANT 局部(探针实测)。发码从来只读比较符与右操作数，
    # 值一直是对的，所以这里钉的是「形状换了、判定没换」: 六个关系符各自的分档(CI-rel/CI-ne)、
    # Is 与普通值混写同一形(CI-mixed，优先级不能把 `, 1` 吃进右操作数)、字符串 Select(CI-word)、
    # Single Select(CI-half，Fix 136 那一档的 Case 侧)、Variant 测试表达式(CI-var)。
    $ciNeedles = @("CI-rel=lt/le/eq/gt/ge/ge", "CI-ne=is4/ne4", "CI-mixed=big/big/two/rest",
        "CI-word=beq/gtm/rest", "CI-half=hi/lo/mid", "CI-var=hi/lo/hi", "CI-DONE")
    Add-BasTest "test_caseis" "$Tests\test_caseis.bas" $ciNeedles
    Add-BasTest "test_caseis_x86" "$Tests\test_caseis.bas" $ciNeedles -Arch "x86"
    # <vbeclipse> Fix 2026-10-06: Err.LastDllError 必须是"Declare 调用那一刻的快照"。
    # VB6 每次由 Declare 发起的 DLL 调用: 1) 先 SetLastError(0), 2) 调 API,
    # 3) 返回后立刻 GetLastError() 存进 Err.LastDllError (MSDN ErrObject.LastDllError +
    # 社区对 VB6 运行时的逆向, tek-tips 222-475113 "VB 里 GetLastError 恒为 0")。
    # 改前 cgen_expr_member_precheck.inc / cgen_expr_with.cpp 把 Err.LastDllError 直接映成
    # **访问那一刻**的 GetLastError(), 三条可观测后果全反; 现在由 cgen_decl_api.cpp 为每个
    # Declare 生成的 static __inline 捕获包装承担 (不能改成"调用点插两条语句" —— Declare
    # 调用可以落在 Do While 条件这种表达式位置, 语句注入会把 C 撕成非法代码, 夹具的
    # LDL-loop 就钉这一点)。LDL-survive 是主判据: 200 次字符串拼接 + 打印之后快照仍是 126。
    # LDL-preclear/LDL-cleared 钉 Err.Clear: VB6 文档 "Clear Method (Err Object)" 的 Clear 后
    # 属性表逐项写着 LastDLLError 0 —— 快照不许跨过一次 Clear 存活。
    $ldlNeedles = @("LDL-ret=0", "LDL-fail=126", "LDL-len=492", "LDL-survive=126",
        "LDL-ok=0", "LDL-raw=0", "LDL-loop=4", "LDL-preclear=126", "LDL-cleared=0", "LDL-DONE")
    Add-BasTest "test_lastdllerror" "$Tests\test_lastdllerror.bas" $ldlNeedles
    Add-BasTest "test_lastdllerror_x86" "$Tests\test_lastdllerror.bas" $ldlNeedles -Arch "x86"
    # <vbeclipse> 账 #220: B / BF 曾被 RTL 当成两枚**外部链接的 C 全局** (const int32_t B = 1;
    # BF = 2;) 去接 Picture.Line 发码原样吐出的语法旗标。用户模块一发 `Public B As Long` 就撞成
    # C2373 + C2166 给 const 赋值 —— 改前探针实测 BUILD-RC=1 / 5 条诊断 / no exe。旗标现在由
    # parser 在 Line 的 style 位置折成字面量, 这两个名字整条归还给用户 (位置那一头见 pcline 的
    # [CODEGEN-NOTE]，两处不许互相覆盖)。
    $ncNeedles = @("NC-B=13 NC-BFLEN=2", "NC-ACC=15", "NC-BOX=3/8", "NC-DONE")
    Add-BasTest "test_nameclash" "$Tests\test_nameclash.bas" $ncNeedles
    Add-BasTest "test_nameclash_x86" "$Tests\test_nameclash.bas" $ncNeedles -Arch "x86"
    # 账 #219: RTL 那枚裸名全局 Changed 撤掉后, 模块级 Public Changed As Long 整个归用户。
    # 改前实测 BUILD-RC=1 / no exe (C2371 重定义), 改后两种架构都要出 exe 并读出这个数。
    $nc219Needles = @("NC219-CHANGED=6", "NC219-VT=3", "NC219-DONE")
    Add-BasTest "test_rtl_naked_changed" "$Tests\test_rtl_naked_changed.bas" $nc219Needles
    Add-BasTest "test_rtl_naked_changed_x86" "$Tests\test_rtl_naked_changed.bas" $nc219Needles -Arch "x86"
    # <vbeclipse>: Join/Filter 的数组槽 (Variant 数组曾按 BSTR* 读 → 段错误; Filter 的
    # VB6 可选参曾不补 → C2198/C2440 编不过)。含 1-based 源数组、零命中空数组、非字符串元素 → 13。
    Add-BasTest "test_joinfilter" "$Tests\test_joinfilter.bas" @("JF-var=[abc|xyz|abd]", "JF-var-def=[abc xyz abd]", "JF-str=[abc|xyz|abd]", "JF-f-lb=0 ub=1", "JF-f=[abc|abd]", "JF-none-ub=-1", "JF-excl-ub=0", "JF-excl=[xyz]", "JF-err=13", "JF-DONE")
    # <vbeclipse>: 内置函数 VB6 合法最小实参形 (Format(x)/Filter(a,x)/GetAttr/SetAttr) 的
    # 编译面护栏 —— 价值在"整条用例能编过", needle 只取 locale 无关读数
    # (AR-fmt2 走小数点, 随区域设置变, 故不入 needle)。
    Add-BasTest "test_builtin_arity" "$Tests\test_builtin_arity.bas" @("AR-fmt=4", "AR-attr=True", "AR-join1=[abc xyz  ]", "AR-filter=[abc]", "AR-DONE")
    # <vbeclipse>: Variant **装箱表本身**的哨兵 (Boolean=11 / Byte=17 / Single=4 / Date=7,
    # 加 CLng/CDbl/CDate/IsDate 的反向档)。装箱以前有两份并行表 + 9 个裸点, 某一档对不对
    # 取决于表达式走哪条路 (DT43 的布尔就是这么漏的) —— 表再漏一档必须红在表上。
    $boxNeedles = @("VB-mod-bool=11/Boolean", "VB-mod-byte=17/Byte", "VB-loc-bool=11/Boolean",
        "VB-loc-byte=17/Byte", "VB-func-bool=11/Boolean", "VB-expr-bool=11", "VB-mod-long=3",
        "VB-mod-int=2", "VB-mod-str=8", "VB-mod-sin=4", "VB-mod-dbl=5", "VB-mod-date=7",
        "VB-asg-bool=11/Boolean", "VB-asg-byte=17/Byte", "VB-asg-date=7", "VB-call-date=7/Date",
        "VB-date-clng=46023", "VB-date-cdbl=46023", "VB-date-isdate=True", "VB-date-cdate=True",
        "VB-sin-cdbl=1.5", "VB-sin-clng=2", "VB-DONE")
    # ⚠ 这里曾经是字面 TAB: `"$Tests<TAB>est_variant_boxing.bas"`。
    # PowerShell 的 "" 里 **`\t` 不是转义**（转义是反引号），所以只要有人把 `\t` 写成真 TAB,
    # 路径就变成 `tests` + TAB + `est_...` —— 门里读到的是 `error VB1006: 无法打开文件`,
    # 而本机 `Get-ChildItem` 永远看不到它（文件明明在）。判据见文件末尾的
    # `Assert-TestRegistry`（注册表自检: 每条 Add-* 的 .bas/.vbp 必须真存在）。
    Add-BasTest "test_variant_boxing" "$Tests\test_variant_boxing.bas" $boxNeedles
    # <vbeclipse>: vbTextCompare 六个入口对表 (InStr 两形/InStrRev/Replace/Split/Filter/
    # StrComp)。修复前 RTL 5 处 (void)compare + InStr 第 4 参被截、三参"字符串优先"形
    # 错接槽位 (实测段错误)；另修 Debug.Print StrComp(...) 被 "vb6_Str" 前缀误判成 BSTR
    # (把返回值 1 当指针 → 段错误)。读数全 ASCII, 两架构应一致。
    Add-BasTest "test_text_compare" "$Tests\test_text_compare.bas" @("TC-instr-b=0", "TC-instr-t=7", "TC-instr4-t=20", "TC-instr4-b=0", "TC-rep-b=[-A-A]", "TC-rep-t=[----]", "TC-split-b-ub=0", "TC-split-t-ub=2 [a|b|c]", "TC-filter-b-ub=0", "TC-filter-t-ub=1 [abc|ABD]", "TC-sc-b=1", "TC-sc-t=-1", "TC-sc-eq=0", "TC-DONE")
    Add-BasTest "test_fileio" "$Tests\test_fileio.bas"
    # --- Fix 197: RTL 文件 I/O 在非 ASCII 路径下必须工作 ---
    # 源码纯 ASCII, 中文文件名在运行时用 ChrW 拼出, 检查 9 项 MkDir/Print#/Line Input/
    # Write#/FileCopy/Kill/Name/ChDir/RmDir 全过 (缺一即 FAIL: 断言 needle 缺失)。
    Add-BasTest "test_nonascii_fileio" "$Tests\test_nonascii_fileio.bas" @("NA1-MKDIR=Y", "NA2-PRINT-SIZE=Y", "NA3-ROUNDTRIP=Y", "NA4-WRITE=Y", "NA5-FILECOPY=Y", "NA6-KILL=Y", "NA7-NAME=Y", "NA8-CHDIR=Y", "NA9-RMDIR=Y", "NONASCII-FILEIO-DONE")
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
    # 账 #172: 日期**字面量** #...# 的编译期折算。改前 parser 只挂原文、发码读没写过的联合体槽
    # (Debug 恰好 0、Release / /RTCu 是垃圾)，所以针面两头钉：与 DateSerial 的运行期算法同值、
    # 带 PM 的那条只比整日期多出不到一天。负控 = 旧编译器上 A/B/C/E/F/G/H/I/J/K 全 False、
    # L-serial=-6.27743597849989e+66。
    Add-BasTest "test_datelit" "$Tests\test_datelit.bas" @("A-eq2=True", "B-epoch=True", "C-maxday=True", "D-plus31=True", "E-pm-after=True", "F-pm-in-day=True", "M-pm-not-next=True", "G-leap=True", "H-y99=True", "I-y49=True", "J-dash-dmy=True", "K-case-hit=True", "L-serial=43832", "DL-DONE")
    Add-BasTest "test_variant" "$Tests\test_variant.bas" @("PASS1a", "PASS1c", "PASS5", "Done")
    # ai/022 W1: Boolean 的类型可见性 + 装箱口径。32 条读数逐条钉: 转字符串的四条路
    # (B1-B8)、类型标记 (B9-B14)、落进 Variant 的那一半 (B15-B20)、反向护栏 —— Integer /
    # Long / Byte 的装箱读数一个都不许跟着动 (B21-B27)、判定语义 (B28-B29)、插值 (B30)、
    # 裸值 Debug.Print (B31-B32)。
    $boolNeedles = @("BOOL-DONE") + (1..30 | ForEach-Object { "B$_=Y" }) + @("B31-rawTrue", "B32-rawFalse")
    Add-BasTest "test_bool_display" "$Tests\test_bool_display.bas" $boolNeedles
    # 账 #238: 与 Variant 比较的那个标量操作数以前**按地址交出去** (`&d`, d 是 double) ——
    # vb6_VarCmp* 于是把一个 8 字节标量当 vb6_VARIANT 的布局读: 相等的两个数答 False (改前
    # 两架构实测 VC01..VC05 / VC10 / VC15 七条全 False), 而且读过头。现在取址只问一处判据
    # (CCodeGen::cmpOperandMayTakeAddr → 声明那几张表 + 符号表)。两头都钉: 每条相等都有
    # 一条近似不等必须 False (否则"恒真"也算修好了); Long / String 那两条本来就走快捷通道与
    # vb6_StrCmp, 一起钉住 —— 把它们改道到 VarCmp 也会红。
    $varcmpNeedles = @("VC01-var-eq-double=True", "VC02-double-eq-var=True", "VC03-var-eq-single=True",
        "VC04-var-eq-currency=True", "VC05-module-var-eq-double=True",
        "VC06-var-ne-double-true=True", "VC07-var-eq-double-false=True",
        "VC08-var-gt-double=True", "VC09-double-lt-var=True", "VC10-rel-reverse-false=True",
        "VC11-var-eq-long=True", "VC12-var-eq-long-false=True",
        "VC13-var-eq-string=True", "VC14-var-eq-string-false=True",
        "VC15-var-eq-date=True", "VC16-var-eq-date-false=True", "VC-DONE")
    Add-BasTest "test_varcmp_scalar" "$Tests\test_varcmp_scalar.bas" $varcmpNeedles
    Write-Host ""

        # --- P5.5 数据类型兼容性测试 ---
    Write-Host "--- Compat Tests (P5.5) ---" -ForegroundColor Yellow

    # Fix 195: Replace 必须返回**完整**结果。原实现无条件用手工 malloc + 字符数长度
    # 前缀构造结果 BSTR, 而 Windows 下 vb6_BSTR_Len 走 SysStringLen (前缀是字节数) ——
    # 任何替换成功的结果都被截成一半 (Replace("abc","b","") 得 "a" 而非 "ac"),
    # 且该内存会被 SysFreeString 释放 → 堆损坏。start>1 时还漏了前缀字符。
    Add-BasTest "test_replace" "$Tests\test_replace.bas" @(
        "R1=OK", "R2=OK", "R3=OK", "R4=OK", "R5=OK",
        "R6=OK", "R7=OK", "R8=OK", "R9=OK", "R10=OK", "REPLACE-DONE")
    Add-BasTest "test_compat" "$Tests\test_compat.bas"
    Add-BasTest "test_types" "$Tests\test_types.bas"
    Add-BasTest "test_control" "$Tests\test_control.bas"
    Add-BasTest "test_declare" "$Tests\test_declare.bas"
    # --- Fix 161b-decl-out: Declare A 版 API 的 String 出参回写 + SDK A/W 宏抢占 ---
    # ByVal String 当可写缓冲 (GetUserName/GetModuleFileName 形态) 必须能回读;
    # VB 名恰是 SDK A/W 宏名 (GetUserName→#define GetUserName GetUserNameW) 时,
    # 必须有显式 Alias 才走 vb6_di_ 桩绕开宏; CreateWindowExA 类名不得乱码。
    # 期望挂在本批自己的夹具上 (避免"期望挂错夹具"的假红)。
    # 判据匹配现在由 Test-NeedleHit 做字面比 (方括号不再是通配符), 故此不必再用无括号的替代标记行。
    #   不得出现方括号 (夹具因此额外打印无括号的稳定标记行)。
    Add-BasTest "test_declare_byval_string_out" "$Tests\declare_out\declare_byval_string_out.bas" @("byval-name-ok=Y")
    Add-BasTest "test_declare_gmn_path_out" "$Tests\declare_out\declare_gmn_path_out.bas" @("gmn-path-ok=Y", "gmn-fixed-ok=Y")
    Add-BasTest "test_declare_byref_string_out" "$Tests\declare_out\declare_byref_string_out.bas" @("byref-name-ok=Y")
    # 对照: ByRef UDT 路径本来就正常 (证明"出参读法不通"不适用于 UDT)
    Add-BasTest "test_declare_byref_udt_out" "$Tests\declare_out\declare_byref_udt_out.bas" @("hr=0")
    Add-BasTest "test_declare_cwex_ansi" "$Tests\declare_out\declare_cwex_ansi.bas" @("hwnd-ok=Y")
    # Fix 161c: Declare A 版 ByVal String 实参是**字面量 / 内联 COM 属性读**。
    # 修复前出参回写只按 AST 种类判左值, MemberAccessExpr 判 true 但生成的是
    # vb6_VariantToString(...) 右值 → &(右值) → C2102, 真实工程 Charts 2020 编译失败。
    Add-BasTest "test_declare_cwex_lit_com" "$Tests\declare_out\declare_cwex_lit_com.bas" @("lit-com-hwnd-ok=Y", "lit-only-hwnd-ok=Y")

    # ai/029:429 那条"未登记的控件属性按数值读漏裸指针"（Fix 161d）。
    # 两条路: Select Case (cgen_select.cpp 无参 resolveComValue ⇒ 默认 BSTR, 而
    # tempType 是 int32_t) 与 Not (cgen_expr.cpp 的 UnaryOp::Not 落 (int32_t)(operand);
    # Fix 092r 当时只补了 Negate)。同族的 Fix 092n(For)/092r(Negate) 早修过。
    # 只有**未登记**属性中招 —— 已登记的走 getControlPropReadFn 专属 getter。
    # 靶子选 TreeView.Caption: 未登记 **且宿主答得出值**(GetWindowTextW ⇒ 真 BSTR,
    # 非 NULL) ⇒ 基线截出的指针低位非 0, 判据能红。(用 Style 会假绿: 宿主答 Empty
    # ⇒ StringProp 给 NULL ⇒ 截成 0, 与修复后同值 —— 实测踩过。)
    # 判据一律取**值**: 基线 `~(指针低位)` 也是非 0, `If Not x` 的真假分不出来。
    $cpNeedles = @("P-SEL0=Y", "P-NOT=Y", "P-NOTVAL=-1",
                   "P-REG-SEL=Y", "P-REG-NOT=Y", "P-GEN-SEL=Y", "CTRLPROP-DONE")
    Add-BasTest "test_ctrlprop" "$Tests\ctrlprop\PropApp.vbp" $cpNeedles
    Add-BasTest "test_ctrlprop_x86" "$Tests\ctrlprop\PropApp.vbp" $cpNeedles -Arch "x86"
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
    # 账 #115: Len() 的"存储宽度"兜底桶把模块级 String 也吞了 (knownBstrVars_ 每过程入口 clear,
    # 只有局部声明/形参进表) => `Len(gS)` 发成 sizeof(gS): x64 读 8、x86 读 4。两条架构各真跑一次,
    # 因为**修复前的读数本身随架构变** —— 只跑默认架构就看不见这半个症状。
    Add-BasTest "test_len_width" "$Tests\test_len_width.bas" @("LW1=7", "LW2=5", "LW3=5", "LW4=2", "LW5=5", "LW6=5", "LW7=4", "LW8=2", "LW9=1", "LW10=10", "LW11=0", "LW12=5")
    Add-BasTest "test_len_width_x86" "$Tests\test_len_width.bas" @("LW1=7", "LW2=5", "LW3=5", "LW4=2", "LW5=5", "LW6=5", "LW7=4", "LW8=2", "LW9=1", "LW10=10", "LW11=0", "LW12=5") -Arch "x86"
    # 账 #113: String 形参/局部变量存进**模块变量**必须是拷贝。
    # 修复前 `gS = s;` 存的是调用方那只 BSTR 的地址，调用方一改写自己的变量
    # 就悬垂 —— 本例先用 Churn 把那块已释放的堆块盖掉，不扰动的话悬垂也会"读对"。
    # BASE 实测：S113-A 读回的是扰动串（0123456789ABCDEF）、S113-C 为空；NEW 两处都对。
    Add-BasTest "test_str_alias" "$Tests\test_str_alias.bas" @("S113-A=PAYLOAD-0123456789", "S113-B=18", "S113-C=SECOND-ONE", "S113-D=36", "S113-E=LITERAL-OK", "S113-F=ABC", "S113-G=FROM-FUNC", "S113-DONE")
    Add-BasTest "test_str_alias_x86" "$Tests\test_str_alias.bas" @("S113-A=PAYLOAD-0123456789", "S113-B=18", "S113-C=SECOND-ONE", "S113-D=36", "S113-E=LITERAL-OK", "S113-F=ABC", "S113-G=FROM-FUNC", "S113-DONE") -Arch "x86"
    # 账 #119: 局部/形参 String 目标的临时串泄漏。判据 = K32GetProcessMemoryInfo 的
    # PagefileUsage **增量**(10 万次 `s = Left(gBig, 64)`, 阈值 2MB; 修复前那 10 万只是纯漏,
    # 量级 15MB), 加上五条所有权语义护栏 (借用/别名/ByRef 写穿/拼接链/反复赋值不互踩)。
    # api1/api2 两条是自带的失效负控: 读不到内存就把判据打成 N, 免得"自己减自己"假绿。
    Add-BasTest "test_str_leak" "$Tests\test_str_leak.bas" @("S119-api1=1", "S119-api2=1", "S119-leak-ok=Y", "S119-borrow=SHARED-ME", "S119-func-len=5", "S119-concat=ABC", "S119-byref-len=3 head=C", "S119-chain=AAABBB", "S119-DONE")
    Add-BasTest "test_str_leak_x86" "$Tests\test_str_leak.bas" @("S119-api1=1", "S119-api2=1", "S119-leak-ok=Y", "S119-borrow=SHARED-ME", "S119-func-len=5", "S119-concat=ABC", "S119-byref-len=3 head=C", "S119-chain=AAABBB", "S119-DONE") -Arch "x86"
    # 账 #122: String 目标赋值的隐式 CStr。修复前整型那半是**静默**把数字当 BSTR 指针存
    # (实测 BASE 二进制 `s = n` 编得过、跑起来 SIGSEGV), 浮点/日期那半才是响的 cl C2440。
    # 19 条读数全部取自真实输出; 针面不带方括号 (原因见 test_str_leak 上面那条注释)。
    # date 那条不比字面文本 (随区域设置变), 只比"同一折算路的两个来源"是否一致。
    $cnumNeedles = @("S122-lit=123", "S122-long=42", "S122-dbl=2.5", "S122-byte=65", "S122-cbyte=65", "S122-asc=65", "S122-bool=True", "S122-expr=43", "S122-len=4", "S122-ubound=3", "S122-date-ok=Y", "S122-str-left=ABC", "S122-str-ucase=XY", "S122-str-func=MMMM", "S122-func-num=21", "S122-concat=n=42 d=2.5", "S122-borrow=SHARED", "S122-loop=20042", "S122-DONE")
    Add-BasTest "test_str_cnum_assign" "$Tests\test_str_cnum_assign.bas" $cnumNeedles
    Add-BasTest "test_str_cnum_assign_x86" "$Tests\test_str_cnum_assign.bas" $cnumNeedles -Arch "x86"
    # 账 #123: 局部/形参/局部Const 的 As Byte 在类型推断里以前完全不可见 (C 型 uint8_t 不匹配
    # 任何登记分支)。三件后果: 赋给 String 崩 (SIGSEGV)、等值比较**静默给错答案**
    # (vb6_VarCmpLongEq(&bt,…) 拿 1 字节地址当 vb6_VARIANT*)、装箱 VT_I4 而非 VT_UI1。
    # 26 条读数全取自真实输出, x64 与 x86 逐字相同; 针面不带方括号。
    $byteNeedles = @("B1-cmp-eq-true=True", "B2-cmp-eq-false=False", "B3-cmp-lt=True", "B4-cmp-module-eq=True", "B5-cmp-const-eq=True", "B6-str-local=65", "B7-str-module=66", "B8-str-const=70", "B9-str-len=2", "B10-pack-local=17", "B11-pack-module=17", "B12-pack-const=17", "B13-pack-cbyte=17", "B14-pack-param=17", "B15-packed-val=67", "B16-param-byval-eq=EQ200", "B17-param-byval-ne=NE200", "B18-byref-writethru=201", "B19-byref-cmp=True", "B20-arith=204", "B21-print=66", "B22-mixed-lt=True", "B23-mixed-sum=1066", "B24-loop-pack=3", "B25-loop-val=202", "B26-DONE")
    Add-BasTest "test_byte_visible" "$Tests\test_byte_visible.bas" $byteNeedles
    Add-BasTest "test_byte_visible_x86" "$Tests\test_byte_visible.bas" $byteNeedles -Arch "x86"
    # Fix 190: Declare "As Any" ByRef 的下标链实参必须取地址, 不能把元素值当指针
    Add-BasTest "test_asany_subscript" "$Tests\test_asany_subscript.bas" @("WITH-SUB=Y", "EXPR-SUB=Y", "SCALAR=Y", "CHAIN=Y", "ASANY-DONE")
    # Delegate (tB extension): typed function pointers, stdcall/cdecl thunks, both arches
    Add-BasTest "test_delegate" "$Tests\test_delegate.bas" @("CALL=Y", "INIT=Y", "REASSIGN=Y", "BITCMP=Y", "APICB=Y", "QSORT=Y", "MODVAR=Y", "DELEGATE-DONE")
    Add-BasTest "test_delegate_x86" "$Tests\test_delegate.bas" @("CALL=Y", "INIT=Y", "REASSIGN=Y", "BITCMP=Y", "APICB=Y", "QSORT=Y", "MODVAR=Y", "DELEGATE-DONE") -Arch "x86"
    # 账 #184: `AddressOf` 取址的过程 = Win32/COM 回调, 交出去的必须是 __stdcall **桩**的地址
    # (VB6 的 AddressOf 给的就是桩, 不是本体)。x86 下本体是 __cdecl ⇒ 每次回调把栈少弹 N*4 字节:
    # VBFlexGrid 的 SUBCLASSPROC 有 6 个形参 (漏弹 24 字节) ⇒ 起窗即堆损坏 0xC0000374, 本地 3/4 复现。
    # 这里钉"回调语义照旧"那半边: 两遍枚举计数相同 / lParam 逐字送到 / hWnd 非 0 / 返回 0 就停 /
    # 没被取址的过程不受影响。判据不依赖"桌面上有几枚窗口"(全用自洽式), 所以两片不会因环境抖。
    # "桩存在"那半边由 addrof_cb_shape 钉 —— 它的负控就是 BASE 的发码: 里面一个 aoThunk_ 都没有。
    $aoNeedles = @("AC1-tag=Y", "AC2-echo=Y", "AC3-hwnd=Y", "AC4-stop=Y", "AC5-ret=Y", "AC6-plain=Y", "AC-DONE")
    Add-BasTest "test_addrof_cb" "$Tests\test_addrof_cb.bas" $aoNeedles
    Add-BasTest "test_addrof_cb_x86" "$Tests\test_addrof_cb.bas" $aoNeedles -Arch "x86"
    # Overloading (tB extension): same-name by-type/arity resolution, Optional span, Variant tier
    Add-BasTest "test_overload" "$Tests\test_overload.bas" @("F-L5", "F-Shi", "F-L6", "3", "5", "O0", "O17", "21", "OVERLOAD-DONE")
    Add-BasTest "test_overload_x86" "$Tests\test_overload.bas" @("F-L5", "F-Shi", "F-L6", "3", "5", "O0", "O17", "21", "OVERLOAD-DONE") -Arch "x86"
    # Generics (tB extension): monomorphizing prepass — generic UDT (nested/multi-arity),
    # generic Function/Sub with explicit instantiation + call-site type inference.
    Add-BasTest "test_generics" "$Tests\test_generics.bas" @("G-A=12", "G-B=hi", "G-C=21 abc!", "G-D=10", "G-E=10", "G-F=ab", "G-G=20", "LEN=2", "G-H=0", "G-I=20", "GENERICS-DONE")
    Add-BasTest "test_generics_x86" "$Tests\test_generics.bas" @("G-A=12", "G-B=hi", "G-C=21 abc!", "G-D=10", "G-E=10", "G-F=ab", "G-G=20", "LEN=2", "G-H=0", "G-I=20", "GENERICS-DONE") -Arch "x86"
    # --- ai/028 V1: 反引号原始多行串 (C3 扩展; 词法出口归一, 计划书 R4) ---
    # 18 条读数逐字比掉「与手写的 VB6 串相等」。其中 RS14 = Const 落点、RS15 = Declare 的
    # Lib 名落点 (DI 桩所属家族按 Lib 串选, 折错就 LNK2019)、RS16/17/18 = 定界计数。
    $rsNeedles = @("RS01=OK", "RS02=OK", "RS03=OK", "RS04=OK", "RS05=OK", "RS06=OK",
                   "RS07=OK", "RS08=OK", "RS09=OK", "RS10=OK", "RS11=OK", "RS12=OK",
                   "RS13=OK", "RS14=OK", "RS15=OK", "RS16=OK", "RS17=OK", "RS18=OK",
                   "RAWSTR-DONE")
    Add-BasTest "test_rawstr" "$Tests\test_rawstr.bas" $rsNeedles
    Add-BasTest "test_rawstr_x86" "$Tests\test_rawstr.bas" $rsNeedles -Arch "x86"    # ai/028 V2: 串内插值 (美元花括号开孔)。21 条读数逐条钉"与手写的 & CStr() / Format$
    # 同读数" —— 展开在词法出口, 编译器下游看到的就是手写形 (计划书 R2/R4)。
    $riNeedles = @("RI01=OK", "RI02=OK", "RI03=OK", "RI04=OK", "RI05=OK", "RI06=OK",
                   "RI07=OK", "RI08=OK", "RI09=OK", "RI10=OK", "RI11=OK", "RI12=OK",
                   "RI13=OK", "RI14=OK", "RI15=OK", "RI16=OK", "RI17=OK", "RI18=OK",
                   "RI19=OK", "RI20=OK", "RI21=OK", "p=1234", "INTERP-DONE")
    Add-BasTest "test_interp" "$Tests\test_interp.bas" $riNeedles
    Add-BasTest "test_interp_x86" "$Tests\test_interp.bas" $riNeedles -Arch "x86"

    # 同一份内容存成 4 种 (编码 x 行尾): 前一对量解码, 后一对量「行界归一成 CRLF」这条口径。
    foreach ($v in @("rs_utf8bom_crlf", "rs_utf8bom_lf", "rs_gbk_crlf", "rs_gbk_lf")) {
        # 参数位置上不能直接写 "..." + $v + "...": PowerShell 会把 + 当独立实参传进来
        # (实测四份变体全成 FAIL (compile)，因为 $Source 只拿到目录)。$($v) 显式界定变量名。
        Add-BasTest ("rs_var_" + $v) "$Tests\rawstr_var\$($v).bas" @("RV1=OK", "RV2=OK", "RV3=OK", "RV-DONE")
    }

    # Interface (tB extension, ai/022 B01): contract-block syntax layer - the blocks
    # parse, the new keywords stay soft, and the codegen path is still untouched.
    Add-BasTest "test_interface" "$Tests\test_interface.bas" @("ITF-SOFT:12", "ITF-1:OK", "ITF-2:OK", "INTERFACE-DONE")
    Add-BasTest "test_interface_x86" "$Tests\test_interface.bas" @("ITF-SOFT:12", "ITF-1:OK", "ITF-2:OK", "INTERFACE-DONE") -Arch "x86"
    Add-BasTest "test_varcmp_scalar_x86" "$Tests\test_varcmp_scalar.bas" $varcmpNeedles -Arch "x86"
    Add-BasTest "test_bool_display_x86" "$Tests\test_bool_display.bas" $boolNeedles -Arch "x86"

    # ai/009 5.10 (P3, 溢出检查): 窄整型收窄赋值越界必须报 Error 6, 边界值 (255 /
    # -32768 / 32767 / 2147483647) 必须**不**报 —— 两个方向都锁, 后者防"把合法
    # 程序改成编不过"这种比不检查更糟的回归。
    $ovfNeedles = @(
        "OVF-OK b-1 err=6", "OVF-OK b-256 err=6", "OVF-OK b-255 err=0", "OVF-OK b-0 err=0",
        "OVF-OK i-40000 err=6", "OVF-OK i-i32max err=6", "OVF-OK i-min err=0", "OVF-OK i-max err=0",
        "OVF-OK l-5e9 err=6", "OVF-OK l-min err=0", "OVF-OK l-max err=0",
        "OVF-OK cb-300 err=6", "OVF-OK cb-200 err=0",
        "OVF-OK ci-40000 err=6", "OVF-OK ci-30000 err=0",
        "OVF-VAL b=255 i=-32768 l=2147483647", "OVF-VAL2 b=0 i=32767",
        "OVF-GOTO err=6", "OVERFLOW-DONE")
    Add-BasTest "test_overflow" "$Tests\test_overflow.bas" $ovfNeedles
    Add-BasTest "test_overflow_x86" "$Tests\test_overflow.bas" $ovfNeedles -Arch "x86"

    # ---- 注册表自检 (2026-09-30, 门 #233 的教训) --------------------------------
    # 为什么单开一道: `test_variant_boxing` 那条注册里 `\t` 被写成了**真 TAB**
    # (`"$Tests<TAB>est_variant_boxing.bas"`), PowerShell 把它原样传给了 C3 ⇒ 门里读到
    # `error VB1006: 无法打开文件: …tests<TAB>est_variant_boxing.bas`, 而本机
    # `Get-ChildItem tests\*.bas` 一定看得到这个文件 —— 两边看到的世界不一样, 于是这条红
    # 只在 GA 上出现、本地怎么点都点不出来。自检放在**分片与执行之前**: 路径不成立就整条
    # 用例队列都不许跑, 而不是等某个分片跑到它才红。
    # 判据两条: ① 路径里不许有控制字符 (`\t`/`\n`/`\r`) —— 这类字符在控制台看起来就是空白;
    #          ② 文件必须真的存在 (相对 $Tests)。
    $regErrors = @()
    foreach ($t in $basQueue) {
        $src = [string]$t.Source
        if ($src -match '[\x00-\x1F]') {
            $shown = $src -replace '[\x00-\x1F]', { '<0x{0:X2}>' -f [int][char]$_.Value }
            $regErrors += "Add-BasTest '$($t.Name)': 路径里有控制字符 (看起来像空白): $shown"
        } elseif (-not (Test-Path -LiteralPath $src)) {
            $regErrors += "Add-BasTest '$($t.Name)': 源文件不存在: $src"
        }
    }
    if ($regErrors.Count -gt 0) {
        Write-Host ""
        Write-Host "  [FATAL] bas 用例注册表自检不通过 ($($regErrors.Count) 条):" -ForegroundColor Red
        foreach ($e in $regErrors) { Write-Host "    - $e" -ForegroundColor Red }
        $script:fail += $regErrors.Count
        exit 1
    }

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
# === ai/022 B13a: ActiveX DLL 工程测试 (产物 + 生成的 COM 服务器入口) ==============
# A DLL has no stdout, so the observable surface is: (1) the project links into a .dll
# at all, (2) the def file exports the COM entry points, (3) the generated COM server
# table (dll_entry.c) carries the identity the project declared. The compiler writes
# those C files to a temp dir and deletes it unless --keep-for-debug is passed, and
# prints "intermediates kept at: <dir>" on stderr -- that line is the only handle.
function Test-VbpDll {
    param(
        [string]$Name,
        [string]$VbpFile,
        [string[]]$Needles = @(),      # asserted against dll_entry.c + activex_dll.def
        [string[]]$Absent = @(),
        [string[]]$LogNeedles = @(),   # asserted against the compiler's own output
        [string]$Arch = ""
    )
    # vbp 分片: 本片不跑这例 —— 闸门必须在任何副作用之前 (建目录/起进程)。
    if (-not (Enter-VbpShard)) { return }
    $script:total++
    Write-Host -NoNewline "  [VBP-DLL] $Name ... "

    if ($Arch) {
        $out = & $C3 $VbpFile --arch $Arch --output-dir $OutDir --keep-for-debug @IncArg 2>&1
    } else {
        $out = & $C3 $VbpFile --output-dir $OutDir --keep-for-debug @IncArg 2>&1
    }
    $exitCode = $LASTEXITCODE
    $logText = (($out | Out-String) -replace '\s+', ' ')

    if ($exitCode -ne 0) {
        $script:fail++
        Write-Host "FAIL (compile)" -ForegroundColor Red
        if ($Verbose) { Write-Host $logText }
        return
    }

    # 账 #166: DLL 那侧同理 —— 驱动命名用的还是 ExeName32，这里以前只认文件名。
    $baseName = Resolve-VbpExeBase -VbpFile $VbpFile
    $dllPath = Join-Path $OutDir "$baseName.dll"
    if (-not (Test-Path $dllPath)) {
        $script:fail++
        Write-Host "FAIL (no dll)" -ForegroundColor Red
        return
    }

    $genDir = $null
    foreach ($line in $out) {
        $t = "$line"
        $at = $t.IndexOf("intermediates kept at: ")
        if ($at -ge 0) { $genDir = $t.Substring($at + 23).Trim(); break }
    }
    if (-not $genDir -or -not (Test-Path $genDir)) {
        $script:fail++
        Write-Host "FAIL (no intermediates)" -ForegroundColor Red
        return
    }
    $gen = ""
    foreach ($f in @("dll_entry.c", "activex_dll.def")) {
        $p = Join-Path $genDir $f
        if (Test-Path $p) { $gen += ((Get-Content $p -Raw) -replace '\s+', ' ') }
    }
    if ($gen.Length -lt 40) {
        $script:fail++
        Write-Host "FAIL (no generated entry)" -ForegroundColor Red
        return
    }

    $detail = @()
    foreach ($n in $Needles)     { if (-not $gen.Contains($n))     { $detail += "missing: $n" } }
    foreach ($a in $Absent)      { if ($gen.Contains($a))          { $detail += "unexpected: $a" } }
    foreach ($n in $LogNeedles)  { if (-not $logText.Contains($n)) { $detail += "log missing: $n" } }
    if ($detail.Count -gt 0) {
        $script:fail++
        Write-Host "FAIL (assert)" -ForegroundColor Red
        foreach ($d in $detail) { Write-Host "    $d" -ForegroundColor DarkGray }
        return
    }
    $script:pass++
    Write-Host "PASS" -ForegroundColor Green
}

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

    # Fix 194: 控件属性的字符串写入值必须转 BSTR (控件数组元素具名属性曾生成
    # `vb6_SetControlText(hwnd, (BSTR)ListCount)` 直接段错误), 且 List(j) 参与
    # 字符串相等比较要按 BSTR 处理 (RTL 声明是 void* → 曾判成 VariantObject, 比较恒假)。
    Test-Vbp "ctrlprop" "$Tests\ctrlprop\CtrlProp.vbp" @("CP1=2", "CP2=2", "CP3=1", "CP4=2", "CP5=2", "CP6=1", "CP7=2", "CP8=0", "CP9=0", "CP10=1", "CP11=0", "CTRLPROP-DONE",
        "CP12=ltt/String", "CP13=xtt/String", "CP14=yes", "CP15=/end",
        "CP16=dtagL/ltt", "CP17=dtxt/xtt", "CP18=/2")
    # ai/029 C29-1a: 接上 Shape / Line 的"创建那一刀" —— 改之前 controlTypeToWin32Class 对
    # 这两个类型返回 nullptr, 控件被当"不可见控件"跳过, 句柄永远是 NULL, 屏幕上什么都没有 (029 §二-2)。
    # 21 条读数: 设计期几何落位 (CS1-CS4)、设计期整数属性落位 (CS5-CS7)、运行期读写回路 (CS8-CS10)、Line 改端点连窗口一起搬 (CS11-CS14)、手册那条"笔宽不是 1 就强制实线"的规则 (CS15-CS16)、容器 (Frame) 内的那条创建路 (CS17-CS19)、控件数组按槽位走 (CS20-CS21)。
    # 窗体自己 Unload Me 退出 => 走 Test-Vbp 拿 stdout 针,
    # 不需要 Test-GuiVbp 那套"起窗不崩"的弱判据。负控: 喂 BASE 二进制 14 条全翻红。
    $csNeedles = @("CTRLSHAPE-DONE") + (1..21 | ForEach-Object { "CS$_=Y" })
    Test-Vbp "ctrlshape" "$Tests\ctrlshape\CsApp.vbp" $csNeedles
    Test-Vbp "ctrlshape_x86" "$Tests\ctrlshape\CsApp.vbp" $csNeedles -Arch "x86"
    # ai/029 C29-1b: 文件系统三控件 (Drive/Dir/File ListBox) 接上"创建那一刀"。
    # 改之前这三类在 controlTypeToWin32Class 里缺格 => 句柄永远 NULL, RTL 那套 P20-37
    # 的填充 helper 从来没被喂过句柄; 而且这些 RTL 入口没有任何头声明, 生成代码按
    # "返回 int" 的隐式原型编译, 字符串句柄被截成 32 位 (真编译真跑才暴露的段错误)。
    # 14 条读数: 三控件都有窗口且填进去过 (CF1-CF3)、Dir 的 [名字] 约定 (CF4)、
    # 设计期 Path/Pattern 落位 (CF5-CF7)、改 Pattern 立刻重刷 (CF8-CF9)、
    # ListIndex/FileName 回路 (CF10-CF11)、目录->文件与盘->目录两条联动 (CF12-CF13)、
    # 与原生 ListBox 的读数口径一致 (CF14)。负控: 喂 BASE 二进制 CF1-CF10/12/13 翻红。
    # 账 #68: CF15/CF16 = 窗体模块里裸 Left(...) / Right(...) 走内置函数 (窗体的 Left 属性要写 Me.Left);
    # 修复前 Left 被属性抢走 (发成 vb6_GetControlLeft(hwnd) 再拼实参), 同一棵树上这两条当场红。
    $cfNeedles = @("CTRLFILES-DONE") + (1..16 | ForEach-Object { "CF$_=Y" })
    Test-Vbp "ctrlfiles" "$Tests\ctrlfiles\CfApp.vbp" $cfNeedles
    Test-Vbp "ctrlfiles_x86" "$Tests\ctrlfiles\CfApp.vbp" $cfNeedles -Arch "x86"
    # ai/029 C29-9 / 决策 D6：CommonDialog 换成原生 comdlg32，不再经 MSComDlg.OCX。
    # 为什么必须换（实测）：那个 OCX 只有 32 位，x64 里 CoCreateInstance 直接失败，于是改之前
    # 这枚控件是静默空转的 —— 探针六条属性读数全空、六个 Show* 一个都不出现，而程序照旧打完
    # 尾针、退出码 0。10 条读数：设计期四行落位（DL1-DL4）/ Filter 竖线原样读回（DL5）/
    # 另一枚不被继承（DL6-DL7）/ 运行期读写回路（DL8）/ 两枚不串（DL9）/ 六个 Show* 的发码形状（DL10）。
    # DL10 用恒假守卫把调用留在源码里：真弹框的判据要一套"起窗 + 自关"的探针（下一小批 C29-9b），
    # 否则用例会在没人点"取消"的地方把门卡死。负控：喂 BASE 二进制直接编不过
    # （C2065: vb6_hwnd_dl2 未声明 —— 那条路上压根没有属性宿主）。
    $dlNeedles = @("CTRLDLG-DONE") + (1..10 | ForEach-Object { "DL$_=Y" })
    Test-Vbp "ctrldlg" "$Tests\ctrldlg\DlApp.vbp" $dlNeedles
    Test-Vbp "ctrldlg_x86" "$Tests\ctrldlg\DlApp.vbp" $dlNeedles -Arch "x86"
    # 发码两面都要钉：原生入口在场，OCX 那一族形状不许还在场（D6：摘一类少一类）。
    Test-EmitcShape "dl_emitc_shape" @("$Tests\ctrldlg\DlApp.vbp") @(
        'vb6_RegisterCommDialogClass((void*)hInstance);',
        '"VB6_COMMONDIALOG", "",',
        'vb6_CdShowOpen((void*)vb6_hwnd_dl1);',
        'vb6_CdShowFont((void*)vb6_hwnd_dl1);',
        'vb6_CdSetFlags((void*)vb6_hwnd_dl1, 528);'
    )
    # 反面断言走新助手 Test-EmitcAbsent：只断原生入口在不够 —— 两条路并存时读数目视全绿、
    # 发码却还在 CoCreateInstance，正是本批要拆掉的东西。
    Test-EmitcAbsent "dl_emitc_no_ocx" @("$Tests\ctrldlg\DlApp.vbp") @(
        'vb6_com_dl1',
        'CLSIDFromProgID',
        'CoCreateInstance'
    )
    # 通知接线这一刀没法在无头环境里真点一下, 所以断的是发码形状: 三条 WM_COMMAND 派发
    # (含 Dir 下钻的前置判定) + 设计期 Path/Pattern 落到初值。少了任何一条, 控件就是
    # "能显示、不联动" —— 而 CF12/CF13 是手工调 Sub 证明的, 不看这里就没人盯接线。
    # ai/029 C29-T: VB.Timer 运行期真触发 + 精度提到 ms 级。
    # 改之前的实测（029 §九 C29-T 那一格）：设计期 Enabled=0 的 Timer 压根不挂表，于是
    # Timer1.Enabled = True 落到 vb6_SetTimerEnabled(vb6_hwnd_<timer>, ...) —— Timer 是无窗口
    # 控件、句柄恒 NULL => SetPropW(NULL,...) 静默丢；Interval 改了也没人重排周期；精度只有
    # SetTimer 那一档 ~15.6 ms 地板（Interval=20 实得 34.5 ms/tick、Interval=5 封顶 ~64/tick）。
    # 现在 Enabled/Interval 真的起停与重排，底层走 winmm timeSetEvent（LoadLibrary 取，
    # 不新增 import lib；取不到才退回 SetTimer）。读数是"秒级墙钟窗口里的 tick 数带区间"：
    # 20 ms 名义 50 次，允许 [40,60]。负控（BASE 二进制）10 条全翻红且每条对上症状：
    # T2=0 开不起来 / T3=32 改了不生效 / T5=16 关掉还在烧 / T6=32 精度地板。
    # 账 #156 加 T11/T12（第二枚窗体 TmForm2，自己一枚 20ms 的 Timer）：计时器 id 原来
    # 取自控件 id 那个计数器，而编译器每建一枚窗体都发一次 vb6_ResetControlId() ⇒ 第二枚
    # 窗体的 Timer 与第一枚的第一枚 Timer 同号，派发按 id 查表又是"先建的那格先命中"，
    # 于是第二枚的事件过程一次都不跑、第一枚的速率翻倍。负控读数：T11=N/0、T12=N/61
    # （第一枚那 1 秒里名义只有 10 拍）。两条故意用差 5 倍的周期，翻红时差的是量级不是抖动。
    $tmNeedles = @("TIMERPROG-DONE") + (1..12 | ForEach-Object { "T$_=Y" })
    Test-Vbp "tmtimer" "$Tests\c29timer\TmApp.vbp" $tmNeedles
    # Fix <vbeclipse> 2026-10-06: LoadRes* 实装回归 — .res 资源段加载
    # (字符串表/裸 DIB/组图标/PNG·WebP 字节签名/CUSTOM 数据/错误 326)。
    $resNeedles = @("STR=ResString-OK|Y", "BMP=Y/0", "ICO=Y/0", "PNG=Y/0", "WEBP=Y/0", "BIN=8/222", "MISS=326")
    Test-Vbp "resload" "$Tests\resload\ResLoad.vbp" $resNeedles
    $resAlphaNeedles = @("PICSET=True", "PAINTED=True")
    Test-Vbp "resalpha" "$Tests\resload\ResAlpha.vbp" $resAlphaNeedles
    Test-Vbp "tmtimer_x86" "$Tests\c29timer\TmApp.vbp" $tmNeedles -Arch "x86"
    # 账 #157: 窗体显示时把焦点交给**这枚窗体里 TabIndex 最小的那枚拿得到焦点的控件**（VB6 口径）。
    # 改之前的实测读数（029 的「s-0 第 2 条改口径」）：窗体确实是活动/前台窗、SetFocus 本身也能落地，
    # 缺的就是"显示时没人给焦点"那一步 —— 焦点停在**窗体自己**那一层。负控读数：M1=N、D1=N，
    # D2..D7 在 BASE 上也是 Y（那六条是防伪证的：把选择改成"第一枚创建的/最后的/Label/禁用的/藏起来的"
    # 都会让其中某枚翻红，而它们不证明本批那一刀）。夹具 ModalDlg 里 cmdA 的 TabIndex=4 却是赢家 ——
    # 前面压着 TabIndex 0(Label)/1(TabStop=False)/2(Enabled=False)/3(Visible=False) 四枚排除，
    # 而且 cmdA 是 .frm 里**最后声明**的那枚 ⇒ "照创建顺序挑"给出的是 cmdB，与正确答案不同。
    # 读数全走 LongPtr 形参（HexEq/NotEq 两个助手）：控件 `.hwnd` 的装箱那条路是坏的（账 #159），
    # 直接写 GetFocus() = ctl.hwnd 会恒假 —— 又一次判据自伤，绕开它才叫测到东西。
    # 账 #83(b) 的第一半（主泵走 IsDialogMessage）用同一个夹具的 MW* 那一相来断：
    # 每拍读一次焦点、再往它 post 一对 VK_TAB。判据**故意不问跳的顺序** —— `IsDialogMessage`
    # 按 z-order（≈创建顺序）走，不是 VB6 的 TabIndex 顺序（实测站点序列
    # txtMain → cmdY → cmdX → txtSecond → 回头），顺序那一刀另记账 #163。这里钉三件事：
    # 四枚可聚焦的**都走到**（MW-new=4）、**回头落在起点**（MW-repeat=txtMain）、
    # 回到起点前走过 3 枚新站（MW-hops=3）。改之前焦点压根不动 ⇒ MW-new=1。
    # 同一份产物还带一条开关侧负控：`C3_OCX_NO_DLGMSG=1` 把泵里这一句关掉 ⇒ MW-new 回到 1
    # （本地实测；开关与模态那条循环共用，先例 Fix 144b）。
    $mdNeedles = @("MODAL-DONE", "M1-startup=Y", "M2-returned=Y", "D1-first=Y",
                    "D2-notB=Y", "D3-notLbl=Y", "D4-notOff=Y", "D5-notDis=Y",
                    "D6-notHidden=Y", "D7-ticks=Y",
"MW-new=4/repeat=txtMain/hops=3",
                    "MW-seq=txtMain,txtSecond,cmdX,cmdY,")
    Test-Vbp "modal" "$Tests\modal\ModalApp.vbp" $mdNeedles
    Test-Vbp "modal_x86" "$Tests\modal\ModalApp.vbp" $mdNeedles -Arch "x86"
    Test-EmitcShape "md_emitc_focus" @("$Tests\modal\ModalApp.vbp") @(
        'vb6_Form_SetInitialFocus((void*)hwnd, (void*)vb6_hwnd_cmdA);',    # 模态窗体：四枚排除之后那枚
        'vb6_Form_SetInitialFocus((void*)hwnd, (void*)vb6_hwnd_txtMain);'  # 启动窗体：跳过 TabIndex=0 的 Label
    )
    Test-EmitcAbsent "md_emitc_skip" @("$Tests\modal\ModalApp.vbp") @(
        'vb6_Form_SetInitialFocus((void*)hwnd, (void*)vb6_hwnd_lblHead',   # Label 拿不到焦点
        'vb6_Form_SetInitialFocus((void*)hwnd, (void*)vb6_hwnd_cmdOff',    # 显式 TabStop=False
        'vb6_Form_SetInitialFocus((void*)hwnd, (void*)vb6_hwnd_cmdDis',    # 设计期禁用
        'vb6_Form_SetInitialFocus((void*)hwnd, (void*)vb6_hwnd_txtHidden', # 设计期藏起来
        'vb6_Form_SetInitialFocus((void*)hwnd, (void*)vb6_hwnd_cmdB',      # 它是第一个声明的，不是 tab 序里的目标
        'vb6_Form_SetInitialFocus((void*)hwnd, (void*)vb6_hwnd_lblMain',   # 启动窗体那枚 Label 同理
        'vb6_Form_SetInitialFocus((void*)hwnd, (void*)vb6_hwnd_t2',        # Timer 不配当焦点目标
        'vb6_Form_SetInitialFocus((void*)hwnd, (void*)vb6_hwnd_tMain'
    )
    # 账 #83(b2) + 账 #165: 容器的扩展样式位 `WS_EX_CONTROLPARENT`(**65536**) 两条创建路都接上 ——
    # 正面四条 = 顶层 Frame（两处）/ 顶层 PictureBox / **容器里的容器**（Frame2 在 Frame1 里，走的
    # 正是第二条创建路）/ 容器里的 PictureBox。反面四条 = 容器里的按钮、顶层按钮、以及账 #164 那两枚
    # 「又立回 WS_TABSTOP 的 PictureBox」——都不许挂这一位（它只属于容器）。
    # ⚠ 这一位**以前抄错了数**：`controlContainerExStyleBit` 里手写的是 0x00040000，那是
    #   `WS_EX_APPWINDOW`；CONTROLPARENT 在 SDK 头（winuser.h:2855）里是 0x00010000=65536。
    #   抄错的那两轮里，运行期读数 `EX-fr / EX-f2 / EX-pic` 打印的是**发码想发的数落到窗口上了**，
    #   证不了「那是对话框管理器认的那一位」⇒ 样式绿、行为坏，没人看出根因 = **账 #165 的根因**。
    #   现在两头都通：`in2 / deep / inpic` 三条**已从缺陷读数翻成判据**（容器里的兄弟真被走到了）。
    # ⚠ 另一课（为什么这三条以前恒 N）：相 2 原来只给 9 拍，而这一圈要 7 站 ⇒ 走不完就被裁掉，
    #   「走不到」其实是「没走到」。现在相 2 的出口改成「七站齐了就收」，拍号只当保险丝（40 拍）。
    # 账 #163 落地之后这一格开始**钉次序**：`TW-order` 钉的是首次访问的次序（不是逐拍读数，
    #   所以某一拍没动不会假红 —— 与「条数是时序不是不变量」那一族分开）。导航器口径下的六站 =
    #   父窗内 `TabIndex` 深度优先：cmdTop1(1) → Frame1(2){cmdIn1(0) cmdIn2(1) Frame2(2){cmdShy 禁用、
    #   cmdDeep(1)、picDeep 不进站}} → optFrame(4){两枚 TabStop=0 ⇒ 一站都不出} → cmdTop2(5) →
    #   Pic1(6){cmdInPic(0)}。起点是夹具自己 SetFocus 到 cmdIn1 ⇒ 读到的圈 =
    #   `cmdIn1,cmdIn2,cmdDeep,cmdTop2,cmdInPic,cmdTop1`。
    # `TW-orenter=N` 钉的是「`TabStop = 0` 的单选组不再占站」—— 那一站本是 OS 把 `WS_TABSTOP`
    #   自己挪到勾选那枚上造成的（探针 `.build/cp2` 八份读数），导航器改读创建时存的 `VB6_TabStop`。
    #   两条负控各红各的针（本地实测）：`C3_OCX_NO_TABNAV=1` ⇒ `TW-order` 退回
    #   `cmdIn1,cmdTop2,optA,cmdTop1,cmdInPic,cmdDeep,cmdIn2` 且 `TW-orenter=Y`；`C3_OCX_NO_DLGMSG=1`
    #   把泵里那一句关掉 ⇒ 一跳都不走。
    # 账 #168 的方向键那一半也在这条夹具上：`AK-pre / down / wrap / up` 四证人 = 同容器同型单选钮
    #   按 `TabIndex` 走（optA→optB）、到尾回绕（optB 再 VK_DOWN 回 optA）、VK_UP 反向回绕；
    #   `optA=N/optB=Y` 钉的是勾选真的交接过去了（走 `BM_CLICK`，组里别人被自动取消）。
    #   `clicks=` 从账 #171 起**钉数值 3**（三步方向键 = 三声，先前实测是 6）：那多出来的三条是
    #   **焦点进入组内另一枚时单选钮自己替父窗发的那条 `BN_CLICKED`**（六轮量到底：两条 SetFocus 路
    #   同形、关掉 `IsDialogMessage` 一字不变、这两枚单选钮一次子类化都没挂 ⇒ 不是 #161 那一族的
    #   「子类重发」，是 OS 按焦点移动发的通知）。口径 = 「来这条通知时它自己勾上了没有」，
    #   落在 RTL 唯一一处 `vb6_RadioClickCounts`，发码侧只对 OptionButton 的 arm 加这一道筛。
    #   负控两头都在：**BASE 编译器**读出 `clicks=6` 且 `BF-noclick=False`（`btnfocus` 的 optA 计数），
    #   自家改动若把筛选过头（真点击也不发了）则 `clicks=0` —— 两个方向都会当场红，不是单侧的绿灯。
    # ⚠ 夹具自己的坑，先写在这条路上别踩第二次：`TabWalkApp.vbp` 的 `ExeName32` **必须等于 vbp 文件名**，
    #   否则 `Test-Vbp` 按 vbp 名去找 exe ⇒ 本地怎么都过（我手动跑的是 TabWalk.exe）、CI 两条架构一起
    #   `FAIL (no exe)`（门 #218 就是这么红的）。要么改名一致，要么显式传 `-ExeName`。
    $twNeedles = @("TABWALK-DONE", "TW-walked=Y", "TW-top=Y", "TW-shy=Y",
                    "TW-in1=Y/in2=Y", "TW-deep=Y/inpic=Y", "TW-picstop=Y",
                    "TW-orenter=N",
                    "TW-order=cmdIn1,cmdIn2,cmdDeep,cmdTop2,cmdInPic,cmdTop1,",
                    "AK-pre=optA/down=optB/wrap=optA/up=optB/optA=N/optB=Y/clicks=3")
    Test-Vbp "tabwalk" "$Tests\tabwalk\TabWalkApp.vbp" $twNeedles
    Test-Vbp "tabwalk_x86" "$Tests\tabwalk\TabWalkApp.vbp" $twNeedles -Arch "x86"
    Test-EmitcShape "tw_emitc_cparent" @("$Tests\tabwalk\TabWalkApp.vbp") @(
        '1409286151L, 65536L,',           # 顶层 Frame（Frame1 与 optFrame 两处）
        '1417675278L, 65536L,',           # 顶层 PictureBox（账 #164：这一串已不含 WS_TABSTOP）
        '1350566414L, 65536L,',           # 容器里的 PictureBox（picDeep）—— 第二条创建路同口径
        '1342177287L, 65536L,'            # 容器里的容器 —— 第二条创建路也给了这一位
    )
    Test-EmitcAbsent "tw_emitc_notplain" @("$Tests\tabwalk\TabWalkApp.vbp") @(
        '1342259200L, 65536L,',           # 容器里的普通按钮不该挂这一位
        '1409368064L, 65536L,',           # 顶层按钮同样不该挂
        '1417740814L, 65536L,',           # 账 #164: 顶层 PictureBox 又立回 WS_TABSTOP
        '1350631950L, 65536L,'            # 容器里那枚又立回来（第二条创建路）
    )
    # 账 #166: 这条用例**本身就是那条红的固化** —— 工程文件叫 `NameProbe.vbp`，而产物叫
    # `RenamedProbe.exe`（`.vbp` 里 `ExeName32="RenamedProbe.exe"`，故意与文件名不同、还带后缀）。
    # 改之前 `Test-Vbp` 按文件名去找 `NameProbe.exe` ⇒ `FAIL (no exe)`；改之后走单一权威
    # `Resolve-VbpExeBase`（读 .vbp 现值，缺键才回落文件名）⇒ 正常。
    # ⇒ **这是一条能红的针**：把 resolver 摘掉/改回文件名口径，它立刻红回去，不用等门。
    Test-Vbp "exename" "$Tests\exename\NameProbe.vbp" @("EXENAME-DONE", "NP-ok=Y")
    # 账 #160 = C29-TI: 设计期 `TabIndex` 下发到窗口。RTL 那一对 (`vb6_GetTabIndex`/`vb6_SetTabIndex`，
    # 存窗口属性 `VB6_TabIndex`，没存过时 getter 答 0) 早就在，缺的只有"创建时没人发"这一刀
    # ⇒ 改之前 `.frm` 写着 `TabIndex = 5`、运行期读回 0。四档判据各管一件事：
    #   TI-explicit = 写了的照发（5/2/1，故意与声明顺序相反 ⇒ "照创建顺序编号"会全读成别的数）
    #   TI-in       = **每个父窗自己从 0 编号**（框架里那两枚读 0/9，不占窗体那一串的号）
    #   TI-fallback = 没写的用同一父窗内的声明序号兜底（picBox=3、lblIn=1、cmdSix=5）——
    #                 兜成 0 会让好几枚同时声称自己是 0，所以这一档必须钉在"能分辨"的位置上
    #   TI-runtime  = 运行期赋值照样生效（setter 那条路没被发码期那一句盖死）
    # 发码面：正面三条（顶层显式值 / 顶层兜底值 / **第二条创建路**容器里那枚），
    # 反面两条（Timer 那类无窗口的不发；窗体自己那枚 `hwnd` 也不该被发）。
    $tiNeedles = @("TABINDEX-DONE", "TI-explicit=5/2/1", "TI-in=0/9",
                    "TI-fallback=3/1/5", "TI-runtime=40/Y")
    Test-Vbp "ctrltabindex" "$Tests\ctrltabindex\TiApp.vbp" $tiNeedles
    Test-Vbp "ctrltabindex_x86" "$Tests\ctrltabindex\TiApp.vbp" $tiNeedles -Arch "x86"
    Test-EmitcShape "ti_emitc_design" @("$Tests\ctrltabindex\TiApp.vbp") @(
        'vb6_SetTabIndex((void*)vb6_hwnd_cmdFive, 5);',   # 显式值（它其实是第 2 个声明的）
        'vb6_SetTabIndex((void*)vb6_hwnd_cmdSix, 5);',    # 没写 ⇒ 兜同一父窗内的声明序号
        'vb6_SetTabIndex((void*)vb6_hwnd_cmdIn9, 9);'      # 容器里那枚 —— 走的正是第二条创建路
    )
    Test-EmitcAbsent "ti_emitc_nowindow" @("$Tests\ctrltabindex\TiApp.vbp") @(
        'vb6_SetTabIndex((void*)vb6_hwnd_t1',             # Timer 没有窗口，号没地方存
        'vb6_SetTabIndex((void*)hwnd,'                     # 窗体自己那枚句柄不该被发号
    )
    # 账 #158 的 Combo 那一半：`_GotFocus` / `_LostFocus` 以前**永远不触发** —— 发码那两支撑的是
    # `code == 1024 / 2048`，而那是 CBEM_*（发给 ComboBoxEx 的消息号），不是 WM_COMMAND 的
    # notification code；真码是 `CBN_SETFOCUS=3` / `CBN_KILLFOCUS=4`。同一段里 EN_=256/512、
    # LBN_=4/5、BN_=6/7 三张表本来就是对的，只有 Combo 这一格抄错。
    # A/B 实测（同一份夹具，两份编译器）：BASE 是 `CF-cb1=0/0`、`CF-cb2=0/0/1`（**Validate 一直是活的**，
    # 它走另一条派发），NEW 是 `1/1` 与 `1/1/1`；两份里 `CF-lb` / `CF-tx` 都是 `1/1` ⇒ 邻座那两张表
    # 没被带坏。`CF-cross` 是串台检查：LBN_SETFOCUS 与 CBN_KILLFOCUS **同为 4**，全靠 `id ==` 过滤
    # 分开 —— 谁把那道过滤挪走，这一条当场红。
    $cbNeedles = @("COMBOFOCUS-DONE", "CF-cb1=1/1", "CF-cb2=1/1/1", "CF-lb=1/1",
                    "CF-tx=1/1", "CF-each=Y", "CF-cross=Y")
    Test-Vbp "combofocus" "$Tests\combofocus\CbApp.vbp" $cbNeedles
    Test-Vbp "combofocus_x86" "$Tests\combofocus\CbApp.vbp" $cbNeedles -Arch "x86"
    Test-EmitcShape "cb_emitc_codes" @("$Tests\combofocus\CbApp.vbp") @(
        'if (id == 104 && code == 3) { extern void vb6_cb2_GotFocus();',              # CBN_SETFOCUS
        'if (id == 105 && code == 4) { extern void vb6_cb1_LostFocus();',             # CBN_KILLFOCUS
        'if (id == 104 && code == 4 && !GetPropW((HWND)lParam, L"VB6_ValidateCancel")) {'  # 带 Validate 的那一份发码
    )
    Test-EmitcAbsent "cb_emitc_wrongtable" @("$Tests\combofocus\CbApp.vbp") @(
        'code == 1024',     # CBEM_SETTBILLOS（或 RTB 的 EN_UPDATE）—— 本工程里没有这些控件
        'code == 2048'
    )

    # C29-BN-a: CommandButton / CheckBox / OptionButton `_GotFocus` / `_LostFocus` never fired
    # -- but the code table they use (BN_SETFOCUS=6 / BN_KILLFOCUS=7) was always right.  What was
    # missing is the creation-time style bit BS_NOTIFY: a BUTTON without it never reports focus
    # changes to its parent (bare-Win32 probe .build/bnnotify/bnnotify.c: two buttons differing
    # only in that bit).  BN_CLICKED does not need it, which is why button _Click always worked.
    # Same measurement found a second gap: the focus code table only walked top-level children
    # (IDs 100+), so controls inside a Frame/PictureBox (IDs 200+) got no arm at all -- now one
    # lambda serves both creation roads.  And because the same id now also carries code=6/7,
    # the _Click arm must filter on code == 0 -- BF-noclick pins exactly that.
    # A/B (same fixture, two compilers; x64 and x86 logs cmp byte-for-byte identical):
    #   BASE: BF-cmd=0/0 BF-chk=0/0 BF-opt=0/0 BF-tx=1/1 BF-in=0/0 BF-intx=0/0 BF-each=N
    #   NEW : all six counters 1/1, BF-both=Y (both roads read the same numbers), BF-noclick=Y,
    #         BF-each=Y
    $bfNeedles = @("BTNFOCUS-DONE", "BF-cmd=1/1", "BF-chk=1/1", "BF-opt=1/1", "BF-tx=1/1",
                    "BF-in=1/1", "BF-intx=1/1", "BF-both=Y", "BF-noclick=Y", "BF-each=Y")
    Test-Vbp "btnfocus" "$Tests\btnfocus\BfApp.vbp" $bfNeedles
    Test-Vbp "btnfocus_x86" "$Tests\btnfocus\BfApp.vbp" $bfNeedles -Arch "x86"
    Test-EmitcShape "bf_emitc_bothroads" @("$Tests\btnfocus\BfApp.vbp") @(
        'if (id == 106 && code == 6) { extern void vb6_cmdA_GotFocus();',
        'if (id == 200 && code == 6) { extern void vb6_cmdIn_GotFocus();',
        'if (id == 201 && code == 256) { extern void vb6_txtIn_GotFocus();',
        'if (id == 106 && code == 0) {',
        'vb6_CreateControl("BUTTON", "IN",1342259200L',
        '1409368064L, 0L,'
    )
    Test-EmitcAbsent "bf_emitc_oldshape" @("$Tests\btnfocus\BfApp.vbp") @(
        'if (id == 106) {',
        'vb6_CreateControl("BUTTON", "IN",1342242816L',
        '1409351680L, 0L,'
    )
    # ai/030 T30-A: 内容寻址 obj store —— 命中/解耦/不改产物三条一起断 (用例自带隔离 store)
    Test-ObjCache "objcache" "$Tests\hello.bas"
    Test-EmitcShape "cf_emitc_shape" @("$Tests\ctrlfiles\CfApp.vbp") @(
        'extern void vb6_drvList_Change(); vb6_drvList_Change();',
        'extern void vb6_dirList_Change(); vb6_dirList_Change();',
        'extern void vb6_fileList_Click(); vb6_fileList_Click();',
        'if (vb6_DirListBoxDescendSelected((void*)vb6_hwnd_dirList)) {',
        'vb6_DirListBoxSetPath((void*)vb6_hwnd_dirList, vb6_BSTR_FromStr(L"C:\\Windows\\System32"));',
        'vb6_FileListBoxSetPattern((void*)vb6_hwnd_fileList, vb6_BSTR_FromStr(L"*.dll"));',
        'vb6_Left(vb6_GetListItem((void*)vb6_hwnd_auxList, 0)'      # 账 #68: 裸 Left 是函数, 不是窗体属性
    )
    Test-EmitcAbsent "cf_emitc_no_stolen_left" @("$Tests\ctrlfiles\CfApp.vbp") @(
        'vb6_GetControlLeft(vb6_hwnd_CfForm, ',
        # C29-SL-d: fileList 的 Click 只有 WM_COMMAND/SELCHANGE 那一条来源。子类化那一路
        # 补上 Click 档之后必须把它排除掉（见 controlClickFromNativeNotify），否则一次点击
        # 双发；这条针钉的就是「那一路没有替它再挂一条」。
        'if (msg == WM_LBUTTONUP) { extern void vb6_fileList_Click();'
    )
    # ai/029 C29-8a: TreeView 的标量属性面换原生 SysTreeView32（D6：不碰 MSCOMCTL.OCX）。
    # 改之前的实测（029 §九 前置测量）：这枚控件**窗口本来就建得出来**（controlTypeToWin32Class
    # 早就有格，几何读数对），缺的是属性面 —— cgen 读写表里 TreeView 零格，属性一律落到
    # "未知属性"的通用兜底 vb6_ComGetObjectProp(裸 HWND, "…")，于是 tv1.CheckBoxes 读回空串、
    # 写进去静默丢，而 .frm 里的 CheckBoxes/LineStyle/Indentation 一个字节都不发。
    # 12 条读数：设计期五值落位（TV1-TV5）/ 没写过就是 VB6 默认（TV6）/ 默认缩进是从真窗口
    # 问出来的（TV7，句柄为空只会读到 0）/ 运行期赋值 + 反向可逆（TV8-TV9）/
    # Indentation 的缇值往返与超界（TV10-TV11）/ 通用属性面没被抢走（TV12）。
    # 负控：把 tv1 的五行设计期值整体取反（CheckBoxes/HotTracking 0、LineStyle 0、
    # Indentation 500、HideSelection -1）后 TV1-TV5 全翻 N，而 tv2 那七条纹丝不动 ——
    # 读数问的是那五个值，不是一句常绿。
    #
    # ai/029 C29-8b: 同一个工程再加 15 条 Nodes/Node 读数（TV13-TV27），复用 C29-3 那套
    # 真 IDispatch 成员对象机制（vb6forms_memberobj.c），**结构一律现问原生树**：
    #   TV13 Count / TV14 Add 返回对象的 Index / TV15 下标取与按 Key 取（同一条 Item 两种实参）
    #   TV16 Child+Children / TV17 Parent+Next+Previous / TV18 Root（顶层返回自己、非顶层返回根祖先）
    #   TV19 Text/Tag 写回后换一枚对象再读（证同步进了控件，不是变量里存的串）
    #   TV20 Checked（原生 state image：TVM_SETITEMW 写、TVM_GETITEMSTATE 问）
    #   TV21-TV23 Expanded 开-关-以及 EnsureVisible 把父节点撑开（TVM_ENSUREVISIBLE 的真行为）
    #   TV24 For Each 走 _NewEnum，且顺序 == 集合序 == 插入序（口径见 029 §九）
    #   TV25 Remove "a" 带走整棵子树（原生删父即删子，表要跟住）/ TV26 Clear / TV27 集合这条路
    #   没把标量属性面抢走（CheckBoxes 读数与另一枚空表控件的 Count）
    # 负控两条，**都实测过各自红在哪几条**（不是推的）：
    #   ① `tv1.Nodes.Remove "a"` 换成一个不存在的 Key "zz" → 只红 **TV25**。TV26 常绿是
    #      设计上就该这样：后面那句 Clear 照样把表清空，Count 仍是 0 —— 别把它算进负控。
    #   ② .frm 里 tv1 的 `CheckBoxes = -1` 取反成 0 → 红 **TV1 与 TV27**，而 **TV20 不红**。
    #      TV20 不红正是这条读数的价值：勾选态是节点的 state image 位，**没开 TVS_CHECKBOXES
    #      也照样写得进、读得出**（只是屏幕上不画方框）—— 与 VB6 "先设 CheckBoxes=True 才
    #      看得到" 的次序口径一致，实测就是这么钉住的。
    # ai/029 C29-8c: 再加 6 条事件读数 (TV28-TV33)。NodeClick / Expand / Collapse 走父窗的
    # WM_NOTIFY (P20-42 那条通道，C29-7/C29-4 已各自接过一半)，通知里的 itemNew.hItem 由 RTL
    # 折成 1 基节点序号，再交 vb6_TreeView_NodeAt 造对象回调 —— 处理器参数就是 Node 对象。
    # 触发用 RTL 的 Sim* 助手 (C29-4 的 SimClick 同一先例，判据专用、不对应 VB6 语义)，
    # 且必须放 Timer: Form_Load 阶段被 block events during form init 拦掉、Form_Activate 无头不来。
    # TV33 是把一件"意外"钉成读数: 光 Form_Load 自己改 Expanded / EnsureVisible 就真发出过
    # 两条 TVN_ITEMEXPANDEDW (实测到这里 gExpands 已是 2) —— 真控件发的通知与 Sim 发的走同一条链。
    # 因此 TV28/TV31 一律问**增量**，不问绝对值 (拿绝对值比会把两件事混在一起，踩过)。
    # 负控两条 (红在哪几条量过): ① Sim 的节点下标换成一个不存在的 (99/98) => 红 TV28、TV29，
    #   其余纹丝不动 (证派发真去表里查过节点)；② 摘掉 tv1_Expand 处理器 => 红 TV33、TV31、TV32，
    #   NodeClick 那几条照旧 Y (分支是按处理器存在性建的)。
    $tvNeedles = @("TREEVIEW-DONE") + (1..33 | ForEach-Object { "TV$_=Y" })
    # 账 #88 (比较上下文里字符串成员被按数值解包): TV34-TV37 四条。分工 ——
    #   TV15 从"先取进局部变量"改回**直接**比 (`ndIx.Text = "子乙"`) ⇒ BASE 上它是 N (红)，
    #   TV35 同理钉链式形态 (`tv1.Nodes(2).Text` / `.Key`)；TV36 是"必须仍然 N"的负向闸
    #   (拿另一个节点的字面量比，防上面那两条变成恒真)；TV37 钉**选择性** —— 同一条链上的
    #   数值成员 (Index/Children) 照旧走 int 档，两档一起改是错的。TV34 保留旧的局部变量
    #   形态，两版都 Y，证升级没把那条路径弄丢。
    $tvNeedles += @("TV34=Y", "TV35=Y", "TV36=N", "TV37=Y")
    # 128-a: CheckBoxes / HideSelection / HotTracking 归 Boolean 档之后的三个面
    $tvNeedles += @("TV38=True/False/Boolean")
    Test-Vbp "ctrltreeview" "$Tests\ctrltreeview\TvfApp.vbp" $tvNeedles
    Test-Vbp "ctrltreeview_x86" "$Tests\ctrltreeview\TvfApp.vbp" $tvNeedles -Arch "x86"
    # 发码面两面都钉：设计期 Init 逐参数钉（含 -999 那条哨兵：VB6 的 True 就是 -1，
    # 用 -1 当"未写"等于设计期永远勾不上复选框），创建样式位钉 TVS_HASLINES|WS_BORDER，
    # 反面断这枚控件的属性不许再走 COM 兜底、工程里不许再出现 CoCreateInstance。
    Test-EmitcShape "tv_emitc_shape" @("$Tests\ctrltreeview\TvfApp.vbp") @(
        'vb6_TreeView_Init((void*)vb6_hwnd_tv1, 1, 300, -1, -1, 0);',
        'vb6_TreeView_Init((void*)vb6_hwnd_tv2, -999, -999, -999, -999, -999);',
        '1417740290L, 0L,',
        # C29-8b: `tv1.Nodes` 必须立成**真 IDispatch 集合对象** (vb6forms_memberobj.c 的
        # NODES 族)。这三条钉的是发码形状里最容易退回的三处: 宿主槽 (真窗口 = vb6_hwnd_
        # 而非 vb6_com_)、Add 的 Missing 打包 (省略实参不能编成 0，否则 relationship 静默
        # 变成 tvwFirst)、以及"下标/Key 都发同一条 Item"。
        'vb6_ComCallObject(vb6_TreeView_Nodes((void*)vb6_hwnd_tv1), L"Item", (void*[]){vb6_ComPackInt(2)}, 1)',
        'L"Add", (void*[]){vb6_ComPackMissing(), vb6_ComPackMissing(), vb6_ComPackBSTR(vb6_BSTR_FromStr(L"a"))',
        'vb6_ComSetProp(ndA, L"Checked", vb6_ComPackBool((-1)))',
        # C29-8c: 派发面的三条针 —— Sim 助手发的是 RTL 直调 (不是 COM 派发)、WM_NOTIFY 分支
        # 按码值 + hwndFrom 双条件建、展开/折回靠 action 三态分流 (999 = 认不出，两边都不接)。
        'vb6_TreeView_SimNodeClick((void*)vb6_hwnd_tv1, 2);',
        'if (pNM42->code == -451 && (void*)pNM42->hwndFrom == vb6_hwnd_tv1) {',
        'vb6_TreeView_NotifyExpanded((void*)lParam) == -1',
        # 账 #88: 比较上下文里字符串成员必须按 BSTR 解包 (TV35 那条链式形态)。
        'vb6_CStr(vb6_VariantFromValue(vb6_ComGetStringProp(vb6_ComCallObject(vb6_TreeView_Nodes(',
        # 同一条判据的选择性那一半: 数值成员 (Index) 照旧走 int 档 —— 这条在修复前后都在,
        # 它的作用是挡住"为了修字符串把两档一起改成 BSTR"那种反向过度修正。
        'vb6_ComGetIntProp(vb6_ComCallObject(vb6_TreeView_Nodes((void*)vb6_hwnd_tv1), L"Item", (void*[]){vb6_ComPackInt(2)}, 1), L"Index")',
        'vb6_CStrBool(vb6_TreeView_GetCheckBoxes('
    )
    Test-EmitcAbsent "tv_emitc_no_com_fallback" @("$Tests\ctrltreeview\TvfApp.vbp") @(
        'vb6_ComGetObjectProp(vb6_hwnd_tv1',
        # C29-8b: 更具体的一条 —— Nodes 这一格一旦被摘掉，就会退回"拿 HWND 当 IDispatch
        # 问它要 Nodes 属性"的假路径 (链接过、运行期整棵树一句话都读不出来)。
        'vb6_ComGetObjectProp(vb6_hwnd_tv1, L"Nodes")',
        # C29-8c: Sim* 这类判据方法一旦被下面那条 axSlotObj 分支先吃掉，就会编成
        # 「取 SimNodeClick 属性 + Item 下标」—— 编得过、跑起来什么都不发 (实测踩过，
        # 修法是把钩子抢在那条分支之前)。这条断言就是别让那个形状再回来。
        'vb6_ComGetObjectProp(vb6_hwnd_tv1, L"SimNodeClick")',
        # 账 #88: 这条是 BASE 上 TV15/TV35/TV36 三处的实际发码形状 —— 字符串成员被按数值
        # 解包后再 CStrLong (拿 BSTR 指针当数字与字面量比，恒 False)。它一回来判据就得红。
        'vb6_CStrLong(vb6_ComGetIntProp(vb6_ComCallObject(vb6_TreeView_Nodes(',
        'CoCreateInstance',
        'vb6_CStrLong(vb6_TreeView_GetCheckBoxes('
    )

    # 账 #125 (设计期 Enabled / Visible / CheckBox·OptionButton 的 Value): 两条创建路各一份 ——
    #   Frame 里的子控件 (cbIn/cbDisIn/cbHidIn) 与顶层同形状 (cbOut/cbDisOut/cbHidOut)。
    #   **判据纪律**: Visible 只能在窗体真显示之后问 —— Form_Load 里父窗还没 Show，
    #   IsWindowVisible 对任何控件都返回假 (第一版探针就这么假绿过)，所以读数在 Timer 里。
    #   DS4/DS5 是**选择性**那一半: 没写过这三项的控件 (cbDef/obDef) 必须照旧 —— 不勾、
    #   不藏、不灰，Frame 仍可见；少了它，把"每个控件都设一遍"的写错了也照样绿。
    #   DS2/DS3 是账 #124 的**主判据**: 控件布尔属性归 Boolean 之后 CStr 打 False/True
    #   （修复前是 -1/0）；DS8/DS9 钉另两个消费面 —— TypeName 给出 "Boolean"、装箱 VarType
    #   是 **11 (VT_BOOL)**（修复前 TypeName=Long、VarType=3）。DS5 仍是数字档：OptionButton
    #   的 Value 在 VB6 也是 Boolean，今天还没登记（记在 029 §九，另格）。
    $cstNeedles = @("CTRLSTATE-DONE", "DS1=11", "DS2=FalseFalse", "DS3=FalseFalse",
                    "DS4=0TrueTrue", "DS5=FalseTrue", "DS6=YYY", "DS7=True",
                    "DS8=Boolean11False", "DS9=11True", "DS10=11False", "DS11=False/Boolean/11", "DS12=Y",
                    # 账 #83(a)（C29-SL-r）: VB6 的 TabStop 默认 True，而两条创建路以前都不立
                    # WS_TABSTOP ⇒ BASE 上 ST1 整条是 `0/0/0/0`、ST2 是 `3/Long`（读数见 029 的
                    # C29-SL-r-0）。四格分别是 顶层 / 容器里 / 容器里那枚 Label（拿不到焦点，不该立）/
                    # 显式写了 `TabStop = 0` 的那枚。读的是窗口 GWL_STYLE，没有自存 ⇒ 不是自洽假绿。
                    "ST1-tab=True/True/False/False", "ST2-bool=11/Boolean")
    Test-Vbp "ctrlstate" "$Tests\ctrlstate\CtrlState.vbp" $cstNeedles
    Test-Vbp "ctrlstate_x86" "$Tests\ctrlstate\CtrlState.vbp" $cstNeedles -Arch "x86"
    Test-EmitcShape "cs_emitc_state" @("$Tests\ctrlstate\CtrlState.vbp") @(
        'vb6_SetCheckValue((void*)vb6_hwnd_cbIn, 1);',    # 子控件 Value=1
        'vb6_SetCheckValue((void*)vb6_hwnd_cbOut, 1);',   # 顶层 Value=1
        'vb6_SetControlEnabled((void*)vb6_hwnd_cbDisIn, 0);',
        'vb6_SetControlEnabled((void*)vb6_hwnd_cbDisOut, 0);',
        'vb6_SetControlVisible((void*)vb6_hwnd_cbHidIn, 0);',
        'vb6_SetControlVisible((void*)vb6_hwnd_cbHidOut, 0);',
        'vb6_SetOptionValue((void*)vb6_hwnd_obOn, 1);',   # OptionButton 的 Value=-1 折成 BST_CHECKED
        # 账 #124: 控件布尔属性按 Boolean 解封 —— CStr 折成 vb6_CStrBool，装箱走 vb6_VariantBool。
        # 修复前这两处分别是 vb6_CStrLong 与 vb6_VariantFromValue（TypeName 也就跟着给 Long）。
        # 2026-09-30: 两条装箱都带 `(int16_t)` 收窄了 —— 以前"赋值点不窄化"是**两份表**的产物
        # (wrapVariantValue 自己的 Boolean 档 vs boxToVariant 的 Boolean 档), 而那正是账 #123/
        # Fix 198 那条"同一类型两条路两种结果"的老坑。现在字面量与属性读都落到 boxToVariant
        # 这一处权威 ⇒ 两形合并; 断言跟着改成实际形状 (收窄在 C 侧也是必需的: vb6_VariantBool
        # 形参是 int16_t, 而属性 getter 返回 int)。
        'vb6_CStrBool(vb6_GetControlEnabled(',
        'vb6_VariantBool((int16_t)(vb6_GetControlEnabled(',
        'v = vb6_VariantBool((int16_t)(vb6_GetControlVisible(',
        'vb6_CStrBool(vb6_GetOptionValue('
    )
    Test-EmitcShape "cs_emitc_tabstop" @("$Tests\ctrlstate\CtrlState.vbp") @(
        # 账 #83(a): 立位是**进创建参数**的（不是建好再 SetWindowLong），所以两条创建路各钉一枚。
        # BASE 上这两串一个都没有（那一趟所有控件都不立位）⇒ 能红。
        '1409368067L, 0L,',        # 顶层 CheckBox：1409286147 + WS_TABSTOP(65536) + BS_NOTIFY(16384)
        '1342259203L, 0L,',        # Frame 里的 CheckBox：同两条 —— 第二条创建路
        # 显式写了 `TabStop = 0` 的那枚**保持不立**（1409286144+3+16384，只挂 BS_NOTIFY、不挂 WS_TABSTOP）。
        # 账 #158 之后这一格的数字也跟着挪了一格 ⇒ 它现在**能红**；防的还是同一件事："默认立"被写成"一律立"。
        '1409302531L, 0L,',
        'vb6_CStrBool(vb6_GetTabStop('      # 读侧归到 Boolean 档（#124 同族）
    )
    Test-EmitcAbsent "cs_emitc_selectivity" @("$Tests\ctrlstate\CtrlState.vbp") @(
        'vb6_SetCheckValue((void*)vb6_hwnd_cbDef',      # 没写 Value 的复选框不许被设
        'vb6_SetControlVisible((void*)vb6_hwnd_lbOut',  # 没写 Visible 的标签不许被藏
        'vb6_SetControlVisible((void*)vb6_hwnd_fr',     # Frame 自己也不许被藏
        # 账 #124: 这两条是修复前的形状（布尔属性被当 Long）—— 一回来判据就得红。
        'vb6_CStrLong(vb6_GetControlEnabled(',
        'vb6_CStrLong(vb6_GetControlVisible(',
        'vb6_SetCheckValue((void*)vb6_hwnd_obOn'
    )
    Test-EmitcAbsent "cs_emitc_tabstop_off" @("$Tests\ctrlstate\CtrlState.vbp") @(
        # 账 #83(a) 的反面：拿不到焦点的这两类**不许**被立上 WS_TABSTOP
        # （Label = 1409286400、Frame = 1409286151，各自 +65536 那两串如果出现就说明排除表被改坏）。
        '1409351936L, 0L,',        # Label 立了位 —— 不该出现
        '1409351687L, 0L,'         # Frame 立了位 —— 不该出现
        # 账 #158: `BS_NOTIFY` 只给 BUTTON 族那三型 —— 拿不到焦点的这两类**不许**被带上那一位
        # （Frame = 1409286151+16384、Label = 1409286400+16384；这两串是排除表的红线）
        '1409302535L, 0L,',        # Frame 挂了 BS_NOTIFY —— 不该出现
        '1409302784L, 0L,'         # Label 挂了 BS_NOTIFY —— 不该出现
    )

    # ai/029 C29-SL-a: Slider (原生 msctls_trackbar32)。登记之前这枚控件**连窗口都没建**
    # (探针 .build/slprobe 实测: CreateControls 里没有它、vb6_hwnd_sld1 恒 NULL,
    #  属性读回全空、写进去静默丢、退出码照旧 0)。
    # 读数取自真跑输出。三枚控件的分工 ——
    #   sld1 (TickFrequency=10, 横): SL1 窗口真在、SL2+SL4 属性与**证人**一致、SL3+SL5 频率与"有没有刻度";
    #   sld2 (Orientation=1, 竖):   SL6=11 —— 创建时那一档样式位与控件自己的滑块形状对上;
    #   sld3 (什么都没写):          SL7=00 ⇒ TickFrequency 自存 0 **且控件确实没画刻度**
    #                               (实测: 光挂 TBS_AUTOTICKS 不给 TBM_SETTICFREQ 是不画刻度的);
    #   SL9=11 / SL10=00:           运行期翻向**真的生效** (第一发探针拿 TBM_GETCHANNELRECT
    #                               判方向, 得出过"改不动"的错结论 —— 那条消息的 rect 永远把
    #                               行程长度放在 x 分量, 横竖答同一组数; 订正见 029 §九 C29-SL-0);
    #   SL11-en=FalseTrue:          #124 的布尔面 + 选择性 (另一枚没动的仍可读回 True)。
    $slidNeedles = @("SLIDER-DONE", "SL1-vis=True", "SL2-ori=0", "SL3-freq=10",
                     "SL4-travel=0", "SL5-tick=-1", "SL6-vertori=11", "SL7-nofreq=00",
                     "SL8-freqset=5", "SL9-flip=11", "SL10-flipback=00", "SL11-en=FalseTrue",
                     # C29-SL-b 值面（读数取自本机真跑; SB3 是 Init 参数顺序的证人,
                     # SB9 是「什么都没写就照原生答」的证人 —— LargeChange 原生默认 20,
                     # VB6 文档那条 5 本机拿不到真值, 所以**不自作主张折成 5**, 见 029）
                     "SB1-range=10/100", "SB2-value=42", "SB3-changes=2/8", "SB4-sel=20/60/True",
                     "SB5-clampmax=100", "SB6-clampmin=10", "SB7-shrink=10/40/10",
                     "SB8-selrange-off=False", "SB9-defaults=0/100/1/20/0/0", "SB10-push=80/80",
                     # C29-SL-c 事件面（SC0 是原始证人行、不进针；分法与理由见 029 §C29-SL-c）。
                     # SC2/SC11 是「Change 按值比、不按码表枚举」的两条鉴别针 —— 拿码表硬枚举
                     # 会在同值的第二条 5 与 SetValue 之后那条上各多发一次。
                     "SC1-track=Y", "SC2-same=Y", "SC3-next=Y", "SC4-posit=Y", "SC5-endtrk=Y",
                     "SC6-line=Y", "SC7-value=54", "SC8-vert=Y", "SC9-isolate=Y",
                     "SC10-setval=Y", "SC11-baseline=Y",
                     # C29-SL-d 常规事件面（SD0 是原始证人行、不进针）。四条各自的
                     # 分法由读数钉住：SD1 钉「抬起只点 Click、不牵连 Change」，SD2 钉
                     # 「双击是单独一档、不顺带再点一次 Click」，SD3 钉「按键那条真的
                     # 动了控件」（KeyCode=37 与 Change 各一次），SD5 钉「抬键不再动值」，
                     # SD6 是认来源那条（没挂 handler 的那枚被点，别人一条都不许涨）。
                     "SD1-click=Y", "SD2-dbl=Y", "SD3-keydown=Y", "SD4-value=49",
                     "SD5-keyup=Y", "SD6-isolate=Y",
                     # SD7 是「装不装那一趟」的证人：sld2 除 Change（走父窗那条通道）外
                     # 只有 Click 一条，改之前它压根不会被子类化 ⇒ 处理器编得出来、没人送消息。
                     "SD7-install=Y",
                     # C29-SL-e 通用字符串属性面。SE1/SE2/SE3/SE4/SE6/SE11 六条是改之前
                     # 就能翻红的（两条 getter 声明成 `void*`，而装箱那张 _Generic 表把
                     # "其他指针"送去 VariantObject：CStr 打空、TypeName 答 Object、
                     # VarType 答 9、Tag 的比较答假）；SE5/SE7/SE8/SE9 四条本来就直接拿
                     # 指针、改之前也对，留着当形状证人。SE10（.frm 里设计期那一条）不登记
                     # —— 它现在读回来是空的，登记等于把错的口径钉死，另记账 #142。
                     "SE1-tip=abc", "SE2-tag=t9", "SE3-tn=String", "SE4-vt=8",
                     "SE5-len=3", "SE6-eqtag=Y", "SE7-eqtip=Y", "SE8-assign=3/abc",
                     "SE9-unset=/", "SE11-cat=v=t9",
                     # 账 #142: 设计期写了非空值就要到窗口上。SE10/SE12/SE13 三条在改之前
                     # 全是空（BASE 实测 `SE13-over=/`），SE14 是"显式写了空串 = 与默认同、
                     # 且不发 setter"那一面的运行时半边（它在 BASE 上也是 /end，只当证人）。
                     "SE10-dt=dtip", "SE12-dttag=dtag", "SE13-over=rt/",
                     "SE14-emptytag=/end",
                     # C29-SL-g 焦点面。改之前 SG1/SG2/SG3/SG4 四条在 BASE 上逐条是 N
                     # （四类白名单不含 Slider，处理器编得出来没人送消息）。SG0 是原始证人行。
                     "SG1-got=Y", "SG2-move=Y", "SG3-back=Y", "SG4-isolate=Y",
                     # C29-SL-h TickStyle 四档（枚举真值读自 OCX 自带的类型库）。
                     # SH2/SH3/SH6 钉的是「控件几何真的跟着那一档动」—— chan.top 与
                     # ts=0 基线比高低；1 与 2 之间不写几何断言（本机两处读数顺序相反，
                     # 探针 40px 答 20/19、夹具 400 缇答 18/19），两档只靠样式位读回区分。
                     # SH4 是 ts=3 唯一的证人（GetNumTicks=0；TickPresent 在那一档照旧答有）。
                     # SH7 钉越界归 0，SH8 钉类型档（Long / VT_I4，与 Orientation 同档）。
                     # SH0-geom 那行是绝对像素，只做证人、不登记 —— 登记等于把本机钉进判据。
                     "SH1-base=0/Y", "SH2-top=1/Y", "SH3-both=2/Y", "SH4-none=3/0/Y",
                     "SH5-back=0/Y/Y", "SH6-dt=2/Y", "SH7-oob=0", "SH8-type=Long/3",
                     # SH9 是**第二条创建路**（容器子控件）的证人：sld7 挂在 Frame 里，
                     # 那条路以前对 Slider 一格样式都不挂（BASE 实测创建参数 1342177280 =
                     # 只有 WS_CHILD|WS_VISIBLE，连 TBS_AUTOTICKS 都没有）。
                     "SH9-child=1/Y",
                     # C29-SL-i 选区面（类型库：SelStart 0x0007 + SelLength 0x0008 + 方法
                     # ClearSel 0x000e；VB6 那一面没有 SelEnd 这个名字）。SI2 钉「写长度时起点
                     # 不动」，SI3 钉「远端由控件夹进量程」，SI4+SI4b 钉「清空之后两端是 0，
                     # 而紧接着写长度不该静默什么都不发生」（空态锚量程下限，与控件自己那一态一致），
                     # SI5/SI6 钉「没挂 SelectRange 就是写不进去，我们不伪造」，SI9 钉设计期那一对
                     # （.frm 只写 SelStart+SelLength 时由 cgen 折成 (起, 止) 下发）。
                     "SI1-len=40", "SI2-set=20/30/10", "SI3-clamp=20/100/80",
                     "SI4-clear=0/0/0", "SI4b-reafter=10/5", "SI5-nobit=0/False",
                     "SI6-nobit2=0", "SI7-neg=0", "SI8-type=Long/3", "SI9-dt=25/40/15",
                     # SJ1 是**账 #148** 那条疑点的答案：ToolTipText 的"存回来"早有判据，
                     # "注册进宿主"这一步以前没人验过。裸编探针（无 v6 manifest ⇒ comctl v5）
                     # 里 TTM_ADDTOOLW 直接返回失败，但那不能定罪产品 —— 产物里问宿主
                     # TTM_GETTEXT 读回来：运行期设过的 sld3、设计期设过的 sld5 都是 Y，
                     # 从没设过的 sld1 是 N（第三格是负控：三条齐 Y 就说明证人问不出东西）。
                     # BASE 上整条是 N/N/N —— 未登记时读回来是对象、折成数值恒假。
                     "SJ1-reg=Y/Y/N",
                     # C29-SL-k 的 Text 气泡。SK3 是这格的心脏：拖一档之后**宿主自己**答
                     # "可见"并且把 "VOL" 原样读回来（TTM_GETTEXT 问的是 tooltip 控件，
                     # 不是我们的窗口属性）；SK4 钉 ENDTRACK 收回去；SK5 钉 TextPosition
                     # 两档摆的位置不同（只比高低、不钉绝对像素 —— 与 TickStyle 同一教训）；
                     # SK6 钉 Text 空着时气泡里就是当前的值；SK7 钉收尾之后一切都归位。
                     # SK1 刻意只当证人不当红针：BASE 上 Text 走的是"拿 HWND 当 IDispatch 问
                     # 它要 Text 属性"那支，串照样回得来，所以这一条两边都是 VOL/String/8。
                     "SK1-text=VOL/String/8", "SK2-off=N/N", "SK3-show=Y/VOL/Y",
                     "SK4-hide=N", "SK5-pos=1/Y", "SK6-num=33/33/Y", "SK7-end=N/0",
                     # C29-SL-l（账 #143）: 控件的**零实参方法**两形。证人问的是焦点自己 ——
                     # sld6 那对 GotFocus/LostFocus 只有 WM_SETFOCUS 真到才涨（每步先把焦点
                     # 挪到没挂处理器的 sld1，再问增量：SetFocus 落在已有焦点的窗口上不重发）。
                     # SN3 钉不带括号的 ClearSel：以前那一形掉 vb6_ComCall(裸 HWND, L"ClearSel")，
                     # 编得过、跑起来一声不响，区段一点没动。
                     "SN1-bare=Y/0", "SN2-paren=Y/Y", "SN3-clearsel=0/0",
                     # C29-SL-m（账 #150）: With 块里的控件方法三形。这一形在 BASE 上是**编译不过**
                     # （不是读数差），所以三条都是整件工程级的红。
                     "SN4-with-bare=Y", "SN5-with-paren=Y/Y", "SN6-with-clearsel=0/0",
                     # C29-SL-n（账 #141）: 一枚**只挂 `_DblClick`** 的滑杆（sld8）。
                     # 第二格钉 arm 传进来的 Cancel —— VB6 的签名是 `Sub X_DblClick(Cancel As Integer)`,
                     # 以前 arm 写死无参调用, 按标准签名写的处理器直接编不过（C2198）。
                     "SN7-dblone=1/0",
                     # C29-SL-o（账 #83 的 Slider 半边）: sld7 在 .frm 里写的这批值面属性，
                     # BASE 那条容器路**一条都不下发** ⇒ 整段读出原生默认档
                     # （BASE 实测 SO1=0/100/0/1 / SO2=20/0 / SO3=0/0/N，与 SB9-defaults 逐字同）。
                     # 判据不是"看着非零"：同一批属性在顶层那枚 sld4 上是 SB1=10/100 / SB2=42 /
                     # SB3=2/8 / SI9=25/40/15，两条路现在必须给出同一串数（SelLength 那格还顺带
                     # 钉住折叠后的 SelectRange 真立起来了）。
                     "SO1-child=10/100/42/2", "SO2-child-page=8/5", "SO3-child-sel=20/15/Y",
                     # C29-SL-p（说明 §4 那条例子量出来的）: FontSize 以前从窗口的整数像素高度**反算**
                     # 点号回来 ⇒ 写 8 读回 8.25、写 10 读回 9.75、写 14 读回 14.25（96 DPI 下 1pt =
                     # 1.3333px，往返必然落格）。现在请求值按窗口自存。SP1 那两条刻意**只钉前半**
                     # （`SP1b-round14=14/` 是个前缀）：后半那枚 txtG 从没设过字体、走的是反算那一支，
                     # 它的数是 DPI 的函数（本机 8.25 = 系统默认 11px），钉死等于把本机写进判据。
                     # SP2 第一格问的是**窗口**（证人 FontPixelHeight，只比相对高低）—— 自存的数
                     # 读回来当然还是自存的数，光问它就是 #148 那条自洽假绿。
                     # SP3 老实标成**行为钉**（BASE 也读 0）：0pt 折算成 lfHeight=0，反算也是 0，
                     # 两边同数 —— 它防的是「SetPropW(0) 等于删属性」那一条（账 #107）被人改回去。
                     "SP1-round=10", "SP1b-round14=14/", "SP2-real=Y", "SP3-zero=0",
                     # C29-SL-q（账 #154）: `.frm` 里写的设计期 `FontName` / `FontSize` 以前整条丢掉
                     # （BASE 同夹具：SQ1=`/8.25`、SQ2=`8.25`、SQ3=`N`），两条创建路现在都发
                     # （txtH 在窗体上、txtI 在 Frame 里 —— 这一族栽过几次"只接一头"）。
                     # SQ1 那格还顺带钉住**读侧**的一条旧坏：`FontName` 没登记成 String 档时
                     # `Print ... & ctl.FontName` 打的是空串，而 `Len(...)` 直接拿指针、答得对
                     # （探针里 "Consolas" 读成 8）—— 同一枚属性两种答案，就是没登记的证状。
                     # SQ3 问的是窗口（证人只比相对高低），BASE 那枚即使字体真换了也读不出名字。
                     "SQ1-dt=Arial/20", "SQ2-child-dt=20", "SQ3-real=Y")
    Test-Vbp "ctrlslider" "$Tests\ctrlslider\SlidApp.vbp" $slidNeedles
    Test-Vbp "ctrlslider_x86" "$Tests\ctrlslider\SlidApp.vbp" $slidNeedles -Arch "x86"
    Test-EmitcShape "sl_emitc_native" @("$Tests\ctrlslider\SlidApp.vbp") @(
        '"msctls_trackbar32", "",',
        '1409351681L, 0L,',                                        # 横杆: TBS_AUTOTICKS, 无 TBS_VERT
        '1409351683L, 0L,',                                        # 竖杆: 多挂 TBS_VERT(0x2)
        'vb6_Slider_Init((void*)vb6_hwnd_sld1, -999, -999, -999, -999, -999, 10L, -999, -999, -999);',
        'vb6_Slider_Init((void*)vb6_hwnd_sld4, 10L, 100L, 42L, 2L, 8L, 5L, 20L, 60L, -1L);',
        'vb6_Slider_Init((void*)vb6_hwnd_sld5, -999, -999, -999, -999, -999, -999, -999, -999, -999);',
        'vb6_Slider_SetOrientation(vb6_hwnd_sld1, 1);',
        'vb6_Slider_TravelIsVert(vb6_hwnd_sld2',
        # SL-b: 运行期值面 setter/getter 直发控件, 没有一格自存
        'vb6_Slider_SetMax(vb6_hwnd_sld4, 40);',
        'vb6_Slider_SetSelectRange(vb6_hwnd_sld4, 0);',
        'vb6_CStrLong(vb6_Slider_GetMin(vb6_hwnd_sld4',
        'vb6_CStrBool(vb6_Slider_GetSelectRange(',
        # C29-SL-e: 归到 String 档之后 CStr 被折掉，两条读侧直接用 getter
        'vb6_BSTR_Concat(vb6_BSTR_FromStr(L"SE1-tip="), vb6_GetToolTipText(vb6_hwnd_sld3)',
        'vb6_BSTR_Concat(vb6_BSTR_FromStr(L"SE2-tag="), vb6_GetControlTag(vb6_hwnd_sld3)',
        'vb6_SetToolTipText((void*)vb6_hwnd_sld5, L"dtip");',
        'vb6_SetControlTag((void*)vb6_hwnd_sld5, L"dtag");'                  # #124 那条布尔口径
        # C29-SL-h: 设计期那一条 TickStyle=2 必须出现在**创建参数**里（1409286145 那枚是
        # TBS_AUTOTICKS，多挂 TBS_BOTH=0x8 才是 1409286153）—— 只写自存不算下发到窗口。
        # 读侧走 CStrLong（登记成 Long 档 ⇒ CStr 不再被折成 Variant 那一趟）；写侧直接 setter。
        '1409351689L, 0L,',
        'vb6_Slider_SetTickStyle(vb6_hwnd_sld3, 2);',
        'vb6_Slider_SetTickStyle(vb6_hwnd_sld3, 9);',
        'vb6_CStrLong(vb6_Slider_GetTickStyle(vb6_hwnd_sld3',
        'vb6_Slider_GetNumTicks(vb6_hwnd_sld3',
        'vb6_Slider_ChannelTop(vb6_hwnd_sld3',
        # 容器里那枚滑杆的创建参数: 1342177280(BASE, 一位不挂) + 0x5 = TBS_AUTOTICKS|TBS_TOP
        '1342242821L, 0L,',
        # C29-SL-i: 三条 Sel 面各钉一枚形状，外加设计期那一对折出来的 Init 参数
        # （BASE 上那一条第 9 参是 -999 —— SelLength 整条被丢，读回来是空区段）。
        'vb6_Slider_ClearSel((void*)vb6_hwnd_sld4)',
        'vb6_Slider_SetSelLength(vb6_hwnd_sld4, 10)',
        'vb6_CStrLong(vb6_Slider_GetSelLength(vb6_hwnd_sld4',
        'vb6_Slider_Init((void*)vb6_hwnd_sld6, -999, -999, -999, 1L, -999, -999, 25L, 40L, -1L);',
        # SJ-j: 证人走登记过的那条 getter（不再被折成「装箱 + 转数值」那一趟）。
        'vb6_ToolTipRegistered(vb6_hwnd_sld3',
        # C29-SL-k: 派发里那一条**只发一次、不按控件展开**（RTL 里按类名筛掉 ScrollBar），
        # 再加 Text / TextPosition 的读写与一条证人。
        'vb6_Slider_BubbleNotify(scrollHwnd, scrollCode, HIWORD(wParam));',
        'vb6_Slider_SetText(vb6_hwnd_sld3',
        'vb6_Slider_SetTextPosition(vb6_hwnd_sld3, 1);',
        'vb6_Slider_BubbleVisible(vb6_hwnd_sld3',
        'vb6_Slider_BubbleText(vb6_hwnd_sld3',
        # C29-SL-l: 两条码头共用一张表 ⇒ 带括号与不带括号都发同一条专桩（裸形多一条注释）。
        'vb6_SetControlFocus((void*)vb6_hwnd_sld5);  /* setfocus */',
        'vb6_SetControlFocus((void*)vb6_hwnd_sld6);',
        'vb6_Slider_ClearSel((void*)vb6_hwnd_sld4);  /* clearsel */',
        # C29-SL-m: With 那一帧的接收者就是入口那枚 HWND（两形都走 pendingChainObj_ 协议）。
        'vb6_SetControlFocus((void*)_vb6_with_0);',
        'vb6_Slider_ClearSel((void*)_vb6_with_2);',
        # C29-SL-o（账 #83 的 Slider 半边）: 容器子控件那条创建路以前**压根不发这一条**
        # ⇒ 这里没有可禁的反面形状（BASE 的 emit 里 sld7 只有 CreateControl 那一行），
        # 红的是这一条正向针 + 三条真跑读数（A/B 实测见 029 那格）。第 9 参是 35L：
        # .frm 只写 SelStart=20 + SelLength=15，终点由同一处出口折出来。
        'vb6_Slider_Init((void*)vb6_hwnd_sld7, 10L, 100L, 42L, 2L, 8L, 5L, 20L, 35L, -1L);',
        # C29-SL-p: 写侧形状没动（发码面看不出这次改的是**读侧自存**），钉这一条是防有人把
        # 请求值又改成按窗口反算的那一支；证人那条读侧必须是登记过的专桩。
        'vb6_SetControlFontSize(vb6_hwnd_txtF, 26);',
        'vb6_ControlFontPixelHeight(vb6_hwnd_txtF',
        # C29-SL-q（账 #154）: 设计期字体三条形状 —— 顶层两条（Size + Name）、容器那一条（Size）。
        # Name 走新的 W 支（`.frm` 里的字体名在生成码里是 C 字面量，不是 BSTR，喂给上面那支
        # 会让 SysStringLen 去量字面量之后的内存）；发序是先 Size 后 Name（Name 那支在窗口没有
        # 字体时先造 -13 的底子，反过来会把字号盖掉）。
        'vb6_SetControlFontSize((void*)vb6_hwnd_txtH, 20.000000f);  /* design FontSize */',
        'vb6_SetControlFontNameW((void*)vb6_hwnd_txtH, L"Arial");  /* design FontName */',
        'vb6_SetControlFontSize((void*)vb6_hwnd_txtI, 20.000000f);  /* design FontSize */',
        # 读侧 String 档：与 ToolTipText / Tag 同一形状（直接进 Concat，不再装箱）。
        'vb6_BSTR_Concat(vb6_BSTR_FromStr(L"SQ1-dt="), vb6_GetControlFontName(vb6_hwnd_txtH)'
    )
    # C29-SL-c: 事件派发那一段的形状。钉的是"认来源的那枚句柄 + 分流那一档"这一整对 ——
    # Slider 与 ScrollBar 共用一扇 case WM_HSCROLL/WM_VSCROLL 的门，写错句柄比对比就是静默不派发。
    Test-EmitcShape "sl_emitc_events" @("$Tests\ctrlslider\SlidApp.vbp") @(
        # C29-SL-n（账 #141）: 「装不装」三份表共用一处判据之后, 只挂 `_DblClick` 的控件
        # 也会被子类化并在退出时拆掉（以前 arm 生成了却没人 install = 处理器永不触发）。
        # 两形各钉一条: 带 Cancel 的按声明签名发 `fn(&vb6_dblcancel)`, 无参声明仍发 `fn()`。
        'vb6_InstallControlSubclass((void*)vb6_hwnd_sld8, (void*)vb6_ctrl_subproc_sld8);',
        'vb6_RemoveControlSubclass((void*)vb6_hwnd_sld8);',
        'int16_t vb6_dblcancel = 0; extern void vb6_sld8_DblClick(int16_t*); vb6_sld8_DblClick(&vb6_dblcancel);',
        'extern void vb6_sld6_DblClick(); vb6_sld6_DblClick();',
        'case WM_HSCROLL:',
        'if (scrollHwnd == (void*)vb6_hwnd_sld1) {',
        'if (scrollHwnd == (void*)vb6_hwnd_sld2) {',   # 竖杆接的是同一扇门（发的是 WM_VSCROLL）
        'vb6_Slider_FireChange((void*)vb6_hwnd_sld1)',
        'if (scrollCode == 4 || scrollCode == 5)',
        'vb6_Slider_SimNotify((void*)vb6_hwnd_sld1, 5, 40)',
        'vb6_Slider_SimNotify((void*)vb6_hwnd_sld2, 5, 30)',
        # C29-SL-d: 常规事件那三档 arm + 「只有 Click 处理器也照样装子类化」那一步。
        # 装不装是第二份判据（frame_menu 那一趟），原先漏了 _Click ⇒ arm 发出来没人送消息。
        'if (msg == WM_LBUTTONUP) { extern void vb6_sld6_Click(); vb6_sld6_Click(); }',
        'if (msg == WM_LBUTTONDBLCLK) { extern void vb6_sld6_DblClick(); vb6_sld6_DblClick(); }',
        'if (msg == WM_KEYDOWN) {',
        'vb6_InstallControlSubclass((void*)vb6_hwnd_sld6, (void*)vb6_ctrl_subproc_sld6);',
        'vb6_Slider_SimStdEvent((void*)vb6_hwnd_sld6, 2, 37)',
        'if (msg == WM_SETFOCUS) { extern void vb6_sld6_GotFocus(); vb6_sld6_GotFocus(); }',
        'if (msg == WM_KILLFOCUS) { extern void vb6_sld6_LostFocus(); vb6_sld6_LostFocus(); }',
        # sld3 只挂着 GotFocus 一条处理器 —— 装与拆那两份表也得认它（同一句判据有三份）
        'if (msg == WM_SETFOCUS) { extern void vb6_sld3_GotFocus(); vb6_sld3_GotFocus(); }',
        'vb6_InstallControlSubclass((void*)vb6_hwnd_sld3, (void*)vb6_ctrl_subproc_sld3);',
        'vb6_RemoveControlSubclass((void*)vb6_hwnd_sld3);',
        'vb6_Slider_SimStdEvent((void*)vb6_hwnd_sld6, 4, 0)',
        'if (msg == WM_LBUTTONUP) { extern void vb6_sld2_Click(); vb6_sld2_Click(); }',
        'vb6_InstallControlSubclass((void*)vb6_hwnd_sld2, (void*)vb6_ctrl_subproc_sld2);'
    )
    Test-EmitcAbsent "sl_emitc_no_com_fallback" @("$Tests\ctrlslider\SlidApp.vbp") @(
        'vb6_ComGetProp(vb6_hwnd_sld1',      # 兜底那条"拿 HWND 当 IDispatch"不许回来
        'vb6_ComSetProp(vb6_hwnd_sld1',
        'vb6_ComGetProp(vb6_hwnd_sld4',      # 值面同样不许回来
        # SL-c: 判据方法一旦被 axSlotObj 那支先吃掉，就编成「取 SimNotify 属性 + Item 下标」
        # —— 编得过、跑起来什么都不发（C29-8c 实测踩过）。
        'vb6_ComGetObjectProp(vb6_hwnd_sld1, L"SimNotify")',
        'vb6_ComCall(vb6_hwnd_sld1, L"SimNotify"',
        # C29-SL-d: 同一条坑的第二枚判据方法（SimStdEvent）。
        'vb6_ComGetObjectProp(vb6_hwnd_sld6, L"SimStdEvent")',
        # C29-SL-e: 改之前的两条形状（BASE 上逐条命中，所以这两条反向针能红）
        'vb6_CStr(vb6_VariantFromValue(vb6_GetToolTipText',
        'vb6_CStr(vb6_VariantFromValue(vb6_GetControlTag',
        'vb6_SetControlTag((void*)vb6_hwnd_sld1',   # 账 #142: Tag="" 不发
        'vb6_ComCall(vb6_hwnd_sld6, L"SimStdEvent"',
        # C29-SL-h: TickStyle / GetNumTicks / ChannelTop 三条登记之前全部落到那扇兜底门
        # —— 拿裸 HWND 当 IDispatch：编得过、跑得动、什么都不发生（BASE 实测逐条命中）。
        'vb6_ComGetProp(vb6_hwnd_sld3, L"TickStyle")',
        'vb6_ComSetProp(vb6_hwnd_sld3, L"TickStyle"',
        'vb6_ComGetIntProp(vb6_hwnd_sld3, L"GetNumTicks")',
        'vb6_ComGetIntProp(vb6_hwnd_sld3, L"ChannelTop")',
        # C29-SL-i: 没登记之前 SelLength 走属性兜底、ClearSel 走方法兜底 —— 三条在 BASE 上
        # 逐条命中，在 NEW 上归零。（第四条候选 vb6_ComGetObjectProp(..., L"ClearSel") 在
        # BASE 上本来就不出现，方法那一形落的是 vb6_ComCall，所以不拿来当针。）
        'vb6_ComGetProp(vb6_hwnd_sld4, L"SelLength")',
        'vb6_ComSetProp(vb6_hwnd_sld4, L"SelLength"',
        'vb6_ComCall(vb6_hwnd_sld4, L"ClearSel"',
        'vb6_ComGetProp(vb6_hwnd_sld3, L"ToolTipRegistered")',
        'vb6_ComGetProp(vb6_hwnd_sld7, L"TickStyle")',
        # C29-SL-k: Text / TextPosition / 三条证人在登记之前全落那扇兜底门（BASE 逐条命中）。
        # 注意 Text 那支在 BASE 上回的是**字符串属性读**，串照样拿得到 —— 所以红的是气泡
        # 那三条，不是 SK1。
        'vb6_ComGetStringProp(vb6_hwnd_sld3, L"Text")',
        'vb6_ComSetProp(vb6_hwnd_sld3, L"Text"',
        'vb6_ComGetProp(vb6_hwnd_sld3, L"BubbleVisible")',
        'vb6_ComGetStringProp(vb6_hwnd_sld1, L"BubbleText")',
        # C29-SL-l: SetFocus 两形在登记之前**都**落这扇兜底门（BASE 逐条命中，NEW 归零）。
        # 不带括号那一形落的是语句路那条，带括号那一形落的是表达式路那条 —— 只接一头
        # 就会留一条「编得过、跑了、什么都没发生」，本线踩过三次。
        'vb6_ComCall(vb6_hwnd_sld5, L"SetFocus"',
        'vb6_ComCall(vb6_hwnd_sld6, L"SetFocus"',
        # C29-SL-q: `FontName` 归到 String 档之前的那支装箱 —— 打印面先装箱再 CStr，打出来是空串，
        # 而 `Len(...)` 那一面直接拿指针、答得对（"Consolas" 读出 8），所以这条旧坏一直没被看见。
        'vb6_CStr(vb6_VariantFromValue(vb6_GetControlFontName(',
        # C29-SL-p: 证人没登记之前，`txtF.FontPixelHeight` 落的就是这扇兜底门 —— 拿裸 HWND 当
        # IDispatch 问它要一个不存在的属性（BASE 实测逐字命中，NEW 归零 ⇒ 这条反向针能红）。
        'vb6_ComGetIntProp(vb6_hwnd_txtF, L"FontPixelHeight")',
        'CoCreateInstance'
    )

    # ai/029:429 + Fix 161d —— 源码面双保险: 生成串里不得再出现"指针返回函数被当数值"。
    # (这一条原先被放在 bas 段, 现挪到 vbp 段与其余 Test-Emitc* 同段 —— 它一直在 PASS,
    #  只是归属类别与惯例不一致; 挪动后由 vbp job 执行。)
    Test-EmitcAbsent "cp_emitc_no_ptr_as_num" @("$Tests\ctrlprop\PropApp.vbp") @(
        '(int32_t)(vb6_ComGetStringProp(',        # Not / 取负那条的 cast
        '(int32_t)(vb6_ComGetObjectProp(',
        '_vb6_select_0 = vb6_ComGetStringProp(',  # Select Case 的 int32_t temp 那条
        '_vb6_select_0 = vb6_ComGetObjectProp('
    )

    # Fix 161e: MsgBox 首参是**数值**时必须转字符串 (VB6 隐式转换: MsgBox 7 就显示 "7")。
    # 此前 cgen_expr_call_conv_output.inc 只包装了 COM 读 / Variant 两类形态，纯数值
    # (Len(...) 返回值 / 数值变量 / 字面量 / 算术表达式 / 浮点 / 布尔) 一条都没包
    # ⇒ int32_t(double/bool) 直接当 BSTR 指针传给 vb6_MsgBox1 ⇒ **弹窗是空的**
    # (地址 7、123 不可读; 真读得到时显示的也是垃圾)。修复: 未被认领的形态统一过
    # wrapToBSTR (cgen_expr_binary_util.cpp:81)，它按 inferExprType 分流且 String 原样返回。
    # ⚠ 判据走**源码面**而非运行时: MsgBox 是模态 MessageBoxW，无头环境会阻塞线程
    #   (见本文件上面那条注释「运行中弹 MsgBox 的用例需获得前台焦点」)。
    #   所以 tests/msgbox_num.bas 只喂 --emit-c，不登记为运行用例。
    Test-EmitcShape "mb_emitc_num_wrapped" @("$Tests\msgbox_num.bas") @(
        'vb6_MsgBox1(vb6_CStrLong(vb6_Len(',    # Len 返回值 (int32_t)
        'vb6_MsgBox1(vb6_CStrLong(n))',         # 数值变量
        'vb6_MsgBox1(vb6_CStrLong(123))',       # 数值字面量
        'vb6_MsgBox1(vb6_CStrLong((n + 1)))',   # 算术表达式
        'vb6_MsgBox1(vb6_CStrDbl(d))',          # 浮点
        'vb6_MsgBox1(vb6_CStrBool(b))',         # 布尔
        'vb6_MsgBox1(s)'                        # 字符串: 必须原样, 不得被误包
    )
    Test-EmitcAbsent "mb_emitc_no_bare_num" @("$Tests\msgbox_num.bas") @(
        'vb6_MsgBox1(vb6_Len(',    # 裸数值函数返回值
        'vb6_MsgBox1(n)',          # 裸数值变量
        'vb6_MsgBox1(123)',        # 裸字面量
        'vb6_MsgBox1((n + 1))',    # 裸算术表达式
        'vb6_MsgBox1(d)',          # 裸浮点
        'vb6_MsgBox1(b)'           # 裸布尔
    )

    # 账 #115: `Len(<裸标识符>)` 的"stor-宽度兜底桶"把**模块级 String**也吞了(knownBstrVars_ 每过程入口 clear，
    # 只有局部声明/形参进表) ⇒ 发成 (int32_t)sizeof(gS)。发码面钉两边：模块级 String 必须走 vb6_Len，
    # 而**数值型变量照旧走 sizeof**(Task #44 的 SSTabEx/ChooseColor 堆越界就是靠它救回来的，别把那一半改坏)。
    Test-EmitcShape "len_emitc_width" @("$Tests\test_len_width.bas") @(
        'vb6_CStrLong(vb6_Len(gS))',                    # 模块级 String：字符数
        'vb6_CStrLong(vb6_Len(gE))',                    # 模块级空串：0
        'vb6_CStrLong(vb6_Len(lS))',                    # 局部 String
        'vb6_CStrLong(((int32_t)sizeof(gL))))',         # Long：存储宽度
        'vb6_CStrLong(((int32_t)sizeof(gI))))',         # Integer：存储宽度
        'vb6_CStrLong(((int32_t)sizeof(gB))))'          # Byte：存储宽度(LEN-BYTE 那条崩溃的反证)
    )
    Test-EmitcAbsent "len_emitc_no_ptr_width" @("$Tests\test_len_width.bas") @(
        '((int32_t)sizeof(gS))',                        # 修复前的形状：x64 读 8、x86 读 4
        '((int32_t)sizeof(gE))'
    )

    # 账 #116: `Dim X As String * N` 的 typeRef 是 FixedStringTypeRef，而三处发码点
    # (cgen_decl_var.cpp / cgen_decl_func.cpp / cgen_decl_prop.cpp) 把它按 SimpleTypeRef 读 ->name
    # —— 读到的其实是节点里的 `ExprPtr length`，那个"长度"是个堆指针 ⇒ 拷字符串时张口要几十 GB
    # ⇒ std::bad_alloc 没人接 → abort()：退出码 3、零诊断，调试版 CRT 还弹一个前台模态框
    # (把跑夹具的人的会话整个卡住)。BASE 上这个文件**一行 C 都不发**，所以下面每条形状读数
    # 本身就是修复前的红点；再补一条 ICE 文案的缺席断言，防这一族以后换个形态回来。
    # 口径：本用例只钉"六个定长串落点都能正常发码"。VB6 那套空格补齐/截断语义还欠着
    # (模块级与 UDT 成员这两处发的仍是普通动态串)，所以刻意不做成运行用例。
    Test-EmitcShape "fixedstr_emitc_decl" @("$Tests\test_fixedstr_decl.bas") @(
        'BSTR gT = NULL;',                              # 模块级：崩溃就崩在这一行发不出来
        'BSTR vb6_F(void) {',                           # Function 返回类型那一处
        'void vb6_S(BSTR x);',                          # 形参那一处
        'BSTR f;',                                    # UDT 成员
        'BSTR lT = vb6_BSTR_FixedSTR(4);',              # 局部仍走补空格那条(唯一补的一条)
        'vb6_DebugWriteLong((int32_t)(vb6_Len(gT)));'   # 模块级 String 的 Len 照旧是字符数
    )
    Test-EmitcAbsent "fixedstr_emitc_no_ice" @("$Tests\test_fixedstr_decl.bas") @(
        '(ICE)',                                        # ICE 兜底文案: 出现即本批又崩了
        '((int32_t)sizeof(gT))'                         # 别把定长串的 Len 折成指针宽
    )
    # 账 #119: AssignMove 铺到局部/形参/ByRef 写穿三类目标 (以前只有模块级目标走这条)。
    # 这四行是"机制本体"的直接断言 —— 真跑那只 S119-leak-ok 只说结果, 这里钉的是发码形状:
    #   136 行那条循环体 (Left 的自有临时) / 定长串 String$ / ByRef 形参解引用目标
    #   / 借用侧 (变量赋给变量) 必须**留在**深拷贝 —— 它错了就是悬垂, 不是泄漏。
    Test-EmitcShape "strleak_emitc_move" @("$Tests\test_str_leak.bas") @(
        'vb6_BSTR_AssignMove(&s, vb6_Left(gBig, 64));',
        'vb6_BSTR_AssignMove(&t, vb6_String((*n), 67));',
        'vb6_BSTR_AssignMove(&(*o), vb6_MakeTag((&(int32_t){3})));',
        'vb6_BSTR_Assign(&d, s);'
    )
    Test-EmitcAbsent "strleak_emitc_no_move_borrowed" @("$Tests\test_str_leak.bas") @(
        'vb6_BSTR_AssignMove(&d, s);',                  # 借来的值绝不能被"移动"走
        'vb6_BSTR_AssignMove(&vb6_ret_MakeTag, t);'     # 返回变量同理 (t 是局部, 还要用)
    )
    # 账 #122: 赋值那一刻的隐式 CStr。四条钉"该折的折了" (字面量/变量/浮点/内置标量函数),
    # 三条钉"不该折的一条都没动": 返回串的工程函数、返回串的内置函数、以及 Fix 140 那条
    # Byte() 数组还原 —— 最后这条是本轮护栏抓出来的真回归 (Charts 2020 的 Caption 那一行
    # 一度被发成 vb6_CStrByte(vb6_ByteArrayToString(..))), 出现即复发。
    Test-EmitcShape "cnum_emitc_wrapped" @("$Tests\test_str_cnum_assign.bas") @(
        'vb6_BSTR_AssignMove(&s, vb6_CStrLong(123));',
        'vb6_BSTR_AssignMove(&s, vb6_CStrDbl(d));',
        'vb6_BSTR_AssignMove(&s, vb6_CStrBool(bo));',
        'vb6_BSTR_AssignMove(&s, vb6_CStrLong(vb6_Len(vb6_BSTR_FromStr(L"abcd"))));'
    )
    Test-EmitcAbsent "cnum_emitc_no_overwrap" @("$Tests\test_str_cnum_assign.bas") @(
        'vb6_CStrLong(vb6_MakeStr(',                    # 返回 String 的工程函数不许当数值折
        'vb6_CStrLong(vb6_Left(',                       # 返回 String 的内置函数同理
        'vb6_CStrByte(vb6_ByteArrayToString('           # Fix 140 那支转完就是字符串了
    )
    # 账 #123: 发码面钉"Byte 可见"。四条钉该出现的形状 (装箱带收窄、赋 String 走 CStrByte、
    # 比较走直接 C 比较而不是取地址), 三条钉那两种**修复前**的形状不许再出现 —— 尤其
    # vb6_VarCmpLongEq(&bt,…) 这一族: 助手签名是 (vb6_VARIANT*, int32_t), 拿 1 字节对象的
    # 地址当 16 字节 VARIANT 传, 它一出现就是"读邻居栈当 vt"的静默错答案回来了。
    Test-EmitcShape "byte_emitc_visible" @("$Tests\test_byte_visible.bas") @(
        'v = vb6_VariantByte((uint8_t)(bt));',
        'v = vb6_VariantByte((uint8_t)(gb));',
        'vb6_BSTR_AssignMove(&s, vb6_CStrByte(bt));',
        'vb6_CStrBool((-(bt == 65)))'
    )
    Test-EmitcAbsent "byte_emitc_no_variant_addr" @("$Tests\test_byte_visible.bas") @(
        'vb6_VarCmpLongEq(&bt,',                        # 局部 Byte 不许再被当 VARIANT 取地址
        'vb6_VarCmpLongEq(&b,',                         # 形参同 (B16 那形修复前就是这条)
        'vb6_VariantLong(bt)'                           # 装箱不许退成 VT_I4
    )
    # ai/029 C29-DT-a: DTPicker 换成原生 SysDateTimePick32（D6：不碰 MSCOMCT2.OCX，32 位进不了 x64）。
    # 改之前这枚控件走的是"第三方 OCX 按 COM 后期绑定"那一组 => 工程没引用类型库时连符号都查不到，
    # 属性读回空、写进去静默丢，而编译与退出码全都好看。24 条读数的分工：
    #   DT1-DT5  创建样式进窗口（三枚各一种设计期组合：长日期+CheckBox / 全默认 / 时间+UpDown）
    #   DT6-DT8  控件侧读数 DTM_GETIDEALSIZE：长日期明显比时间宽，且三枚各自算各自的
    #   DT9-DT11 运行期切档时"样式位"与"控件自己算的宽度"**一起**跟着改（只动一边就是假绿）
    #   DT12-DT14 CustomFormat 与 dtpCustom 那一位（原生只有 DTM_SETFORMATW、没有 Get 对称项 ⇒ 串自存）
    #   DT15-DT21 下拉月历五色：逐格读回 + 改一格其余四格不动（序号撞车的负控）+ 两枚控件各自一格
    #   DT22-DT24 设计期 CustomFormat 那条字符串路（.frm 里带引号 → 去引号 → 当 C 字面量发）
    #   DT25-DT36（C29-DT-b）Date 值面：Value/MinDate/MaxDate 往返、改一端不动另一端、越界被控件拒绝
    #     （值保持原样，不是钳到边界）、以及未勾那一态（原生仍回填内部日期 ⇒ GetValue 认返回标志回 0）
    #   DT37-DT42（C29-DT-c）DTN_* 事件面：三条各派发一次（Change/DropDown/CloseUp 码值 -759/-754/-753，
    #     两两撞车的负控就是"发 DropDown 时 Change 计数不许动"）、Sim 从 dt2 发而 dt1 的 handler 不许动
    #     （hwndFrom 认来源）、以及**控件自己**发的通知也走同一条链（DT41 拨日期看得见增量、DT42 那条
    #     空转的 CheckBox 写一下都不许加）。计数一律比增量：Form_Load 半截就写过好几回值。
    # 两条量出来的口径（原文写在夹具头注释里）：① CheckBox / UpDown 只能在**创建时**给（子窗口
    # 的创建参数，事后写 GWL_STYLE 会被控件抹回 —— DT11 就是钉这条边界的针，做成真运行期切换时
    # 它必须翻红）；Format 切档运行期倒是有效（DT9/DT10）。② 光读 GWL_STYLE 会自洽地假绿（SDK 的
    # DTS_TIMEFORMAT=0x9 自带 bit0=UPDOWN），所以格式类判据一律配一条 DTM_GETIDEALSIZE 控件侧读数。
    # 本机读数：x64 与 x86 各 42/42（宽度 143/64/95/121 与事件计数两架构逐字相同）；
    # 拿 DT-a 之前的编译器（c298c_base_C3.exe，那时无原生 DTPicker）跑同一件夹具 = 37 红 / 5 绿，
    # 绿的正是五条"负向"读数 (DT11/DT21/DT33/DT35/DT42 —— 都成立在"什么都没发生"上)。
    # DT-c 还有一条更紧的负控：Sim 钩子接上前（同一件夹具、同一个新编译器，只是调用点少写
    # 一对括号 ⇒ 整条语句被当属性读丢掉），DT37-DT40 四条当场红、接上就绿。
    # ⚠ 判据方法的写法有讲究：`dt1.SimChange`（不带括号）在语义层是**属性读**，发码一条都不发；
    # 必须写 `dt1.SimChange()` 才走调用路。这条由上面那条针 (vb6_DTP_SimChange…) 钉住。
    $dtNeedles = @("CTRLDATETIME-DONE") + (1..42 | ForEach-Object { "DT$_=Y" })
    # 128-a: CheckBox 的 CStr / TypeName / VarType 三个面 (VB6 要 True/Boolean/11)
    $dtNeedles += @("DT43=True/Boolean/11")
    Test-Vbp "ctrldatetime" "$Tests\ctrldatetime\DtfApp.vbp" $dtNeedles
    Test-Vbp "ctrldatetime_x86" "$Tests\ctrldatetime\DtfApp.vbp" $dtNeedles -Arch "x86"
    # 发码面两面都钉：创建样式位逐枚钉（1409286150 = 长日期+复选框；1409286153 = 时间位+UpDown 位，
    # 合起来恰好就是 SDK 的 DTS_TIMEFORMAT 0x9 —— 那条撞车在发码里留个可见的痕迹），
    # 反面断这枚控件的属性不许再走 COM 兜底、工程里不许再出现 CoCreateInstance。
    Test-EmitcShape "dt_emitc_shape" @("$Tests\ctrldatetime\DtfApp.vbp") @(
        '"SysDateTimePick32", "",',
        '1409351686L, 0L,',
        '1409351689L, 0L,',
        'vb6_DTP_Init((void*)vb6_hwnd_dt4, L"yyyy-MM-dd HH:mm");',
        'vb6_DTP_SetCheckBox(vb6_hwnd_dt2, (-1));',
        'vb6_DTP_SetCustomFormat(vb6_hwnd_dt2, vb6_BSTR_FromStr(L"yyyy-MM-dd"));',
        # C29-DT-b：Date 走的是**裸 double 变量**（`Dim d As Date` 发成 `double d`），
        # 不经过任何装箱 —— 这条针就是别让值面哪天退回 VARIANT 形状而没人察觉。
        'vb6_DTP_SetValue(vb6_hwnd_dt2, dReq);',
        'vb6_DTP_GetValue(vb6_hwnd_dt2',
        'vb6_DTP_SetHasDate(vb6_hwnd_dt1, 0);',
        # C29-DT-c：派发那三分支的形状（码值撞在同一个负数段里，写错一位就静默不派发，
        # 所以三条各钉一条，且钉的是"码值 + 认来源的那枚句柄"这一整对）。
        'pNM42->code == -759 && (void*)pNM42->hwndFrom == vb6_hwnd_dt1',
        'pNM42->code == -754 && (void*)pNM42->hwndFrom == vb6_hwnd_dt1',
        'pNM42->code == -753 && (void*)pNM42->hwndFrom == vb6_hwnd_dt1',
        'pNM42->code == -759 && (void*)pNM42->hwndFrom == vb6_hwnd_dt2',
        'vb6_DTP_SimChange((void*)vb6_hwnd_dt1)',
        'vb6_DTP_SimCloseUp((void*)vb6_hwnd_dt1)',
        # 处理器调用名的解析也钉一条：形参表必须是空的（VB6 这三条都没有参数），
        # 写错成带参就会在链接期 LNK2019、而编 C 阶段看不出任何异常。
        'extern void vb6_dt1_Change();',
        'vb6_CStrBool(vb6_DTP_GetCheckBox('
    )
    Test-EmitcAbsent "dt_emitc_no_com_fallback" @("$Tests\ctrldatetime\DtfApp.vbp") @(
        'vb6_ComGetObjectProp(vb6_hwnd_dt1',
        'vb6_ComSetObjectProp(vb6_hwnd_dt1',
        # DT-b 的 Value 也不许再走 COM 兜底（那正是它改之前整枚控件的默认下场）
        'vb6_ComGetObjectProp(vb6_hwnd_dt2, L"Value")',
        # DT-c 的三条判据方法一旦被下面那条 axSlotObj 分支先吃掉，就会编成
        # 「取 SimChange 属性 + Item 下标」—— 编得过、跑起来什么都不发（C29-8c 实测踩过）。
        'vb6_ComGetObjectProp(vb6_hwnd_dt1, L"SimChange")',
        'CoCreateInstance',
        'vb6_CStrLong(vb6_DTP_GetCheckBox('
    )
    # ai/029 C29-MV-a: MonthView 换成原生 SysMonthCal32（D6：不碰 MSCOMCT2.OCX，32 位进不了 x64）。
    # 22 条读数的分工：
    #   MV1-MV4   创建样式进窗口（四枚各一种设计期组合：MultiSelect+MaxSelCount+多月 / 全默认 /
    #             周号 / 不要今天），mv2 那条同时是"没写的不被继承"的对照
    #   MV5-MV6   多月平铺问**控件自己**：MCM_GETCALENDARCOUNT 说 mv1 眼下画了 2 个月、mv2 画 1 个
    #   MV7-MV10  MCM_GETMINREQRECT 的控件侧尺寸：周号加宽、今天那一行加高、多月不改单月的最小尺寸
    #   MV11-MV17 六色逐格读写 + 改一格其余五格不动（MCSC_ 序号撞车的负控）+ 两枚控件各自一格
    #   MV18-MV20 MaxSelCount 真往返过控件；**没挂 MCS_MULTISELECT 的那枚写不进去**（写完读回还是
    #             原生默认的 1）—— 这一条比 GWL_STYLE 硬，问的是控件按没按那位办事
    #   MV21      运行期改 ShowToday：样式位落得下**且**控件的最小尺寸跟着变（与 DTPicker 那两位
    #             被抹回去相反，MV 这一族运行期是有效的 —— 两枚控件的边界各量各的，不互相外推）
    #   MV22      通用属性面 + 两枚不串台
    #   MV23-MV31（C29-MV-b）Date 值面。两条量出来的硬边界决定了这一批的形状：
    #     ① **两张表互斥** —— 没挂 MCS_MULTISELECT 时 MCM_GET/SETCURSEL 有效而 GET/SELRANGE
    #        一律失败，挂了正好反过来 ⇒ MV28 钉非多选那侧读不出范围、MV31 钉多选那侧读不出
    #        Value（两条合起来才是这条互斥，单钉一条会放过"只实现了一半"的改动）；
    #     ② SETSELRANGE 收的是**闭区间**、控件内部存成半开、GETSELRANGE 原样吐内部值 ⇒
    #        不折回来止端恒比写入值多一天（GetSelEnd 折一天；MV26/MV27 钉往返与"改一端不动
    #        另一端"）；MV32/MV33 钉 MaxSelCount 真夹得住范围、而且夹的是"往里扩"那头、
    #        不挤掉已经选好的那端。
    #     判据一律以"本月 1 号"为基准推日期，不写死 —— 原生只让在**当前显示的那一个月**里选，
    #     写死就会在 CI 的未知日期上于月初/月末随机红（同一条纪律见 022 账 #79）。
    #   MV33-MV37（C29-MV-c）DateClick（原生 MCN_SELCHANGE = -749）：MV34 走完整派发链**且参数
    #     就是负载里那一天**（这条同时是 ABI 针 —— ByVal 参数按值传，写成指针形状时 x64 侥幸对、
    #     x86 当场错值，实测踩过）；MV35 认来源（从 mv2 发的不许叫 mv1）；MV36 钉"程序化改选
    #     不叫 DateClick"（原生只由用户交互驱动，与 DT-c 的 Change 同型口径）；MV37 钉没写
    #     处理器的那枚控件发了通知也不许把别人的 handler 顺带叫起来。
    # 本机读数：x64 与 x86 各 37/37；原始读数两架构逐字相同
    # （`R=1/1 218/242/178/159` = 默认上限/写后仍、单月最小尺寸 218×178、带周号 242 宽、
    #   去掉今天那一行 159 高；`E=2/1/46271` = 两条事件计数 + 最后收到的日期序列）。
    # 负控 = 拿 DT/MV 之前的二进制（.build/c298c_base_C3.exe）跑同一件夹具 = 32 红 / 5 绿。
    # 留绿的五条（MV4/MV9/MV17/MV28/MV31）全是"应当为 0 / 应当相等"那类**边界针** ——
    # 什么都不实现的空控件也满足它们，所以这几条不承担"验货"，只承担"别把边界改回去"；
    # 真正盘货的是另外 28 条。（记下来是免得下一个人把"BASE 有 5 绿"读成判据松。）
    $mvNeedles = @("CTRLMONTHVIEW-DONE") + (1..40 | ForEach-Object { "MV$_=Y" })
    # 128-a: MultiSelect / ShowToday 的 CStr 与 TypeName (ShowToday 此时已被 MV21 打开)
    $mvNeedles += @("MV41=True/True/Boolean")
    Test-Vbp "ctrlmonthview" "$Tests\ctrlmonthview\MvfApp.vbp" $mvNeedles
    Test-Vbp "ctrlmonthview_x86" "$Tests\ctrlmonthview\MvfApp.vbp" $mvNeedles -Arch "x86"
    # 发码两面都钉：类名 + 四条创建样式位逐枚钉（1409286146 = 基+MULTISELECT / 1409286148 = 基+
    # WEEKNUMBERS / 1409286160 = 基+NOTODAY，注意 ShowToday 与原生那位是**反**的：.frm 写 False
    # 才挂上去），设计期 Init 连多月与 MaxSelCount 一起钉；反面断这枚控件不许再走 COM 兜底。
    Test-EmitcShape "mv_emitc_shape" @("$Tests\ctrlmonthview\MvfApp.vbp") @(
        '"SysMonthCal32", "",',
        '1409351682L, 0L,',
        '1409351684L, 0L,',
        '1409351696L, 0L,',
        'vb6_MV_Init((void*)vb6_hwnd_mv1, 1, 2, 7);',
        'vb6_MV_Init((void*)vb6_hwnd_mv2, 1, 1, -999);',
        'vb6_MV_SetMaxSelCount(vb6_hwnd_mv1, 3);',
        'vb6_MV_SetShowToday(vb6_hwnd_mv4, (-1));',
        'vb6_MV_GetMonthCount(vb6_hwnd_mv2',
        # MV-b：Date 走**裸 double / 裸算术式**，一个装箱都不过（同 DT-b 那条纪律）。
        'vb6_MV_SetValue(vb6_hwnd_mv2, 44562.75);',
        'vb6_MV_SetSelStart(vb6_hwnd_mv1, (d0 + 1))',
        'vb6_MV_GetSelEnd(vb6_hwnd_mv1',
        # MV-c：派发那一条钉"码值 + 认来源的那枚句柄"这一整对，处理器签名钉**按值收 Date**
        # （ByVal 的 ABI；写成指针形状时 x64 侥幸能跑、x86 错值，所以两头都得钉）。
        'pNM42->code == -749 && (void*)pNM42->hwndFrom == vb6_hwnd_mv1',
        'extern void vb6_mv1_DateClick(double);',
        'vb6_mv1_DateClick(vb6_MV_NotifyDate((void*)lParam));',
        'vb6_MV_SimDateClick((void*)vb6_hwnd_mv1, (d0 + 4))',
        'vb6_CStrBool(vb6_MV_GetMultiSelect('
    )
    Test-EmitcAbsent "mv_emitc_no_com_fallback" @("$Tests\ctrlmonthview\MvfApp.vbp") @(
        'vb6_ComGetObjectProp(vb6_hwnd_mv1',
        'vb6_ComSetObjectProp(vb6_hwnd_mv2',
        'vb6_ComGetObjectProp(vb6_hwnd_mv3, L"MaxSelCount")',
        # 值面那三格也不许退回 COM 兜底
        'vb6_ComSetObjectProp(vb6_hwnd_mv2, L"Value"',
        'vb6_ComGetObjectProp(vb6_hwnd_mv1, L"SelEnd")',
        # MV-c：判据方法一旦被 axSlotObj 那条分支先吃掉，就编成「取 SimDateClick 属性 +
        # Item 下标」—— 编得过、跑起来什么都不发（C29-8c / DT-c 各踩过一次）。
        'vb6_ComGetObjectProp(vb6_hwnd_mv1, L"SimDateClick")',
        'CoCreateInstance',
        'vb6_CStrLong(vb6_MV_GetMultiSelect('
    )
    # ai/029 C29-RT-a: RichTextBox 换成原生 Msftedit.dll 的 RICHEDIT50W（D6：不碰 RICHTX32.OCX）。
    # 改之前这枚控件同样**连窗口都没有**（controlTypeToWin32Class 缺格）⇒ 属性读全靠"什么都不写
    # 也满足"的那副样子：负控（.build/c298c_base_C3.exe，DT/MV 之前的二进制）跑同一件夹具
    # = **22 红 / 10 绿**，留绿的十条（RT2/4/6/7/12/15/16/21/26/27）全是"读回 0 / 空 / 反向判断"
    # 那一类；其中 rt1.VScrollRange 在 BASE 下**连数都没打出来**（原始行 `S=0/0//0/0` 那个空位
    # = 账 #82 那条既有缺陷：未登记的控件属性按数值读会漏裸指针）。真跑：x64 与 x86 各 32/32。
    #
    # 三条量出来的口径决定了这一批为什么长这样（三轮 C 探针 `.build/rtprobe3.c`，读数在 029 §九）：
    #   ① **滚动条归控件管，样式位会被它自己抹掉** —— 创建时给了 WS_VSCROLL，空文本下 GWL_STYLE
    #      那两位就没了（不需要滚动就把 bar 连样式位一起拆），灌 60 行才自己回来。所以 cgen 一并挂
    #      ES_DISABLENOSCROLL(0x2000) 让位稳定；而**事后** SetWindowLong 加那两位只有外观、量程停在
    #      默认 0..100（内容真高 0..1281）⇒ 这四位刻意**没有写口**，判据除样式位还配一条控件自己的
    #      读数 VScrollRange/HScrollRange（RT11-RT15：挂 bar 的两枚随内容长过 100、没挂的那枚不动）。
    #   ② ReadOnly 走 EM_SETREADONLY 事后有效 —— 与 ① 正相反，同一枚控件里两种属性各有各的
    #      "事后行不行"，不许互相外推。
    #   ③ WordWrap 的原生方向与网上那段经典片段**相反**：EM_SETTARGETDEVICE 的 lParam 才是目标 DC，
    #      NULL = 折到本窗客户区宽（= VB6 的 True），传 GetDC(本窗) = 目标宽变成整屏（不折）。
    #      原生从没调过这条时是"不折"，与 VB6 默认相反 ⇒ 设计期没写也要显式下发一次 True。
    #      EM_SET/GETWRAPMODE 在这枚上问不出也设不动（mode 恒读 0）⇒ 只能自存窗口属性。
    # 顺带修了一条同族既有缺陷：vb6_Get/SetBorderStyle 只认 "Edit" 类，RICHEDIT50W 落到"存属性"
    # 那条兜底分支，而 SetPropW(0) 等于**删属性**（GetPropW 回 NULL ⇒ 恒读默认 1）⇒ BorderStyle=None
    # 永远设不上（RT5/RT6 就是它的正负两面）。
    # ---- C29-RT-b（同一件夹具往后接 RT33..RT58）= Sel* 的格式面 ----
    # 三态问法（第三/四轮探针）：EM_GETCHARFORMAT / EM_GETPARAFORMAT 把"选区内不一致"那一位从返回的
    # dwMask 里清掉（只问 italic：全一致 0xFFFFFFFF、跨界 0xFFFFFFFD；字符面连返回值都等于那张掩码）。
    # 本项目没有 Null 可回 ⇒ 混合一律按"没有"那一头（False / 0 / 空串）—— 与 VB6 教的 `= True` 等价。
    # 两条折算也是这格量出来的：字号原生单位 1/20 磅（RT46 读回 14）、对齐 VB6 0左/1中/2右 对原生
    # 1/3/2（不折算就会把"居中"读成"右对齐"，RT49/RT50 钉住）；悬挂缩进用 PFM_OFFSET 取负，
    # PFM_OFFSETINDENT 会把整段推走（实测 720 → 960），不是悬挂。
    # BASE（本批之前的编译器）同一件夹具 = 20 绿 / 38 红：RT33-RT58 里 16 条当场红，剩下 10 条是
    # "应当为 0 / 应当相等"那类反向针（什么都不实现也满足它们）—— 与 DT/MV 每次的分布同型。
    $rtNeedles = @("CTRLRICHTEXT-DONE") + (1..89 | ForEach-Object { "RT$_=Y" })
    # 128-a: ReadOnly / WordWrap 的 CStr 与 TypeName (rt1 设计期就没开换行)
    $rtNeedles += @("RT90=True/False/Boolean")
    # 账 #161: 只动选区不该发 Change (RT91) / 一次赋值恰好一条 (RT92)
    $rtNeedles += @("RT91=Y", "RT92=Y")
    Test-Vbp "ctrlrichtextbox" "$Tests\ctrlrichtextbox\RtfApp.vbp" $rtNeedles
    Test-Vbp "ctrlrichtextbox_x86" "$Tests\ctrlrichtextbox\RtfApp.vbp" $rtNeedles -Arch "x86"
    # 发码正面：类名 + 四位创建样式逐枚钉（1409286148 = 基+ES_MULTILINE，rt2 全默认；
    # 1411391492 = 基+VSCROLL+DISABLENOSCROLL / 1410342916 = 基+HSCROLL+DISABLENOSCROLL /
    # 1412440068 = 基+两个+DISABLENOSCROLL），设计期 Init 连 WordWrap/ReadOnly 一起钉（-999 = 没写），
    # 初值文本钉"句柄赋值之后才发"那一条，选区/上限/量程读数钉走的是原生 getter 而不是 COM 兜底。
    Test-EmitcShape "rt_emitc_shape" @("$Tests\ctrlrichtextbox\RtfApp.vbp") @(
        '"RICHEDIT50W", "",',
        '1412505604L, 0L,',
        '1409351684L, 0L,',
        '1411457028L, 0L,',
        '1410408452L, 0L,',
        'vb6_RTB_Init((void*)vb6_hwnd_rt1, 0, -999);',
        'vb6_RTB_Init((void*)vb6_hwnd_rt2, -999, -999);',
        'vb6_RTB_Init((void*)vb6_hwnd_rt3, -999, 1);',
        'vb6_SetBorderStyle((void*)vb6_hwnd_rt1, 1);',
        'SendMessageW((HWND)vb6_hwnd_rt1, WM_SETTEXT, 0, (LPARAM)L"DesignText-Alpha");',
        'vb6_RTB_SetSelText(vb6_hwnd_rt2, vb6_BSTR_FromStr(L"XY"));',
        'vb6_RTB_SetWordWrap(vb6_hwnd_rt3, (-1));',
        'vb6_RTB_SetMaxLength(vb6_hwnd_rt2, 20);',
        'vb6_RTB_GetVScrollRange(vb6_hwnd_rt3',
        # C29-RT-c: TextRTF 读写两面 + 三条方法的发码形状（含两个可选实参缺省填什么）
        'vb6_RTB_GetTextRTF(vb6_hwnd_rt4)',
        'vb6_RTB_SetTextRTF(vb6_hwnd_rt3, sRtf);',
        'vb6_RTB_SaveFile((void*)vb6_hwnd_rt4, sPath, 0);',
        'vb6_RTB_LoadFile((void*)vb6_hwnd_rt3, sPath, 0);',
        'vb6_RTB_SaveFile((void*)vb6_hwnd_rt4, sPath2, 1);',
        'vb6_RTB_LoadFile((void*)vb6_hwnd_rt3, sPath2, 1);',
        'vb6_RTB_Find((void*)vb6_hwnd_rt4, vb6_BSTR_FromStr(L"alph"), 0, (-1), 1)',
        'vb6_RTB_Find((void*)vb6_hwnd_rt4, vb6_BSTR_FromStr(L"nope-not-here"), -1, -1, 0)',
        'vb6_RTB_GetHScrollRange(vb6_hwnd_rt4',
        # RT-b：四条效果走 CHARFORMAT2W 的同一族 setter（布尔按 VB6 的 -1/0 发），
        # 颜色/字体名/字号/对齐/三缩进各一条 —— 全部钉"裸调用 + 裸数值"，不许出现装箱。
        'vb6_RTB_SetSelBold(vb6_hwnd_rt2, (-1));',
        'vb6_RTB_SetSelUnderline(vb6_hwnd_rt2, (-1));',
        'vb6_RTB_SetSelStrikethru(vb6_hwnd_rt2, (-1));',
        'vb6_RTB_GetSelItalic(vb6_hwnd_rt2',
        'vb6_RTB_SetSelColor(vb6_hwnd_rt2, col);',
        'vb6_RTB_SetSelFontName(vb6_hwnd_rt2, vb6_BSTR_FromStr(L"Courier New"));',
        'vb6_RTB_SetSelFontSize(vb6_hwnd_rt2, 14);',
        'vb6_RTB_GetSelFontSize(vb6_hwnd_rt2',
        'vb6_RTB_SetSelAlignment(vb6_hwnd_rt2, 2);',
        'vb6_RTB_SetSelIndent(vb6_hwnd_rt2, 720);',
        'vb6_RTB_SetSelRightIndent(vb6_hwnd_rt2, 1440);',
        'vb6_RTB_SetSelHangingIndent(vb6_hwnd_rt2, 360);',
        'vb6_RTB_GetSelHangingIndent(vb6_hwnd_rt2',
        # C29-RT-d: 两条事件的派发形状。Change 走 WM_COMMAND/1024(EN_UPDATE，**不是** EDIT 那
        # 条 768)；SelChange 两条通道各一条 arm（WM_COMMAND/1815 与 WM_NOTIFY/1794），
        # 判据侧 SimNotify 的发码形状也钉死 —— 标记没被语句路消费掉的症状就是运行期一声不响。
        'if (id == 100 && code == 1024) {',
        'if (id == 102 && code == 1024) {',
        'if (id == 100 && code == 1815) {',
        'if (pNM42->code == 1794 && (void*)pNM42->hwndFrom == vb6_hwnd_rt1) {',
        'extern void vb6_rt1_Change(); vb6_rt1_Change();',
        'extern void vb6_rt3_SelChange(); vb6_rt3_SelChange();',
        '{ vb6_RTB_SimNotify((void*)vb6_hwnd_rt1, (int32_t)1794); }',
        'vb6_CStrBool(vb6_RTB_GetReadOnly('
    )
    # 反面：这枚控件不许再走 COM 后期绑定；而 ScrollBars 那四位**不许有写口** ——
    # 事后写只有外观、没有量程，发一条"写得动但什么都不改"的 setter 比不发更难查。
    Test-EmitcAbsent "rt_emitc_no_com_fallback" @("$Tests\ctrlrichtextbox\RtfApp.vbp") @(
        'vb6_ComGetObjectProp(vb6_hwnd_rt1',
        'vb6_ComSetObjectProp(vb6_hwnd_rt2',
        'vb6_ComGetObjectProp(vb6_hwnd_rt3, L"MaxLength")',

        # C29-RT-c 的反面：三条方法与 TextRTF 都不许再落回拿 HWND 当 IDispatch 那条假路
        # （打了标记却没在语句路消费掉的症状就是编得过、运行期一声不响 —— 本线踩过三次）。
        'vb6_ComCall(vb6_hwnd_rt4, L"Find"',
        'vb6_ComGetObjectProp(vb6_hwnd_rt4, L"TextRTF")',
        'vb6_ComSetObjectProp(vb6_hwnd_rt3, L"TextRTF"',
        'vb6_ComCall(vb6_hwnd_rt3, L"LoadFile"',
        # RT-d：SimNotify 也不许落回"把 HWND 当 IDispatch"那条假路
        'vb6_ComCall(vb6_hwnd_rt1, L"SimNotify"',
        'vb6_RTB_SetScrollBars',
        # RT-b：格式面也不许退回 COM 兜底（一条都不许）
        'vb6_ComGetObjectProp(vb6_hwnd_rt2, L"SelBold")',
        'vb6_ComSetObjectProp(vb6_hwnd_rt2, L"SelColor"',
        'vb6_ComSetObjectProp(vb6_hwnd_rt2, L"SelAlignment"',
        'CoCreateInstance',
        'vb6_CStrLong(vb6_RTB_GetReadOnly('
    )
    # ai/029 C29-5a: Toolbar 换成原生 ToolbarWindow32（D6：不碰 MSCOMCTL.OCX）。
    # 改之前这枚控件**连窗口都没有**：controlTypeToWin32Class 缺格，而且被"ImageList || Toolbar
    # 走 CoCreateInstance"那一组扣住 (直接 continue) => vb6_hwnd_tb1 压根不声明 —— 实测读一个
    # tb1.Visible 就是 `C2065: vb6_hwnd_tb1 未声明的标识符`，整件工程编不过。12 条读数：
    # 设计期三条按钮进了控件 (TB1，问的是原生 TB_BUTTONCOUNT 不是我那张表)、没写的不被继承 (TB2)、
    # 矩形按 .frm (TB3-TB4，证 CCS_NORESIZE|CCS_NOPARENTALIGN 那两位)、标量属性设计期与默认
    # (TB5-TB7)、运行期可逆赋值 (TB8-TB9)、两枚不串 (TB10)、通用属性面 (TB11-TB12)。
    # Visible 刻意不问: 无头跑里父窗从未 ShowWindow，任何子窗口的 Visible 读数都是假 (TreeView 那
    # 条 TV1 同因)，换成与父窗无关的 Enabled 才问得出东西。
    # 路上量到两条 v6 主题的坑，都写进了 RTL 注释: TB_RESET 之后 TB_ADDBUTTONSW 会返回 TRUE
    # 却一个按钮都不加 (所以只 append/insert)；C3 的产物嵌了 Common-Controls 6.0 的 manifest，
    # v6 工具栏**没被告知结构体尺寸就静默吞按钮** (不嵌 manifest 的独立 C 探针在 v5 下是好的)
    # => 发按钮前先 TB_BUTTONSTRUCTSIZE。
    # ai/029 C29-5b: 同一个工程再加 15 条 Buttons/Button 读数 (TB13-TB27)。集合是**真 IDispatch**
    # (vb6forms_memberobj.c 的 BUTTONS 族，与 8b 的 Nodes 同一条 cheapest route)。分界: Caption /
    # Image / Enabled / Visible / Value 现问控件 (TB_GET/SETBUTTONINFOW)，Key / Tag / ToolTipText /
    # Style / Width 住 5a 那张表 (原生 fsStyle 分不出「占位符」那一档，ToolTipText 的原生面要
    # TTN_GETDISPINFO)。TB15 那条尤其值钱: 它证的是**设计期 caption 真进了控件的字符串表**
    # (iString 往返)，5a 只数过按钮个数、没验过文字。
    # 路上量到一条 v6 主题坑: TB_ADDBUTTONSW **不吃调用方给的 fsState** (实测建完读回 0，
    # 连 TBSTATE_ENABLED 都没有) ⇒ Button.Enabled 的默认读数会是 False，与 VB6 相反；
    # 建完补一条 TB_SETBUTTONINFOW 把启用位打上去 (TB19 就是这条的读数)。
    # 负控两条 (红在哪几条是**量出来的**，不是推的): ① `tb1.Buttons.Remove "open"` 换成
    #   一个不存在的 Key => **只红 TB26** (TB27 常绿是对的: 后面 Clear 照样把两边清空)。
    #   ② .frm 里把 tb1 的 `TextStyle` 从 1 改成 0 => 红 **TB5、TB10、TB27** —— 前两条是
    #   5a 的设计期/默认对照，第三条正是"集合那条路没把标量属性面抢走"的读数。
    # C29-5c 加 TB28-TB34 (两条按钮事件)。**两条不在同一条通道上**: ButtonClick 走
    #   WM_COMMAND(id=控件的 idCommand, code=0, lParam=工具栏)，ButtonMenuClick 走
    #   WM_NOTIFY(TBN_DROPDOWN=-710, hdr.idFrom=同一个 id)，且只有 Style 5 那颗发得出 (TB33)。
    $tbNeedles = @("CTRLTOOLBAR-DONE") + (1..34 | ForEach-Object { "TB$_=Y" })
    Test-Vbp "ctrltoolbar" "$Tests\ctrltoolbar\TbApp.vbp" $tbNeedles
    Test-Vbp "ctrltoolbar_x86" "$Tests\ctrltoolbar\TbApp.vbp" $tbNeedles -Arch "x86"
    # ai/029 C29-M: 工程自带一份**不含清单**的 .res（ResFile32="no_manifest.res"，里面只有一条
    # 对话框模板）时，内置那份 comctl v6 清单必须照样进产物。旧判据只看"给没给 ResFile"，
    # 于是这种 VB6 里很常见的工程连内置的一起让掉 ⇒ 产物静默退回 v5.82（实测：BASE 编译器编
    # 同一件夹具，产物 0 份清单；改后 1 份且是内置那份）。清单数=1 同时挡住"两份 #1 打架"。
    $mfNeedles = @("CTRLMANIFEST-DONE", "CM1=Y", "CM2=Y", "CM3=Y")
    Test-Vbp "ctrlmanifest" "$Tests\ctrlmanifest\MfApp.vbp" $mfNeedles -ShardGroup "mfapp"
    # 主判据：这份工程的 .res 里没有清单 ⇒ 产物必须仍然带**内置那一份**（改前 BASE 编出来是 0 份）。
    # "恰好 1 份"同时挡住最坏的那种错：两份 #1 会让加载器直接报错。
    Test-ProductManifest "ctrlmanifest_builtin" "$OutDir\MfApp.exe" 1 "Microsoft.Windows.Common-Controls" "dpiAware" -ShardGroup "mfapp"
    Test-Vbp "ctrlmanifest_x86" "$Tests\ctrlmanifest\MfApp.vbp" $mfNeedles -Arch "x86" -ShardGroup "mfapp"
    Test-ProductManifest "ctrlmanifest_builtin_x86" "$OutDir\MfApp.exe" 1 "Microsoft.Windows.Common-Controls" "dpiAware" -ShardGroup "mfapp"
    # 发码面: 设计期四条逐参数钉 (含 -999 哨兵那条没写过的控件)、创建样式那个常量、
    # 反面断这枚控件不再走 vb6_com_ 槽 / CoCreateInstance / Buttons 的 COM 兜底。
    Test-EmitcShape "tb_emitc_shape" @("$Tests\ctrltoolbar\TbApp.vbp") @(
        'vb6_Toolbar_Init((void*)vb6_hwnd_tb1, -999, 1, -999, 2);',
        'vb6_Toolbar_Init((void*)vb6_hwnd_tb2, -999, -999, -999, -999);',
        'vb6_Toolbar_AddButton((void*)vb6_hwnd_tb1, 2, NULL, L"", 3, -1, NULL, 8);',
        # C29-5b: 三条发码形状针 —— 集合对象本体 (真 IDispatch 的入口，宿主槽必须是
        # vb6_hwnd_ 而不是 vb6_com_)、按 Key 取下标、Button 属性写落到 COM 派发上。
        # 第二条还钉住「省略的实参要发成 Missing 而不是 0」—— 0 会被当成插到第 1 格前面。
        'vb6_ComCallObject(vb6_Toolbar_Buttons((void*)vb6_hwnd_tb1), L"Item", (void*[]){vb6_ComPackBSTR(vb6_BSTR_FromStr(L"save"))}, 1)',
        'vb6_ComCallObject(vb6_Toolbar_Buttons((void*)vb6_hwnd_tb1), L"Add", (void*[]){vb6_ComPackMissing(), vb6_ComPackBSTR(vb6_BSTR_FromStr(L"cut"))',
        'vb6_ComSetProp(b1, L"Enabled", vb6_ComPackBool(0))',
        # C29-5c: 两条按钮事件各钉一条发码形状 —— 第一条是 WM_COMMAND 那一路按 lParam 认
        # 来源、按 id 取按钮；第二条是 WM_NOTIFY 那一路的 TBN_DROPDOWN 分支；第三条钉 Sim
        # 走的是 RTL 直调 (不是"取属性 + Item"那条被吞成静默空转的路，见 8c 的教训)。
        'void* vb6_tbBtn5c = vb6_Toolbar_ButtonAt((void*)vb6_tbSrc5c, id);',
        'if (pNM42->code == -710 && (void*)pNM42->hwndFrom == vb6_hwnd_tb1) {',
        'vb6_Toolbar_SimButtonClick((void*)vb6_hwnd_tb1, 1);',
        '1409353996L, 0L,'
    )
    Test-EmitcAbsent "tb_emitc_no_ocx" @("$Tests\ctrltoolbar\TbApp.vbp") @(
        'vb6_com_tb1',
        'CoCreateInstance',
        'vb6_ComGetObjectProp(vb6_hwnd_tb1, L"Buttons")',
        'vb6_ComGetObjectProp(vb6_hwnd_tb1, L"SimButtonClick")'
    )
    # ai/028 V1 的另两个 R4 落点: 模块头 Attribute 的值与 CreateObject 的工程内 ProgID
    # 都写成反引号串 —— 前者折错则模块名对不上 .vbp, 后者折错则没有改写、运行期变查注册表。
    $rsProjNeedles = @("RP1=OK", "RP2=OK", "RP3=OK", "RP4=OK", "RP5=OK", "RP-DONE")
    Test-Vbp "rawstr_proj" "$Tests\rawstr_proj\RsApp.vbp" $rsProjNeedles
    Test-Vbp "rawstr_proj_x86" "$Tests\rawstr_proj\RsApp.vbp" $rsProjNeedles -Arch "x86"


    # P20-38: ProgressBar 复刻 (msctls_progress32, 不加载 mscomctl.ocx)。
    # PB11 盯 SetPropW 存 0 被当成"未设置"回落默认值的坑。
    # P20-43 修正: 期望串必须与夹具 Debug.Print 的**完整标签**逐字一致
    # (此前登记成缩写 "PB1=100", 夹具打的是 "PB1-MAX=100" -> GA 报 output mismatch,
    #  内容其实全对 —— 纯粹是注册表与夹具漂移)。
    Test-Vbp "ctrlprogress" "$Tests\ctrlprogress\CtrlProgress.vbp" @(
        "PB1-MAX=100", "PB2-MIN=0", "PB3-VALUE=0", "PB4-SCROLLSTD=1",
        "PB5-ORIENTH=0", "PB6-SET40=40", "PB7-MAX200=200", "PB8-D2MAX=10",
        "PB9-D2MIN=-10", "PB10-D2ORI=1", "PB11-D2SCR=0", "CTRLPROGRESS-DONE")

    # P20-39: ImageList 复刻 (comctl32 ImageList_* API, 不加载 mscomctl.ocx)。
    # 图片三路来源: ①设计期 .frx 裸 DIB ②运行期 LoadPicture (VB6 StdPicture = 活着的
    # IPicture) ③Remove/Clear 对混插集合的 key 表搬动。`ListImages("Key")` 按 Key 取项
    # 也在这里 (Item 实参是宽字符串不是下标)。
    # IL1~IL3 盯 .frx 记录布局: 16B GUID + magic + imgSize 必须让 readPicture 读到 808,
    # GUID 少写一个字节就会整条记录后移、静默丢掉两张图 (不报任何错)。
    # 三张 .bmp 是**运行期**由 LoadPicture 从 App.Path 读的, 必须拷进 $OutDir。
    Copy-Item "$Tests\ctrlimagelist\*.bmp" $OutDir -Force
    Test-Vbp "ctrlimagelist" "$Tests\ctrlimagelist\CtrlImageList.vbp" @(
        "IL1-DTCOUNT=2", "IL2-DTKEY1=dt1", "IL3-DTKEY2=dt2", "IL4-ADDRT=3",
        "IL5-COUNT=3", "IL6-AFTERRM=2", "IL7-KEY1=dt2", "IL8-AFTERRM2=1",
        "IL9-KEY1=rt", "IL10-COUNT=2", "IL11-KEY1=first", "IL12-BYKEY=first",
        "IL13-BYKEYIDX=1", "IL14-W=16", "IL15-SETW=32", "IL16-H=16",
        "IL17-AFTERCLR=0", "CTRLIMAGELIST-DONE")
    # C29-3 顺手补的覆盖: 这条以前**从来没有 x86 版本**。而 C29-3 撞出来的那个越界写
    # 恰恰只在 x86 暴露 (sizeof(vb6_VARIANT)=24 vs Windows VARIANT=16 → 写坏堆),
    # x64 因为两个尺寸恰好相等而"绿得可疑"。控件类的判据必须双架构。
    Test-Vbp "ctrlimagelist_x86" "$Tests\ctrlimagelist\CtrlImageList.vbp" @(
        "IL1-DTCOUNT=2", "IL2-DTKEY1=dt1", "IL3-DTKEY2=dt2", "IL4-ADDRT=3",
        "IL5-COUNT=3", "IL6-AFTERRM=2", "IL7-KEY1=dt2", "IL8-AFTERRM2=1",
        "IL9-KEY1=rt", "IL10-COUNT=2", "IL11-KEY1=first", "IL12-BYKEY=first",
        "IL13-BYKEYIDX=1", "IL14-W=16", "IL15-SETW=32", "IL16-H=16",
        "IL17-AFTERCLR=0", "CTRLIMAGELIST-DONE") -Arch "x86"

    # --- ai/029 C29-3: 控件"成员对象"机制立样 (ImageList 的 ListImages / ListImage) ---
    # 四条验收 (计划书原文): ① Set img = ListImages.Add(, "Open", LoadPicture(..))
    # ② img.Key ③ ListImages.Count ④ For Each。
    # 口径 = 计划书 D1: 集合与成员对象都是**真 IDispatch** (vb6forms_memberobj.c),
    # 所以 `As Object` 的晚绑定吃的是同一个对象 (MO4 专测这条), `For Each` 由
    # _NewEnum 走标准 IEnumVARIANT。老写法 `n = .Add(..)` 仍按 VB6 取默认属性 Index
    # (Let 侧 memObjLetScalar_ 转换) —— 旧夹具 ctrlimagelist 的 IL4/IL10 盯这条。
    # 三张 bmp 复用 ctrlimagelist 那批 (上面已 Copy-Item 进 $OutDir)。
    $c29imgExpected = @(
        "MO1-ADD-KEY=Open", "MO2-ADD-IDX=1", "MO3-COUNT=1",
        "MO4-KEY2=Close", "MO5-COUNT=2", "MO6-ITEM1-KEY=Open",
        "MO7-ITEMKEY-IDX=2", "MO8-FOREACH=Open,Close,",
        "MO9-AFTERRM=1", "MO10-FOREACH2=Close,", "MO11-AFTERCLEAR=0")
    Test-Vbp "c29imgobj" "$Tests\c29imagelistobj\C29ImgObj.vbp" $c29imgExpected
    Test-Vbp "c29imgobj_x86" "$Tests\c29imagelistobj\C29ImgObj.vbp" $c29imgExpected -Arch "x86"

    # --- ai/029 C29-4: StatusBar 事件面 (PanelClick/PanelDblClick) ---
    # 数据面 P20-40 已有 ctrlstatusbar 42 条; 这里盯事件: SimClick 是判据专用方法
    # (RTL 程序化发真 WM_NOTIFY, 走完整派发链), handler 内真读/真写 Panel 对象成员。
    $c29sbevtExpected = @(
        "SB1-COUNT=3",
        "EVT1-CLICK-INDEX=1", "EVT2-CLICK-TEXT=one", "EVT3-CLICK-KEY=p1", "EVT4-CLICK-AFTER=one!",
        "EVT1-CLICK-INDEX=3", "EVT2-CLICK-TEXT=three", "EVT3-CLICK-KEY=p3", "EVT4-CLICK-AFTER=three!",
        "EVT5-DBL-INDEX=2")
    Test-Vbp "c29sbevt" "$Tests\c29statusbarevt\SbEvent.vbp" $c29sbevtExpected
    Test-Vbp "c29sbevt_x86" "$Tests\c29statusbarevt\SbEvent.vbp" $c29sbevtExpected -Arch "x86"


    # --- ai/029 C29-7: ListView (数据面 + 事件面) ---
    # 数据面: ColumnHeaders.Add (标题) / ListItems.Add (数据) / SubItems(i) **1 基, 1 就是
    # 第 2 列** / 两个集合的 Count 与 For Each / 按 Key 与按下标取项 / 成员属性读写 /
    # ListView 自身的 View(3=报表) 与 GridLines。
    # 口径: 集合与成员对象都走 C29-3 立起来的**真 IDispatch** (vb6forms_memberobj.c),
    # 所以 `Set itm = .ListItems.Add(..)` 之后 itm.Text / itm.SubItems(1) / itm.Selected
    # 全走晚绑定; ListView 是**真窗口**, owner 是 vb6_hwnd_X (ImageList 那族是 vb6_com_X)。
    $c29lvExpected = @(
        "LV1-COL-KEY=c1", "LV2-COL-TEXT=姓名", "LV3-COL-IDX=1", "LV4-COLWIDTH=1200",
        "LV5-COLCOUNT=2", "LV6-ITEM-KEY=r1", "LV7-ITEM-TEXT=张三", "LV8-ITEM-IDX=1",
        "LV9-SUB1=销售部", "LV10-ITEMCOUNT=2", "LV11-FOREACH=张三/销售部,李四/技术部,",
        "LV12-COLS=姓名,部门,", "LV13-VIEW=3", "LV14-GRID=True", "LV15-BYKEY=李四",
        "LV16-ITEM1=张三", "LV17-SEL=-1", "LV18-COLW=900", "LV19-AFTERRM=1",
        "LV20-AFTERCLEAR=0", "LV23=True/Boolean/True", "LV24=EQ", "LV25=False/False")
    Test-Vbp "c29listview" "$Tests\c29listview\C29ListView.vbp" $c29lvExpected
    Test-Vbp "c29listview_x86" "$Tests\c29listview\C29ListView.vbp" $c29lvExpected -Arch "x86"
    # 事件接线: 无头环境点不了鼠标 (生成代码里 LV21/LV22 只有真点击才会打),
    # 但"WM_NOTIFY 分支 → OnNotify 换算 → 取成员对象 → 回调"这条链必须在**生成代码里
    # 看得见** —— 少了任何一环都是"接线了却没生效", 而运行期完全静默。
    Test-EmitcShape "lv_emitc_events" @("$Tests\c29listview\C29ListView.vbp") @(
        "pNM42->code == -114",
        "vb6_ListView_OnNotify((void*)vb6_hwnd_ListView1, -114",
        "vb6_ListView_ListItemAt((void*)vb6_hwnd_ListView1",
        "pNM42->code == -108",
        "vb6_ListView_OnNotify((void*)vb6_hwnd_ListView1, -108",
        "vb6_ListView_ColumnHeaderAt((void*)vb6_hwnd_ListView1",
        "_ItemClick(vb6_lvItem7)", "_ColumnClick(vb6_lvHdr7)",
        'vb6_CStrBool(vb6_ListView_GetGridLines(',
        'vb6_CStrBool(vb6_ListView_GetMultiSelect(')
        # C29-4: 事件回调改**传值** (ByVal 对象语义)。旧形状传 &obj 是 void**, 与
        # handler 形参 void* 不符 —— 成员读拿"指针的地址"当 IDispatch, 必然 not found。
        # StatusBar 的 WM_NOTIFY 派发形状由 c29sbevt (SbEvent.vbp, 真有 StatusBar) 覆盖 ——
        # 本 ListView 夹具没有 StatusBar, 在此断言恒缺 (曾误挂于此致 GA vbp 红)。

    # Fix 195: .frx 三种 blob 的真实布局 —— 字符串 (Text) / 字符串表 (List) /
    # 整数表 (ItemData)。旧 readIntList 按"每项 2B 整数"读 ItemData, 读到的是
    # 结构的字节本身, 任何工程都解出 1/304/12288 这串恒定假值 → 设计期 ItemData
    # 编进 exe 一直是垃圾。本用例的 ItemData 取 5/300/-7, 旧实现必错。
    # 账 #207: 同一份 blob 接第二类控件。设计期 List/ItemData 的 .frx 引用以前只认
    # ListBox 一档，ComboBox 那 17 处（Charts 2020 四个 demo 工程的 "Number of Series"）
    # 全落空 ⇒ 下拉框是空的，而 ucTreeMaps 的 demo 正是用 Combo1.ListIndex 决定画不画。
    # FD8 是真判据（三项 List + 两项 ItemData 各自对上），FD9 留原始读数，
    # FD10 只钉前缀（留原始读数，不钉值）：`Combo1.ListIndex = 2` 之后 clicks=0。
    # 这一形对不对**本机量不了**（没有 VB6）。唯一能引的旁证是自家 ctrlfiles 的 CF14/CF15：
    # 那两条是**手工调 fileList_Click 来模拟一次点击**的，也就是说这套夹具一直按
    # "程序改 ListIndex 不发 Click" 写 —— 所以这里不把任何一种口径钉成断言（台账 §B43 / 账 #208）。
    Test-Vbp "frxdata" "$Tests\frxdata\FrxData.vbp" @(
        "FD1=alpha|beta", "FD2=3", "FD3=1234", "FD4=5", "FD5=300", "FD6=-7",
        "FD7=OK", "FD8-COMBO=True", "FD9-COMBO-RAW count=3 item0=1234 id0=5 id2=-7",
        "FD10-COMBO-LI-RAW", "FRXDATA-DONE")

    # P20-40: StatusBar 复刻 (msctls_status32, 不加载 mscomctl.ocx)。
    # comctl32 v5.82 / v6 都不注册 msctls_status32, 连 dwICC=0xFFFFFFFF 全开也补不上,
    # 所以 RTL 自己注册一个同名真窗口类 (见 vb6_StatusBar_RegisterClass)。
    # SB8/SB15 盯 sbrNum 面板的设计期 Text 不能被"系统自动显示"覆盖 (text / shown 两个字段);
    # SB29/SB31 盯 Panels.Add 插到中间时 memmove 留下的悬垂副本 (双重释放 → ClearPanels 崩)。
    # P20-43 修正: 同 ctrlprogress —— 期望串与夹具完整标签逐字一致。
    # SB0 是已知局限 (vb6_GetControlHwnd 直返入参, Me.hwnd 在 Debug.Print 里为空),
    # SB34-LASTKEY= 为空是对的 (第 3 格是没给 Key 的时间面板)。
    Test-Vbp "ctrlstatusbar" "$Tests\ctrlstatusbar\CtrlStatusBar.vbp" @(
        "SB0-HWND= CAP=CtrlStatusBar CL=", "SB1-COUNT=3", "SB2-KEY1=pr",
        "SB3-TEXT1=Ready", "SB4-STYLE1=0", "SB5-AUTOSZ1=1", "SB6-MINW1=40",
        "SB7-KEY2=tp", "SB8-TEXT2=Tip", "SB9-STYLE2=2", "SB10-W2=120",
        "SB11-AUTOSZ2=0", "SB12-TIP2=NumLock state", "SB13-STYLE3=5",
        "SB14-IDXBYKEY=2", "SB15-TEXTBYKEY=Tip", "SB16-ALIGN=2", "SB17-STYLE=0",
        "SB18-SIMPLE2=Simple text here", "SB19-SIMPLE2B=Changed", "SB20-STYLE2B=0",
        "SB21-SETTEXT=Busy", "SB22-ADDIDX=4", "SB23-COUNT2=4", "SB24-NEWKEY=extra",
        "SB25-NEWTEXT=Extra", "SB26-INSERT=2", "SB27-COUNT3=5", "SB28-IDXP2=ins",
        "SB29-IDXP3=tp", "SB30-AFTERRM=4", "SB31-P2KEY=tp", "SB32-COUNT4=4",
        "SB33-AFTERRM2=3", "SB34-LASTKEY=", "SB35-SETMINW=77", "SB36-SETW=123",
        "SB37-SETAUTOSZ=0", "SB38-SETTIP=hello", "SB39-SETSTYLE=6",
        "SB40-AFTERCLR=0", "CTRLSTATUSBAR-DONE")

    # 账 #205: 状态条那两处问字体（sbrContents 的排版宽 + 画文字）以前都裸问窗口，而状态条是
    # RTL 自注册的窗口类 ⇒ 恒回 NULL，实测同一串文字在 8pt 与 20pt 两枚上量出同一个宽(140)。
    # 排版结果从 VB 没有现成的门（Panels(i).Width 读的是**请求值**，不是排版后的宽），
    # 所以判据问窗口本人：SB_GETPARTS(WM_USER+6) 交回各格右边界。
    # SF02/SF04 两头钉"字体参与排版"（改前都 False），SF05 钉"显式给过 Width 的那格不许挪"。
    # RAW 三行只钉前缀 —— 110/260 是 DPI 的函数，不钉绝对数（#147 那条口径）。
    $sbFontExpected = @(
        "SF00-RAW pf8=", "SF01-RAW fix=", "SF02-CONTENTS-TRACKS-FONT=True",
        "SF03-RAW after=", "SF04-RUNTIME-FONT=True",
        "SF05-FIXED-EDGE-UNCHANGED=True", "SBFONT-DONE")
    Test-Vbp "sbfont" "$Tests\sbfont\SbFont.vbp" $sbFontExpected
    Test-Vbp "sbfont_x86" "$Tests\sbfont\SbFont.vbp" $sbFontExpected -Arch "x86"
    # <vbeclipse> 账 #221 = C29-PL-a: Picture.Line 落原生 GDI。这条判据只能**画完再问像素** ——
    # 改前 8 处 Line 全发成 ComGetObjectProp(hwnd,"Line") 再取 Item, 两跳在 RTL 都登记成
    # "认识但什么都不做" ⇒ 编得过、跑得起、一笔不画、一条诊断也不打 (CLINE 那族静默空转的第三例)。
    # 四形各钉两头: 画过的那枚像素 == 交进去的颜色, 没画过的那枚 != 它 —— 只钉前者会放过
    # "整片刷成红", 只钉后者会放过"根本没落笔"。PL03 同时证明 BF(=1|4) 与 B(=1) 不是同一个数。
    # 画与问都放在 Timer 第一拍 (dcsurf 那条口径: 窗口问题在窗口活着的时候问);
    # 先试 Form_Load 时 GetPixel 一律 -1 (CLR_INVALID), 那不是本刀的靶子, 记在夹具注释里。
    # 账 #239: 控件那户 Print 的笔位。改前 picP.Print 用本族自存的像素光标 (VB6_PrintX/Y)，
    # 既不读 picP.CurrentX/CurrentY 也不推进它们 —— 同一份判据在改前的编译器上实测
    # 推进 = 0 (而同一枚控件答 TextHeight = 13)、笔位放到 60 而墨落在第 2 行、Cls 之后
    # CurrentY 仍是 400。PL08 钉"增量 == 自己量的 TextHeight"、PL09 钉 Cls 复位笔位、
    # PL10 钉缇档同一问 (一份存储不许两种单位)、PL11 钉 Cls 用的是带 Fix 187 哨兵那份背景色、
    # PL12 钉两行叠两行高。PL13/PL14 只钉前缀 (13/195/26 都是 DPI 的函数，不钉绝对数)。
    $pcDrawExpected = @("PL01-LINE=True", "PL02-BOX=True", "PL03-FILL=True", "PL04-CIRCLE=True",
        "PL08-PEN=True", "PL09-CLSPEN=True", "PL10-TWIPADV=True",
        "PL11-BLACKCLS=True", "PL12-STACK=True",
        "PL13-RAW dp=", "PL14-RAW twip=", "PL-DONE")
    Test-Vbp "pclinedraw" "$Tests\pcline\PcDraw.vbp" $pcDrawExpected
    Test-Vbp "pclinedraw_x86" "$Tests\pcline\PcDraw.vbp" $pcDrawExpected -Arch "x86"

    # 账 #233 = Form 的绘图状态属性。这一组改前**压根编不出来**：`Me.DrawWidth = 3` 发成
    # `vb6_Form_DrawGetWidth(vb6_hwnd_X) = 3;` (C2106) —— 写侧从没登记进 cgen 的写表，而
    # 那四条读侧硬编码住在 cgen_expr_member_form_builtin.inc 的一份侧表里（注释还写着"写侧
    # 不需要"）。两头各钉：FD01/02/03 = 写进去的数读回来一样（钉 +1 编码与写侧接线）；
    # FD07/08 = 一枚像素证人（画到的那一点问得到、旁边那一点不是那个颜色）；
    # FD10 = 两份编码合一的证人 —— `Print` 推进笔位之后，绘图那一路读得到同一个数
    # （改前那里读回的是 float 位图案当整数的天文数，或 0）。
    # FD04/05/06 钉的是自己算出来的定值，可以直接当判据；FD09 那行 **AFTERPRINT 的 y 与字号/DPI 有关**
    # ⇒ 只打印不当判据（本线口径：读数留档，判据换成"落在 30..3000 之间"那一枚布尔）。
    $fdrawExpected = @("FD01-drawwidth=True", "FD02-curxy=True", "FD03-pset2=True",
        "FD04-RAW xy=300,130 dw=3", "FD05-sm0=1", "FD06-sm1=3",
        "FD07-PIXEL=True", "FD08-neg=True",
        "FD11-forecolor=True", "FD12-RAW pen=16711680 blue=16711680 first=255",
        "FD10-printstore=True",
        # 237: the pen advance and the paint position must use this window's own unit.
        # The RAW lines carry DPI-dependent numbers, so only their prefixes and the
        # two scale-mode readings are pinned; the True/False judges are ratios.
        "FD13-printadvance-twips=True", "FD14-printadvance-points=True",
        "FD15-RAW", "sm0=1 sm1=2",
        "FD16-paint-units=True", "FD17-RAW",
        "FD-DONE")
    Test-Vbp "fdrawstate" "$Tests\fdraw\FDemo.vbp" $fdrawExpected
    Test-Vbp "fdrawstate_x86" "$Tests\fdraw\FDemo.vbp" $fdrawExpected -Arch "x86"

    # --- P20-42: SSTab (SysTabControl32 复刻) ---
    # 期望串取自夹具真实输出 (别缩写标签)。TS25..TS28 是切页显隐: vb6_GetControlVisible
    # 走 IsWindowVisible 沿父链传播, 所以断言放在 Timer 里 (窗体已显示之后)。
    # 账 #124 起这几条打 True/False —— `.Visible` 在类型 oracle 里归 Boolean (VB6 口径),
    # 以前按 Long 打 -1/0。TS20/TS21 (TabVisible) 与 TS25B (Tabs) 不受影响: 它们是 int 属性。
    # TS30 是 Click(PreviousTab), 由 RTL 在程序化改 Tab 时补发的 TCN_SELCHANGE 触发。
    Test-Vbp "ctrlsstab" "$Tests\ctrlsstab\CtrlSSTab.vbp" @(
        "TS1-TABS=3", "TS2-TAB=1", "TS3-ORIENT=0", "TS4-STYLE=0", "TS5-PERROW=3",
        "TS6-WRAP=False", "TS7-SET0=0", "TS8-SET2=2", "TS9-OOR=2", "TS10-ORIENT=1",
        "TS11-STYLE=1", "TS12-PERROW=4", "TS13-WRAP=True", "TS14-TABS5=5",
        "TS15-TABAFTERGROW=2", "TS16-TABS2=2", "TS17-TABAFTERSHRINK=1",
        "TS18-CAP0=常规", "TS19-CAP0B=改过", "TS20-VIS1=-1", "TS21-VIS1B=0",
        "TS22-P0LEFT=240", "TS23-P1LEFT=240", "TS24-P2LEFT=240",
        "TS31-TIP0=tip0", "TS32-TIP2=tip2", "TS33-TIP1=",
        "TS25-TABVIS=True", "TS26-AT0-P0VIS=True P1VIS=False P2VIS=False",
        "TS27-AT1-P0VIS=False P1VIS=True P2VIS=False", "TS28-AT2-P0VIS=False P1VIS=False P2VIS=True",
        "TS29-SETTAB2=2", "TS29B-SETTAB0=0",
        "TS34=True/Boolean/11",
        # 账 #167：容器清单收成**一处权威**之前的读数 —— BASE 上两步都是 `0/0`（容器子控件在
        # 创建侧与派发侧数到不同的 id，BN_SETFOCUS 带的那个号找不到自己的 arm）。两步反向都钉，
        # 免得「全指到同一枚」那种假绿蒙过去。
        "TS35-ARM-TAB=1/0", "TS36-ARM-FRAME=0/1",
        "CTRLSSTAB-DONE", "CTRLSSTAB-VISDONE", "CTRLSSTAB-CLICKDONE")

    # --- P20-43/44: 窗体事件面 + OLE 拖放 (目标侧 Drop + 源侧 OLEDrag) ---
    # **要 -Env**: OLE 的那几条断言靠 C3_OLEDDB_TEST=1 驱动 —— 无头环境没法真拖
    # (DoDragDrop 要真实鼠标键状态), RTL 在该变量下改走"直接 fire IDropTarget 方法 /
    # 只跑源事件链"的联测路径。EV24 的 X 坐标随屏幕布局变, 所以只断言到 EFF=1。
    Test-Vbp "frmevents" "$Tests\frmevents\FrmEvents.vbp" @(
        "EV01-INIT", "EV02-LOAD", "EV03-RESIZE", "EV04-ACTIVATE", "EV05-PAINT",
        "EV06-GOTFOCUS", "EV19-TIMER-FIRED", "EV20-LOOPDONE", "EV21-QUERYUNLOAD",
        "EV22-UNLOAD", "EV23-TERMINATE",
        "EV24-OLE-DROP=OLE-TEST-DROP EFF=1", "EV25-OLE-OVER",
        "EV26-DRAG-DONE", "EV27-STARTDRAG", "EV29-COMPLETE=3") -Env "C3_OLEDDB_TEST=1"

    # --- 账 #185: PictureBox/Image 的事件面 vs RTL 自绘子类的同名槽位碰撞 ---
    # 两层子类以前把原始窗口过程存在同一个窗口属性名 (VB6_OrigProc) 上, 又都写「属性已存在就不装」
    # ⇒ 后装的那层被静默丢掉: Paint / MouseDown / MouseUp / Click 一次也不响。
    # 判据两头 (缺一头就是假绿):
    #   ① 事件臂要响 —— PB02..PB05 只有事件层真装上才有; PB06 是 Label 的证人 (同为 STATIC,
    #      但不含 SS_BITMAP ⇒ 没人跟它抢), 修复前后都该有 —— 它排除「夹具自己没驱动起来」这种红。
    #   ② 表面不能丢 —— PB01 里那颗像素必须是设计期 BackColor 的红 (255)。画序定成「最外层
    #      BeginPaint/EndPaint 一次, DC 经 VB6_PaintDC 交下去让自绘层画表面, 再抬用户的 _Paint」。
    # 负控 = 修复前的编译器编同一份夹具: x64 与 x86 都只剩 PB00/PB06 两条,
    #   汇总行读 mdown=0 mup=0 click=0 img=0 lb=1 paint_ok=0。
    Test-Vbp "pbsub" "$Tests\pbsub\PbSub.vbp" @(
        "PB01-PAINT px=255", "PB02-PIC-MDOWN x=20", "PB03-PIC-MUP", "PB04-PIC-CLICK",
        "PB05-IMG-CLICK", "PB06-LB-MDOWN",
        "PB-CNT mdown=1 mup=1 click=1 img=1 lb=1 paint_ok=1", "PB-WAIT done=True")
    Test-Vbp "pbsub_x86" "$Tests\pbsub\PbSub.vbp" @(
        "PB01-PAINT px=255", "PB02-PIC-MDOWN x=20", "PB03-PIC-MUP", "PB04-PIC-CLICK",
        "PB05-IMG-CLICK", "PB06-LB-MDOWN",
        "PB-CNT mdown=1 mup=1 click=1 img=1 lb=1 paint_ok=1", "PB-WAIT done=True") -Arch "x86"

    # --- P20-46: 控件字符串全程 W / 支持多国语言（用户要求）的源码面反例断言 ---
    # 判据两条:
    #   ① RTL 源码里不得**直接调 A 版 Win32 API** —— 编译已带 /DUNICODE /D_UNICODE /utf-8
    #      (msvc_driver.cpp:172), 但不带后缀的宏与 A 版调用仍可能混进来。
    #      `src/rtl/core/di/` 下的 **DI 桩除外**: 那是用户 `Declare ... Alias "xxxA"` 时
    #      C3 提供的转发桩, 本来就该给 A 版。
    #   ② 字体 charset 必须是 DEFAULT_CHARSET —— 写死 GB2312_CHARSET 之类会让系统在字体里
    #      找不到韩文/俄文字形, 显示成方框 (Fix 190 的教训)。
    if (Enter-VbpShard) {
        $script:total++
        Write-Host -NoNewline "  [SRC] widechar_only ... "
        $rtlRoot = Join-Path (Split-Path $PSScriptRoot -Parent) "src\rtl\core"
        $ansiBad = @()
        if (Test-Path $rtlRoot) {
            # 递归取文件再 Select-String —— `-Path '...\*\*.c'` 这种通配是匹配不到文件的,
            # 那会让断言永远 PASS (假绿)。这里两向都验过: 不过滤 di/ 时必须命中。
            $rtlFiles = Get-ChildItem -Path $rtlRoot -Recurse -File -Include *.c, *.cpp -ErrorAction SilentlyContinue |
                        Where-Object { $_.FullName -notmatch '\\di\\' }
            if ($rtlFiles) {
                # 用 @() 强制数组: Select-String 只命中 1 条时返回标量, 直接 `$x += ...` 会
                # "MatchInfo 没有 op_Addition" 而抛异常 (断言崩掉而不是判红) —— 两向验证抓到过。
                $ansiBad = @(Select-String -Path $rtlFiles.FullName -ErrorAction SilentlyContinue `
                    -Pattern '\b(SendMessageA|PostMessageA|CreateWindowExA|RegisterClassA|DefWindowProcA|CallWindowProcA|GetWindowTextA|SetWindowTextA|GetClassNameA|DrawTextA|CreateFontA|LoadCursorA|LoadIconA)\s*\(' |
                    Where-Object { $_.Line -notmatch '^\s*(//|\*)' })
                $ansiBad += @(Select-String -Path $rtlFiles.FullName -ErrorAction SilentlyContinue `
                    -Pattern '(GB2312_CHARSET|SHIFTJIS_CHARSET|HANGEUL_CHARSET|CHINESEBIG5_CHARSET)' |
                    Where-Object { $_.Line -notmatch '^\s*(//|\*)' })
            }
        }
    }
    if ($ansiBad.Count -eq 0) {
        $script:pass++
        Write-Host "PASS" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL (ANSI/非默认 charset 残留)" -ForegroundColor Red
        $ansiBad | Select-Object -First 5 | ForEach-Object {
            Write-Host ("    " + $_.Filename + ":" + $_.LineNumber + "  " + $_.Line.Trim()) -ForegroundColor Red
        }
    }

    # --- 账 #167: 「谁是容器」只许有一处权威（结构性哨兵，钉的是**不许再抄**）---
    # 这一族的病根不是某一格写错，是同一张清单在好几处各抄一遍、每处抄的还不齐：
    # 创建那两路认 SSTab，事件派发那两路只认 Frame/PictureBox，拆子类那一路连 PictureBox 都漏。
    # 抄漏的后果不是崩，是**容器里的孩子收不到自己的事件**（子控件在两侧数到不同的 id）。
    # 判据：源码里任何一条布尔表达式中同时出现 {Frame, PictureBox, SSTab} 里的两型以上
    # = 有人又把容器清单手写了一遍。唯一权威 `controlIsContainerType` 自己是 switch-case 写的，
    # 这个形状扫不到它 ⇒ 不误伤。
    # 两向验过：收口前（HEAD）这份扫出 **10 处**，收口后 **0 处**。
    # ⚠ 它管不到「只写了一型」的那种漏（dispatch 里 `!= FrmControlType::Frame` 就是这一形）——
    #   那一格由上面的运行期证人 TS35/TS36 兜：容器少认一类，容器里的孩子就没有 _GotFocus。
    if (Enter-VbpShard) {
        $script:total++
        Write-Host -NoNewline "  [SRC] container_list_single_authority ... "
        $beRoot = Join-Path (Split-Path $PSScriptRoot -Parent) "src\backend"
        $copyBad = @()
        if (Test-Path $beRoot) {
            # 先把 `child.controlType` 这种「对象.成员」折成一个记号 CTL，这样两条比较之间
            # 只剩空白 / || / && / 括号 / ! 就说明它们同属一条布尔表达式（跨行续行也算）。
            $objRe = [regex] '\b[A-Za-z_][A-Za-z0-9_]*(?:\.[A-Za-z_][A-Za-z0-9_]*)*\.controlType\b'
            $cmpRe = [regex] '\bCTL\b\s*(?:==|!=)\s*FrmControlType::(Frame|PictureBox|SSTab)\b'
            $glueRe = [regex] '^[\s()!&|]*$'
            foreach ($f in @(Get-ChildItem -Path $beRoot -Recurse -File -Include *.c, *.cpp, *.h, *.inc)) {
                $txt = $objRe.Replace([IO.File]::ReadAllText($f.FullName), 'CTL')
                $ms = @($cmpRe.Matches($txt))
                $i = 0
                while ($i -lt $ms.Count) {
                    $j = $i + 1
                    $kinds = New-Object System.Collections.Generic.HashSet[string]
                    [void]$kinds.Add($ms[$i].Groups[1].Value)
                    while ($j -lt $ms.Count) {
                        # ⚠ `$ms[$k].End` 在 PS 5.1 里取回来是**空**（Match.End 没被绑定上），
                        #   Substring 就以 0 起、gap 变成整个文件 ⇒ 永远不判红。改成 Index+Length。
                        $prevEnd = $ms[$j - 1].Index + $ms[$j - 1].Length
                        $gap = $txt.Substring($prevEnd, $ms[$j].Index - $prevEnd)
                        if (-not $glueRe.IsMatch($gap)) { break }
                        [void]$kinds.Add($ms[$j].Groups[1].Value)
                        $j++
                    }
                    if ($kinds.Count -ge 2) {
                        $ln = ($txt.Substring(0, $ms[$i].Index) -split "`n").Count
                        $copyBad += ($f.Name + ":" + $ln + " [" + (($kinds | Sort-Object) -join ",") + "]")
                    }
                    $i = $j
                }
            }
        }
        if ($copyBad.Count -eq 0) {
            $script:pass++
            Write-Host "PASS" -ForegroundColor Green
        } else {
            $script:fail++
            Write-Host "FAIL (容器清单被手抄了 $($copyBad.Count) 处)" -ForegroundColor Red
            $copyBad | Select-Object -First 8 | ForEach-Object { Write-Host ("    " + $_) -ForegroundColor Red }
        }
    }

    # --- Fix 195: 资源引用缺失不得静默, 且 --extract-frx 能把 .frx 取值导成 VB 代码 ---
    # 背景: VB6 把多行文本/图片甩进同名 .frx, .frm 里只留 `属性 = "X.frx":含偏移`。
    # .frx 缺失时属性设计期取值被静默丢弃, 编出的 exe 与 VB6 不一致却没提示
    # (VB6 IDE 自己会写 <窗体>.log 报"文件引用无效")。
    $frxNoFile = Join-Path $OutDir "frx_nofile"
    if (Test-Path $frxNoFile) { Remove-Item $frxNoFile -Recurse -Force }
    New-Item -ItemType Directory -Path $frxNoFile | Out-Null
    Copy-Item "$Tests\frxdata\FrxData.frm" $frxNoFile
    Test-CliOk "frx_missing_warn" @(
        "`"$frxNoFile\FrxData.frm`"", "--output-dir", "`"$frxNoFile`"") "VB4004"

    $frxExDir = Join-Path $OutDir "frx_extract"
    if (Test-Path $frxExDir) { Remove-Item $frxExDir -Recurse -Force }
    New-Item -ItemType Directory -Path $frxExDir | Out-Null
    Copy-Item "$Tests\frxdata\FrxData.frm" $frxExDir
    Copy-Item "$Tests\frxdata\FrxData.frx" $frxExDir
    Test-CliOk "frx_extract" @("`"$frxExDir\FrxData.frm`"", "--extract-frx") ".frx.bas" -ShardGroup "frx"

    # 导出内容: 文本(多行拼接) / 列表项 / 列表项数据 三类都要落到普通 VB 语句上
    $frxExOut = Join-Path $frxExDir "FrxData.frx.bas"
    if (Enter-VbpShard -Group "frx") {
        $script:total++
        Write-Host -NoNewline "  [CLI] frx_extract_content ... "
        if ((Test-Path $frxExOut) -and
            (Select-String -Path $frxExOut -Pattern 'Text1\.Text = "alpha" & vbCrLf & "beta"' -Quiet) -and
            (Select-String -Path $frxExOut -Pattern 'List1\.AddItem "1234"' -Quiet) -and
            (Select-String -Path $frxExOut -Pattern 'List1\.ItemData\(1\) = 300' -Quiet)) {
            $script:pass++
            Write-Host "PASS" -ForegroundColor Green
        } else {
            $script:fail++
            Write-Host "FAIL" -ForegroundColor Red
            if ($Verbose) { Get-Content $frxExOut -ErrorAction SilentlyContinue }
        }
    }

    # --- Fix 196: 非 ASCII (中文) 路径下的完整编译 + 运行 ---
    # 真因: cl.exe / link.exe 读 @rsp 响应文件时按**系统 ANSI 代码页**解释字节, 而
    # 命令行本身就是 UTF-8。于是 /Fe"…\新建文件夹\out\x.exe" 被解成 GBK 乱码
    # (且 GBK 双字节会吃掉后面的 '\'), 链接期 LNK1104「无法打开文件」/ LNK1117。
    # 修法 = 响应文件写 UTF-16LE+BOM (见 msvc_driver.hpp 的实测矩阵)。
    #
    # 目录名用 [char] 拼出来, 让本脚本保持**纯 ASCII**: Windows PowerShell 5.1 读
    # 无 BOM 的 UTF-8 .ps1 会按 ANSI 解码, 直接写字面量会让路径本身先烂掉。
    $cnName = [string]::Join('', [char]0x4E2D, [char]0x6587, [char]0x8DEF, [char]0x5F84,
                                  [char]0x6D4B, [char]0x8BD5)   # 中文路径测试
    $cnDir = Join-Path $OutDir $cnName
    if (Test-Path $cnDir) { Remove-Item $cnDir -Recurse -Force }
    New-Item -ItemType Directory -Path $cnDir | Out-Null
    Copy-Item "$Tests\frxdata\*" $cnDir
    if (Enter-VbpShard) {
        $script:total++
        Write-Host -NoNewline "  [VBP] nonascii_path ... "
        $cnOut = Join-Path $cnDir "out"
        $cnCompile = & $C3 (Join-Path $cnDir "FrxData.vbp") --output-dir $cnOut @IncArg 2>&1
        $cnExe = Join-Path $cnOut "FrxData.exe"
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path $cnExe)) {
            $script:fail++
            Write-Host "FAIL (compile in non-ASCII path)" -ForegroundColor Red
            if ($Verbose) { Write-Host ($cnCompile | Out-String) }
        } else {
            # 光能链接还不够: 跑起来核对取值, 顺带证明同目录下的 .frx 也按宽路径读到了
            $cnRun = Invoke-TestExe -ExePath $cnExe -WorkDir $cnOut -Name "FrxDataCn"
            $cnOk = $cnRun.Ok
            foreach ($needle in @("FD1=alpha|beta", "FD4=5", "FD6=-7", "FRXDATA-DONE")) {
                if (-not (Test-NeedleHit -Lines $cnRun.Output -Needle $needle)) { $cnOk = $false }
            }
            if ($cnOk) {
                $script:pass++
                Write-Host "PASS" -ForegroundColor Green
            } else {
                $script:fail++
                Write-Host "FAIL (non-ASCII path run)" -ForegroundColor Red
                if ($Verbose) { Write-Host ("    " + $cnRun.Detail); Write-Host ($cnRun.Output -join "`n") }
            }
        }
    }

    # --- Fix 196b: 非 ASCII 路径 + **缺 .frx** 不得让编译器 abort() ---
    # 崩溃机理 (实测栈, driver_frontend.cpp 的 VB4004 告警分支):
    #   std::filesystem::path(utf8String) 的**窄串**重载按系统 ACP(中文机 936/GBK) 解释
    #   char*, 而该处拿到的是 UTF-8。路径字节凑不成合法 GBK 序列时
    #   _Convert_narrow_to_wide 抛 filesystem_error -> 无人接 -> std::terminate -> abort()
    #   => 退出码 3 + 模态「Debug Error / abort() has been called」对话框。
    #   "非 ASCII 字符数为奇数" 几乎必然落进这条 (首字符起按 2 字节分组会剩半个)。
    #   目录名取「中文叉」(3 字 = 9 字节, 奇) 精确命中; 「新建目录」(4 字 = 12 字节, 偶)
    #   反而侥幸不崩 —— 所以上一用例覆盖不到, 必须单列。
    # 仍用 [char] 拼名字保持本脚本纯 ASCII (见上一段注释)。
    $cnOddName = [string]::Join('', [char]0x4E2D, [char]0x6587, [char]0x53C9)   # 中文叉
    $cnOddDir = Join-Path $OutDir $cnOddName
    if (Test-Path $cnOddDir) { Remove-Item $cnOddDir -Recurse -Force }
    New-Item -ItemType Directory -Path $cnOddDir | Out-Null
    Copy-Item "$Tests\frxdata\FrxData.frm" $cnOddDir      # 故意**不**拷 .frx
    if (Enter-VbpShard) {
        $script:total++
        Write-Host -NoNewline "  [VBP] nonascii_missing_frx ... "
        $oddOut = Join-Path $cnOddDir "out"
        $oddLog = (& $C3 (Join-Path $cnOddDir "FrxData.frm") --emit-c --output-dir $oddOut 2>&1 | Out-String)
        $oddRc = $LASTEXITCODE
        if ($oddRc -eq 3) {
            $script:fail++
            Write-Host "FAIL (abort() 复现: exit=3)" -ForegroundColor Red
            if ($Verbose) { Write-Host $oddLog }
        } elseif ($oddRc -ne 0) {
            $script:fail++
            Write-Host "FAIL (exit=$oddRc)" -ForegroundColor Red
            if ($Verbose) { Write-Host $oddLog }
        } elseif ($oddLog -notlike "*VB4004*") {
            $script:fail++
            Write-Host "FAIL (缺 .frx 却未报 VB4004)" -ForegroundColor Red
            if ($Verbose) { Write-Host $oddLog }
        } else {
            $script:pass++
            Write-Host "PASS" -ForegroundColor Green
        }
    }

    Test-Vbp "M6Test" "$Tests\M6Test.vbp" @("M6A:OK", "M6B:OK", "M6C:OK", "M6D:OK", "M6 PASSED")
    Test-Vbp "modulemethod" "$Tests\test_modulemethod.vbp" @("30", "21")

    # QR code project (tests\VbQRCodegen-master): form loads, sets Image1.Picture via Stretch
    Test-GuiVbp "VbQRCodegen" "$Tests\VbQRCodegen-master\test\Project1.vbp"
    # BalloonTooltips: form loads with controls + creates its common-controls tooltip windows (x64).
    Test-GuiVbp "BalloonTooltips" "$Tests\BalloonTooltips\prjBalloonTooltips.vbp" -ExeName "BalloonTooltips" -ShardGroup "balloon"
    # ai/029 C29-M 的反面：这枚工程的 .res **自带** #1 清单（BalloonTooltips.rc 里
    # `1 RT_MANIFEST "BalloonTooltips.exe.manifest"`，那段含 dpiAware/compatibility）
    # ⇒ 必须"用他的、且只有一份"。`dpiAware` 只有用户那份里有，所以这条同时钉住
    # "让位生效"与"内置那份没叠上去"（两份 #1 会让加载器直接报错）。
    # 路径注意：Test-GuiVbp 的产物落在 `output\<用例名>\` 下（不是 $OutDir 根），
    # 第一版我按 $OutDir\BalloonTooltips.exe 断 ⇒ CI 直接 FAIL (no exe) —— 记下来。
    Test-ProductManifest "balloon_manifest_is_user_supplied" "$OutDir\BalloonTooltips\BalloonTooltips.exe" 1 "dpiAware" -ShardGroup "balloon"
    # Charts 2020 demo (3rd-party UserControl charts): windowless chart controls (x86 first;
    # x64 after LongPtr port of API pointers/handles in the .ctl/.cls sources).
    # 账 #175: 除"起窗口 + 退出码 0"外, 再验每控件的绘制缓冲真画了东西 (>=40 色)
    Test-GuiVbp "Charts2020" "$Tests\Charts 2020\Proyecto1.vbp" -Arch "x86" -AutoExitSec 3 -DumpMinColors 40
    # czUI (czForm): 自定义 GDI+ UserControl (.ctl) 无边框窗体 demo, 需 -Arch x86 (32 位)
    Test-GuiVbp "czUI" "$Tests\czUI-main\czFormDemo.vbp" -Arch "x86" -AutoExitSec 3
    # 账 #184 / B23: 这份 demo 每枚网格都经 comctl32 的 SetWindowSubclass 装了 SUBCLASSPROC,
    # 而那个指针以前是 __cdecl 本体地址 ⇒ x86 起窗期堆损坏 (改前本地 3/4 崩、改后 14/14 起窗)。
    # B23 当时记的是"这份 demo 完全在门禁之外 ⇒ 缺陷藏了很久", 修好之后就按它说的挂上来。
    Test-GuiVbp "FlexGridX86" "$Tests\VBFlexGridDemo\VBFlexGridDemo.vbp" -Arch "x86" -AutoExitSec 3
    # NewTab: 第三方 OCX 控件 (NewTab01.ocx, 32 位) 真宿主验证. 免注册便携部署 (OCX 在工程目录, 由 harness 复制到 exe 旁, 不依赖本机注册);
    # 无边框窗体无关闭按钮/无自动退出逻辑, 用 -AutoExitSec 3 收尾避免阻塞后续测试
    Test-GuiVbp "NewTab" "$Tests\NewTab-test\Test.vbp" -Arch "x86" -AutoExitSec 3
    # axcontrolsadd: 运行期 Me.Controls.Add 免注册创建第三方 OCX 控件 (NewTabCtl.NewTab), 验证
    # Object= 免注册表在运行时路径被查询 (regfree hit); OCX 在工程目录, 由 harness 复制到 exe 旁,
    # 不依赖本机注册。断言 ADD=OK (实例化+激活不崩)。x86: NewTab01.ocx 是 32 位, 必须与 32 位 exe 同架构。
    Test-Vbp "axcontrolsadd" "$Tests\axcontrolsadd\AxAdd.vbp" @("ADD=OK") -Arch "x86"
    # ExtShow: 跨模块窗体默认实例"无参" Show (Fix 146 回归靶, 2026-09-20 vbman C2198):
    # .bas caller 调 Form2.Show, 定义侧签名 (hMDIClient, modal) 后调用侧须补 modal=0
    Test-GuiVbp "ExtShow" "$Tests\ext_show_test\test_ext_show.vbp" -AutoExitSec 3
    Write-Host ""

    Test-Vbp "test_implements" "$Tests\test_implements.vbp" @("IMPL1:OK", "IMPL2:OK", "Implements test PASSED")
    Test-Vbp "test_events" "$Tests\test_events\test_events.vbp" @("Events test PASSED")
    Test-Vbp "M7Test" "$Tests\m7_test\M7Test.vbp" @("4/4 PASSED")
    # ai/022 B03: cross-module new-style contract reached through an Interface head-line host
    Test-Vbp "itf_xmod_writer" "$Tests\itf_xmod\XWriter.vbp" @("XMOD1:OK", "XMOD2:OK", "IFV1:OK", "IFV2:OK", "IFV3:OK", "LIFE1:OK", "LIFE2:OK", "LIFE3:OK", "LIFE9:OK", "TERM last=bye", "TERM last=scoped", "QI1:OK", "QI2:OK", "QI3:OK", "QI4:OK", "TOF1:OK", "TOF2:OK", "TOF3:OK", "DN0:OK", "DN1:OK", "DN2:OK", "DN3:OK", "TOC1:OK", "TOC2:OK", "TOC3:OK")
    # ai/022 B10: `Implements IViaNamed Via m_h` -- six slots served by generated
    # adapters over the holder's own interface table (no forwarding member written).
    # Both architectures: the adapter dereferences a field whose type is another
    # class's struct, so a layout slip would only surface on x86 (022 D40 rule).
    $viaExpected = @("VIA1:OK", "VIA2:OK", "VIA3:OK", "VIA4:OK", "VIA5:OK",
        "VIA6:OK", "VIA7:OK", "VIA8:OK", "VIA9:OK", "VIA-DONE")
    Test-Vbp "itf_via_pair" "$Tests\itf_via\Via.vbp" $viaExpected
    Test-Vbp "itf_via_x86" "$Tests\itf_via\Via.vbp" $viaExpected -Arch "x86"
    # ai/022 B11/C05 (ai/026 section 5, items 3-5): a CoClass block name used AS A TYPE --
    # `As Circle` / `New Circle` / `CreateObject("ActApp.Circle")` all bind to the block's
    # [Implementation] class. CC2/CC3 walk the other type positions (module field, parameter,
    # return type); CC6 proves the group view keeps the virtual table (an overridden Area on an
    # Inherits chain must answer), CC4/CC5/CC7 the ProgID rewrite including case.
    # Both architectures: the rewritten variable's C type is a class struct pointer, so a layout
    # slip would show up on x86 only (022 D40 rule).
    $ccActExpected = @("CC1:OK", "CC2:OK", "CC3:OK", "CC4:OK", "CC5:OK", "CC6:OK", "CC7:OK",
        "CC8:OK", "CC9:OK", "CC10:OK", "CC-DONE")
    # Fix 192: 元素类型为项目类的数组 (arr(i).Method)。此前接收者推断不出类 →
    # `VB6_SA_AT(void*, arr, i).Move(...)` → MSVC C2224。AC1/AC2 是静态数组 ——
    # 缺口不是 ReDim 专属; AC3 是 ReDim As <类名>; AC5 是 ReDim Preserve。
    # 基线(改动前)实测 28 条 C2224, 修后 0。两个架构都跑: 元素槽是 void*,
    # 转成 vb6_cls_X* 的布局只在 x86 上才会暴露对齐问题。
    $arrClsExpected = @("AC1:OK", "AC2:OK", "AC3:OK", "AC4:OK", "AC5:OK",
        "AC6:OK", "ARRCLS-DONE")
    Test-Vbp "arr_cls_elem" "$Tests\arr_cls\ArrCls.vbp" $arrClsExpected
    Test-Vbp "arr_cls_elem_x86" "$Tests\arr_cls\ArrCls.vbp" $arrClsExpected -Arch "x86"

    # Fix 193: `TypeOf lhs Is <项目类>`。此前落进 vb6_TypeOf —— 那是 vb6rtl_conv.c 里
    # 一个恒返 0 的桩, 于是项目类这一位一律答"否" (连 `TypeOf raw Is ShapeAct`
    # 都是 False)。改按"声明类 + 祖先链"静态判定后: TOF1-TOF3 自身/祖先 True,
    # TOF4-TOF7 兄弟/子类 False, TOF8-TOF9 Nothing False, TOF10 无虚槽的普通类,
    # TOF11 类数组元素 (与 Fix 192 联动), TOF12 进 If 分支不是只在 Print 里对。
    # 基线(改动前)实测 8 FAIL / 4 OK —— 那 4 个 OK 只是"本该 False"被恒假蒙对。
    $tofExpected = @("TOF1:OK", "TOF2:OK", "TOF3:OK", "TOF4:OK", "TOF5:OK", "TOF6:OK",
        "TOF7:OK", "TOF8:OK", "TOF9:OK", "TOF10:OK", "TOF11:OK", "TOF12:OK", "TOF-DONE")
    Test-Vbp "typeof_class" "$Tests\typeof\Tof.vbp" $tofExpected
    Test-Vbp "typeof_class_x86" "$Tests\typeof\Tof.vbp" $tofExpected -Arch "x86"
    # <vbeclipse>: `Option Compare Text` 按模块生效 (修复前 module.options 从没进代码生成层,
    # 整模块仍按二进制比)。同工程里再放一个**没有** Option Compare 的 Binary 模块
    # (OCBINARY: 读数是 BINARY) —— 模式若靠进程唯一的全局切, 这条就会变 TEXT。
    if (Test-Path "$Tests\optcmp\oc_ok.vbp") {
        $optCmpExpected = @("OC-eq=True", "OC-lt=False", "OC-instr=1", "OC-instr-exp-b=0",
            "OC-instr3=1", "OC-sc=0", "OC-rep=[----]", "OC-split-ub=2", "OC-filter-ub=1",
            "OC-like=True", "OC-case=hit", "OC-binary-mod=BINARY", "OC-DONE")
        Test-Vbp "optcmp_text_module" "$Tests\optcmp\oc_ok.vbp" $optCmpExpected
    }

    # <vbeclipse>: CC2 走 **ByRef 组名形参** (`Sub UseCircle(c As Circle)`), CC10 走
    # **真 ByRef Variant 形参** —— 同一个类实参、两种槽。这一对是"别过度修复"的夹子:
    # CC2 曾被装箱成 vb6_VARIANT* (与 `vb6_cls_ShapeAct**` 形参 ABI 不符 → 段错误),
    # 而把 CC10 也一起改直传则会把裸类指针交给按 vb6_VARIANT 解析的被调方。
    Test-Vbp "cc_act_pair" "$Tests\cc_act\Act.vbp" $ccActExpected
    Test-Vbp "cc_act_x86" "$Tests\cc_act\Act.vbp" $ccActExpected -Arch "x86"

    # <vbeclipse> 回归夹子 (ve_list): 真工程 VbEclipse 的 Classes/List.cls 全文 + 它的两类
    # "类名/变体槽"坑。每一条都对应一个已修的根因, 且都曾**编得过 (MSVC 只给警告/rc=0)
    # 而运行期错** —— 只靠生成码逐处对照才发现, 所以必须有运行期夹子盯着:
    #   · idx/lit/bykey = Alpha/Beta/Gamma  → rev16: 类成员 **Variant 数组元素**赋值曾是
    #     identity 浅拷贝, 槽与 ByRef 实参共用同一只 BSTR ⇒ 每槽读出"最后写入的值"
    #     (修前这里三条全是 Gamma)。
    #   · objidB/takeid                     → rev17b: `As <工程类>` 的 ByRef 实参曾**不装箱**,
    #     `vb6_cls_X**` 被当 `vb6_VARIANT*` 直接传 (C4133 警告而已) ⇒ 取回 NULL + av read 0x4。
    #   · selfnil=False / pophits=1         → rev17: 成员与**类同名**的类符号被语义层 erase 掉,
    #     进不了 coclass 表 ⇒ 装箱出 vt=9/pdisp=NULL 的空壳 Variant (修前 selfnil=True 后崩)。
    #   · fwdnil=False/fwdid                → rev17b 的取用侧 (`Set x = L.Item(...)` 再读属性)。
    # x86 + x64 双跑: 这一族问题在 x86 多数"侥幸能过" (指针同宽/隐式声明截断), x64 才露。
    $veListExpected = @(
        "LOCAL a0=[A] a1=[B]", "LOCAL k0=[K0] k1=[K1]",
        "count=2",
        "idx=0 len=5 s=[Alpha]", "idx=1 len=4 s=[Beta]", "idx=2 len=5 s=[Gamma]",
        "lit0=[Alpha] lit1=[Beta] lit2=[Gamma]",
        "bykeyAlpha=[Alpha]", "bykeyGamma=[Gamma]",
        "vt0=8 vt2=8",
        "indexOfGamma=2",
        "objnil=False", "objidB=[Props]", "takeid=[Props]",
        "selfnil=False", "selfid=[SelfId]", "popnil=False", "pophits=1",
        "selfid2=[SelfId] pophits2=1",
        "fwdvt=9", "fwdnil=False", "fwdid=[FwdId]")
    Test-Vbp "ve_list" "$Tests\ve_list\VeList.vbp" $veListExpected
    Test-Vbp "ve_list_x86" "$Tests\ve_list\VeList.vbp" $veListExpected -Arch "x86"

    # <vbeclipse> 回归夹子 (iface_wrap): 接口宿主**类型映射**的最小夹子, GA run 36732948503
    # (BalloonTooltips 红) 的直接切片。IFoo 显式 `Attribute VB_Creatable = True` ⇒ 进不了
    # ivref 登记表 (Pass A2 判据), `Dim f As IFoo` 必须经 Class/isInterface 分支发
    # `vb6_iface_IFoo` 包装类型 (阶段 3.6 对 Implements 宿主逐模块打标)。rev7 首版把
    # projectClassNameOf 短路放在符号真值之前, 把这路劫成裸 `vb6_cls_IFoo*` ⇒ 与 Set 侧
    # `vb6_iface_IFoo_wrap(...)` 对不上 → C2440 整工程编译红。`Set f = v` (Variant 往返)
    # 是 BalloonTooltips mdlSubclass.SubclassWnd 的原形态, Bar() 走接口 vtable 钉住运行期。
    $ifaceWrapExpected = @("IFACE-CALL-OK", "IFACE-DONE")
    Test-Vbp "iface_wrap" "$Tests\iface_wrap\IfaceWrap.vbp" $ifaceWrapExpected
    Test-Vbp "iface_wrap_x86" "$Tests\iface_wrap\IfaceWrap.vbp" $ifaceWrapExpected -Arch "x86"

    # <vbeclipse> 回归夹子 (ve_units): 控件**自己的坐标系**读数。两枚 UC 同尺寸、同字体,
    # 只差 .ctl 声明的 ScaleMode (1=缇 / 3=像素)。判据写成「缇型 = 像素型 x TwipsPerPixelX」
    # ⇒ 与 DPI 无关, 换机器不会漂。钉的是账 #175 (ScaleWidth) 与账 #177 (TextWidth/TextHeight)
    # 同一族口径: 控件宿主交出去的每一个量纲都得跟着它声明的 ScaleMode 走。
    # 改前实测: U-SW=True 而 U-TW/U-TH=False (文字量纲交的是设备像素)。
    # 账 #189 加的两条针 (ve_units 里塞一枚 UC 控件数组 uArr(0..2)): 整体成员 Count/LBound/UBound
    # 必须走 vb6_arr_<名>。修前这夹具**编不过** —— Count 发成对句柄变量的 COM 属性读, 跨窗体时
    # 是 undeclared identifier (C2065), 同窗体时恒答 0。负控实测 (a4e3b574 的编译器 x86):
    #   frmUnits.c(214)/(220): error C2065 "vb6_hwnd_uArr": 未声明的标识符, BUILD rc=1。
    $veUnitsExpected = @("U-SW=True", "U-TW=True", "U-TH=True", "U-CTX=True", "U-HW=True", "U-CNT=True",
        "U-ARR-RAW count=3 lb=0 ub=2", "U-ARRM-methods=3", "U-ARR=True",
        "U-ARREVT-RAW i1=1 h1=1 i2=2 h2=2 ret=7", "U-ARREVT=True",
        "U-ARRCLICK-RAW hw=True idx=2 hits=1 ret=0", "U-ARRCLICK=True",
        "U-ARRDBL-RAW hw=True idx=2 dbl=1 hits=0 ret=0", "U-ARRDBL=True",
        "U-ARREXT-RAW l=3600 t=1320 w=1200 h=1140", "U-ARREXT=True", "U-DONE")
    Test-Vbp "ve_units" "$Tests\ve_units\Units.vbp" $veUnitsExpected
    Test-Vbp "ve_units_x86" "$Tests\ve_units\Units.vbp" $veUnitsExpected -Arch "x86"

    # <vbeclipse> 回归夹子 (evtcase) 账 #190: 控件事件臂调用的函数名必须按 **Sub 自己的拼写** 发。
    # VB6 的标识符大小写不敏感、C 敏感: 以前臂里那个名字是拿控件的设计期拼写现拼的, 于是
    # "改了控件名没改过程名" (VB6 完全合法) 就变成引用一个没人定义的函数 —— 链接期 LNK2019。
    # 三枚控件: cmdRun/CmdRun_Click 与 chkOpt/ChkOpt_Click 只差大小写 (本账的两条臂),
    # cmdSame/cmdSame_Click 拼写一致 = **证人** (它排除"BM_CLICK 这条路自己没驱动起来"那种假红)。
    # 负控 = 修复前的编译器 (ea5155af) 真编译同一份夹具:
    #   CaseForm.obj : error LNK2019 无法解析的外部符号 _vb6_cmdRun_Click / _vb6_chkOpt_Click,
    #   fatal LNK1120: 2 个  —— 恰好两枚, 证人那枚没有第三条。
    $evtCaseExpected = @("EC01-RUN-CLICK", "EC02-CHK-CLICK chk=1", "EC03-SAME-CLICK",
        "EC-CNT run=1 chk=1 same=1", "EC-DONE")
    Test-Vbp "evtcase" "$Tests\evtcase\EvtCase.vbp" $evtCaseExpected
    Test-Vbp "evtcase_x86" "$Tests\evtcase\EvtCase.vbp" $evtCaseExpected -Arch "x86"

    # <vbeclipse> 回归夹子 (piccur) 账 #192: With 块里读写的那一枚必须是**它自己**的状态。
    # 真工程那一形 = Charts 2020/ucTreeMaps 的 PropPagFMR.pag:256-266（`With Picture1 : .CurrentX /
    # .CurrentY`）与 pag:618（`With lstFonts : Call .Clear`）—— 后端两张表从没登记过这三条，
    # 于是全部落到 cgen_expr_with.cpp 那条 "tempVar + "." + 成员名" 的兜底 = C2039。
    # 三头: PC01 写了读得回；PC02 **另一枚没被写坏**（RTL 按 HWND 存窗口属性，全局一份就露馅，
    # 同 #156 那条双向钉）；PC03 清空前那个数确实是 2（拦住"Clear 什么都不做、本来就 0"那种假绿）。
    # 负控 = 修复前那台编译器 (40bbef67) 真编同一份夹具：PicForm.c 150/151/160 三条
    #   error C2039 "CurrentX"/"CurrentY"/"Clear" 不是 "HWND__" 的成员，BUILD rc=1、一条读数都不出。
    $picCurExpected = @("PC01-CUR=True", "PC02-SEP=True aX=350 bX=0",
        "PC03-CLEAR=True before=2 after=0", "PC-DONE")
    Test-Vbp "piccur" "$Tests\piccur\PicCur.vbp" $picCurExpected
    Test-Vbp "piccur_x86" "$Tests\piccur\PicCur.vbp" $picCurExpected -Arch "x86"

    # 账 #197 (§B32): `VB6_ScaleMode` 这个窗口属性以前**读点齐全、写者为零** —— 缺省 1=缇，
    # 而 ScaleWidth / 控件几何 (#175 那一族换算) / 文字量纲 (#177) 全按这一档折算，
    # 所以它错起来不是"某一枚控件长歪"，是**量出来的数全是错的**。
    # 两头判据（同 #153/#154 那条口径：存请求值 + 一枚问窗口的证人）：
    #   SM01 设计期声明各自读回（修前三枚全答 1 = 缺省）；SM02/SM03 问窗口 —— 同尺寸的缇框
    #   量出来必须比像素框大一个单位比（比值只取 >3，不钉绝对数，DPI 变了也不假红）；
    #   SM04 运行期换档后两枚**逐数相等**（证明这数是活的，不是创建时烘死的）；
    #   SM05 另一枚没被写坏（RTL 按 HWND 存，全局一份就露馅，同 #192 双向钉）；
    #   SM06/SM07 Frame 里那枚 —— **第二条创建路**也发了这一句（账 #83/#151 那条"两条路都要打"）。
    # 负控 = 改前那台真编同一份夹具：SmForm.c 166/182 两条 error C2039 "ScaleMode" 不是 "HWND__"
    #   的成员 + C2198，rc=1、一条读数都不出（物证 .build/b197out/b64/c3-error.log）。
    $scaleModeExpected = @("SM01-MODE=3/3/1", "SM02-UNIT=True", "SM03-CHILD=True",
        "SM04-SWITCH=True", "SM05-OTHER=True", "SM06-NEST=True", "SM07-NESTCHILD=True",
        "SM-DONE")
    Test-Vbp "scalemode" "$Tests\scalemode\ScaleMode.vbp" $scaleModeExpected
    Test-Vbp "scalemode_x86" "$Tests\scalemode\ScaleMode.vbp" $scaleModeExpected -Arch "x86"

    # 账 #196 (绘图面那一半): RTL 早有一处「这枚控件的绘图 DC 从哪儿来」的口径 (Print/Cls 走它)，
    # 但**句柄交不回 VB 代码** —— `.hDC` 没出口，只能撞 cgen_expr_with.cpp 那条 "hwnd.成员" 兜底
    # = C2039（真工程物证 Charts 2020/ucTreeMaps 的 PropPagFMR.pag:258 `With Picture1 : TextOut .hDC`）。
    # 四头判据（缺一头就是假绿）：
    #   DS01 同一枚反复读 + With 那一形三个数彼此相等且非零（不缓存就每次换个句柄，没出口就全是 0）；
    #   DS02 两枚互不相等（出口写成全局一份当场红，同 #192/#156 那条双向钉）；
    #   DS03 GetDeviceCaps(hDC, LOGPIXELSX) > 0 —— 问的是 GDI，证明交回来是一张**活的 DC**，
    #        不是数字或野句柄（DS05 只钉前缀：dpi 是函数的数，不钉绝对值）；
    #   DS04 GetPixel 各自等于**自己那枚**的设计期底色（红/蓝）—— 这才证明这张 DC 指的是这一枚
    #        控件的表面，而不是屏幕或别人的窗口。底色取设计期值，不与"谁最后落色"抢次序。
    # 负控 = 改前那台真编同一份夹具: DcForm.c(167) error C2039 "hDC" 不是 "HWND__" 的成员，
    #   BUILD rc=1、一条读数都不出（物证 .build/b196neg/）。
    $dcSurfExpected = @("DS01-SAME=True", "DS02-SEP=True", "DS03-LIVE=True",
        "DS04-PIXEL=True", "DS05-RAW",
        "TH01-TWOFORMS=True", "TH02-UNITS=True", "TH03-FONT=True", "TH04-WIDTH=True",
        "TH05-RAW", "TH06-FONTRAW", "FR01-DEFAULT=True", "FR02-RAW",
        "SX10-FOURFORMS=True", "SX11-RAW", "DS-DONE")
    Test-Vbp "dcsurf" "$Tests\dcsurf\DcSurf.vbp" $dcSurfExpected
    Test-Vbp "dcsurf_x86" "$Tests\dcsurf\DcSurf.vbp" $dcSurfExpected -Arch "x86"

    # 账 #187: Charts 2020 的三枚 UC 子工程真编真链 (x64 + x86)，只要求"编得过、出得了 exe"，不跑。
    # 本轮实测字节数 (当前这台)：x64 696,320 / 576,000 / 576,000；x86 496,128 / 498,176，
    # ucChartBar x86 = .build/b191new/ucChartBar/Proyecto1.exe。
    # ucProgressCircular **刻意不列** —— 它今天还红着 (B29③④)，列进来就是把已知的红当门禁基线；
    # 等编过之后再加，那一加就是"这个工程从此不许退回编不过"。ucTreeMaps 原本也在这条黑名单里，
    # 账 #201 手写三枚 DI 桩之后它编得过、链得出 exe ⇒ 移到下面那组去了。
    Test-VbpBuild "charts_ucChartBar"  "$Tests\Charts 2020\ucChartBar\Proyecto1.vbp"
    Test-VbpBuild "charts_ucChartBar_x86" "$Tests\Charts 2020\ucChartBar\Proyecto1.vbp" -Arch "x86"
    Test-VbpBuild "charts_ucChartArea"  "$Tests\Charts 2020\ucChartArea\Proyecto1.vbp"
    Test-VbpBuild "charts_ucChartArea_x86" "$Tests\Charts 2020\ucChartArea\Proyecto1.vbp" -Arch "x86"
    Test-VbpBuild "charts_ucPieChart"  "$Tests\Charts 2020\ucPieChart\Proyecto1.vbp"
    Test-VbpBuild "charts_ucPieChart_x86" "$Tests\Charts 2020\ucPieChart\Proyecto1.vbp" -Arch "x86"
    # 账 #201: ucTreeMaps 是 #187 立这条门禁时**刻意不列**的那两枚之一（当时卡在 B29③④ 与
    # #192/#196 的编译错上）。现在两刀的编译面全清、只剩三枚 DI 桩没登记 ⇒ 手写单桩之后
    # 第一次真编真链出 exe（本轮实测 x64 659,968 / x86 562,688 字节，.build/b201out/）。
    # 这一加就是那句口径："这个工程从此不许退回编不过"。
    Test-VbpBuild "charts_ucTreeMaps"  "$Tests\Charts 2020\ucTreeMaps\Proyecto1.vbp"
    Test-VbpBuild "charts_ucTreeMaps_x86" "$Tests\Charts 2020\ucTreeMaps\Proyecto1.vbp" -Arch "x86"

    # <vbeclipse> 回归夹子 (optdef) 账 #194: VB 的整数类型后缀是**词法**，不许抄进生成 C。
    # 语义层那份 Optional 默认值求值以前直接 return rawText，于是 `Optional ... As Long = 0&` 发成
    # `(*FontIndex) = 0&;` = C2059 (真工程证据: Charts 2020/ucTreeMaps 的 PropPagFMR.pag:740/886)。
    # 两头钉: OD-VAL 证明四个默认值各按声明落地 (十六进制那枚同时证明"按数值重打"没把值改错)，
    # OD-EXPL 证明显式实参照样赢 —— 只钉前头那条的话，"恒取默认值"也是绿的。
    # 负控 = 修复前的那台编译器 (ea3eeeec) 真编同一份夹具: OptForm.c 116/118/119 三条
    #   error C2059 语法错误 ";" + C2065 "H10"，BUILD rc=1、一条读数都不出。
    $optDefExpected = @("OD-RAW l=12 i=3 h=16 z=0", "OD-VAL=True",
        "OD-SET l=1 i=5 h=32 z=9", "OD-EXPL=True", "OD-DONE")
    Test-Vbp "optdef" "$Tests\optdef\OptDef.vbp" $optDefExpected
    Test-Vbp "optdef_x86" "$Tests\optdef\OptDef.vbp" $optDefExpected -Arch "x86"

    # <vbeclipse> 回归夹子 (erase_sub) 账 #186: `Erase m_items(1).bvData` —— 销毁 **UDT 数组那一格里**
    # 的动态数组。修前 parser 一律 VB2001 "Erase 不支持带下标的形式"，把真工程挡在最前面
    # (Charts 2020/ucTreeMaps 的 PropPagFMR.pag:720 `Erase m_tvFiles(lIndex).bvData` 就卡在这)。
    # 两头钉 (缺一头就是假绿):
    #   ①EA01-KEEP 别的格不受影响 (s0=30 / s2=50) —— 拦"静默退化成销毁整个数组";
    #   ②EA02-RECYCLE 那一格还能重新分配并用 (s1=20 len=1) —— 拦"没置 NULL"：ReDim 自己会先销毁旧数组,
    #     没置 NULL 就是双释放 (现场直接崩)；反过来"没真销毁"这里会读到残留的 40。
    # 负控 = 修复前的编译器 (a159b85d) 编同一份夹具: VB2001 + VB2003, BUILD rc=1, 一条读数都不出现。
    $eraseSubExpected = @("EA01-KEEP s0=30 s2=50", "EA02-RECYCLE s1=20 len=1",
        "EA-CNT keep_ok=True recycle_ok=True", "EA-DONE")
    Test-Vbp "erase_sub" "$Tests\erase_sub\EraseSub.vbp" $eraseSubExpected
    Test-Vbp "erase_sub_x86" "$Tests\erase_sub\EraseSub.vbp" $eraseSubExpected -Arch "x86"

    # <vbeclipse> 账 #212 的两面判据 (夹子 tests/pberr, 一枚 .cls + 一个 Sub Main): 运行时错误 9 从
    # **实例方法体内**抛出时 —— ① 方法自己写的 On Error GoTo 有没有接住它 (而不是越过它落到调用方,
    # 或干脆落不回); ② 抛过之后那枚对象的成员还读不读得出数 (跳走时实例栈必须已经解链, 否则现场就是
    # 「错误被吃掉了」那一族: 错误还在、对象已坏)。来源是 #209/#214: 越界元素访问改成抛 9 之后,
    # VbEclipse 那批调用点全在 UC 的实例方法里 (`With m_Serie(Index)` 取的就是 &(me->m_arr(k)) 当对象),
    # 而此前只量过「没处理器 ⇒ Unhandled error 9 + 进程按 9 退出」这一面。
    # PE-1d/PE-2d/PE-null = 处理器在同一枚方法里 (一维越界 / 二维越界 / 从没 ReDim 的成员动态数组);
    # PE-caller = 方法内**没有**处理器, 9 必须继续往上交回调用方的处理器; PE-state = 三次抛出之后
    # 那枚对象的成员照旧读得出 20-40。真 VB6 五个数全同 (9/9/9/9/20-40) —— 实测两台逐行相同,
    # 所以这一格没有产品改动, 是把「不回归」钉住的一枚夹子 (见 §B46 末段与 §B47)。
    $pbErrExpected = @("PE-1d=9", "PE-2d=9", "PE-null=9", "PE-caller=9", "PE-state=20-40", "PE-DONE")
    Test-Vbp "pberr" "$Tests\pberr\PbErr.vbp" $pbErrExpected
    Test-Vbp "pberr_x86" "$Tests\pberr\PbErr.vbp" $pbErrExpected -Arch "x86"

    # 账 #186 的**边界**负例: `Erase m_list(1)` 这形 VB6 只在元素是 Variant(装着数组) 时合法,
    # 本仓没那条通路 —— 静默降级成"销毁整个数组"会改变语义, 所以必须继续报诊断。
    # 针面取消息里的英文片段 (GBK 控制台下中文会被折行/转码, 与既有 Test-CompileFail 同一口径)。
    Test-CompileFail "erase_neg_indexed" @("$Tests\erase_neg\EraseNeg.bas") `
        "indexed Erase target not supported here"

    # 发码形状针 (账 #186): 元素成员数组的销毁必须走 VB6_SA_AT(...) 那一形, 不能退回整数组,
    # 也不能漏掉置 NULL 那半 (双释放探测器)。
    Test-CodegenNote "erase_member_shape" @("$Tests\erase_sub\EraseSub.vbp") @(
        "vb6_SafeArrayDestroy1D(VB6_SA_AT(vb6_type_TElem, m_items, 1).bvData); VB6_SA_AT(vb6_type_TElem, m_items, 1).bvData = NULL;") @(
        "vb6_SafeArrayDestroy1D(m_items);")

    # test_vbman 用于验证外部 COM 组件 VBMANLIB (x86 DLL, 供 32 位程序调用)
    # ai/022 B07b: INH2..INH11 cover the merged member face + prefix-copied fields +
    # forwarding stubs (private Long/UDT/BSTR fields, Optional params, Property Get/Let,
    # 3-level chain, child-wins shadowing, base/derived instance isolation).
    # ai/022 B09b: run the same project on BOTH architectures. An inherited field is embedded by
    # value, so the derived struct's prefix must be byte-exact against the base struct; a wrong
    # field type is invisible on x64 when sizeof(void*) happens to equal that type's size, and only
    # the x86 build/run catches it (022 D39: INH35/INH36/INH39 failed on x86 for exactly that).
    $inhExpected = @(
        "INH0:derived", "INH1:OK", "INH2:OK", "INH3:OK", "INH4:OK", "INH5:OK",
        "INH6:OK", "INH7:OK", "INH8:OK", "INH9:OK", "INH10:OK", "INH11:OK",
        "INH12:OK", "INH13:OK", "INH14:OK", "INH15:OK", "INH16:OK",
        "INH17:OK", "INH18:OK", "INH19:OK", "INH20:OK", "INH21:OK",
        "INH22:OK", "INH23:OK", "INH24:OK", "INH25:OK", "INH26:OK",
        # ai/022 B09: INH44/INH45 = MyBase de-virtualized (the same members answer "derived"
        # through Me./obj.), INH49/INH50 = construction chain root->leaf + MyBase.Class_Initialize,
        # INH51 = Overrides returning a project class (com_entry.c forward-decl ordering).
        "INH44:OK", "INH45:OK", "INH46:OK", "INH47:OK", "INH48:OK", "INH49:OK", "INH50:OK",
        "INH51:OK", "INH52:OK",
        # ai/022 B09c: INH53/INH54 = `Set MyBase.<Property Set>` target side, INH55 = property
        # write through a UDT object field, INH56/INH57 = immediate-base Class_Initialize,
        # INH58 = root's Private UDT field used two levels down (the x86 stride case).
        "INH53:OK", "INH54:OK", "INH55:OK", "INH56:OK", "INH57:OK", "INH58:OK", "INH59:OK")
    Test-Vbp "cls_inh_pair" "$Tests\cls_inh\Inh.vbp" $inhExpected
    Test-Vbp "cls_inh_x86" "$Tests\cls_inh\Inh.vbp" $inhExpected -Arch "x86"
    # --- ai/022 B13a: the ActiveX DLL pipeline enters the gate for the first time ---
    # Both projects under tests\test_activex_dll have existed since P6 but were never
    # registered, so nothing in the regression had ever LINKED a .dll -- every claim on
    # the COM server side was unmeasured. These three cases make that surface observable
    # (product + exports + the generated coclass table) without changing compiler code.
    Test-VbpDll "ax_dll_calc" "$Tests\test_activex_dll\test_activex_dll.vbp" @(
        "const vb6_CoClassDesc g_vb6_coclasses[]",
        '"TestAXDLL.Calc"',
        '"{D84F362F-8EF1-D16D-8814-C16ADB700BAB}"',
        'L"SetValue", 1, 1',
        "DllGetClassObject", "DllRegisterServer")
    Test-VbpDll "ax_dll_event" "$Tests\test_activex_dll\test_event_dll.vbp" @(
        '"EventCalc"', '"{E1F2A3B4-C5D6-7890-ABCD-123456789ABC}"',
        "vb6_disp_EventCalc_Increment_invoke", "DllCanUnloadNow", "DllUnregisterServer")
    Test-VbpDll "ax_dll_calc_x86" "$Tests\test_activex_dll\test_activex_dll.vbp" @(
        "const vb6_CoClassDesc g_vb6_coclasses[]", '"TestAXDLL.Calc"',
        '"{D84F362F-8EF1-D16D-8814-C16ADB700BAB}"',
        "DllGetClassObject", "DllRegisterServer") -Arch "x86"
    # B13b: the two identity channels are merged for the DLL product. This case used to
    # PIN THE FORK (needle 0x0AD9CBC7 / absent CoDll.PG); flipping it is the batch's
    # acceptance evidence, so the needles are now exactly the other way round:
    #   - the group ProgID CoDll.PG is in the product (a second row, same CLSID) because the
    #     block writes [ComCreatable(True)] -- the first product-level consequence that bit has
    #   - IID_vb6iface_IProbe == the stage-2.7 value (0xF5CEF988), so the dllentry minter no
    #     longer answers for an interface the resolver already resolved
    #   - 0AD9CBC7 (what generateIid derived here before B13b) must be gone
    # The legacy <Proj>.<Class> row stays: existing DLL clients keep working.
    # B13e: IID_vb6def_CImpl is back to the *default dispinterface* GUID (0x7CA8CD81, written
    # back by the TypeLib builder) instead of the interface's own -- the server's member
    # surface (desc->methods) is that one, so table and typelib now advertise what they answer.
    Test-VbpDll "cc_dll_identity_single_source" "$Tests\cc_dll\CoDll.vbp" @(
        '"CoDll.CImpl"',
        '"CoDll.PG"',
        "const int g_vb6_coclassCount = 2;",
        "{11112222-3333-4444-5555-666677778888}",
        "0xF5CEF988",
        "0x7CA8CD81",
        # ai/022 B17: CImpl 多了个公有成员 Twice（外部晚绑定要有东西可点），methodCount 0→1。
        # 这条针同时钉住"表里的成员面与库里 `_CImpl` 那一档同源"（B13e 的广告==应答）。
        "1, /* methodCount */") @(
        "0AD9CBC7") @(
        "CoClass 'PG' identity: CLSID={11112222-3333-4444-5555-666677778888} (vbp)",
        "IID={F5CEF988-3217-6173-94B7-BB99C4B8CB81} (minted)",
        "ProgID=CoDll.PG (minted)",
        "impl='CImpl' comCreatable=True")

    # ai/022 B13c/B13e: two invariants read out of the *products* (dll_entry.c / CImpl.h /
    # CoDll.tlb): (甲) one interface == one GUID across table / vtable QI / typelib (B13c);
    # (乙) the typelib's coclass DEFAULT ref == the IID the server actually answers with, i.e.
    # the class's own default dispinterface (B13e, after B13c's redirect was reverted).
    # B15: 甲's third channel now reads the interface's own row (TKIND_INTERFACE, named IProbe)
    # instead of the empty `_IProbe` dispinterface, and the helper also asserts the two rows this
    # batch deleted stay deleted -- so reverting the shape turns this same case red.
    Test-TlbIdentitySingleSource "cc_dll_tlb_matches_table" "$Tests\cc_dll\CoDll.vbp" "CoDll" "CImpl" "IProbe" "{11112222-3333-4444-5555-666677778888}"
    # ai/022 B16: 库里那一档真接口的**成员面**。B15 只摆正了形状与身份（cFuncs 仍是 0），
    # 本批起它要如实发契约，两条读数（x64 / x86）各钉自己的槽偏移 —— 同一个 vtable，
    # x86 客户端按 4 字节一步读到的槽 3/4 在 12/16，x64 客户端在 24/32；
    # 「建库 flag 与目标位数不配」或「发成员却不发偏移」都会让其中一条红。
    Test-TlbIfaceContract "cc_dll_tlb_contract_x64" "$Tests\cc_dll\CoDll.vbp" "CoDll" @(
        "TYPE 0 kind=interface name=IProbe guid={F5CEF988-3217-6173-94B7-BB99C4B8CB81} cFuncs=2 cVars=0 cImplTypes=0 cbSizeInstance=8",
        "FUNC 0 memid=1 name=Ping invkind=1 funckind=purevirtual callconv=stdcall oVft=24 cParams=1 ret=0x0018",
        "PARAM 0 flags=0x0001 vt=0x0003",
        "FUNC 1 memid=2 name=get_Got invkind=2 funckind=purevirtual callconv=stdcall oVft=32 cParams=0 ret=0x0003") @(
        "name=IProbe guid={F5CEF988-3217-6173-94B7-BB99C4B8CB81} cFuncs=0",
        "oVft=12",
        "oVft=16")
    Test-TlbIfaceContract "cc_dll_tlb_contract_x86" "$Tests\cc_dll\CoDll.vbp" "CoDll" @(
        "TYPE 0 kind=interface name=IProbe guid={F5CEF988-3217-6173-94B7-BB99C4B8CB81} cFuncs=2 cVars=0 cImplTypes=0 cbSizeInstance=4",
        "FUNC 0 memid=1 name=Ping invkind=1 funckind=purevirtual callconv=stdcall oVft=12 cParams=1 ret=0x0018",
        "FUNC 1 memid=2 name=get_Got invkind=2 funckind=purevirtual callconv=stdcall oVft=16 cParams=0 ret=0x0003") @(
        "oVft=24",
        "oVft=32") "x86"
    # ai/022 B14: the DLL product finally gets a real caller. TestAXDLL.Calc is the legacy
    # face (Public members exist, so IDispatch must answer); cc_dll's CImpl only satisfies a
    # modern interface, so its IDispatch member surface must be EMPTY (B13c ruling (b)).
    # B16 DID flip the second half of that pin on purpose: the interface IID now answers with
    # the THIN pointer (same=no -- it is no longer the IDispatch wrapper), and the new
    # CREATE_IFACE/VTBL_* readings call its contract slots by the library's shape. The slot
    # numbers [3]=Ping, [4]=Got are the library's oVft=24/32 read back by tests\tools -- if the
    # row's shape and the shipped vtable ever drift apart, VTBL_GET_AFTER stops being 42.
    Test-DispatchInvoke "ax_dll_dispatch_invoke" "$Tests\test_activex_dll\test_activex_dll.vbp" `
        "test_activex_dll" "{D84F362F-8EF1-D16D-8814-C16ADB700BAB}" @(
        "CREATE hr=0x00000000 ptr=OK",
        "QI_IUNKNOWN hr=0x00000000 same=yes",
        "TYPEINFOCOUNT=1 hr=0x00000000",
        "CALL=ADD hr=0x00000000 result=42",
        "CALL=GETVALUE hr=0x00000000 result=7",
        "NAMES=bogus hr=0x80020006 dispid=-1")
    Test-DispatchInvoke "cc_dll_dispatch_iface_only" "$Tests\cc_dll\CoDll.vbp" `
        "CoDll" "{11112222-3333-4444-5555-666677778888}" @(
        "CREATE hr=0x00000000 ptr=OK",
        "QI_IUNKNOWN hr=0x00000000 same=yes",
        "EXTRA_IID={F5CEF988-3217-6173-94B7-BB99C4B8CB81}",
        "QI_EXTRA hr=0x00000000 same=no",
        "CREATE_IFACE hr=0x00000000 ptr=OK",
        "VTBL_GET_BEFORE=0",
        "VTBL_GET_AFTER=42") @(
        "CALL=ADD") "{F5CEF988-3217-6173-94B7-BB99C4B8CB81}"
    # ai/022 B16: 同一批断言在 x86 上再真跑一遍 —— 这是布局改动（vtable 槽 + 库里的偏移 +
    # 调用约定）唯一的 x86 侧端到端读数：x86 探针按库里那形状直调槽 3/4，
    # 若槽位置/调用约定/返回值任何一处对不上，VTBL_GET_AFTER 就不是 42。
    Test-DispatchInvoke "cc_dll_dispatch_iface_only_x86" "$Tests\cc_dll\CoDll.vbp" `
        "CoDll" "{11112222-3333-4444-5555-666677778888}" @(
        "CREATE hr=0x00000000 ptr=OK",
        "QI_IUNKNOWN hr=0x00000000 same=yes",
        "QI_EXTRA hr=0x00000000 same=no",
        "CREATE_IFACE hr=0x00000000 ptr=OK",
        "VTBL_GET_BEFORE=0",
        "VTBL_GET_AFTER=42",
        "DONE") @(
        "CALL=ADD") "{F5CEF988-3217-6173-94B7-BB99C4B8CB81}" "x86"

    # --- ai/022 B17: 外部激活（注册 -> 系统那条路 -> 反注册后不留键）---
    # 上面两条走的是**进程内**独立客户端（LoadLibrary + DllGetClassObject，刻意不查注册表）；
    # 这两条走的是系统那条路：DllRegisterServer 真写注册表，再按 CLSID/ProgID 让 COM 自己去找
    # InprocServer32、装载 DLL —— 也就是 CreateObject / CoCreateInstance 客户实际走的链路。
    # 断言的形状与理由（读数见 ai/022 D64）：
    #   - REG_* 四项 + REG_TYPELIB_PATH：CLSID/ProgID/InprocServer32/ThreadingModel 与
    #     "按 LIBID 找得到类型库"都成立（B17 修掉两条真缺陷才走到这一步：rc.exe 发现面太窄
    #     ⇒ 库根本没嵌进 DLL；反注册的实参顺序写反 ⇒ 整棵 TypeLib 键静默留着）
    #   - NAMES=Twice + CALL=Twice result=42：晚绑定按名调用**类的公有成员**（CreateObject 那半）
    #   - NAMES=Ping hr=0x80020006：契约成员是 Private，类的默认面上点不到 —— 这一条是
    #     **VB6 语义的应有读数**，不是缺陷；接口成员走下面的早绑定那条
    #   - COCREATE_IFACE + VTBL_GET_AFTER=42：接口 IID 的 QI/激活交薄指针，按库里 oVft 直调契约槽
    #   - CLEAN_*=gone：反注册之后 CLSID/ProgID/TypeLib 三类键都不留（用例可重复跑、不脏机器）
    $comActNeedles = @(
        "REGSVR hr=0x00000000",
        "REG_INPROC_PATH=OK",
        "REG_THREADING=OK Apartment",
        "REG_PROGID_CLSID=OK {11112222-3333-4444-5555-666677778888}",
        "REG_TYPELIB_PATH=OK",
        "PROGID_LOOKUP hr=0x00000000 same=yes",
        "COCREATE_DISP hr=0x00000000 ptr=OK",
        "NAMES=Twice hr=0x00000000 dispid=1",
        "CALL=Twice hr=0x00000000 result=42",
        "NAMES=Ping hr=0x80020006 dispid=-1",
        "COCREATE_IFACE hr=0x00000000 ptr=OK",
        "VTBL_GET_AFTER=42",
        "UNREGSVR hr=0x00000000",
        "CLEAN_CLSID=gone",
        "CLEAN_PROGID=gone",
        "CLEAN_TYPELIB=gone",
        "DONE")
    Test-ComActivate "cc_dll_external_activate" "$Tests\cc_dll\CoDll.vbp" "CoDll" `
        "{11112222-3333-4444-5555-666677778888}" "CoDll.CImpl" `
        "{F5CEF988-3217-6173-94B7-BB99C4B8CB81}" $comActNeedles
    Test-ComActivate "cc_dll_external_activate_x86" "$Tests\cc_dll\CoDll.vbp" "CoDll" `
        "{11112222-3333-4444-5555-666677778888}" "CoDll.CImpl" `
        "{F5CEF988-3217-6173-94B7-BB99C4B8CB81}" $comActNeedles "x86"
    # 真正的外部客户是个**别的进程**：C3 编译的 `tests\cc_dll_client` 不引用 DLL，
    # CreateObject + 按名调用全走注册表 → IDispatch 晚绑定（test_p613_typelib 那一族的做法）。
    Test-ComActivateClient "cc_dll_late_client" "$Tests\cc_dll\CoDll.vbp" "CoDll" `
        "{11112222-3333-4444-5555-666677778888}" "CoDll.CImpl" `
        "$Tests\cc_dll_client\LateClient.vbp" @("EXT1:OK", "EXT2:OK", "EXT-DONE")

    # --- ai/022 B18 端到端示例（收口批）: 同一个源集合编成 EXE 与 DLL 两种形态 ---
    # EXE 形态: 语言层全用一遍（Interface/Implements(+Via 委托)/Inherits/Overrides/Protected/
    # MyBase/CoClass 块/`As <块名>`/`New <块名>`/工程内 CreateObject 改写/TypeOf），
    # DEMO1..DEMO12 各自钉一个行为；x64 与 x86 各跑一遍（继承来的字段与槽布局只有 x86 才暴露）。
    # DLL 形态: 同一个源集合 + [ComCreatable(True)] 的块，注册后由**另一个进程**的 C3 客户
    # CreateObject 激活（B17 的外部那条路）。
    $demoNeedles = @(
        "DEMO1:OK", "DEMO2:OK", "DEMO3:OK", "DEMO4:OK", "DEMO5:OK", "DEMO6:OK",
        "DEMO7:OK", "DEMO8:OK", "DEMO9:OK", "DEMO10:OK", "DEMO11:OK", "DEMO12:OK", "DEMO-DONE")
    Test-Vbp "cc_demo_exe" "$Tests\cc_demo\DemoExe.vbp" $demoNeedles
    Test-Vbp "cc_demo_exe_x86" "$Tests\cc_demo\DemoExe.vbp" $demoNeedles -Arch "x86"
    Test-ComActivateClient "cc_demo_dll_external" "$Tests\cc_demo\DemoDll.vbp" "DemoDll" `
        "{993BE038-BBA4-7804-FEB0-E65927384CA7}" "DemoDll.Shape" `
        "$Tests\cc_demo\DemoClient.vbp" @("DEMOEXT1:OK", "DEMOEXT-DONE")
    Test-ComActivateClient "cc_demo_dll_external_x86" "$Tests\cc_demo\DemoDll.vbp" "DemoDll" `
        "{993BE038-BBA4-7804-FEB0-E65927384CA7}" "DemoDll.Shape" `
        "$Tests\cc_demo\DemoClient.vbp" @("DEMOEXT1:OK", "DEMOEXT-DONE") "x86"
    # --- ai/022 B19: 控制台输出的编码（四语种）---
    # 控制台那条路走 WriteConsoleW，渲染与 chcp 无关（实测 936/65001/437 逐字相同）；
    # 重定向那条走“控制台代码页”的字节（装不下才退回 UTF-8）。两条都真跑真读。
    $cnSample = Join-Path $Tests "cc_cn\CnMain.bas"
    $cnNeedles = @("中文", "日本語 テスト", "한국어 테스트", "English test")
    Test-CnConsoleOutput "cc_cn_console" $cnSample $cnNeedles
    Test-CnConsoleOutput "cc_cn_console_x86" $cnSample $cnNeedles "x86"
    Test-CnRedirectOutput "cc_cn_redirect_gbk" $cnSample 936
    Test-Vbp "test_vbman" "$Tests\test_vbman\test_vbman.vbp" @("P24-04a:OK", "P24-04b:OK", "P24-04:2/2") -Arch "x86" -RequiresCom "VBMANLIB.cVBMAN"
    # ai/029 C29-9b: 这两条放在**整组最后**。它们会多开两个窗体 + 真模态对话框，而
    # frmevents 的拖放点是**按窗口位置现算**的 —— 实测每次启动级联偏移约 26 px
    # (单跑 X=-40；把我的探针用例跑在它前面 → -118；连跑三次 → -144/-170/-196)。
    # 偏移累积到 GA 的桌面几何上就足以让拖放落不进目标窗 ⇒ EV24/EV25 整条不出现。
    # 跑在最后 = 我引入的偏移不再影响任何用例。测试本体的脆弱点(没把窗口位置钉住)
    # 不在本批范围，已记进 029 §九 C29-9b 那一格。
    # ai/029 C29-9b：上面那两条**不弹框**（恒假守卫），真弹框由这两条补 —— 环境变量
    # C3_CDPROBE=1 才走弹框那条路（夹具里 `If Environ("C3_CDPROBE")="1"`），RTL 侧的一次性
    # 线程只认本线程创建的 #32770，发 WM_COMMAND/IDCANCEL 等价于"用户点了取消"。于是
    # DL11(取消报 32755) / DL12(取消不改进数) / DL13(模态循环真跑过 ≥30ms) / DL14
    # (CancelError=False 时静默返回) 四条能断。不设 env 时这四行根本不打印，上面那 10 条
    # 的形状逐字不变 —— 卡死风险也只在这两条里，而它们由 -RunTimeoutSec 兜底。
    $dlProbeNeedles = @("CTRLDLG-DONE") + (1..14 | ForEach-Object { "DL$_=Y" })
    Test-Vbp "ctrldlg_probe" "$Tests\ctrldlg\DlApp.vbp" $dlProbeNeedles -Env "C3_CDPROBE=1"
    Test-Vbp "ctrldlg_probe_x86" "$Tests\ctrldlg\DlApp.vbp" $dlProbeNeedles -Env "C3_CDPROBE=1" -Arch "x86"

    # ai/029 C29-WS-a/b/c/d：Winsock 走原生 Winsock2（不加载 MSWINSCK.OCX）。判据 = 同进程几枚控件
    # 的 **UDP 回环一来一回** + **TCP 一整轮**（Listen / ConnectionRequest / Accept / Connect），
    # 端口一律交给系统挑：Bind 0 后读 LocalPort；地址写死 127.0.0.1 / localhost，
    # 所以既不碰外网、也不跟 CI 上别的作业抢固定端口。等事件一律"一步一个 Timer tick"
    # （第一版拿 DoEvents 连泵 60 次等包到，本机就假红过一次 —— 见 029 §九 本格）。
    # 注册放在整个 vbp 块**最后**：这条会真的建窗（不可见的身份窗 + 一枚窗体），而 C29-9b
    # 量过"多开一窗就让后面的按位置算点心的用例翻红"，排最后就不会再影响任何用例。
    $wsNeedles = @("CTRLWINSOCK-DONE") + (1..65 | ForEach-Object { "WS$_=Y" })
    Test-Vbp "ctrlwinsock" "$Tests\ctrlwinsock\WsApp.vbp" $wsNeedles
    Test-Vbp "ctrlwinsock_x86" "$Tests\ctrlwinsock\WsApp.vbp" $wsNeedles -Arch "x86"
    # 发码正面：类名 + 不可见 0x0 的创建参数（与 Timer 同一枚 style 值）+ 设计期 Create/Init
    # 三条 + 事件回调注册的序号 + 四条方法的发码形状（含 GetData 出参取址、缺省 type/maxLen）。
    Test-EmitcShape "ws_emitc_shape" @("$Tests\ctrlwinsock\WsApp.vbp") @(
        '"VB6_WINSOCK", "",',
        '1140850688L, 0L,',
        'vb6_RegisterWinsockClass((void*)hInstance);',
        'vb6_Ws_Create((void*)vb6_hwnd_wsA);',
        'vb6_Ws_SetProtocol((void*)vb6_hwnd_wsA, 1);',
        'vb6_Ws_SetProtocol((void*)vb6_hwnd_wsC, 0);',
        'vb6_Ws_SetRemoteHost((void*)vb6_hwnd_wsA, L"");',
        'vb6_Ws_SetEventHandler((void*)vb6_hwnd_wsA, 2, (void*)vb6_wsA_DataArrival); }',
        'vb6_Ws_SetEventHandler((void*)vb6_hwnd_wsB, 6, (void*)vb6_wsB_StateChanged); }',
        'vb6_Ws_Bind((void*)vb6_hwnd_wsA, (int32_t)0, L"");',
        'vb6_Ws_SendData((void*)vb6_hwnd_wsA, vb6_BSTR_FromStr(L"hello-B"));',
        'vb6_Ws_GetData((void*)vb6_hwnd_wsB, &gGotB, 0, (-1));',
        'vb6_Ws_PeekData((void*)vb6_hwnd_wsB, &gPeek, 0, (-1));',
        'vb6_Ws_Close((void*)vb6_hwnd_wsA);',
        'vb6_Ws_Listen((void*)vb6_hwnd_wsS);',
        'vb6_Ws_Connect((void*)vb6_hwnd_wsT);',
        'vb6_Ws_Accept((void*)vb6_hwnd_wsS, (int32_t)gReq1);',
        'vb6_Ws_SetEventHandler((void*)vb6_hwnd_wsS, 3, (void*)vb6_wsS_ConnectionRequest); }',
        'vb6_Ws_GetState(vb6_hwnd_wsC',
        'vb6_Ws_GetLocalPort(vb6_hwnd_wsB',
        # Error 那一格七参数，形状本身就是一条针：ByRef 的两格（Description / CancelDisplay）
        # 必须是指针 —— 写成按值在 x64 上照样"读得像对的"，只有这条发码针 + x86 真跑拦得住。
        'static void vb6_wsD_Error(int16_t Number, BSTR* Description, int32_t Scode, BSTR Source, BSTR HelpFile, int32_t HelpContext, int16_t* CancelDisplay)',
        'vb6_Ws_SetEventHandler((void*)vb6_hwnd_wsD, 7, (void*)vb6_wsD_Error); }',
        # WS-d：发送侧两条事件的注册序号（4 = SendComplete、5 = SendProgress）
        'vb6_Ws_SetEventHandler((void*)vb6_hwnd_wsT, 4, (void*)vb6_wsT_SendComplete); }',
        'vb6_Ws_SetEventHandler((void*)vb6_hwnd_wsT, 5, (void*)vb6_wsT_SendProgress); }',
        'vb6_Ws_SetEventHandler((void*)vb6_hwnd_wsE, 3, (void*)vb6_wsE_ConnectionRequest); }',
        'vb6_Ws_GetData((void*)vb6_hwnd_wsB, &gBytes, (int32_t)8209, (-1))',
        'vb6_Ws_GetData((void*)vb6_hwnd_wsB, &gPart, (int32_t)8209, (int32_t)3)',
        'vb6_Ws_Accept((void*)vb6_hwnd_wsE, (int32_t)gReqEId2)',
        'vb6_Ws_Accept((void*)vb6_hwnd_wsE, (int32_t)gReqE3)',
        'vb6_Ws_Connect((void*)vb6_hwnd_wsT2)',
        'vb6_Ws_GetData((void*)vb6_hwnd_wsT2, &gGotG, 0, (-1))',
        'vb6_Ws_SetEventHandler((void*)vb6_hwnd_wsE, 1, (void*)vb6_wsE_Close); }'
    )
    # 反面：一条都不许落回"把 HWND 当 IDispatch 用"那条假路（本线踩过三次的那声不响），
    # LocalPort / LocalIP 也**不许有写口** —— 那两格在 VB6 就是运行期只读（端口归 Bind 管、
    # 地址归系统定），发一条"写得动但什么都不改"的 setter 比不发更难查（RT-a 的 ScrollBars 同口径）。
    Test-EmitcAbsent "ws_emitc_no_com_fallback" @("$Tests\ctrlwinsock\WsApp.vbp") @(
        'vb6_ComCall(vb6_hwnd_wsA, L"close"',
        'vb6_ComCall(vb6_hwnd_wsA, L"SendData"',
        'vb6_ComCall(vb6_hwnd_wsB, L"GetData"',
        'vb6_ComCall(vb6_hwnd_wsS, L"Listen"',
        'vb6_ComCall(vb6_hwnd_wsT, L"Connect"',
        'vb6_ComCall(vb6_hwnd_wsS, L"Accept"',
        'vb6_ComGetObjectProp(vb6_hwnd_wsB, L"LocalPort")',
        'vb6_Ws_SetLocalPort',
        'vb6_Ws_SetLocalIP',
        'vb6_Ws_TraceCmd',
        'vb6_ComCall(vb6_hwnd_wsD, L"Bind"'
    )
    $vbpSw.Stop()
    Write-Host "  (vbp/gui tests took $([Math]::Round($vbpSw.Elapsed.TotalSeconds))s)"
# 分片时把"本片跑了几个/让给别人几个"打出来: 四个 shard 的 ran+skipped 之和必须等于不分片时的
# TOTAL。哪一片的数对不上, 一眼能看出来 —— 分片漏跑是静默的, 没有这行就只能靠猜。
if ($VbpShardTotal -gt 1) {
    Write-Host "  (vbp shard ${VbpShard}/${VbpShardTotal}: ran $script:VbpCaseNo, skipped $script:VbpSkipped)" -ForegroundColor Cyan
}
    Write-Host ""
}

if ($Category -in @("all", "compile")) {
    # --- 静态对表: DI 桩 census (只许增, 不许静默减) ---
    # 见 scripts/check_di_stubs.ps1 头注释: 重生成 di 桩是按"跑它那次会话的引用面"发桩的,
    # 换个工程会话重跑就会**静默丢掉**上次的桩 (本批实测丢 9 个 + winspool.lib), 而只有
    # 某个工程恰好 Declare 到它时才在链接期以 LNK2019 暴露 —— 那是门里"看运气"才发现的红。
    # 这里改成每次 compile 段都静态对一遍基线 (不编译、不跑程序)。
    Write-Host "--- Static Checks ---" -ForegroundColor Yellow
    Test-DiStubCensus
    Test-UcScaleUnitsCensus
    Test-HostPseudoTableCensus
    Test-AddressOfThunkSites
    Test-FloatLiteralShape
    Test-IntLiteralShape
    Test-IntSuffixSites
    Test-ScaleModeWriters
    Test-ControlDC
    Test-SaAccess
    Test-BitwiseAuthority
    Test-BodyDeclShape
    Test-CaseIsShape
    Test-DocHostAuthority
    Test-RtlNakedNames
    Test-RtlResourceIds
    Test-UcArrayEventSites
    Test-SubclassSlotSites
    Test-CtrlArrayMemberSites
    Test-CtrlPropTypeAuthority
    Test-FormDrawState
    Test-FixtureTimerClose
    Test-VariantCmpBoxing
    Test-EventHandlerNames
    Test-UcInstanceExit

    Write-Host ""

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
    # tB extension (ai/022 B01): Interface contract-block diagnostics must fire.
    $itfNeg = @(
        @("itf_n01_member_body", "$Tests\itf_neg\n01_member_body.bas", "must not contain an implementation body"),
        @("itf_n02_visibility", "$Tests\itf_neg\n02_visibility.bas", "must not carry an access modifier"),
        @("itf_n03_field", "$Tests\itf_neg\n03_field.bas", "accepts only Sub/Function/Property signatures"),
        @("itf_n04_missing_end", "$Tests\itf_neg\n04_missing_end.bas", "expected 'End Interface'"),
        @("itf_n05_unknown_attr", "$Tests\itf_neg\n05_unknown_attr.bas", "Unrecognized attribute line [NotAnAttr]"),
        @("itf_n06_attr_no_target", "$Tests\itf_neg\n06_attr_no_target.bas", "Attribute line must precede an Interface or CoClass declaration"),
        @("itf_n07_generic", "$Tests\itf_neg\n07_generic.bas", "does not support generic type parameters"),
        # ai/022 B02 (semantic layer): stage 2.7 contract registry + Implements checks
        @("itf_n08_missing_slot", "$Tests\itf_neg\n08_missing_slot.cls", "is not implemented by class"),
        @("itf_n09_sig_mismatch", "$Tests\itf_neg\n09_sig_mismatch.cls", "signature mismatch"),
        @("itf_n10_unknown_parent", "$Tests\itf_neg\n10_unknown_parent.bas", "extends unknown interface"),
        @("itf_n11_extends_cycle", "$Tests\itf_neg\n11_extends_cycle.bas", "Circular Extends chain"),
        @("itf_n12_dup_member", "$Tests\itf_neg\n12_dup_member.bas", "cannot be overloaded"),
        @("itf_n13_module_collision", "$Tests\itf_neg\n13_module_collision.bas", "collides with a module of the same name"),
        # ai/022 B02b: member-level Implements I.M[, I.N] trailing clause
        @("itf_n14_clause_no_member", "$Tests\itf_neg\n14_clause_no_member.cls", "has no member"),
        @("itf_n15_clause_not_implemented", "$Tests\itf_neg\n15_clause_not_implemented.cls", "does not implement interface"),
        @("itf_n16_clause_in_bas", "$Tests\itf_neg\n16_clause_in_bas.bas", "only valid in a class module"),
        @("itf_n17_clause_unqualified", "$Tests\itf_neg\n17_clause_unqualified.bas", "needs a qualified name"),
        @("itf_n18_clause_sig_mismatch", "$Tests\itf_neg\n18_clause_sig_mismatch.cls", "signature mismatch"),
        @("itf_n19_iface_in_generic", "$Tests\itf_neg\n19_iface_in_generic.cls", "not allowed inside a generic class template"),
        # ai/022 B03: Interface head-line host form (.cls named after its single block)
        @("itf_n20_host_extra_decl", "$Tests\itf_neg\n20_host_extra_decl.cls", "may contain only the Interface block"),
        # ai/022 B10 (Implements .. Via): the delegate clause is resolved in stage 2.7
        # Pass D, and both rejection reasons are error-level.
        @("itf_n21_via_no_field", "$Tests\itf_neg\n21_via_no_field.cls", "is not a module-level field"),
        @("itf_n23_via_not_iface", "$Tests\itf_neg\n23_via_not_iface.cls", "not an Interface block"),
        @("itf_n24_via_in_bas", "$Tests\itf_neg\n24_via_in_bas.bas", "only allowed in a class module"),
        # ai/022 B11/C01 (CoClass block, syntax layer): malformed clauses must be refused
        # at the parser -- C01 does no validation beyond block structure.
        @("itf_n25_coclass_no_name", "$Tests\itf_neg\n25_coclass_no_name.bas", "expected CoClass name"),
        @("itf_n26_coclass_missing_end", "$Tests\itf_neg\n26_coclass_missing_end.bas", "expected 'End CoClass'"),
        @("itf_n27_coclass_bad_member", "$Tests\itf_neg\n27_coclass_bad_member.bas", "CoClass block accepts only attribute lines"),
        @("itf_n28_coclass_ref_no_name", "$Tests\itf_neg\n28_coclass_ref_no_name.bas", "expected interface name after 'Interface' in CoClass block"),
        # ai/022 B11/C03a (CoClass shape + name validation, stage 2.7 Pass F): every refused
        # shape must report its own reason, and none of them may be silent any more.
        @("itf_n29_coclass_dup_name", "$Tests\itf_neg\n29_coclass_dup_name.bas", "is declared twice"),
        @("itf_n30_coclass_unknown_iface", "$Tests\itf_neg\n30_coclass_unknown_iface.bas", "is not an Interface block in this project"),
        @("itf_n31_coclass_two_defaults", "$Tests\itf_neg\n31_coclass_two_defaults.bas", "marks 2 interfaces [Default]"),
        @("itf_n32_coclass_dup_entry", "$Tests\itf_neg\n32_coclass_dup_entry.bas", "more than once (the contract set is a set)"),
        @("itf_n33_coclass_impl_missing", "$Tests\itf_neg\n33_coclass_impl_missing.bas", "is not a class module of this project"),
        @("itf_n34_coclass_exe_creatable", "$Tests\itf_neg\n34_coclass_exe_creatable.bas", "marks [ComCreatable(True)] in an EXE project"),
        # ai/022 B11/C03b (CoClass contract aggregation, stage 3.4c): the block binds a
        # class, and every slot of every listed interface must be met by that class or an
        # ancestor -- same VB3012/VB3017 family the Implements checker uses.
        @("itf_n37_coclass_missing_slot", "$Tests\itf_neg\n37_coclass_missing_slot.cls", "is not implemented by class"),
        @("itf_n38_coclass_sig_mismatch", "$Tests\itf_neg\n38_coclass_sig_mismatch.cls", "signature mismatch (interface:")
    )
    foreach ($c in $itfNeg) {
        if (Test-Path $c[1]) { Test-SyntaxFail $c[0] $c[1] $c[2] }
        else { Write-Host "  [SYNTAX-FAIL] $($c[0]) ... SKIP (missing case file)" -ForegroundColor DarkGray }
    }
    # ai/022 B10: the Via holder field must name a class that implements the interface
    # itself -- a same-named member is not enough (there is no slot field to delegate to).
    Test-SyntaxFailMulti "itf_n22_via_holder_not_impl" @("$Tests\itf_neg\n22_via_base.cls", "$Tests\itf_neg\n22_via_deleg.cls") "does not implement interface"
    # ai/022 B11/C03a: two more refusals only exist once a second module is in scope --
    # a class used as an interface (VB6 habit), and a block name shadowing another module.
    Test-SyntaxFailMulti "itf_n35_coclass_legacy_cls" @("$Tests\itf_neg\n35_legacy_cls_iface.bas", "$Tests\itf_neg\n35_cls.cls") "but that is a class module"
    Test-SyntaxFailMulti "itf_n36_coclass_vs_module" @("$Tests\itf_neg\n36_coclass_vs_module.bas", "$Tests\itf_neg\n36_other.bas") "collides with a module of the same name"
    # ai/022 B11/C03b: inheriting a CoClass block name was already refused, but the sentence
    # blamed a name that does exist in the project; the new wording names the real mistake.
    Test-SyntaxFailMulti "itf_n39_inherits_coclass" @("$Tests\itf_neg\n39_coclass_as_base.bas", "$Tests\itf_neg\n39_coclass_as_base_der.cls") "which is a CoClass block"
    # <vbeclipse>: `UBound(vbNull)` 首参是 VarType 常量 —— 真 VB6 编译期就拒 (提示缺少数组)。
    if (Test-Path "$Tests\arr_neg\n01_ubound_vbnull.bas") { Test-SyntaxFail "arr_n01_ubound_vbnull" "$Tests\arr_neg\n01_ubound_vbnull.bas" "VB3043" }
    if (Test-Path "$Tests\arr_neg\n02_join_vbnull.bas") { Test-SyntaxFail "arr_n02_join_vbnull" "$Tests\arr_neg\n02_join_vbnull.bas" "VB3043" }
    # Positive guard (ai/022 B02): a class satisfying a new-style contract (
    # Extends-inherited slot + property tri-slot keys) must stay silent.
    if (Test-Path "$Tests\itf_pos\p01_contract_ok.cls") { Test-Syntax "itf_p01_contract_ok" "$Tests\itf_pos\p01_contract_ok.cls" }
    # ai/022 B02b positive guard: explicit clause binding (cross-interface slot, inherited
    # slot named via child interface, property tri-slot keys, arbitrary member names).
    if (Test-Path "$Tests\itf_pos\p02_clause_binding.cls") { Test-Syntax "itf_p02_clause_binding" "$Tests\itf_pos\p02_clause_binding.cls" }
    # ai/022 B03 positive guard: a host module name is no longer a name collision.
    if (Test-Path "$Tests\itf_pos\p03_headline_host.cls") { Test-Syntax "itf_p03_headline_host" "$Tests\itf_pos\p03_headline_host.cls" }
    # ai/022 B10 positive guard: Via is registered as a SOFT keyword, so an existing
    # program that uses Via as a variable name still parses.
    if (Test-Path "$Tests\itf_pos\p04_via_soft_ident.bas") { Test-Syntax "itf_p04_via_soft_ident" "$Tests\itf_pos\p04_via_soft_ident.bas" }
    # ai/022 B11/C01 positive guards: the CoClass block form parses end to end, and CoClass
    # stays usable as an ordinary identifier (soft keyword registration).
    if (Test-Path "$Tests\itf_pos\p05_coclass_block.bas") { Test-Syntax "itf_p05_coclass_block" "$Tests\itf_pos\p05_coclass_block.bas" }
    if (Test-Path "$Tests\itf_pos\p06_coclass_soft_ident.bas") { Test-Syntax "itf_p06_coclass_soft_ident" "$Tests\itf_pos\p06_coclass_soft_ident.bas" }
    if (Test-Path "$Tests\itf_pos\p07_coclass_cls_host.cls") { Test-Syntax "itf_p07_coclass_cls_host" "$Tests\itf_pos\p07_coclass_cls_host.cls" }
    # ai/022 B11/C03b positive guards: the contract may be met by an ancestor that merely
    # declares the members (p08), or delegated whole to a holder object (p09, B10 x B11).
    if (Test-Path "$Tests\itf_pos\p08_coclass_base.cls") {
        Test-SyntaxMulti "itf_p08_coclass_via_ancestor" @("$Tests\itf_pos\p08_coclass_ancestor_contract.cls", "$Tests\itf_pos\p08_coclass_base.cls", "$Tests\itf_pos\p08_coclass_der.cls")
    }
    if (Test-Path "$Tests\itf_pos\p09_coclass_holder.cls") {
        Test-SyntaxMulti "itf_p09_coclass_via_delegated" @("$Tests\itf_pos\p09_coclass_holder.cls", "$Tests\itf_pos\p09_coclass_impl.cls")
    }
    # ai/022 B11/C04 (026 section 6, D52): a VB6 class module that never says "CoClass".
    # Its header attribute lines fold into one CoClass record, solved by the SAME Pass E, and
    # the fold must stay read-only: VB_Creatable=True in an EXE project is the corpus' normal
    # shape (134 of 143 lines) and must NOT inherit C03a's VB3033; a folded name used as an
    # Inherits base must NOT get the "that is a CoClass block" wording (D52-3). C05 adds the
    # third file: the folded name used as a TYPE must still mean the class, so it must also
    # produce no activation line (D54-3).
    if (Test-Path "$Tests\itf_pos\p10_coclass_fold_base.cls") {
        Test-SyntaxNote "itf_p10_coclass_fold" @("$Tests\itf_pos\p10_coclass_fold_base.cls", "$Tests\itf_pos\p10_coclass_fold_der.cls", "$Tests\itf_pos\p10_coclass_fold_use.bas") @(
            "C3: CoClass 'FoldBase' identity: CLSID={96466C30-E240-55A4-9434-24F1C884C2AB} (minted) IID=- (missing) ProgID=VB6EXE.FoldBase (minted) impl='FoldBase' comCreatable=True folded-from-legacy: VB_Creatable=True VB_Exposed=False VB_PredeclaredId=False VB_GlobalNameSpace=False",
            "C3: CoClass 'FoldDer' identity:") @("VB_VarHelpID", "VB_Description", "is a CoClass block", "VB3033", "activated in-project")
    }
    # Both shapes in one module: the hand-written block wins, so the record carries no fold
    # tag and [ComCreatable(False)] -- not VB_Creatable=True -- is what reaches the identity.
    if (Test-Path "$Tests\itf_pos\p11_coclass_block_wins.cls") {
        Test-SyntaxNote "itf_p11_coclass_block_wins" @("$Tests\itf_pos\p11_coclass_block_wins.cls") @(
            "C3: class 'FoldWins' has both a CoClass block and 4 legacy header attribute line(s): the block wins, the attributes are not folded",
            "CoClass 'FoldWins' identity: CLSID={EA2B2FD6-E5C6-5192-D0C9-A13BC6FC7859} (minted) IID=- (missing) ProgID=VB6EXE.FoldWins (minted) impl='' comCreatable=False") @("folded-from-legacy")
    }
    # ai/022 B11/C05 (ai/026 section 5, items 3-5): the observable face of in-project
    # activation is one information line per block that is ACTUALLY USED as a type -- and no
    # line at all for a block nobody binds to (cc_id declares three and uses none, see the
    # byte guard). Asserting the line rather than the exit code, because the stage succeeds.
    Test-SyntaxNote "cc_act_group_names" @("$Tests\cc_act\Act.vbp") @(
        "CoClass 'Circle' activated in-project: type name -> class 'ShapeAct'",
        "CoClass 'Ring' activated in-project: type name -> class 'RingAct'",
        "ProgID=ActApp.Circle") @("declares no [Implementation]", "VB3039")
    # A block without [Implementation] stays legal, but binding a variable to it has no
    # answer -- the use site is where the refusal lands (D54-2: today that shape is a silent
    # late-bound call on a null pointer, which is worse than an error).
    Test-SyntaxFail "itf_n40_coclass_type_no_impl" "$Tests\itf_neg\n40_coclass_type_no_impl.bas" "declares no [Implementation] class"
    # ai/022 B11/C02: the three identity tiers, asserted against expected GUIDs computed by an
    # independent FNV-1a re-implementation of the seed strings "coc:<proj>.<coclass>" and
    # "itf:<proj>.<iface>" (both lowered) -- NOT scraped from this compiler's own output, or the
    # test could only ever re-state the implementation (ai/022 D46).
    $ccShapes = "$Tests\cc_id\Id.vbp"
    $ccOther = "$Tests\cc_id\Id2.vbp"
    Test-IdentityNote "cc_id_explicit_tier" $ccShapes @(
        "CoClass 'CCCircle' identity: CLSID={11111111-1111-1111-1111-111111111111} (explicit) IID={22222222-3333-4444-5555-666666666666} (explicit) ProgID=Shapes.Circle (explicit) impl='CircleImpl' comCreatable=False",
        "CoClass 'CCVbp' identity: CLSID={33333333-4444-5555-6666-777777777777} (vbp) IID={62D63A9A-7316-DAB9-B2D1-5DDFB57EDB17} (minted) ProgID=ShapesApp.CCVbp (minted) impl='VbpImpl' comCreatable=False",
        "CoClass 'CCMint' identity: CLSID={A5B36375-4E3A-D5F9-D2A2-1912F3EDC498} (minted) IID={62D63A9A-7316-DAB9-B2D1-5DDFB57EDB17} (minted) ProgID=ShapesApp.CCMint (minted) impl='' comCreatable=False")
    # Same sources, only the vbp Name= differs: the minted tier moves while the explicit line
    # stays character-identical to the case above -- that split IS the reproducibility claim.
    Test-IdentityNote "cc_id_project_name_scope" $ccOther @(
        "CoClass 'CCCircle' identity: CLSID={11111111-1111-1111-1111-111111111111} (explicit)",
        "CoClass 'CCVbp' identity: CLSID={33333333-4444-5555-6666-777777777777} (vbp) IID={28519764-65C8-D639-C831-604BAD706603} (minted) ProgID=OtherApp.CCVbp (minted) impl='VbpImpl' comCreatable=False",
        "CoClass 'CCMint' identity: CLSID={CE88DE91-E77D-563D-D74F-5CA0DF3902D2} (minted) IID={28519764-65C8-D639-C831-604BAD706603} (minted) ProgID=OtherApp.CCMint (minted)")
    Test-IdentityStable "cc_id_repeatable" $ccShapes
    # ai/022 B11/C04: a class that writes NO block but is listed in the .vbp the VB6 way
    # (Class=Name; file.cls; {CLSID}). Folding has to hand that entry to the same resolver,
    # so the vbp tier lights up for legacy projects too -- the proof that there is still only
    # one identity channel (D47) rather than a legacy side door.
    Test-IdentityNote "cc_id_fold_vbp_tier" "$Tests\cc_id\IdFold.vbp" @(
        "CoClass 'FoldVbp' identity: CLSID={77777777-8888-9999-AAAABBBBBBBBBBBB} (vbp) IID=- (missing) ProgID=FoldApp.FoldVbp (minted) impl='FoldVbp' comCreatable=True folded-from-legacy: VB_Creatable=True VB_Exposed=False VB_PredeclaredId=False VB_GlobalNameSpace=False")
    Test-IdentityStable "cc_id_fold_repeatable" "$Tests\cc_id\IdFold.vbp"
    # ai/022 B13d: 规范 IUnknown —— 薄指针的 QI 认 IID_IUnknown 时必须回“本类实现序里第一个
    # 接口”的薄指针（两处 QI 回同一个值），不再是各自的 self。XWriter 的 CWriter 同时实现
    # IWriter + ILog = 最小可用形状；真跑那一半由 itf_xmod_writer（QI1..QI4）钉住兄弟/本接口分支没坏。
    Test-CanonicalIUnknownShape "itf_canonical_iunknown" @("$Tests\itf_xmod\XWriter.vbp") "void* canon = &me->__iv_IWriter;" 2
    # ai/022 B07a (class Inherits, P3): chain diagnostics must fire. Single-file cases ride
    # the existing Test-SyntaxFail path; the two-module cases need Test-SyntaxFailMulti.
    $clsInhNeg = @(
        @("ci_n01_unknown_base", "$Tests\cls_neg\ci_n01_unknown_base.cls", "inherits unknown base class"),
        @("ci_n02_not_class", "$Tests\cls_neg\ci_n02_not_class.bas", "only allowed in a class module"),
        @("ci_n03_duplicate_clause", "$Tests\cls_neg\ci_n03_duplicate_clause.cls", "more than one Inherits clause"),
        @("ci_n04_self_cycle", "$Tests\cls_neg\ci_n04_self_cycle.cls", "Circular Inherits chain"),
        @("ci_n05_generic_template", "$Tests\cls_neg\ci_n05_generic_template.cls", "not allowed inside a generic class template"),
        @("ci_n11_protected_in_interface", "$Tests\cls_neg\ci_n11_protected_in_interface.bas", "must not carry an access modifier")
    )
    foreach ($c in $clsInhNeg) {
        if (Test-Path $c[1]) { Test-SyntaxFail $c[0] $c[1] $c[2] }
        else { Write-Host "  [SYNTAX-FAIL] $($c[0]) ... SKIP (missing case file)" -ForegroundColor DarkGray }
    }
    if (Test-Path "$Tests\cls_neg\ci_n06_pair_a.cls") {
        Test-SyntaxFailMulti "ci_n06_pair_cycle" @("$Tests\cls_neg\ci_n06_pair_a.cls", "$Tests\cls_neg\ci_n06_pair_b.cls") "Circular Inherits chain"
    }
    if (Test-Path "$Tests\cls_neg\ci_n07_iface_host.cls") {
        Test-SyntaxFailMulti "ci_n07_base_is_iface_host" @("$Tests\cls_neg\ci_n07_iface_host.cls", "$Tests\cls_neg\ci_n07_derives_host.cls") "inherits unknown base class"
    }
    # ai/022 B07b (v1 boundaries): unqualified inherited call, inherited field redeclared,
    # and an event-bearing base. All three are two-module cases -> Test-SyntaxFailMulti.
    if (Test-Path "$Tests\cls_neg\ci_n08_base.cls") {
        Test-SyntaxFailMulti "ci_n08_bare_inherited_call" @("$Tests\cls_neg\ci_n08_base.cls", "$Tests\cls_neg\ci_n08_derived.cls") "cannot be called unqualified"
    }
    if (Test-Path "$Tests\cls_neg\ci_n09_base.cls") {
        Test-SyntaxFailMulti "ci_n09_redeclared_field" @("$Tests\cls_neg\ci_n09_base.cls", "$Tests\cls_neg\ci_n09_derived.cls") "redeclares inherited field"
    }
    if (Test-Path "$Tests\cls_neg\ci_n10_base.cls") {
        Test-SyntaxFailMulti "ci_n10_event_base" @("$Tests\cls_neg\ci_n10_base.cls", "$Tests\cls_neg\ci_n10_derived.cls") "declares an Event"
    }
    # ai/022 B08b/B08d (Overridable/Overrides): the contract shapes are two-module cases,
    # the three placement cases are single-file. All ride the existing helpers.
    $ovNeg = @(
        @("ci_n12_override_ghost", "ci_n12", "no matching member in the inherited class chain"),
        @("ci_n13_not_overridable", "ci_n13", "not declared Overridable"),
        @("ci_n14_override_signature", "ci_n14", "does not match the Overridable member"),
        @("ci_n19_propertylet_virtual", "ci_n19", "is a Property Let/Set, which this build cannot dispatch"),
        @("ci_n20_bare_virtual_call", "ci_n20", "Bare (unqualified) call to overridable member")
    )
    foreach ($c in $ovNeg) {
        $a = "$Tests\cls_neg\" + $c[1] + "_base.cls"
        $b = "$Tests\cls_neg\" + $c[1] + "_derived.cls"
        if ((Test-Path $a) -and (Test-Path $b)) {
            Test-SyntaxFailMulti $c[0] @($a, $b) $c[2]
        } else {
            Write-Host "  [SYNTAX-FAIL] $($c[0]) ... SKIP (missing case files)" -ForegroundColor DarkGray
        }
    }
    if (Test-Path "$Tests\cls_neg\ci_n16_override_no_inherits.cls") {
        Test-SyntaxFail "ci_n16_override_no_base" "$Tests\cls_neg\ci_n16_override_no_inherits.cls" "has no base class to override"
    }
    if (Test-Path "$Tests\cls_neg\ci_n17_overridable_in_bas.bas") {
        Test-SyntaxFail "ci_n17_overridable_in_bas" "$Tests\cls_neg\ci_n17_overridable_in_bas.bas" "only allowed in a class module"
    }
    if (Test-Path "$Tests\cls_neg\ci_n18_overridable_in_interface.bas") {
        Test-SyntaxFail "ci_n18_overridable_in_iface" "$Tests\cls_neg\ci_n18_overridable_in_interface.bas" "must not carry a virtual modifier"
    }
    # ai/028 V1 负例: 未闭合的反引号串 (表达式位与 Attribute 行两个入口) 与三枚反引号
    # (想写一枚字面反引号但少写闭合符) 都报 1007。最后一条是 R1 的钉子 —— 普通的 "" 串
    # 一律不许跨行、不许插值, 老诊断 1002 必须继续报, 否则就是新语法吃掉老语法。
    $rsNeg = @(
        @("rs_n1_unclosed", "rs_n1_unclosed.bas", "VB1007"),
        @("rs_n2_three_backticks", "rs_n2_three_backticks.bas", "VB1007"),
        @("rs_n3_plain_quote_oneline", "rs_n3_plain_quote_still_oneline.bas", "VB1002"),
        @("rs_n4_unclosed_on_attr", "rs_n4_unclosed_on_attr.bas", "VB1007")
    )
    foreach ($c in $rsNeg) {
        $rsNegPath = "$Tests\rawstr_neg\" + $c[1]
        if (Test-Path $rsNegPath) {
            Test-SyntaxFail $c[0] $rsNegPath $c[2]
        } else {
            Write-Host "  [SYNTAX-FAIL] $($c[0]) ... SKIP (missing case file)" -ForegroundColor DarkGray
        }
    }
    # 发码形状判据: 字面量独占一行、行界以 \r\n 转义出现、非 ASCII 一律 \uXXXX
    # (⇒ 与 cl.exe 的源码编码假设无关)。
    # ai/028 V2 负例。前两条在词法层 (1008 = 孔没等到闭合的右花括号, 含"孔跨行"这种写法;
    # 1009 = 空孔), --syntax-only 就够。后两条是 R2 的钉子: 孔里的表达式就是普通表达式,
    # 未声明的名字照报既有的 VB3001。in_n4 特意让**第二个孔**出错 —— 报在 (8,7) 才证明
    # 子扫描的窗口把行列播种做对了 (指回原文件, 不需要事后平移 AST)。
    $inNeg = @(
        @("in_n1_unclosed_hole", "in_n1_unclosed_hole.bas", "VB1008"),
        @("in_n2_empty_hole", "in_n2_empty_hole.bas", "VB1009")
    )
    foreach ($c in $inNeg) {
        $inNegPath = "$Tests\interp_neg\" + $c[1]
        if (Test-Path $inNegPath) {
            Test-SyntaxFail $c[0] $inNegPath $c[2]
        } else {
            Write-Host "  [SYNTAX-FAIL] $($c[0]) ... SKIP (missing case file)" -ForegroundColor DarkGray
        }
    }
    Test-CodegenNote "in_n3_undeclared_in_hole" @("$Tests\interp_neg\in_n3_undeclared_in_hole.bas") @("VB3001", "nopeHere")
    Test-CodegenNote "in_n4_second_hole_line" @("$Tests\interp_neg\in_n4_second_hole_line.bas") @("VB3001", "alsoNope", "(8,7)")
    # <vbeclipse> 账 #215 的另一面: 数值针看不出 Variant 与 Long 的差别(TypeName 装箱后都打 Long)，
    # 所以这里钉发码里的声明行，并用 Absent 钉诊断 —— 改前同一份夹具是 7 条 VB3001。
    Test-CodegenNote "bodydecl_one_per_declarator" @("$Tests\test_bodydecl.bas") @(
        "int32_t u1 = 0;", "int32_t u2 = 0;",
        "const int32_t c2 = 4;", "static int32_t st2 = 0;") @("VB3001")
    # <vbeclipse> 账 #217 的诊断两面。neg 这份同时写着 `Case Is > 2` 和一枚真没声明过的名字:
    # 'Is' 必须不再出现(改前每形一条)，nopeHereIsNotAName 必须继续出现 —— 拦「把整条诊断关掉」
    # 那种修法。implicit 那份故意不写 Option Explicit，钉的是改前那枚没人引用的隐式局部。
    Test-CodegenNote "caseis_is_not_an_identifier" @("$Tests\test_caseis_neg.bas") @(
        "VB3001", "nopeHereIsNotAName") @("'Is'")
    Test-CodegenNote "caseis_no_phantom_local" @("$Tests\test_caseis_implicit.bas") @(
        "int32_t _vb6_select_0", "BSTR _vb6_select_1") @("vb6_VARIANT Is")
    # <vbeclipse> 账 #220 的位置那一头 + 账 #221 的改道: Line 尾部的 B/C/F 只在 **style 格**
    # (color 之后) 折成位数值 (B=1 / C=2 / F=4 ⇒ BF=5), 写在 color 格时必须还是用户自己的变量;
    # 折好之后整条调用必须由 controlCanvasMethod 那张表改成原生 vb6_ControlLine(...)。
    # 改前两格都发裸名、由 RTL 的 const int32_t B / BF 接住 (语料实测 vb6_ComPackValue(BF)),
    # 而 Line 本身没人应答: Absent 那两条钉的就是"又退回兜底 / 名字又回来"这两件。
    Test-CodegenNote "pcline_flag_folded" @("$Tests\pcline\PcForm.frm") @(
        "vb6_ControlLine((void*)vb6_hwnd_picA, (double)10, (double)10, (double)50, (double)50, (int32_t)255, (int32_t)5)",
        "vb6_ControlLine((void*)vb6_hwnd_picA, (double)0, (double)0, (double)20, (double)20, (int32_t)16711680, (int32_t)1)",
        "vb6_ControlLine((void*)vb6_hwnd_picA, (double)5, (double)5, (double)15, (double)15, (int32_t)B, (int32_t)0)",
        "vb6_ControlLine((void*)vb6_hwnd_picA, (double)6, (double)6, (double)16, (double)16, (int32_t)BF, (int32_t)0)") @(
        "vb6_ComGetObjectProp(vb6_hwnd_picA, L" + [char]34 + "Line" + [char]34 + ")",
        "vb6_ComPackValue(BF)", "vb6_ComPackValue(B)")
    # <vbeclipse> 账 #217 第二刀: 文档隐式对象 (UserControl / PropertyPage / Extender / Ambient)
    # 与全局库前缀 VBA 现在语义层也认识 (判据 = Module::docKind 一处写、两处读)。两份真工程的
    # 发码语料实测: VB3001 全仓 2934→158 (grid 499→19、Charts 主工程 499→34)，而宿主符号一个没动
    # (vb6_UserControl_ScaleWidth 436=436 / vb6_Ambient_UserMode 94=94 / vb6_PropertyPage_hWnd 16=16)。
    # dhExp 钉"诊断清零 + 符号照旧"，dhImp 故意不写 Option Explicit 钉"不再发死局部"。
    Test-CodegenNote "dochost_ctl_no_undeclared" @("$Tests\dochost\dhExp.ctl") @(
        "vb6_UserControl_ScaleWidth", "vb6_UserControl_hWnd",
        "vb6_Ambient_UserMode", "vb6_Extender_Tag", "vb6_UCase") @("VB3001")
    Test-CodegenNote "dochost_no_phantom_locals" @("$Tests\dochost\dhImp.ctl") @(
        "vb6_UserControl_ScaleWidth", "vb6_Ambient_UserMode", "vb6_UCase") @(
        "vb6_VARIANT UserControl", "vb6_VARIANT Ambient", "vb6_VARIANT VBA")
    # 账 #219 (裸写的文档成员): 语义层以前不问 kHostPseudoRows, 于是每条合法裸写配一句 VB3001
    # (语料实测 Changed 24 条 / hDC 24 条), 而发码那侧从来是对的 (vb6_PropertyPage_Changed 186 处、
    # 裸名 0 处)。判据收到那张表的 HPF_BARE 一列: 表里有的一个都不许报 (两份夹子), 表里没有的
    # 必须照报 (dhTypo) —— 否则"放行"就成了"文档里写什么都行"。
    Test-CodegenNote "dochost_bare_ctl_symbols" @("$Tests\dochost\dhBare.ctl") @(
        "vb6_UserControl_ScaleWidth", "vb6_UserControl_hDC", "vb6_UserControl_hWnd") @("VB3001")
    Test-CodegenNote "dochost_bare_pag_dirty_flag" @("$Tests\dochost\dhBare.pag") @(
        "vb6_PropertyPage_Changed = (-1);") @("VB3001")
    Test-CodegenNote "dochost_bare_typo_still_reported" @("$Tests\dochost\dhTypo.pag") @(
        "'Changd'") @("'Changed'", "vb6_PropertyPage_Changed")
    # 账 #184 的形状面: 被 AddressOf 取址的过程 → Private 的发 static 桩、Public 的发外链桩,
    # 桩体逐字转调本体 (参数个数/宽度/顺序与本体一致, 只有约定不同); AddressOf 站点取桩地址;
    # 本体**保持 cdecl** (直接调用那条路与 RTL 的 cdecl 登记面都不受牵连)。
    # 两条 Absent 就是负控: BASE 的发码里取址点写的是 (void*)vb6_CountWin, 本体没有 __stdcall。
    Test-CodegenNote "addrof_cb_shape" @("$Tests\test_addrof_cb.bas") @(
        "static int32_t __stdcall aoThunk_vb6_CountWin(intptr_t hWnd, intptr_t lParam) { return vb6_CountWin(hWnd, lParam); }",
        "int32_t __stdcall aoThunk_vb6_StopWin(intptr_t hWnd, intptr_t lParam) { return vb6_StopWin(hWnd, lParam); }",
        "EnumWindows((void*)aoThunk_vb6_CountWin, 4242);") @(
        "EnumWindows((void*)vb6_CountWin, 4242);",
        "__stdcall vb6_CountWin(intptr_t hWnd, intptr_t lParam) {")
    # 账 #188: 设计期字号是**整数**时, 发码以前会拼出 C 里的非法字面量 `12f` (C2059 bad suffix),
    # 而 VB6 的默认字号 8.25 带小数点 ⇒ 门禁里从来没响过。钉在真实工程上 (Charts 2020/ucChartBar 的
    # BeginProperty TitleFont Size=12): 改后必须是 `12.0f`, 且 8.25 那一族**一个字节都不许变**
    # (证人不抖)。Absent 那条就是负控 —— HEAD 的发码里正是 `->Size = 12f;`。
    Test-CodegenNote "fontsize_literal" @("$Tests\Charts 2020\ucChartBar\Proyecto1.vbp") @(
        "->Size = 12.0f;", "->Size = 8.25f;") @(
        "->Size = 12f;")

    # 账 #191 的发码形状针: 控件数组元素的方法调用必须与单枚同走"直发生成函数"那条出口,
    # 不能落 vb6_ComCall 兜底 (宿主值不是 IDispatch ⇒ 交回空值, 编得过、静默不干活)。
    # 修前实测同一条语句发成 vb6_ComCall(vb6_CtrlArr_GetAt(&vb6_arr_uArr, aj), L"SW", NULL, 0)。
    Test-CodegenNote "ucarr_member_call" @("$Tests\ve_units\Units.vbp") @(
        "vb6_ucUnitPix_SW((vb6_cls_ucUnitPix*)vb6_UC_InstanceOf(vb6_CtrlArr_GetAt(&vb6_arr_uArr, aj)))") @(
        "vb6_ComCall(vb6_CtrlArr_GetAt(&vb6_arr_uArr, aj)"
    # 账 #191-b 的发码形状针 (钉在真实工程 Charts 2020/ucChartBar 的 Form2.frm:223 那一条上):
    # 「对象型属性再往下一层的写」`ucChartBar1(i).Font.Size = ...` 必须发成
    #   vb6_ComSetProp(vb6_ucChartBar_prop_get_Font(直发的实例), L"Size", pack(...)) —— 取回那枚
    # StdFont 再对它写一层 (StdFont 本身是真 IDispatch)。修前实测两形都坏: 左值折叠那一路只取
    # prop_get_ 那一段、把尾巴 `.Size` 整段丢掉 ⇒ 发成 prop_let_Font(elem, <double>), 值实参的槽
    # 是 vb6_ComIface_Font*, Form2.c 471-477 四条 C2440, 整个工程编不过; 基线那台走 ComGetObjectProp
    # 虽然编得过, 但 UC 不可分发时那条恒 NULL, 写上去是静默无效 (#191 要治的正是这一半)。
    Test-CodegenNote "ucobj_chain_write" @("$Tests\Charts 2020\ucChartBar\Proyecto1.vbp") @(
        "vb6_ComSetProp(vb6_ucChartBar_prop_get_Font((vb6_cls_ucChartBar*)vb6_UC_InstanceOf(vb6_CtrlArr_GetAt(&vb6_arr_ucChartBar1, i)))") @(
        "vb6_ucChartBar_prop_let_Font((vb6_cls_ucChartBar*)vb6_UC_InstanceOf(vb6_CtrlArr_GetAt(&vb6_arr_ucChartBar1, i)), ((double")

    # 账 #194 的发码形状针: Optional 默认值那几条必须按**数值**发。Absent 是抄源码文本那一形。
    Test-CodegenNote "optdef_default_shape" @("$Tests\optdef\OptDef.vbp") @(
        "aL = 12L;",
        "aH = 16L;",
        "aZ = 0L;") @(
        "aL = 12&",
        "aH = &H10&",
        "aZ = 0&")

    # 同一刀钉在**真工程**上 (ucTreeMaps 的 PropPagFMR：`Optional ... As Long = 0&` 两处)。
    # 这个工程今天还红着 (#192 那一族 C2039)，但发码面已过 —— 所以这条针现在就能红能绿。
    Test-CodegenNote "optdef_default_shape_real" @("$Tests\Charts 2020\ucTreeMaps\Proyecto1.vbp") @(
        "if (!_has_FontIndex) (*FontIndex) = 0L;") @(
        "(*FontIndex) = 0&;")

    # 形状针 (账 #192): With 那枚控件的属性读写与方法必须走两张表各自的出口，
    # 不许再出现 tempVar + "." + 成员名 那一形（HWND__ 是结构体指针，那是编译错）。
    Test-CodegenNote "with_ctrl_cursor_clear" @("$Tests\piccur\PicCur.vbp") @(
        "vb6_SetCurrentY(_vb6_with_",
        "vb6_ClearList((void*)_vb6_with_",
        "vb6_GetCurrentX(vb6_hwnd_picA") @(
        "_vb6_with_0.CurrentX",
        "_vb6_with_0.CurrentY",
        "_vb6_with_1.Clear()")

    # 形状针 (账 #197): 设计块的 ScaleMode 必须发成那一句；读的那一头必须落回 #175 的单位权威
    # (vb6_WindowScaleModeSelf)，不许再出现 HWND 取成员那一形（那是编译错，不是静默空转）。
    Test-CodegenNote "scalemode_design_write" @("$Tests\scalemode\ScaleMode.vbp") @(
        "vb6_SetScaleMode((void*)vb6_hwnd_picPix, 3);",
        "vb6_SetScaleMode((void*)vb6_hwnd_picTwip, 1);",
        "vb6_SetScaleMode((void*)hwnd, 3);",
        "vb6_WindowScaleModeSelf(_vb6_with_") @(
        "_vb6_with_0.ScaleMode",
        "vb6_hwnd_picPix.ScaleMode",
        "vb6_hwnd_picNest.ScaleMode")

    # 同一刀钉在**真工程**上: ucTreeMaps 的 PropPagFMR.pag:256-266 `With Picture1 : .ScaleMode` ——
    # 这一形以前是 PropPagFMR.c(74) 两条 C2039 的来源（本轮实测 C2039 4→2，剩 hDC/TextHeight = #196）。
    Test-CodegenNote "scalemode_read_real" @("$Tests\Charts 2020\ucTreeMaps\Proyecto1.vbp") @(
        "vb6_WindowScaleModeSelf(_vb6_with_0)") @(
        "_vb6_with_0.ScaleMode")

    # 形状针 (账 #196): `.hDC` 两形都必须落在那一个绘图 DC 出口上，不许再出现 HWND 取成员
    # (那是编译错，With 形那一半的真红)，也不许再走晚绑定那条 `ComGetLongPtrProp(X, L"hDC")` ——
    # 那条**实测是能用的 DC**（改前那台探针：HDC=671159075、DPI=96），但每读一次新取一张、从不归还，
    # 与 VB6 的"一个对象一张"不符（.build/b196probe）。两条都钉住 = 两形同归一处口径。
    Test-CodegenNote "dcsurf_dc_shape" @("$Tests\dcsurf\DcSurf.vbp") @(
        "vb6_GetControlHDC(vb6_hwnd_picA)",
        "vb6_GetControlHDC(_vb6_with_") @(
        "_vb6_with_0.hDC",
        "vb6_hwnd_picA.hDC",
        'vb6_ComGetLongPtrProp(vb6_hwnd_picA, L"hDC")')

    # 同一刀钉在**真工程**上: ucTreeMaps 的 PropPagFMR.pag:258 `With Picture1 : TextOut .hDC` ——
    # 这一形以前是 PropPagFMR.c(74) 那条 C2039 的来源 (与 #197 那条 ScaleMode 同处，本轮实测
    # C2039 由 2 条降到 1 条，只剩 TextHeight = #196 欠的那一半)。
    Test-CodegenNote "dc_read_real" @("$Tests\Charts 2020\ucTreeMaps\Proyecto1.vbp") @(
        "vb6_GetControlHDC(_vb6_with_0)") @(
        "_vb6_with_0.hDC")

    # 形状针 (账 #196 第二条): 按 HWND 的 TextHeight/TextWidth 两形都落在那两个出口上，
    # 实参表由码头拼 —— With 形走「裸名 + pendingChainObj_」那套协议(前置句柄)，裸形直接成串。
    Test-CodegenNote "dcsurf_text_measure" @("$Tests\dcsurf\DcSurf.vbp") @(
        "vb6_ControlTextHeight((void*)vb6_hwnd_picB,",
        "vb6_ControlTextWidth((void*)vb6_hwnd_picB,",
        "vb6_ControlTextHeight((void*)_vb6_with_1") @(
        "_vb6_with_0.TextHeight",
        "vb6_hwnd_picB.TextHeight",
        'vb6_ComCall(vb6_hwnd_picB, L"TextHeight"')

    # 同一刀钉在真工程上: ucTreeMaps 的 PropPagFMR.pag:265 `With Picture1 : .CurrentY + .TextHeight(Text)`
    # —— 这是 #196 收完 hDC 之后 ucTreeMaps 仅剩的那条 C2039。
    Test-CodegenNote "text_measure_real" @("$Tests\Charts 2020\ucTreeMaps\Proyecto1.vbp") @(
        "vb6_ControlTextHeight((void*)_vb6_with_0,") @(
        "_vb6_with_0.TextHeight")

    # 形状针 (账 #196 第三条): 窗体型接收者的 ScaleX/ScaleY **四形**都落在同一处换算上。
    # absent 那一头钉的是接上之前的三种真形状: 假 IDispatch 调用 (编得过、跑起来回 0)、
    # `hwnd.ScaleX(...)` (编译不过)、裸写 `= ScaleX(` (隐式声明, 真工程里今天靠 /OPT:REF 才没响)。
    Test-CodegenNote "dcsurf_scale_units" @("$Tests\dcsurf\DcSurf.vbp") @(
        "sx1 = vb6_ScaleUnitX(1440, 1, 3)",
        "sx2 = vb6_ScaleUnitX(1440, 1, 3)",
        "sx3 = vb6_ScaleUnitX(1440, 1, 3)",
        "sx4 = vb6_ScaleUnitX(1440, 1, 3)",
        "sy1 = vb6_ScaleUnitY(1440, 1, 3)") @(
        'vb6_ComCallDouble(vb6_hwnd_picB, L"ScaleX"',
        'vb6_ComCallDouble(vb6_hwnd_DcForm, L"ScaleX"',
        ".ScaleX(1440, 1, 3)",
        "= ScaleX(1440, 1, 3)")

    # 同一刀钉在真工程上: PropPagFMR.pag:259 那个 TextOutW 的坐标换算, 以前是裸 `ScaleX(`。
    Test-CodegenNote "scale_units_real" @("$Tests\Charts 2020\ucTreeMaps\Proyecto1.vbp") @(
        "vb6_ScaleUnitX(vb6_GetCurrentX(_vb6_with_0)") @(
        '", ScaleX(vb6_GetCurrentX(_vb6_with_0)"',
        'L"ScaleX"')

    # 账 #195: VB 的 Integer 类型后缀 `%` 以前在词法层就被拒 (case '%' 那一支只吃字符不置标志,
    # 于是 `3%` 落回「无后缀十进制按数值大小定档」那一段, 残留的 % 让 parseIntLit 报「超出 64 位」并级联出
    # 9-13 条诊断)。补上标志 + 三种进制扫描器认下这个后缀 + radixDigits 的剥离表加 %。
    # 正例: 十进制/%H%/O%/B%/负数/表达式/参数默认值 七形都要过得了语法 (修前连一条 `y = 3%` 都过不去)。
    Test-SyntaxMulti "int_suffix_forms" @("$Tests\intsuffix\IntSuffix.bas")

    # 形状针: 三种进制的 % 必须按**数值**落地 (255 / 15 / 5), 参数默认值那形要接上 #194 那一处权威。
    # Absent 是后缀漏进生成 C 那一形 (#194 修前实测) —— 两条规则钉在同一枚夹具上。
    Test-CodegenNote "int_suffix_shape" @("$Tests\intsuffix\IntSuffix.bas") @(
        "z = vb6_ChkInt(255);",
        "z = vb6_ChkInt(15);",
        "z = vb6_ChkInt(5);",
        "if (!_has_a) a = 3;") @(
        "a = 3%")

    # 边界负例: 显式 `%` 就是 Integer 档 (-32768..32767)，超出必须报 —— 按 int32 收下再让 int16
    # 去截是把值改错 (2147483648 静默回绕成 -2147483648)。针取消息里的 ASCII 片段 (GBK 口径)。
    Test-SyntaxFailMulti "int_suffix_neg_range" @("$Tests\intsuffix_neg\IntSuffixNeg.bas") `
        "(-32768..32767)"
)

    # 账 #189 的发码形状针: 控件数组的整体成员必须发成 vb6_CtrlArr_*, 一条都不许再落在
    # 句柄变量的 COM 属性读上 (修前 `uArr.Count` 就是 vb6_ComGetIntProp(vb6_hwnd_uArr, L"Count"))。
    Test-CodegenNote "ctrlarr_member" @("$Tests\ve_units\Units.vbp") @(
        "vb6_CtrlArr_GetCount(&vb6_arr_uArr)",
        "vb6_CtrlArr_LBound(&vb6_arr_uArr)",
        "vb6_CtrlArr_UBound(&vb6_arr_uArr)") @(
        "vb6_ComGetIntProp(vb6_hwnd_uArr")

    # 账 #222/#226 的发码形状针: UC 控件数组的事件臂必须**每元素一套** thunk 与 sink, 并把设计期
    # Index 交给那枚共享处理器; 处理器原型全产物只许一份, 且按事件 ABI (ByVal) 发。
    # 三条 Absent 都是改前产物里真实存在的形状 (零形参声明 / 同名 thunk / 共享 sink 槽),
    # 任何一条回来这枚针就红 —— 实测 ucProgressCircular 那 25 条 C2198+C2084 就是这么来的。
    Test-CodegenNote "ucarr_evt_thunk_per_element" @("$Tests\ve_units\Units.vbp") @(
        "static void vb6_frmUnits_uArr_Hit(int16_t Index);",
        "vb6_frmUnits_evtThunk_uArr_Hit_1(void* handler) { vb6_frmUnits_uArr_Hit(1); }",
        "((vb6_cls_ucUnitPix*)vb6_UC_InstanceOf(vb6_CtrlArr_GetAt(&vb6_arr_uArr, 1)))->events = &_sink_uArr_1;",
        "vb6_ucUnitPix_ucHostClick") @(
        "static void vb6_frmUnits_uArr_Hit();",
        "_sink_uArr.onHit",
        "vb6_frmUnits_evtThunk_uArr_Hit(void* handler)")
    # 账 #222 的门 #350 回归 (charts_ucTreeMaps 两档 C2440) 的形状针: 事件形参的 C 类型
    # 三处必须同源 —— 回调 typedef / 容器侧处理器原型 / prelude 那枚 thunk。修法是把 prelude
    # 自带的第二张类型表撤掉、改问 mapTypeRef (见 check_uc_array_event_sites.ps1 的 E5)。
    # Absent 那条是**改前产物里真实存在**的形状, 表一回来它就红。
    Test-CodegenNote "ucevt_thunk_type_same_authority" @("$Tests\Charts 2020\ucTreeMaps\Proyecto1.vbp") @(
        "typedef void (*vb6_evt_ucTreeMaps_ItemClick_cb)(void* handler, vb6_VARIANT Key);",
        "static void vb6_Form2_evtThunk_ucTreeMaps1_ItemClick(void* handler, vb6_VARIANT a0) { vb6_Form2_ucTreeMaps1_ItemClick(a0); }") @(
        "static void vb6_Form2_evtThunk_ucTreeMaps1_ItemClick(void* handler, int32_t a0)")

    # 账 #228: 同一个 VB 类型写成两种拼法必须给出**同一种 C 类型**。库里限定的
    # `stdole.OLE_COLOR` 以前一路掉到 mapTypeRef 末尾的兜底 void*, 而裸名 OLE_COLOR 答
    # int32_t —— 同一枚 VBFlexGrid 事件 (.ctl 写裸名 / 容器写限定名) 于是发送侧交 4 字节、
    # 处理器收 8 字节指针。修完反过来还要钉住**不该折的不折**: StdFont 是真外部类型
    # (工程里没有同名符号), 两种拼法都得留 void* —— 只钉前一半的话, "把所有点号都剥掉"
    # 这种过折照样绿。
    Test-CodegenNote "varcmp_scalar_boxed" @("$Tests\test_varcmp_scalar.bas") @(
        "vb6_VariantFromValue(d)",
        "vb6_VariantFromValue(gD)",
        "vb6_VarCmpEq(&v, ") @(
        "vb6_VarCmpEq(&v, &d)",
        "vb6_VarCmpEq(&d, &v)",
        "vb6_VarCmpEq(&gV, &gD)",
        "vb6_VarCmpGt(&v, &d)",
        "vb6_VarCmpLt(&d, &v)")

    Test-CodegenNote "alias_type_spelling_same_ctype" @("$Tests\test_alias_spellings.bas") @(
        "void vb6_BareColor(int32_t c);",
        "void vb6_QualColor(int32_t c);",
        "void vb6_BareFont(void* f);",
        "void vb6_QualFont(void* f);") @(
        "void vb6_QualColor(void* c)",
        "void vb6_BareFont(int32_t f)")


    # ai/028 V2 的发码形状: 插值必须** literally ** 发成手写的 & CStr() / Format$ 形状 ——
    # 注意第二枚读数挑的是 vb6_CStrLong (按实参类型改发专用 CStr), 这正是"降级成真 AST"
    # 才继承得到的东西 (计划书 R2/R3 的实测面)。
    # de67e089 之后 vb6_Format 走 VB 完整 4 形 (expr, fmt, fdow, fwoy), 2 参 VB 调用
    # 会补 `, 1, 1` (vbSunday / vbFirstJan1) —— 与 VBFlexGrid.ctl:20425 的
    # `Format$(Text, Col.Format, vbUseSystemDayOfWeek, vbUseSystem)` 共用一张签名, 否则
    # cl 19.51 C2197 too many args。这里第二枚针一起带四参尾巴, 才与 emit 对齐。
    Test-EmitcShape "ri_emitc_shape" @("$Tests\test_interp.bas") @(
        'vb6_BSTR_Concat(vb6_BSTR_FromStr(L"n="), vb6_CStrLong(n))',
        'vb6_Format(vb6_VariantLong(n), vb6_BSTR_FromStr(L"#,##0"), 1, 1)')
    Test-EmitcShape "rs_emitc_shape" @("$Tests\test_rawstr.bas") @(
        'vb6_BSTR_FromStr(L"line1\r\nline2")',
        '#define RS_CONST (vb6_BSTR_FromStr(L"k1\r\nk2 = \"v\""))',
        'vb6_BSTR_FromStr(L"\u59D3\u540D: \u5F20\u4E09\r\n\u5907\u6CE8: \"vip\"")',
        'vb6_BSTR_FromStr(L"C:\\note\\{x}\\n")')

    # 账 #174: 属性页里裸名 `SelectedControls(i)` 必须走宿主内建那张表 (cgen_expr_ident_builtin.inc 的
    # kPropertyPageHostMembers)，发射成 RTL 出口 vb6_PropertyPage_SelectedControls(...)。
    # 钉这一格的理由不是「值对不对」，而是 **发射成裸名 = 隐式函数声明**：cl 只给 C4013
    # 「未定义；假设外部返回 int」就编过，而 x86 cdecl 的 double 返回值放在 ST0 —— 调用方
    # 按 int 取值就**永不弹 x87 栈**，八次之后栈满，之后任何浮点取值得 QNaN（账 #173 的
    # Charts 2020 启动即 error 6「Overflow」就是这么炸的，见 vb6com_internal.h 那条注释）。
    # 发射面单文件就能验，所以用 emitc 形状针而不是运行产物。**红侧实测过**：基线编译器
    # 编同一份 .pag 发射 `= SelectedControls(0);` 且没有这根针。
    Test-EmitcShape "pp_selctrl_hostmember" @("$Tests\Charts 2020\LabelPlus\PropPagLP.pag") @(
        'vb6_PropertyPage_SelectedControls(0)')

    # B07a positive guard: a base class in another module resolves and stays silent.
    if (Test-Path "$Tests\cls_neg\ci_pos_base.cls") {
        Test-SyntaxMulti "ci_pos_pair" @("$Tests\cls_neg\ci_pos_base.cls", "$Tests\cls_neg\ci_pos_derived.cls")
    }
    # ai/022 B08c (Protected access from outside the class family): the three out-of-family
    # shapes below must be rejected at the call site, while in-family access (a base-typed
    # variable used from the derived class, and from the declaring class itself) stays silent.
    $protNeg = @(
        @("ci_n21_prot_field_write", "ci_n21_base.cls", "ci_n21_stranger.cls"),
        @("ci_n22_prot_sub_stdmod", "ci_n22_base.cls", "ci_n22_outsider.bas"),
        @("ci_n23_prot_property_get", "ci_n23_base.cls", "ci_n23_stranger.cls")
    )
    foreach ($c in $protNeg) {
        $a = "$Tests\cls_neg\" + $c[1]
        $b = "$Tests\cls_neg\" + $c[2]
        if ((Test-Path $a) -and (Test-Path $b)) {
            Test-SyntaxFailMulti $c[0] @($a, $b) "is Protected"
        } else {
            Write-Host "  [SYNTAX-FAIL] $($c[0]) ... SKIP (missing case files)" -ForegroundColor DarkGray
        }
    }
    if (Test-Path "$Tests\cls_neg\ci_pos2_base.cls") {
        Test-SyntaxMulti "ci_pos2_prot_in_family" @("$Tests\cls_neg\ci_pos2_base.cls", "$Tests\cls_neg\ci_pos2_derived.cls")
    }
    # ai/022 B08e-6 (class-vtable dispatch map, codegen stage): the three shapes below carry a
    # call inside the receiver, so a virtual call there would need the receiver twice. Sites 12
    # and 10 used to compile and bind statically to the base body; they now report VB3027. Site
    # 8 already reported it through the priority-2 dispatcher -- pinned here so a refactor
    # cannot lose it. Site 9's receiver is a plain return variable, so it must still compile.
    $cgenVirtNeg = @(
        @("ci_n24_chain_receiver", "ci_n24", "cannot be dispatched in this build"),
        @("ci_n25_prop_chain_receiver", "ci_n25", "cannot be dispatched in this build"),
        @("ci_n26_prop_receiver", "ci_n26", "cannot be dispatched in this build"),
        # ai/022 B09: `MyBase` where there is no base class, and `MyBase.<name>` the base face
        # does not have -> VB3028. Both used to compile into an undefined `vb6_MyBase_<name>`.
        @("ci_n27_mybase_no_inherits", "ci_n27", "has no Inherits clause"),
        @("ci_n28_mybase_no_such_member", "ci_n28", "has no such member")
    )
    foreach ($c in $cgenVirtNeg) {
        $a = "$Tests\cls_neg\" + $c[1] + "_base.cls"
        $b = "$Tests\cls_neg\" + $c[1] + "_derived.cls"
        if ((Test-Path $a) -and (Test-Path $b)) {
            Test-CompileFail $c[0] @($a, $b) $c[2]
        } else {
            Write-Host "  [COMPILE-FAIL] $($c[0]) ... SKIP (missing case files)" -ForegroundColor DarkGray
        }
    }
    if (Test-Path "$Tests\cls_neg\ci_pos3_base.cls") {
        Test-Compile "ci_pos3_retval_receiver" @("$Tests\cls_neg\ci_pos3_base.cls", "$Tests\cls_neg\ci_pos3_derived.cls")
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

# ai/023 S01: package/project-reference hard checks (diagnostics only, no codegen).
# Negatives ride --syntax-only (stage-0 checks fire before the pipeline); positives
# build+run the host to prove the checks don't break a normal build (D7: warnings
# never reject a build).
if ($Category -in @("all", "pkg")) {
    Write-Host "--- Package Reference Tests (ai/023 S01) ---" -ForegroundColor Yellow
    Test-VbpFail "pkg_n01_not_found"    "$Tests\pkg_neg\n01_not_found.vbp"    "package not found in any search root"
    Test-VbpFail "pkg_n02_escape_name"  "$Tests\pkg_neg\n02_escape_name.vbp"  "package name contains path characters"
    Test-VbpFail "pkg_n03_conflict"     "$Tests\pkg_neg\n03_conflict.vbp"     "package name conflicts with project name"
    Test-VbpFail "pkg_n04_bad_manifest" "$Tests\pkg_neg\n04_bad_manifest.vbp" "unknown key in [C3Package]"
    Test-VbpFail "pkg_n05_name_mismatch" "$Tests\pkg_neg\n05_name_mismatch.vbp" "manifest Name mismatch"
    # ai/023 S02: package source loading — a package module may not collide with a
    # host module name (checked at load time, before the pipeline).
    Test-VbpFail "pkg_n06_module_conflict" "$Tests\pkg_neg\n06_module_conflict.vbp" "package module name conflicts with host module"
    if (Test-Path "$Tests\pkg_pos\p01_ok.vbp") {
        Test-Vbp "pkg_p01_ok" "$Tests\pkg_pos\p01_ok.vbp" @("PKG-S01:OK")
    }
    if (Test-Path "$Tests\pkg_pos\p02_missing_file_warn.vbp") {
        Test-Vbp "pkg_p02_missing_file_warn" "$Tests\pkg_pos\p02_missing_file_warn.vbp" @("PKG-S01:OK")
    }
    # ai/023 S02: host calls a package Function and Property Get across modules.
    if (Test-Path "$Tests\pkg_xmod\pkg_xmod_ok.vbp") {
        Test-Vbp "pkg_s02_xmod" "$Tests\pkg_xmod\pkg_xmod_ok.vbp" @("PKG-XMOD:42", "PKG-XMOD:XM")
    }
    # ai/023 S03: package export boundary — Friend member invisible to host (VB7006),
    # visible again with manifest Friend=True.
    if (Test-Path "$Tests\pkg_xmod\friend_bad.vbp") {
        Test-VbpBuildFail "pkg_s03_friend_blocked" "$Tests\pkg_xmod\friend_bad.vbp" "is not exported by package"
    }
    if (Test-Path "$Tests\pkg_xmod\friend_open_ok.vbp") {
        Test-Vbp "pkg_s03_friend_open" "$Tests\pkg_xmod\friend_open_ok.vbp" @("PKG-FRIEND-OK")
    }
    # ai/023 S04: class modules across the package boundary. An exported class binds
    # statically (PKG-C1 = New + method + property); a class the manifest does not
    # export must be a hard VB7006 error, not a silent late-bound COM fallback
    # (codegen emits vb6_NewObject(L"Cls") when the class symbol is missing).
    if (Test-Path "$Tests\pkg_cls\cls_ok.vbp") {
        Test-Vbp "pkg_s04_cls_ok" "$Tests\pkg_cls\cls_ok.vbp" @("PKG-C1:42")
    }
    if (Test-Path "$Tests\pkg_cls\cls_neg.vbp") {
        Test-VbpBuildFail "pkg_s04_cls_blocked" "$Tests\pkg_cls\cls_neg.vbp" "does not export it"
    }
    if (Test-Path "$Tests\pkg_cls\dimonly.vbp") {
        # `Dim x As Cls` alone (no New) must also be rejected — otherwise the type
        # silently degrades to Object/void*.
        Test-VbpBuildFail "pkg_s04_cls_dim_only" "$Tests\pkg_cls\dimonly.vbp" "does not export it"
    }
    # ai/084a M3: member-level Friend boundary on exported package classes —
    # obj.<Friend member> from the host is a hard 7008 unless the manifest
    # declares Friend=True (same-package / package-to-package stay unrestricted).
    if (Test-Path "$Tests\pkg_cls\cls_friend_obj_bad.vbp") {
        Test-VbpBuildFail "acc_m3_pkg_friend_obj_blocked" "$Tests\pkg_cls\cls_friend_obj_bad.vbp" "Friend"
    }
    if (Test-Path "$Tests\pkg_cls\cls_friend_obj_open_ok.vbp") {
        Test-Vbp "acc_m3_pkg_friend_obj_open" "$Tests\pkg_cls\cls_friend_obj_open_ok.vbp" @("FR-OPEN:12")
    }
    if (Test-Path "$Tests\pkg_cls\samepkg_ok.vbp") {        # Same-package module uses the non-exported class: must NOT be blocked.
        Test-Vbp "pkg_s04_cls_same_package" "$Tests\pkg_cls\samepkg_ok.vbp" @("PKG-C3:OK")
    }
    # ai/023 S05: per-file sha1/size verification. Clean package = no warning;
    # tampered byte = warning VB7007 and build CONTINUES (D7 校验不拒收);
    # --check-packages prints resolution read-only (OK / MISMATCH) and exits.
    if (Test-Path "$Tests\pkg_chk\chk_ok.vbp") {
        Test-Vbp "pkg_s05_chk_ok" "$Tests\pkg_chk\chk_ok.vbp" @("PKG-S05:55")
        Test-CliOk "pkg_s05_check_packages_ok" @('"' + "$Tests\pkg_chk\chk_ok.vbp" + '"', "--check-packages") "sha1=OK"
    }
    if (Test-Path "$Tests\pkg_chk\tamper.vbp") {
        Test-VbpWarn "pkg_s05_tamper_warn_builds" "$Tests\pkg_chk\tamper.vbp" "hash mismatch"
        Test-CliOk "pkg_s05_check_packages_tampered" @('"' + "$Tests\pkg_chk\tamper.vbp" + '"', "--check-packages") "sha1=MISMATCH"
    }
    # ai/023 S06: pack -> unpack roundtrip must be byte-identical.
    if (Test-Path "$Tests\pkg_chk\packages\ChkPkg-1.0\package.c3d") {
        Test-PackRoundtrip "pkg_s06_pack_roundtrip" "$Tests\pkg_chk\packages\ChkPkg-1.0"
    }
    Write-Host ""
}

# ai/084a M1/M2: class member access levels. Positives prove same-class (Me. /
# other instance), Friend-same-project and inherited-Private-field access stay
# legal; the negative proves Private proc access from outside the class is a
# hard 3028 (previously a silent C2129 at MSVC stage).
if ($Category -in @("all", "acc")) {
    Write-Host "--- Member Access Level Tests (ai/084a M1/M2) ---" -ForegroundColor Yellow
    if (Test-Path "$Tests\acc\acc_ok.vbp") {
        Test-Vbp "acc_m2_ok" "$Tests\acc\acc_ok.vbp" @("ACC-USE:6", "ACC-OK")
    }
    if (Test-Path "$Tests\acc\acc_fam.vbp") {
        Test-Vbp "acc_m2_family_field" "$Tests\acc\acc_fam.vbp" @("ACC-FAM:3")
    }
    if (Test-Path "$Tests\acc\acc_neg.vbp") {
        Test-VbpBuildFail "acc_m2_private_blocked" "$Tests\acc\acc_neg.vbp" "Private"
    }
    # ai/084a: Protected — family-internal access (Me. and obj.) stays legal, a
    # stranger reaching a Protected member is a hard 3023 (placeholder now landed).
    if (Test-Path "$Tests\acc\acc_prot_ok.vbp") {
        Test-Vbp "acc_prot_family_ok" "$Tests\acc\acc_prot_ok.vbp" @("PROT:fam")
    }
    if (Test-Path "$Tests\acc\acc_prot_neg.vbp") {
        Test-VbpBuildFail "acc_prot_stranger_blocked" "$Tests\acc\acc_prot_neg.vbp" "Protected"
    }
    Write-Host ""
}

# =============================================
# ctor: ai/084c 类构造函数 — 带参 (New Cls(args) → _NewParams) 与无参 (Class_Initialize)
# =============================================
if ($Category -in @("all", "ctor")) {
    Write-Host "--- ctor (ai/084c class constructors) ---" -ForegroundColor Yellow
    if (Test-Path "$Tests\ctor\ctor_ok.vbp") {
        Test-Vbp "ctor_ok" "$Tests\ctor\ctor_ok.vbp" @("amt:42", "CNT:7")
    }
    if (Test-Path "$Tests\ctor\ctor_neg.vbp") {
        Test-VbpBuildFail "ctor_neg_arity" "$Tests\ctor\ctor_neg.vbp" "3035"
    }
    if (Test-Path "$Tests\ctor\ctor_neg2.vbp") {
        Test-VbpBuildFail "ctor_neg2_no_params" "$Tests\ctor\ctor_neg2.vbp" "3035"
    }
    Write-Host ""
}

# =============================================
# asm: ai/vb-asm-extension-spec — Asm 块完整形态
#   x64: 生成 .asm → ml64 → 链接 (基本块/ByRef/Naked/自动保存/Clobber/单行)
#   x86: __asm{} 内联块 (同一份语法换后端)
#   混排 (项2/项3): Asm 片段与 VB 语句混排 + 片段内引用 VB 局部变量
#   负例: 3037 (类方法里的 Asm) / 3036 (x86 <Naked> 引用参数) / 2012 (<Naked> 修饰非过程) /
#         2014 (Clobber 参数非字符串) / 3040 ([X] 解析不到) / 3041 (x64 引用 >4 个变量)
# =============================================
if ($Category -in @("all", "asm")) {
    Write-Host "--- Asm Block Tests (ai/vb-asm-extension-spec) ---" -ForegroundColor Yellow
    if (Test-Path "$Tests\asm\asm_ok.vbp") {
        Test-Vbp "asm_ok" "$Tests\asm\asm_ok.vbp" @(
            "ASM-ADD:42", "ASM-ATOMIC-OLD:10", "ASM-ATOMIC-NEW:15",
            "ASM-KEEP-RBX:37", "ASM-CLOBBER:12", "ASM-NAKED:100", "ASM-ONELINE:1234",
            # 项4 x64 浮点 (XMM0/xmm1 + xmm0 返回) / 项5 x64 栈传参 (shadow space 之后) /
            # 项6 x64 int64 返回 (RAX) —— 见 spec §12.1/12.2
            "ASM-DBL:3.75", "ASM-SUM6:21", "ASM-BIG64:4000000000", "ASM-DONE")
    }
    if (Test-Path "$Tests\asm\asm_x86.vbp") {
        Test-Vbp "asm_x86_inline" "$Tests\asm\asm_x86.vbp" @(
            "X86-ADD:42", "X86-KEEP-EBX:37", "X86-CLOBBER:12",
            "X86-NAKED:5", "X86-ONELINE:1234",
            # 项4 x86 浮点返回 (ST0 → fstp) / 项5 x86 全栈参数按名解析 /
            # 项6 x86 int64 返回 (EDX:EAX + [Function+4]) —— 见 spec §12.1/12.3
            "X86-DBL:3.75", "X86-SUM5:15", "X86-BIG:4000000000",
            "X86-MAKE64:4294967297", "X86-DONE") -Arch "x86"
    }
    # --- 项2/项3 混排 + 片段引用 VB 局部变量 (spec §12.5) ---
    if (Test-Path "$Tests\asm\asm_mixed_basic.vbp") {
        # 最小混排: 片段 + 一条 VB 语句; VB 语句覆盖片段写的返回值
        Test-Vbp "asm_mixed_basic" "$Tests\asm\asm_mixed_basic.vbp" @("MIXED:1")
    }
    if (Test-Path "$Tests\asm\asm_mixed.vbp") {
        # x64: 片段降级为独立 MASM 过程, 变量以地址传入 ([X] → [rcx])
        #   局部变量读写 / 多片段 / ByRef 参数 / [Function] / callee-saved 自动保存 /
        #   Clobber / [X+4] 偏移形态
        Test-Vbp "asm_mixed_x64" "$Tests\asm\asm_mixed.vbp" @(
            "MIX-ACC:175", "MIX-TWO:22", "MIX-BUMP:15", "MIX-RET:42",
            "MIX-KEEP-RBX:80", "MIX-CLOBBER:11", "MIX-OFFSET:4294967297",
            "MIX-GLOBAL:1005", "MIX-DONE")
    }
    if (Test-Path "$Tests\asm\asm_mixed_x86.vbp") {
        # x86: 片段就地发 __asm{} 内联块, [X] 按名解析 (ByRef 参数经本地副本对齐 x64 语义);
        #      另加一例引用 5 个变量 (x86 无 4 个上限)
        Test-Vbp "asm_mixed_x86_inline" "$Tests\asm\asm_mixed_x86.vbp" @(
            "XMIX-ACC:175", "XMIX-TWO:22", "XMIX-BUMP:15", "XMIX-RET:42",
            "XMIX-KEEP-EBX:80", "XMIX-CLOBBER:11", "XMIX-OFFSET:4294967297",
            "XMIX-SUMMIX:1006", "XMIX-GLOBAL:1005", "XMIX-DONE") -Arch "x86"
    }
    if (Test-Path "$Tests\asm\asm_mixed_neg.vbp") {
        # 3040: [X] 既不是寄存器也不是可见的 VB 变量 (两个架构都触发)
        # 3041: x64 单个片段引用 >4 个 VB 变量 (x64 专属, 见下)
        Test-VbpBuildFail "asm_mixed_neg_unresolved_ref" "$Tests\asm\asm_mixed_neg.vbp" "3040"
    }
    if (Test-Path "$Tests\asm\asm_mixed_neg.vbp") {
        # 3041 只在 x64 出现 (x86 名字解析走栈帧, 不占参数寄存器)
        $script:total++
        Write-Host -NoNewline "  [VBP-BUILD-FAIL] asm_mixed_neg_ref_limit ... "
        $result = & cmd /c ('"' + $C3 + '" "' + "$Tests\asm\asm_mixed_neg.vbp" + '" --output-dir "' + $OutDir + '" 2>&1')
        $text = (($result | Out-String) -replace '\s+', ' ')
        if ($LASTEXITCODE -ne 0 -and $text.Contains("3041")) {
            $script:pass++
            Write-Host "PASS" -ForegroundColor Green
        } else {
            $script:fail++
            Write-Host "FAIL" -ForegroundColor Red
            Write-Host "  expected failing build containing: 3041" -ForegroundColor DarkGray
            if ($Verbose) { Write-Host $text }
        }
    }
    if (Test-Path "$Tests\asm\asm_alias_neg.vbp") {
        # 项1: cmpxchg/mul 的隐含累加器与指针/基址同族 (RAX/EAX) → 3042
        # (实测的静默死循环/段错误, 现在编译期拦住)
        Test-VbpBuildFail "asm_neg_accum_alias" "$Tests\asm\asm_alias_neg.vbp" "3042"
    }
    if (Test-Path "$Tests\asm\asm_alias_x86_neg.vbp") {
        # 同上, x86 内联块 (Test-VbpBuildFail 不带自定义参数, 就地内联判据)
        $script:total++
        Write-Host -NoNewline "  [VBP-BUILD-FAIL] asm_neg_accum_alias_x86 ... "
        $result = & cmd /c ('"' + $C3 + '" "' + "$Tests\asm\asm_alias_x86_neg.vbp" + '" --arch x86 --output-dir "' + $OutDir + '" 2>&1')
        $text = (($result | Out-String) -replace '\s+', ' ')
        if ($text -match "3042") { $script:passed++; Write-Host "PASS" -ForegroundColor Green }
        else { $script:failed++; Write-Host "FAIL (expected 3042)" -ForegroundColor Red }
    }
    if (Test-Path "$Tests\asm\asm_width_neg.vbp") {
        # 宽度不一致 (mov rax, edx) → 3038 (宽度校验前移, 不再漏到 ml64 的 A2022)
        Test-VbpBuildFail "asm_neg_operand_width" "$Tests\asm\asm_width_neg.vbp" "3038"
    }
    if (Test-Path "$Tests\asm\asm_cls_neg.vbp") {
        # 类方法里的 Asm 块 → 3037 (v2 边界: 只支持标准模块过程)
        Test-VbpBuildFail "asm_neg_class_method" "$Tests\asm\asm_cls_neg.vbp" "3037"
    }
    if (Test-Path "$Tests\asm\asm_attr_neg.vbp") {
        Test-VbpBuildFail "asm_neg_naked_on_nonproc" "$Tests\asm\asm_attr_neg.vbp" "2012"
        Test-VbpBuildFail "asm_neg_clobber_nonstring" "$Tests\asm\asm_attr_neg.vbp" "2014"
    }
    if (Test-Path "$Tests\asm\asm_x86_neg.vbp") {
        # x86 负例需要额外 --arch x86, Test-VbpBuildFail 不带自定义参数, 就地内联同款判据
        $script:total++
        Write-Host -NoNewline "  [VBP-BUILD-FAIL] asm_x86_neg ... "
        $result = & cmd /c ('"' + $C3 + '" "' + "$Tests\asm\asm_x86_neg.vbp" + '" --arch x86 --output-dir "' + $OutDir + '" 2>&1')
        $text = (($result | Out-String) -replace '\s+', ' ')
        if ($LASTEXITCODE -ne 0 -and $text.Contains("3036")) {
            $script:pass++
            Write-Host "PASS" -ForegroundColor Green
        } else {
            $script:fail++
            Write-Host "FAIL" -ForegroundColor Red
            Write-Host "  expected failing build containing: 3036" -ForegroundColor DarkGray
            if ($Verbose) { Write-Host $text }
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
