# =====================================================================
# regress_all.ps1 - tests 全量用例对照回归 (排除 vbman)
#
# 用 C3 与本机微软真 VB6 (VB6Mini 安装, VB6.EXE /make) 对同一批用例
# 各自编译/运行, 按对照矩阵给出结论, 用于发现 C3 侧回归。
#
# 用法 (在 tests 目录):
#   .\regress_all.ps1                          # 全量
#   .\regress_all.ps1 -Filter test_iif*        # 只跑匹配用例
#   .\regress_all.ps1 -List                    # 只列出用例, 不执行
#   .\regress_all.ps1 -CompileOnly             # 只比编译, 不运行
#
# 环境: 与 run_tests.ps1 相同 -- 环境变量 C3_VCVARSALL 优先,
#       否则 vswhere 自动探测 vcvarsall (VS2022 C++ 工作负载)。
#
# 判定矩阵:
#   PASS(对照通过)  : 双方编译 OK, 运行退出码一致(且都有输出时输出一致)
#   REG-C3(红)      : 真 VB6 能编译, C3 编译失败   <-- 重点怀疑回归
#   C3-EXT(蓝)      : C3 能编译, 真 VB6 不支持 (C3 扩展语法/能力)
#   BOTH-FAIL       : 双方都编译失败
#   RUN-DIFF        : 双方编译 OK 但运行退出码/输出不一致
#   RUN-TIMEOUT     : 运行超时被杀 (单侧或双侧)
#   SKIP-AUX        : 非独立程序 (无 Sub Main 的模块/被 vbp 引用的成员文件)
#
# 已知既有差异 (非本次拆分引入, 见工作日志):
#   test_concat_leak / test_concat_stress / test_p613_typelib (C3 侧
#   LNK2019 vb6_Form_Print), test_sysfunc (vb6_CurDir 签名), winhttp
#   (vb6_vsink_http_create), test_implements_qi / vbp_project\TestVBP
#   (双方都超时)。
# =====================================================================
param(
    [string]$Filter = "*",          # 用例名通配过滤
    [string]$C3 = "",               # C3.exe 路径, 默认自动找 .build / publish
    [string]$VB6 = "",              # VB6.EXE 路径, 默认自动找 VB6Mini
    [string]$WorkRoot = "",         # 工作目录, 默认 tests\_vb6ref
    [int]$RunTimeoutSec = 30,       # 单个用例运行超时
    [int]$BuildTimeoutSec = 180,    # 单次编译超时
    [switch]$CompileOnly,           # 只比编译不运行
    [switch]$List,                  # 只列出用例
    [int]$MaxCases = 0              # >0 时限制用例数 (调试用)
)

$ErrorActionPreference = "SilentlyContinue"

# === 路径 ===
$Root  = Split-Path -Parent $PSScriptRoot
$Tests = $PSScriptRoot
if (-not $WorkRoot) { $WorkRoot = $env:C3_TESTS_WORKROOT }   # 环境变量可覆盖
if (-not $WorkRoot) { $WorkRoot = Join-Path $Tests "_vb6ref" }

if (-not $C3) {
    # 环境变量 C3_EXE 优先, 其次按仓库相对路径推导
    $C3 = $env:C3_EXE
}
if (-not $C3) {
    foreach ($c in @((Join-Path $Root ".build\C3.exe"), (Join-Path $Root "publish\C3.exe"))) {
        if (Test-Path $c) { $C3 = $c; break }
    }
}
if (-not $C3 -or -not (Test-Path $C3)) {
    Write-Host "[ERROR] 未找到 C3.exe, 请用 -C3 或环境变量 C3_EXE 指定" -ForegroundColor Red; exit 1
}

if (-not $VB6) {
    # 环境变量 VB6_EXE 优先, 其次在 ProgramFiles 下探测 VB6Mini
    $VB6 = $env:VB6_EXE
}
if (-not $VB6) {
    foreach ($base in @(${env:ProgramFiles(x86)}, $env:ProgramFiles)) {
        $cand = Join-Path $base "VB6Mini\bin\VB6.EXE"
        if (Test-Path $cand) { $VB6 = $cand; break }
    }
}
if (-not $VB6 -or -not (Test-Path $VB6)) {
    Write-Host "[ERROR] 未找到 VB6.EXE (VB6Mini), 请用 -VB6 或环境变量 VB6_EXE 指定" -ForegroundColor Red; exit 1
}

# === 用例发现 ===
# vbp 成员解析: 从 vbp 文本里抓 Module=/Class=/Form= 等行引用的源文件
function Get-VbpMembers {
    param([string]$VbpPath)
    $members = @()
    $lines = [IO.File]::ReadAllLines($VbpPath, [Text.Encoding]::Default)
    foreach ($ln in $lines) {
        # Module=Module1; test_iif.bas / Form=Form1.frm / Class=Cls; a.cls
        if ($ln -match '^\s*(Module|Class|Form|UserControl|PropertyPage)\s*=\s*[^;]*;\s*(.+?)\s*$') {
            $members += $matches[2].Trim('"')
        }
    }
    return ,$members
}

# 判断源文件是否含可启动入口
function Test-HasMain {
    param([string]$Path)
    $txt = [IO.File]::ReadAllText($Path, [Text.Encoding]::Default)
    return ($txt -match '(?im)^\s*(public\s+|private\s+)?sub\s+main\s*\(')
}

# 判断 frm 的窗体名
function Get-FormName {
    param([string]$Path)
    $txt = [IO.File]::ReadAllText($Path, [Text.Encoding]::Default)
    if ($txt -match '(?m)^\s*Begin\s+VB\.Form\s+(\S+)') { return $matches[1] }
    return $null
}

# 收集用例: 排除 vbman* / 工作目录 / 输出目录
$excludeDirPat = '^(.*vbman.*|_vb6ref.*|output|node_modules)$'
$allVbp   = Get-ChildItem $Tests -Recurse -Filter *.vbp | Where-Object {
    $rel = $_.DirectoryName.Substring($Tests.Length).TrimStart('\')
    $top = ($rel -split '\\')[0]
    $top -notmatch $excludeDirPat
}
$referenced = New-Object System.Collections.Generic.HashSet[string]
foreach ($v in $allVbp) {
    foreach ($m in (Get-VbpMembers $v.FullName)) {
        [void]$referenced.Add((Join-Path $v.DirectoryName $m).ToLower())
    }
}

$cases = @()
foreach ($v in $allVbp) {
    if ($v.Name -notlike $Filter) { continue }
    $members = @(Get-VbpMembers $v.FullName | ForEach-Object { Join-Path $v.DirectoryName $_ })
    $hasForm = ($members | Where-Object { $_ -like '*.frm' }).Count -gt 0
    $cases += [pscustomobject]@{
        Name = $v.BaseName; Type = "vbp"; Source = $v.FullName
        Members = $members; HasForm = $hasForm; Dir = $v.DirectoryName
    }
}
$allBasFrm = Get-ChildItem $Tests -Recurse -Include *.bas, *.frm | Where-Object {
    $rel = $_.DirectoryName.Substring($Tests.Length).TrimStart('\')
    $top = ($rel -split '\\')[0]
    ($top -notmatch $excludeDirPat) -and (-not $referenced.Contains($_.FullName.ToLower()))
}
foreach ($f in $allBasFrm) {
    if ($f.Name -notlike $Filter) { continue }
    if ($f.Extension -eq ".frm") {
        $cases += [pscustomobject]@{
            Name = $f.BaseName; Type = "frm"; Source = $f.FullName
            Members = @($f.FullName); HasForm = $true; Dir = $f.DirectoryName
        }
    } else {
        if (-not (Test-HasMain $f.FullName)) {
            $cases += [pscustomobject]@{
                Name = $f.BaseName; Type = "aux"; Source = $f.FullName
                Members = @(); HasForm = $false; Dir = $f.DirectoryName
            }
            continue
        }
        $cases += [pscustomobject]@{
            Name = $f.BaseName; Type = "bas"; Source = $f.FullName
            Members = @($f.FullName); HasForm = $false; Dir = $f.DirectoryName
        }
    }
}
$cases = @($cases | Sort-Object Type, Name)
if ($MaxCases -gt 0) { $cases = $cases | Select-Object -First $MaxCases }

if ($List) {
    Write-Host "共 $($cases.Count) 个用例 (排除 vbman):"
    $cases | ForEach-Object { Write-Host ("  [{0,4}] {1,-4} {2}" -f $_.Type, "", $_.Source.Substring($Tests.Length + 1)) }
    exit 0
}

# === MSVC 环境 (与 run_tests.ps1 同一套约定; -List 模式不需要, 放在枚举之后) ===
# 除设置进程环境外, 还把变量存进 $script:MsvcEnv, 由 Invoke-Proc 显式传给子进程 --
# 显式传env 可规避个别机器上进程环境块存在重复 PATH 条目 (PATH/Path) 导致
# 子进程解析到旧值的问题 (实测 cl.exe 找不到)。
$script:MsvcEnv = @{}
# 通用 MSVC 环境解析 (不依赖 cmd.exe / vcvarsall 解析, 不硬编码 VS/SDK 版本)
# 优先 C3_VCVARSALL 推导 VS 根; 否则 vswhere 取最新已装 VS; toolset 与 SDK 版本动态发现。
# 解析结果存入 $script:MsvcEnv (供 Invoke-Proc 显式传给子进程) 并同步设置进程环境。
function Get-MsvcToolset {
    $vsRoot = $null
    if ($env:C3_VCVARSALL -and (Test-Path $env:C3_VCVARSALL)) {
        $p = $env:C3_VCVARSALL
        for ($i = 0; $i -lt 4; $i++) { $p = Split-Path -Parent $p }
        $vsRoot = $p
    }
    if (-not $vsRoot) {
        $vswhere = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
        if (Test-Path $vswhere) {
            $vsPath = & $vswhere -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath 2>$null
            if ($vsPath) { $vsRoot = $vsPath }
        }
    }
    if (-not $vsRoot -or -not (Test-Path $vsRoot)) {
        Write-Host "[WARN] 未找到 Visual Studio (含 VC.Tools), 编译可能失败; 可设置 C3_VCVARSALL" -ForegroundColor Yellow
        return $null
    }
    $msvcRoot = Join-Path $vsRoot "VC\Tools\MSVC"
    $toolVer = Get-ChildItem $msvcRoot -Directory -ErrorAction SilentlyContinue |
        Sort-Object Name -Descending | Select-Object -First 1 -ExpandProperty Name
    if (-not $toolVer) { Write-Host "[WARN] 未找到 MSVC toolset ($msvcRoot)" -ForegroundColor Yellow; return $null }
    $kitRoot = $null
    foreach ($base in @(${env:ProgramFiles(x86)}, $env:ProgramFiles)) {
        if ($base -and (Test-Path (Join-Path $base "Windows Kits\10\Include"))) { $kitRoot = Join-Path $base "Windows Kits\10"; break }
    }
    if (-not $kitRoot) {
        foreach ($rp in @("HKLM:\SOFTWARE\Microsoft\Windows Kits\Installed Roots",
                          "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows Kits\Installed Roots")) {
            if (Test-Path $rp) {
                $kr = (Get-ItemProperty -Path $rp -ErrorAction SilentlyContinue).KitsRoot10
                if ($kr) { $kr = $kr.TrimEnd('\'); if (Test-Path (Join-Path $kr "Include")) { $kitRoot = $kr; break } }
            }
        }
    }
    if (-not $kitRoot) {
        foreach ($d in @("D:","E:","F:")) {
            $cand = Join-Path $d "Windows Kits\10"
            if (Test-Path (Join-Path $cand "Include")) { $kitRoot = $cand; break }
        }
    }
    if (-not $kitRoot) { Write-Host "[WARN] 未找到 Windows SDK (Windows Kits\10\Include)" -ForegroundColor Yellow; return $null }
    $sdkVer = Get-ChildItem (Join-Path $kitRoot "Include") -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^\d+\.\d+\.\d+\.\d+$' } |
        Sort-Object Name -Descending | Select-Object -First 1 -ExpandProperty Name
    if (-not $sdkVer) { Write-Host "[WARN] 未找到 Windows SDK Include 版本 ($kitRoot\Include)" -ForegroundColor Yellow; return $null }

    $tool   = Join-Path $msvcRoot $toolVer
    $binX64 = Join-Path $tool "bin\Hostx64\x64"
    $binX86 = Join-Path $tool "bin\Hostx64\x86"
    $incDir = Join-Path $kitRoot "Include"
    $libDir = Join-Path $kitRoot "Lib"
    $Include = "$tool\include;$incDir\$sdkVer\um;$incDir\$sdkVer\ucrt;$incDir\$sdkVer\shared;$incDir\$sdkVer\winrt;$incDir\$sdkVer\cppwinrt"
    $LibX64  = "$tool\lib\x64;$libDir\$sdkVer\um\x64;$libDir\$sdkVer\ucrt\x64"
    $LibX86  = "$tool\lib\x86;$libDir\$sdkVer\um\x86;$libDir\$sdkVer\ucrt\x86"
    return [pscustomobject]@{ VsRoot=$vsRoot; ToolVer=$toolVer; SdkVer=$sdkVer;
        BinX64=$binX64; BinX86=$binX86; Include=$Include; LibX64=$LibX64; LibX86=$LibX86 }
}

$msvc = Get-MsvcToolset
if ($msvc) {
    $script:MsvcEnv["PATH"]    = "$($msvc.BinX64);$($msvc.BinX86);$env:PATH"
    $script:MsvcEnv["INCLUDE"] = $msvc.Include
    $script:MsvcEnv["LIB"]     = "$($msvc.LibX64);$($msvc.LibX86)"
    $env:PATH    = $script:MsvcEnv["PATH"]
    $env:INCLUDE = $script:MsvcEnv["INCLUDE"]
    $env:LIB     = $script:MsvcEnv["LIB"]
    $clCmd = Get-Command cl.exe -ErrorAction SilentlyContinue
    Write-Host ("  MSVC 环境: VS=$($msvc.VsRoot) toolset=$($msvc.ToolVer) SDK=$($msvc.SdkVer) cl={0}" -f $clCmd.Source) -ForegroundColor Gray
} else {
    Write-Host "[WARN] 未成功配置 MSVC 环境, 编译可能失败" -ForegroundColor Yellow
}

# === 工具函数 ===
# 带超时运行进程, 返回 @{ ExitCode; Stdout; TimedOut }
# 注意: 参数名不能用 $Args (PowerShell 自动变量)
function Invoke-Proc {
    param([string]$Exe, [string[]]$ArgList, [string]$Cwd, [int]$TimeoutSec, [string]$StdoutFile = "", [switch]$ShellExecute)
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $Exe
    if ($ArgList) {
        # 仅含空格/引号的参数才加引号 (VB6.EXE 的 "/make" 带引号会不被识别)
        $psi.Arguments = ($ArgList | ForEach-Object {
            if ($_ -match '[\s"]') { '"{0}"' -f $_ } else { $_ }
        }) -join " "
    }
    $psi.WorkingDirectory = $Cwd
    # VB6.EXE 清单要求管理员, UseShellExecute=false 会报"需要提升", 这类进程走 ShellExecute
    $psi.UseShellExecute = [bool]$ShellExecute
    $redirOut = ($StdoutFile -ne "" -and -not $ShellExecute)
    $psi.RedirectStandardOutput = $redirOut
    $psi.RedirectStandardError = $false
    $psi.CreateNoWindow = -not $ShellExecute
    # 显式下发 MSVC 环境 (含 PATH), 不依赖进程环境继承 (仅 UseShellExecute=false 支持)
    if (-not $ShellExecute -and $script:MsvcEnv -and $script:MsvcEnv.Count -gt 0) {
        foreach ($kv in $script:MsvcEnv.GetEnumerator()) {
            $psi.EnvironmentVariables[$kv.Key] = $kv.Value
        }
    }
    $p = $null; $startErr = ""
    try {
        $p = [System.Diagnostics.Process]::Start($psi)
    } catch {
        $startErr = $_.Exception.Message
    }
    if (-not $p) {
        return @{ ExitCode = -99; Stdout = ""; TimedOut = $false; Err = $startErr }
    }
    if ($redirOut) {
        $waitTask = $p.StandardOutput.ReadToEndAsync()
    }
    if (-not $p.WaitForExit($TimeoutSec * 1000)) {
        try { $p.Kill() } catch {}
        $p.WaitForExit(5000) | Out-Null
        return @{ ExitCode = -1; Stdout = ""; TimedOut = $true; Err = "" }
    }
    $so = ""
    if ($redirOut) {
        try { $so = $waitTask.Result } catch { $so = "" }
    }
    return @{ ExitCode = $p.ExitCode; Stdout = $so; TimedOut = $false; Err = "" }
}

# C3 编译: 成功时返回产物 exe 路径 (编译 cwd 用隔离工作区, 避免错误日志污染源目录)
function Invoke-C3Build {
    param([string]$CaseDir, [string]$Src, [string]$OutDir)
    New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
    $r = Invoke-Proc -Exe $C3 -ArgList @($Src, "--output-dir", $OutDir) -Cwd $OutDir -TimeoutSec $BuildTimeoutSec
    if ($r.Err) { return @{ Ok = $false; Note = "无法启动 C3: $($r.Err)" } }
    if ($r.TimedOut) { return @{ Ok = $false; Note = "编译超时" } }
    if ($r.ExitCode -ne 0) { return @{ Ok = $false; Note = "exit=$($r.ExitCode)" } }
    $exe = Get-ChildItem $OutDir -Filter *.exe | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $exe) { return @{ Ok = $false; Note = "无产物 exe" } }
    return @{ Ok = $true; Exe = $exe.FullName; Note = "" }
}

# 真 VB6 编译: 为独立 bas/frm 生成 vbp, 用 VB6.EXE /make; 成功标准 = 产出 exe
function Invoke-VB6Build {
    param([string]$CaseDir, [string]$WorkDir, [string]$Src, [string]$CaseType, [string]$CaseName)
    New-Item -ItemType Directory -Force -Path $WorkDir | Out-Null
    if ($CaseType -eq "vbp") {
        Copy-Item $Src $WorkDir -Force
        foreach ($m in (Get-VbpMembers $Src)) {
            $mp = Join-Path (Split-Path -Parent $Src) $m
            if (Test-Path $mp) { Copy-Item $mp $WorkDir -Force }
        }
        $vbp = Join-Path $WorkDir (Split-Path -Leaf $Src)
    } else {
        $srcName = Split-Path -Leaf $Src
        Copy-Item $Src $WorkDir -Force
        $lines = @("Type=Exe",
            "Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#$env:SystemRoot\System32\stdole2.tlb#OLE Automation")
        if ($CaseType -eq "frm") {
            # vbp 的 Form= 行不带名称前缀 (窗体名在 frm 内部)
            # 启动对象用空 Sub Main 模块: MDI 子窗体等不能直接当启动窗体, 但编译对照只需编译通过
            $mainBas = Join-Path $WorkDir "__modmain.bas"
            [IO.File]::WriteAllText($mainBas, "Attribute VB_Name = `"ModMain`"`r`nSub Main()`r`nEnd Sub`r`n", [Text.Encoding]::ASCII)
            $lines += "Module=ModMain; __modmain.bas"
            $lines += 'Startup="Sub Main"'
            $lines += "Form=$srcName"
        } else {
            $lines += "Module=Module1; $srcName"
            $lines += 'Startup="Sub Main"'
        }
        $vbp = Join-Path $WorkDir "$CaseName.vbp"
        [IO.File]::WriteAllLines($vbp, $lines, [Text.Encoding]::ASCII)
    }
    $log = Join-Path $WorkDir "vb6_build.log"
    $r = Invoke-Proc -Exe $VB6 -ArgList @("/make", (Split-Path -Leaf $vbp), "/out", "vb6_build.log") -Cwd $WorkDir -TimeoutSec $BuildTimeoutSec -ShellExecute
    if ($r.Err) { return @{ Ok = $false; Note = "无法启动 VB6.EXE: $($r.Err)" } }
    if ($r.TimedOut) { return @{ Ok = $false; Note = "编译超时" } }
    $exe = Get-ChildItem $WorkDir -Filter *.exe | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    $logText = ""
    if (Test-Path $log) { $logText = [IO.File]::ReadAllText($log, [Text.Encoding]::Default) }
    if (-not $exe) {
        $why = "无产物"
        if ($logText -match '([^\r\n]*(?:错误|error|无法|不能|失败)[^\r\n]*)') { $why = $matches[1].Trim() }
        return @{ Ok = $false; Note = $why }
    }
    return @{ Ok = $true; Exe = $exe.FullName; Note = "" }
}

# 规范化输出用于比较
function Normalize-Out([string]$s) {
    return (($s -replace "`r`n", "`n").Trim())
}

# === 主循环 ===
if (-not (Test-Path $WorkRoot)) { New-Item -ItemType Directory -Path $WorkRoot | Out-Null }
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$csvPath = Join-Path $WorkRoot "report-$stamp.csv"
$rows = @()
$n = 0

# CSV 增量写入: 全量运行较长, 中断时已跑部分的结果也保留
$script:csvInit = $false
function Add-Result {
    param([hashtable]$Row)
    $script:rows += [pscustomobject]$Row
    if (-not $script:csvInit) {
        $script:rows | Export-Csv $script:csvPath -NoTypeInformation -Encoding UTF8
        $script:csvInit = $true
    } else {
        $script:rows | Select-Object -Last 1 | Export-Csv $script:csvPath -NoTypeInformation -Encoding UTF8 -Append
    }
}

Write-Host ""
Write-Host "=== tests 全量对照回归 (C3 vs 真 VB6) ===" -ForegroundColor Cyan
Write-Host "  C3 : $C3"
Write-Host "  VB6: $VB6"
Write-Host "  用例: $($cases.Count) (排除 vbman)  超时: 运行 ${RunTimeoutSec}s / 编译 ${BuildTimeoutSec}s"
Write-Host ""

foreach ($case in $cases) {
    $n++
    $tag = "[$n/$($cases.Count)]"
    $safeName = ($case.Source.Substring($Tests.Length + 1) -replace '[\\/]', '__')

    if ($case.Type -eq "aux") {
        Write-Host ("{0} {1,-40} SKIP-AUX (非独立程序)" -f $tag, $case.Name) -ForegroundColor DarkGray
        Add-Result @{ Case = $safeName; Type = $case.Type; C3Compile = "SKIP"; VB6Compile = "SKIP";
            C3Run = ""; VB6Run = ""; C3Exit = ""; VB6Exit = ""; Verdict = "SKIP-AUX"; Note = "无 Sub Main" }
        continue
    }

    $c3Dir  = Join-Path $WorkRoot "work\$safeName.c3"
    $vb6Dir = Join-Path $WorkRoot "work\$safeName.vb6"
    Remove-Item $c3Dir, $vb6Dir -Recurse -Force -ErrorAction SilentlyContinue

    $c3b  = Invoke-C3Build  -CaseDir $case.Dir -Src $case.Source -OutDir $c3Dir
    $vb6b = Invoke-VB6Build -CaseDir $case.Dir -WorkDir $vb6Dir -Src $case.Source -CaseType $case.Type -CaseName $case.Name

    $verdict = ""; $note = ""
    $c3c = "OK"; if (-not $c3b.Ok)  { $c3c = "FAIL" }
    $v6c = "OK"; if (-not $vb6b.Ok) { $v6c = "FAIL" }

    if (-not $c3b.Ok -and -not $vb6b.Ok) {
        $verdict = "BOTH-FAIL"; $note = "C3:$($c3b.Note) | VB6:$($vb6b.Note)"
        Write-Host ("{0} {1,-40} BOTH-FAIL" -f $tag, $case.Name) -ForegroundColor DarkYellow
    } elseif (-not $c3b.Ok -and $vb6b.Ok) {
        $verdict = "REG-C3"; $note = "真VB6可编译, C3失败: $($c3b.Note)"
        Write-Host ("{0} {1,-40} REG-C3  <-- 可疑回归" -f $tag, $case.Name) -ForegroundColor Red
    } elseif ($c3b.Ok -and -not $vb6b.Ok) {
        $verdict = "C3-EXT"; $note = "C3可编译, 真VB6不支持: $($vb6b.Note)"
        Write-Host ("{0} {1,-40} C3-EXT" -f $tag, $case.Name) -ForegroundColor Blue
    } else {
        # 双方编译 OK
        if ($CompileOnly -or $case.HasForm) {
            $verdict = "COMPILE-PASS"
            if ($case.HasForm) { $note = "含窗体, 只比编译" }
            Write-Host ("{0} {1,-40} COMPILE-PASS{2}" -f $tag, $case.Name, $(if ($note) { " (窗体)" } else { "" })) -ForegroundColor Green
        } else {
            # 运行 cwd 用原用例目录 (部分用例按相对路径读文件)
            $c3r  = Invoke-Proc -Exe $c3b.Exe  -ArgList @() -Cwd $case.Dir -TimeoutSec $RunTimeoutSec -StdoutFile (Join-Path $c3Dir  "stdout.txt")
            $vb6r = Invoke-Proc -Exe $vb6b.Exe -ArgList @() -Cwd $case.Dir -TimeoutSec $RunTimeoutSec
            $c3run = "OK"; if ($c3r.TimedOut)  { $c3run = "TIMEOUT" }
            $v6run = "OK"; if ($vb6r.TimedOut) { $v6run = "TIMEOUT" }

            if ($c3r.TimedOut -or $vb6r.TimedOut) {
                if ($c3r.TimedOut -and $vb6r.TimedOut) { $verdict = "RUN-TIMEOUT"; $note = "双方都超时" }
                elseif ($c3r.TimedOut) { $verdict = "RUN-TIMEOUT"; $note = "C3 侧超时(VB6 正常退出 $($vb6r.ExitCode))" }
                else { $verdict = "RUN-TIMEOUT"; $note = "VB6 侧超时(C3 正常退出 $($c3r.ExitCode))" }
                Write-Host ("{0} {1,-40} RUN-TIMEOUT" -f $tag, $case.Name) -ForegroundColor Yellow
            } elseif ($c3r.ExitCode -ne $vb6r.ExitCode) {
                $verdict = "RUN-DIFF"; $note = "退出码 C3=$($c3r.ExitCode) vs VB6=$($vb6r.ExitCode)"
                Write-Host ("{0} {1,-40} RUN-DIFF" -f $tag, $case.Name) -ForegroundColor Red
            } else {
                # 退出码一致; 若双方都有 stdout 则比内容 (真 VB6 的 Debug.Print 静默, 通常不比)
                $vOut = Normalize-Out $vb6r.Stdout
                $cOut = Normalize-Out $c3r.Stdout
                if ($vOut -ne "" -and $cOut -ne "" -and $vOut -ne $cOut) {
                    $verdict = "RUN-DIFF"; $note = "输出不一致"
                    Write-Host ("{0} {1,-40} RUN-DIFF (输出)" -f $tag, $case.Name) -ForegroundColor Red
                } else {
                    $verdict = "PASS"
                    Write-Host ("{0} {1,-40} PASS (exit=$($c3r.ExitCode))" -f $tag, $case.Name) -ForegroundColor Green
                }
            }
            Add-Result @{ Case = $safeName; Type = $case.Type; C3Compile = $c3c; VB6Compile = $v6c;
                C3Run = $c3run; VB6Run = $v6run; C3Exit = $c3r.ExitCode; VB6Exit = $vb6r.ExitCode; Verdict = $verdict; Note = $note }
            continue
        }
    }
    Add-Result @{ Case = $safeName; Type = $case.Type; C3Compile = $c3c; VB6Compile = $v6c;
        C3Run = ""; VB6Run = ""; C3Exit = ""; VB6Exit = ""; Verdict = $verdict; Note = $note }
}

# === 汇总 ===
Write-Host ""
Write-Host "=== 汇总 ===" -ForegroundColor Cyan
$rows | Group-Object Verdict | Sort-Object Count -Descending | ForEach-Object {
    Write-Host ("  {0,-14} {1}" -f $_.Name, $_.Count)
}
Write-Host ""
Write-Host "报告: $csvPath"

$reg = @($rows | Where-Object { $_.Verdict -eq "REG-C3" })
if ($reg.Count -gt 0) {
    Write-Host ""
    Write-Host "!! 发现 $($reg.Count) 个 REG-C3 (真VB6可编译而C3失败), 需要排查:" -ForegroundColor Red
    $reg | ForEach-Object { Write-Host ("   {0}  {1}" -f $_.Case, $_.Note) -ForegroundColor Red }
    exit 2
}
exit 0
