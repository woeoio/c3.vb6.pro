# ============================================================
# ai/022 B14: 真客户端用例（run_tests.ps1 用 `. $PSScriptRoot\disp_invoke.ps1` 引入）
#
# B13a-B13e 五格把 DLL 那条对外管线验到了**字节**（表里的 GUID、类型库里的引用），
# 但从来没有一个调用者真的拿着产出的指针走过一次 `GetIDsOfNames`/`Invoke`。
# 本格补的就是这一层：`tests\tools\disp_probe.c` 是个不查注册表的进程内客户端 ——
# `LoadLibrary` + `DllGetClassObject` + `CreateInstance(IID_IDispatch)`，然后把
# IDispatch 四件套各问一遍。CI 上不需要 regsvr32，也不依赖机器上装过什么组件。
#
# 断言只断"点得通 / 点不通"这两件事实，不断 dispid 的具体数值（那是 typelib 回写的
# 分配顺序，改收集口径就会动，属 B15 的地盘）。
# ============================================================

function Get-DispProbeExe {
    param([string]$Arch = "")     # "" = 默认(x64) | x86
    $src = Join-Path $Tests "tools\disp_probe.c"
    $name = "disp_probe.exe"
    if ($Arch -eq "x86") { $name = "disp_probe_x86.exe" }
    $exe = Join-Path $OutDir $name
    if ((Test-Path $exe) -and ((Get-Item $src).LastWriteTime -le (Get-Item $exe).LastWriteTime)) { return $exe }
    $prev = (Get-Location).Path
    Set-Location $OutDir          # cl 把 .obj 落在当前目录，跟着产物走就不会脏工作树
    if ($Arch -eq "x86") {
        # LIB 顺序坑：$env:LIB 是 x64 在前，x86 链接会挑走 x64 的 ole32.lib（LNK4272）⇒ 临时换掉
        $cl = Join-Path $msvc.BinX86 "cl.exe"
        $savedLib = $env:LIB
        $env:LIB = $msvc.LibX86
        & $cl /nologo /W3 /O1 /utf-8 $src /Fe:$exe /link ole32.lib oleaut32.lib *> $null
        $env:LIB = $savedLib
    } else {
        & cl /nologo /W3 /O1 /utf-8 $src /Fe:$exe /link ole32.lib oleaut32.lib *> $null
    }
    Set-Location $prev
    if (-not (Test-Path $exe)) { return $null }
    return $exe
}

function Test-DispatchInvoke {
    param(
        [string]$Name,
        [string]$VbpFile,
        [string]$DllName,             # 产出 DLL 的文件名主体（= 工程输出名）
        [string]$Clsid,               # 传给探针的 CLSID（带花括号）
        [string[]]$Needles,
        [string[]]$Absent = @(),
        [string]$ExtraIid = "",       # 非空 = 再问一次这枚 IID，看服务器答的是哪份指针
        [string]$Arch = "",           # "" = 默认(x64) | x86 —— B16 起同一批断言在 x86 上也真跑一遍
        [string[]]$Tail = @()         # 账 #305 SB186: 追在探针命令行之后的臂参数 (CREATE=<clsid,...>)
    )
    # vbp 分片: 本片不跑这例 (闸门在 run_tests.ps1 里定义)。
    if (-not (Enter-VbpShard)) { return }
    $script:total++
    Write-Host -NoNewline "  [DISPATCH] $Name ... "

    $c3args = @($VbpFile, "--output-dir", $OutDir)
    if ($Arch) { $c3args = @($VbpFile, "--arch", $Arch, "--output-dir", $OutDir) }
    $out = & $C3 @c3args 2>&1
    if ($LASTEXITCODE -ne 0) {
        $script:fail++
        Write-Host "FAIL (compile)" -ForegroundColor Red
        if ($Verbose) { Write-Host (($out | Out-String) -replace '\s+', ' ') }
        return
    }
    $dllPath = Join-Path $OutDir "$DllName.dll"
    if (-not (Test-Path $dllPath)) {
        $script:fail++
        Write-Host "FAIL (no dll)" -ForegroundColor Red
        return
    }
    $probe = Get-DispProbeExe $Arch
    if (-not $probe) {
        $script:fail++
        Write-Host "FAIL (no probe)" -ForegroundColor Red
        return
    }
    # 账 #305 SB186 (第七十八刀): $Tail 里放的是探针自己的臂参数 (CREATE=<clsid[,clsid]>)，
    # 不是第二枚主 CLSID。空数组按位置展开成零个参数 ⇒ 既有各格的命令行一个字节没变。
    if ($ExtraIid) { $text = ((& $probe $dllPath $Clsid $ExtraIid @Tail) | Out-String) }
    else { $text = ((& $probe $dllPath $Clsid @Tail) | Out-String) }
    if ($LASTEXITCODE -ne 0) {
        $script:fail++
        Write-Host "FAIL (probe exit=$LASTEXITCODE)" -ForegroundColor Red
        Write-Host (($text | Out-String) -replace '\s+', ' ') -ForegroundColor DarkGray
        return
    }
    $flat = $text -replace '\s+', ' '

    $detail = @()
    foreach ($n in $Needles) { if (-not $flat.Contains($n)) { $detail += "missing: $n" } }
    foreach ($a in $Absent)  { if ($flat.Contains($a))     { $detail += "unexpected: $a" } }
    if ($detail.Count -eq 0) {
        $script:pass++
        Write-Host "PASS (real in-proc IDispatch client)" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        if ($Verbose) { foreach ($d in $detail) { Write-Host "        $d" -ForegroundColor DarkGray } }
        else { Write-Host ("        " + ($detail -join '; ')) -ForegroundColor DarkGray }
        Write-Host ("        " + $flat.Trim()) -ForegroundColor DarkGray
    }
}

# ------------------------------------------------------------
# 账 #305 §B188 (第九十刀的判据升格, 任务 #414): 「出厂那一腿的尾巴拆完了没有」是**运行期**的事实,
# 发码形状针只能钉住「尾巴在不在必经路径上」。`tests\tools\leak_probe.c` 就是不查注册表的进程内客户:
# 对一枚 terminate 会抛错的类与一枚干净的类各造各放 N 枚, 量这一段窗口里本进程 commit 的涨量。
# 判据吃的是工具自己给的 `verdict=OK` (容差一枚页) 而不是字节数 —— 改前那枚漏在 N=5000 上是 175 KB
# (x64) / 85 KB (x86), 差三个数量级, 所以既红得下来又不会被 CI 的堆噪声误红。
# ------------------------------------------------------------
function Get-LeakProbeExe {
    param([string]$Arch = "")
    $src = Join-Path $Tests "tools\leak_probe.c"
    $name = "leak_probe.exe"
    if ($Arch -eq "x86") { $name = "leak_probe_x86.exe" }
    $exe = Join-Path $OutDir $name
    if ((Test-Path $exe) -and ((Get-Item $src).LastWriteTime -le (Get-Item $exe).LastWriteTime)) { return $exe }
    $prev = (Get-Location).Path
    Set-Location $OutDir
    if ($Arch -eq "x86") {
        # LIB 顺序坑 (同 Get-DispProbeExe): $env:LIB 是 x64 在前 ⇒ 临时换掉, 否则 LNK4272
        $cl = Join-Path $msvc.BinX86 "cl.exe"
        $savedLib = $env:LIB
        $env:LIB = $msvc.LibX86
        & $cl /nologo /W3 /O1 /utf-8 $src /Fe:$exe /link ole32.lib oleaut32.lib psapi.lib *> $null
        $env:LIB = $savedLib
    } else {
        & cl /nologo /W3 /O1 /utf-8 $src /Fe:$exe /link ole32.lib oleaut32.lib psapi.lib *> $null
    }
    Set-Location $prev
    if (-not (Test-Path $exe)) { return $null }
    return $exe
}

function Test-InstanceLeak {
    param(
        [string]$Name,
        [string]$VbpFile,
        [string]$DllName,
        [string]$RaisingClsid,     # 体里会抛错的类: 本刀的主角
        [string]$CleanClsid,       # 出厂干净的类: 反面证人 (它必须一直 OK, 否则读数不是本刀的形状)
        [long]$Cycles = 5000,
        [string[]]$Needles,
        [string[]]$Absent = @(),
        [string]$Arch = ""
    )
    if (-not (Enter-VbpShard)) { return }
    $script:total++
    Write-Host -NoNewline "  [LEAK] $Name ... "

    $c3args = @($VbpFile, "--output-dir", $OutDir)
    if ($Arch) { $c3args = @($VbpFile, "--arch", $Arch, "--output-dir", $OutDir) }
    $out = & $C3 @c3args 2>&1
    if ($LASTEXITCODE -ne 0) {
        $script:fail++
        Write-Host "FAIL (compile)" -ForegroundColor Red
        if ($Verbose) { Write-Host (($out | Out-String) -replace '\s+', ' ') }
        return
    }
    $dllPath = Join-Path $OutDir "$DllName.dll"
    if (-not (Test-Path $dllPath)) {
        $script:fail++
        Write-Host "FAIL (no dll)" -ForegroundColor Red
        return
    }
    $probe = Get-LeakProbeExe $Arch
    if (-not $probe) {
        $script:fail++
        Write-Host "FAIL (no probe)" -ForegroundColor Red
        return
    }
    $text = ((& $probe $dllPath $RaisingClsid $CleanClsid $Cycles) | Out-String)
    if ($LASTEXITCODE -ne 0) {
        $script:fail++
        Write-Host "FAIL (probe exit=$LASTEXITCODE)" -ForegroundColor Red
        Write-Host (($text | Out-String) -replace '\s+', ' ') -ForegroundColor DarkGray
        return
    }
    $flat = $text -replace '\s+', ' '
    $detail = @()
    foreach ($n in $Needles) { if (-not $flat.Contains($n)) { $detail += "missing: $n" } }
    foreach ($x in $Absent)  { if ($flat.Contains($x))     { $detail += "unexpected: $x" } }
    if ($detail.Count -eq 0) {
        $script:pass++
        Write-Host "PASS (way-out tail released, both classes)" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        if ($Verbose) { foreach ($d in $detail) { Write-Host "        $d" -ForegroundColor DarkGray } }
        else { Write-Host ("        " + ($detail -join '; ')) -ForegroundColor DarkGray }
        Write-Host ("        " + $flat.Trim()) -ForegroundColor DarkGray
    }
}

# ------------------------------------------------------------
# 账 #305 第330刀 (账 #605): 「一枚 ActiveX DLL 里的 App 对象答谁」—— 形状与行为在同一趟编译里各读一次。
# VB6 实测三条 (ai/probes/app_title_dll/client_out.txt，宿主 = 32 位 cscript)：App.Title 答本工程
# 的 Title (P1)、App.EXEName 答 DLL 自己的模块名 (P3)、App.Path 答 DLL 自己的目录 (P4)。改前 C3
# 三条一起答**宿主那枚 EXE**：播种只发在 Exe 的五个入口模板里 (Type=OleDll 一条都不进)，而
# App.Path / App.EXEName 问的是 GetModuleFileNameW(NULL, …)。
# 两把尺缺一不可 —— dll_entry.c 的形状针只证「播种那一行在不在」，进程内那一点才证「读到的到底是谁」。
# 目录坑：探针 exe 与产物 DLL 若同住一个目录，App.Path 改前改后同值、那一发根本不成立 ⇒ DLL 单独
# 编进一个带 b605_self 标记的子目录，探针住在外面；标记在不在答案里就是「答 DLL / 答宿主」的判别位。
# 两台读数 (同一枚夹具、同一枚探针)：
#   改前 READ_ANS=605-title=[disp_probe]    exe=[disp_probe]    tailOk=False
#   改后 READ_ANS=605-title=[B605 Dll Title]  exe=[B605k]       tailOk=True
# ------------------------------------------------------------
function Test-DllAppSelf {
    param(
        [string]$Name,
        [string]$VbpFile,
        [string]$DllName,
        [string]$Clsid,
        [string]$Member,             # 按名点的那一发零实参 Function (探针的 CALL= 那一臂)
        [string]$SeedNeedle,         # dll_entry.c 里应当出现的播种那一行
        [string[]]$Needles,
        [string[]]$Absent = @()
    )
    # vbp 分片: 本片不跑这例 (闸门在 run_tests.ps1 里定义)。
    if (-not (Enter-VbpShard)) { return }
    $script:total++
    Write-Host -NoNewline "  [DLLAPP] $Name ... "

    $sub = Join-Path $OutDir "b605_self"
    New-Item -ItemType Directory -Force -Path $sub | Out-Null
    # dll_entry.c 不进 --emit-c 的 stdout，只活在驱动那句 "intermediates kept at: <dir>"
    # 指到的临时目录里 —— 与 Test-VbpDll 用的是同一把手。
    $out = & $C3 $VbpFile --output-dir $sub --keep-for-debug 2>&1
    if ($LASTEXITCODE -ne 0) {
        $script:fail++
        Write-Host "FAIL (compile)" -ForegroundColor Red
        if ($Verbose) { Write-Host (($out | Out-String) -replace '\s+', ' ') }
        return
    }
    $dllPath = Join-Path $sub "$DllName.dll"
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
    if (-not $genDir) {
        $script:fail++
        Write-Host "FAIL (no intermediates)" -ForegroundColor Red
        return
    }
    $entryPath = Join-Path $genDir "dll_entry.c"
    if (-not (Test-Path $entryPath)) {
        $script:fail++
        Write-Host "FAIL (no dll_entry.c)" -ForegroundColor Red
        return
    }
    # Test-Path 认旧产物 (账 #c3-build-test 那一坑)：上一趟留下的同名 dll 会被当成这一趟的读数，
    # 所以 $dllPath 那一步之前已经先验过 $LASTEXITCODE。
    $entry = ((Get-Content $entryPath -Raw) -replace '\s+', ' ')

    $probe = Get-DispProbeExe ""
    if (-not $probe) {
        $script:fail++
        Write-Host "FAIL (no probe)" -ForegroundColor Red
        return
    }
    $text = ((& $probe $dllPath $Clsid "CALL=$Member") | Out-String)
    if ($LASTEXITCODE -ne 0) {
        $script:fail++
        Write-Host "FAIL (probe exit=$LASTEXITCODE)" -ForegroundColor Red
        Write-Host (($text | Out-String) -replace '\s+', ' ') -ForegroundColor DarkGray
        return
    }
    $flat = ($text -replace '\s+', ' ') + "|" + $entry

    $detail = @()
    if (-not $entry.Contains($SeedNeedle)) { $detail += ("dll_entry missing seed: " + $SeedNeedle) }
    foreach ($n in $Needles) { if (-not $flat.Contains($n)) { $detail += "missing: $n" } }
    foreach ($a in $Absent) { if ($flat.Contains($a)) { $detail += "unexpected: $a" } }
    if ($detail.Count -eq 0) {
        $script:pass++
        Write-Host "PASS (App answers this product, not the host)" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        if ($Verbose) { foreach ($d in $detail) { Write-Host "        $d" -ForegroundColor DarkGray } }
        else { Write-Host ("        " + ($detail -join '; ')) -ForegroundColor DarkGray }
        Write-Host ("        " + $flat.Trim()) -ForegroundColor DarkGray
    }
}
