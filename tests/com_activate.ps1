# ============================================================
# ai/022 B17: 外部激活用例（run_tests.ps1 用 `. $PSScriptRoot\com_activate.ps1` 引入）
#
# B13–B16 四格把 DLL 那条对外管线验到了"进程内真客户端"（disp_probe 不查注册表）与"库里的
# 形状"（tlb_slots）。**没有一个真注册客户碰过它**：CreateObject 走注册表、早绑定客户按库里
# 的 oVft 直调 —— 这两条路此前只在本机临时产物上被探针模拟过。本文件的两个助手就是那条外部
# 通路：`tests\tools\com_act_probe.c` 自己调 DllRegisterServer，然后
#   ① 按 CLSID `CoCreateInstance` 拿 IDispatch（= CreateObject 走的那条路）+ 按名调用公有成员；
#   ② 按**接口 IID** `CoCreateInstance` 拿薄指针 + 按库里 oVft 直调契约槽；
#   ③ 反注册后三类键（CLSID / ProgID / TypeLib）都不许留 —— 用例要能重复跑、不脏机器。
# 第二个助手把"另一个进程才是被测客户"那半接上：探针只做注册/反注册，
# 中间跑一个 C3 编译出来的客户 EXE（`tests\cc_dll_client\`，CreateObject + 晚绑定调用）。
#
# 为什么注册能免管理员跑：写的是 HKCR，非管理员被 Windows 重定向到 HKCU\Software\Classes；
# 反注册走同一个入口（`DllUnregisterServer`），两边对称。
# ============================================================

function Get-ComActProbeExe {
    param([string]$Arch = "")     # "" = 默认(x64) | x86
    $src = Join-Path $Tests "tools\com_act_probe.c"
    $name = "com_act_probe.exe"
    if ($Arch -eq "x86") { $name = "com_act_probe_x86.exe" }
    $exe = Join-Path $OutDir $name
    if ((Test-Path $exe) -and ((Get-Item $src).LastWriteTime -le (Get-Item $exe).LastWriteTime)) { return $exe }
    $prev = (Get-Location).Path
    Set-Location $OutDir          # cl 把 .obj 落在当前目录，跟着产物走就不会脏工作树
    if ($Arch -eq "x86") {
        # LIB 顺序坑：$env:LIB 是 x64 在前，x86 链接会挑走 x64 的 ole32.lib（LNK4272）⇒ 临时换掉
        $cl = Join-Path $msvc.BinX86 "cl.exe"
        $savedLib = $env:LIB
        $env:LIB = $msvc.LibX86
        & $cl /nologo /W3 /O1 /utf-8 $src /Fe:$exe /link ole32.lib oleaut32.lib advapi32.lib *> $null
        $env:LIB = $savedLib
    } else {
        & cl /nologo /W3 /O1 /utf-8 $src /Fe:$exe /link ole32.lib oleaut32.lib advapi32.lib *> $null
    }
    Set-Location $prev
    if (-not (Test-Path $exe)) { return $null }
    return $exe
}

# 编译 DLL 工程, 返回 DLL 路径（外部激活的靶子）。失败返回 $null。
function Build-ComActDll {
    param([string]$VbpFile, [string]$DllName, [string]$Arch = "")
    $c3args = @($VbpFile, "--output-dir", $OutDir)
    if ($Arch) { $c3args = @($VbpFile, "--arch", $Arch, "--output-dir", $OutDir) }
    $out = & $C3 @c3args 2>&1
    if ($LASTEXITCODE -ne 0) { return @{ Dll = $null; Log = ($out | Out-String) } }
    $dll = Join-Path $OutDir "$DllName.dll"
    if (-not (Test-Path $dll)) { return @{ Dll = $null; Log = "no dll: $dll" } }
    return @{ Dll = $dll; Log = ($out | Out-String) }
}

function Test-ComActivate {
    param(
        [string]$Name,
        [string]$VbpFile,
        [string]$DllName,             # 产出 DLL 的文件名主体（= 工程输出名）
        [string]$Clsid,               # 带花括号
        [string]$ProgId,
        [string]$IfaceIid,            # 新式接口的 IID
        [string[]]$Needles,
        [string[]]$Absent = @(),
        [string]$Arch = ""            # "" = 默认(x64) | x86
    )
    # vbp 分片: 本片不跑这例 (闸门在 run_tests.ps1 里定义)。
    if (-not (Enter-VbpShard)) { return }
    $script:total++
    Write-Host -NoNewline "  [COMACT] $Name ... "

    $built = Build-ComActDll -VbpFile $VbpFile -DllName $DllName -Arch $Arch
    if (-not $built.Dll) {
        $script:fail++
        Write-Host "FAIL (compile)" -ForegroundColor Red
        if ($Verbose) { Write-Host $built.Log }
        return
    }
    $probe = Get-ComActProbeExe $Arch
    if (-not $probe) {
        $script:fail++
        Write-Host "FAIL (no probe $Arch)" -ForegroundColor Red
        return
    }
    $text = ((& $probe $built.Dll $Clsid $ProgId $IfaceIid) | Out-String)
    $flat = $text -replace '\s+', ' '

    $detail = @()
    foreach ($n in $Needles) { if (-not $flat.Contains($n)) { $detail += "missing: $n" } }
    foreach ($a in $Absent)  { if ($flat.Contains($a))     { $detail += "unexpected: $a" } }
    if ($detail.Count -eq 0) {
        $script:pass++
        Write-Host "PASS (registered, activated, cleaned up)" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        if ($Verbose) { foreach ($d in $detail) { Write-Host "        $d" -ForegroundColor DarkGray } }
        else { Write-Host ("        " + ($detail -join '; ')) -ForegroundColor DarkGray }
        Write-Host ("        " + $flat.Trim()) -ForegroundColor DarkGray
    }
}

function Test-ComActivateClient {
    param(
        [string]$Name,
        [string]$VbpFile,
        [string]$DllName,
        [string]$Clsid,
        [string]$ProgId,
        [string]$ClientVbp,           # 客户 EXE 工程（不引用 DLL, 走真实注册表链路）
        [string[]]$Needles,
        [string]$Arch = ""
    )
    # vbp 分片: 本片不跑这例 (闸门在 run_tests.ps1 里定义)。
    if (-not (Enter-VbpShard)) { return }
    $script:total++
    Write-Host -NoNewline "  [COMACT] $Name ... "

    $built = Build-ComActDll -VbpFile $VbpFile -DllName $DllName -Arch $Arch
    if (-not $built.Dll) {
        $script:fail++
        Write-Host "FAIL (compile dll)" -ForegroundColor Red
        if ($Verbose) { Write-Host $built.Log }
        return
    }
    $probe = Get-ComActProbeExe $Arch
    if (-not $probe) {
        $script:fail++
        Write-Host "FAIL (no probe $Arch)" -ForegroundColor Red
        return
    }

    $c3args = @($ClientVbp, "--output-dir", $OutDir)
    if ($Arch) { $c3args = @($ClientVbp, "--arch", $Arch, "--output-dir", $OutDir) }
    $clientLog = & $C3 @c3args 2>&1
    $clientExe = Join-Path $OutDir ([System.IO.Path]::GetFileNameWithoutExtension($ClientVbp) + ".exe")
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path $clientExe)) {
        $script:fail++
        Write-Host "FAIL (compile client)" -ForegroundColor Red
        if ($Verbose) { Write-Host ($clientLog | Out-String) }
        return
    }

    # 注册 -> 跑客户 -> 反注册: 中间那步失败也必须走到反注册, 否则机器上留键
    $detail = @()
    $clientOut = @()
    try {
        $regText = ((& $probe $built.Dll $Clsid $ProgId "reg") | Out-String)
        if (-not ($regText -replace '\s+', ' ').Contains("REGSVR hr=0x00000000")) {
            $detail += "register failed: $($regText.Trim())"
        } else {
            $run = Invoke-TestExe -ExePath $clientExe -WorkDir $OutDir -Name ([System.IO.Path]::GetFileNameWithoutExtension($ClientVbp))
            $clientOut = $run.Output
            if (-not $run.Ok) { $detail += "client run error: $($run.Detail)" }
            # #44: 客户进程"跑了却一条输出都没有"是这类红最难查的一种 —— 退出码/通道/
            # 原始字节一律进 detail, 免得只看到"missing: xxx"三行却不知道进程发生了什么。
            # (实证: 宿主把 Remove-Item 换成 fail-closed 版本时, 客户一次都没被启动,
            #  detail 里原先连异常痕迹都没有。)
            if ($clientOut.Count -eq 0 -or $run.ExitCode -ne 0) {
                $rawOut = Join-Path $OutDir "$([System.IO.Path]::GetFileNameWithoutExtension($ClientVbp)).out"
                $rawErr = Join-Path $OutDir "$([System.IO.Path]::GetFileNameWithoutExtension($ClientVbp)).err"
                $snip = "out=<no file>"
                if (Test-Path $rawOut) { $snip = "out=[$(((Get-Content $rawOut -Raw) -replace '\s+',' ').Trim())]" }
                if (Test-Path $rawErr) { $snip += " err=[$(((Get-Content $rawErr -Raw) -replace '\s+',' ').Trim())]" }
                $detail += "client silent: exit=$($run.ExitCode) via=$($run.Detail) lines=$($clientOut.Count) $snip"
            }
        }
    } finally {
        $unregText = ((& $probe $built.Dll $Clsid $ProgId "unreg") | Out-String)
        $unregFlat = $unregText -replace '\s+', ' '
        foreach ($needle in @("UNREGSVR hr=0x00000000", "CLEAN_CLSID=gone", "CLEAN_PROGID=gone", "CLEAN_TYPELIB=gone")) {
            if (-not $unregFlat.Contains($needle)) { $detail += "cleanup: missing $needle" }
        }
    }

    $joined = ($clientOut -join "`n")
    foreach ($n in $Needles) { if (-not $joined.Contains($n)) { $detail += "missing: $n" } }

    if ($detail.Count -eq 0) {
        $script:pass++
        Write-Host "PASS (external CreateObject client + cleanup)" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        if ($Verbose) { foreach ($d in $detail) { Write-Host "        $d" -ForegroundColor DarkGray } }
        else { Write-Host ("        " + ($detail -join '; ')) -ForegroundColor DarkGray }
    }
}
