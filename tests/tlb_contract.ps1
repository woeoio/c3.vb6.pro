# ============================================================
# ai/022 B16: 类型库"契约面"读数（run_tests.ps1 用 `. $PSScriptRoot\tlb_contract.ps1` 引入）
#
# B15 把新式接口在库里的那一档摆成了真接口（TKIND_INTERFACE + Pass E 那枚 IID），但成员面
# 仍是 0 —— 库里"广告"的契约与服务器"发出"的 vtable 之间没有读数把两者对上。B16 把这一档
# 发成如实契约，本文件就是它的读数：`tests\tools\tlb_slots.cpp` 把每个 FUNC 的 vtable 面
# （cFuncs / oVft / callconv / funckind / 返回 vt / 每个参数的 vt 与标志）打回来。
# 与 `tests\tools\tlbprobe.cpp` 刻意分开：后者的打印格式是 identity 判据的 needle 来源，
# 改它等于改判据。
#
# oVft 的单位口径（B16 实测）：**建库时的位数 flag 决定存量，读端按自己的指针宽换算** ——
#   x64 建 + x64 读 → oVft=(3+槽)*8；x86 建 + x86 读 → oVft=(3+槽)*4；
#   而 x86 建 + x64 读会把 12 读成 24。所以 -Arch x86 那条**必须**用 x86 读端，它不是补充，
#   它就是"32 位客户端拿到的槽偏移对不对"的唯一读数（库里发成员却不发对偏移 = 发错）。
#
# x86 读端怎么建（LIB 顺序有坑）：套件的 $env:LIB 是 x64 在前，用 x86 的 cl 链接时会把
# x64 的 ole32.lib 挑走（实测 LNK4272 + 十个未解析外部）⇒ 这里临时把 LIB 换成 $msvc.LibX86。
# 交叉 cl（Hostx64\x86）本来就在 PATH 上（Get-MsvcToolset 装好的），x86 目标由它决定。
# ============================================================

function Get-TlbSlotsExe {
    param([string]$Arch = "x64")     # x64 | x86
    $src = Join-Path $Tests "tools\tlb_slots.cpp"
    $name = "tlb_slots.exe"
    if ($Arch -eq "x86") { $name = "tlb_slots_x86.exe" }
    $exe = Join-Path $OutDir $name
    if ((Test-Path $exe) -and ((Get-Item $src).LastWriteTime -le (Get-Item $exe).LastWriteTime)) { return $exe }
    $prev = (Get-Location).Path
    Set-Location $OutDir          # cl 把 .obj 落在当前目录，跟着产物走就不会脏工作树
    if ($Arch -eq "x86") {
        $cl = Join-Path $msvc.BinX86 "cl.exe"
        $savedLib = $env:LIB
        $env:LIB = $msvc.LibX86
        & $cl /nologo /EHsc /O1 /utf-8 $src /Fe:$exe /link ole32.lib oleaut32.lib *> $null
        $env:LIB = $savedLib
    } else {
        & cl /nologo /EHsc /O1 /utf-8 $src /Fe:$exe /link ole32.lib oleaut32.lib *> $null
    }
    Set-Location $prev
    if (-not (Test-Path $exe)) { return $null }
    return $exe
}

function Test-TlbIfaceContract {
    param(
        [string]$Name,
        [string]$VbpFile,
        [string]$TlbName,             # 产出类型库的文件名主体（= 工程输出名）
        [string[]]$Needles,
        [string[]]$Absent = @(),
        [string]$Arch = ""            # "" = 默认(x64) | x86
    )
    # vbp 分片: 本片不跑这例 (闸门在 run_tests.ps1 里定义)。
    if (-not (Enter-VbpShard)) { return }
    $script:total++
    Write-Host -NoNewline "  [TLB-CONTRACT] $Name ... "

    $c3args = @($VbpFile, "--output-dir", $OutDir, "--keep-for-debug")
    if ($Arch) { $c3args = @($VbpFile, "--arch", $Arch, "--output-dir", $OutDir, "--keep-for-debug") }
    $out = & $C3 @c3args 2>&1
    if ($LASTEXITCODE -ne 0) {
        $script:fail++
        Write-Host "FAIL (compile)" -ForegroundColor Red
        if ($Verbose) { Write-Host (($out | Out-String) -replace '\s+', ' ') }
        return
    }
    # DLL 产物的 .tlb 只活在 --keep-for-debug 的临时目录里（--emit-c 不含它）
    $genDir = ""
    $at = ($out | Out-String).IndexOf("intermediates kept at: ")
    if ($at -ge 0) { $genDir = ($out | Out-String).Substring($at + 23).Trim() }
    if (-not $genDir -or -not (Test-Path $genDir)) {
        $script:fail++
        Write-Host "FAIL (no intermediates)" -ForegroundColor Red
        return
    }
    $tlbPath = Join-Path $genDir "$TlbName.tlb"
    if (-not (Test-Path $tlbPath)) {
        $script:fail++
        Write-Host "FAIL (no tlb)" -ForegroundColor Red
        return
    }
    $readerArch = "x64"
    if ($Arch -eq "x86") { $readerArch = "x86" }
    $reader = Get-TlbSlotsExe $readerArch
    if (-not $reader) {
        $script:fail++
        Write-Host "FAIL (no reader $readerArch)" -ForegroundColor Red
        return
    }
    $text = ((& $reader $tlbPath) | Out-String)
    $flat = $text -replace '\s+', ' '

    $detail = @()
    foreach ($n in $Needles) { if (-not $flat.Contains($n)) { $detail += "missing: $n" } }
    foreach ($a in $Absent)  { if ($flat.Contains($a))     { $detail += "unexpected: $a" } }
    if ($detail.Count -eq 0) {
        $script:pass++
        Write-Host "PASS (typelib advertises the shipped vtable)" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        if ($Verbose) { foreach ($d in $detail) { Write-Host "        $d" -ForegroundColor DarkGray } }
        else { Write-Host ("        " + ($detail -join '; ')) -ForegroundColor DarkGray }
        Write-Host ("        " + $flat.Trim()) -ForegroundColor DarkGray
    }
}
