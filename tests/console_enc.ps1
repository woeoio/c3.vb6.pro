# ============================================================
# 控制台/管道输出的编码用例（run_tests.ps1 用 `. $PSScriptRoot\console_enc.ps1` 引入）
#
# 背景：C3 内部是 UTF-8、程序内部是 UTF-16，而 Windows 这一侧有两个消费者：
#   * **控制台**（cmd / Windows Terminal / ConPTY）：走 WriteConsoleW 交 UTF-16 最稳 ——
#     控制台渲染宽字符**与 chcp 无关**（实测同一份输出在 936/65001/437 下逐字相同）；
#   * **管道/文件**（PowerShell 接走 native 输出、用户 `> file`）：只能按**字节**约定，
#     而 shell 按它启动时的 [Console]::OutputEncoding（默认 = 当时的控制台代码页）解码
#     ⇒ 运行库按"控制台输出代码页"写字节（装不下就退回 UTF-8），这是任何语言都能自洽的那条。
#
# 两条用例分别钉这两条路：
#   [ENC] cc_cn_console  —— 真控制台里跑四语种样本，读回**屏幕缓冲**逐字断言（x64/x86 各一遍）；
#   [ENC] cc_cn_pipe     —— `chcp 936 & prog > file`，断言落盘字节是 GBK 里的那四段文本。
# 这两条合起来覆盖"en/zh/ja/ko 在 cmd 与重定向下都不乱码"；而"在 shell 里显示任意语言"的
# 通用姿势是先把控制台切成 UTF-8（chcp 65001）再起 shell —— 那时运行库按 65001 输出 UTF-8，
# shell 也按 UTF-8 解码，四语种同时成立（见 ai/022 的 D66 读数）。
# ============================================================

function Get-ConCaptureExe {
    # tests\tools\con_capture.c：把命令行跑在真控制台里再读回屏幕缓冲
    $src = Join-Path $Tests "tools\con_capture.c"
    $name = "con_capture.exe"
    $exe = Join-Path $OutDir $name
    if ((Test-Path $exe) -and ((Get-Item $src).LastWriteTime -le (Get-Item $exe).LastWriteTime)) { return $exe }
    $prev = (Get-Location).Path
    Set-Location $OutDir
    & cl /nologo /W3 /O1 /utf-8 $src /Fe:$exe /link user32.lib *> $null
    Set-Location $prev
    if (-not (Test-Path $exe)) { return $null }
    return $exe
}

# 编译四语种样本（UTF-8 BOM 源），返回 exe 路径
function Build-CnSample {
    param([string]$Source, [string]$Arch = "")
    $c3args = @($Source, "--output-dir", $OutDir)
    if ($Arch) { $c3args = @($Source, "--arch", $Arch, "--output-dir", $OutDir) }
    $out = & $C3 @c3args 2>&1
    if ($LASTEXITCODE -ne 0) { return @{ Exe = $null; Log = ($out | Out-String) } }
    $exe = Join-Path $OutDir ([System.IO.Path]::GetFileNameWithoutExtension($Source) + ".exe")
    if (-not (Test-Path $exe)) { return @{ Exe = $null; Log = "no exe: $exe" } }
    return @{ Exe = $exe; Log = "" }
}

# 按 UTF-8 显式读子进程的 stdout：探针写的是 UTF-8，而 PowerShell 对 native 子进程
# 默认按 [Console]::OutputEncoding（=控制台代码页）解码 —— 直接 `& probe` 会把读数自己弄成乱码。
function Invoke-CapturedUtf8 {
    param([string]$Exe, [string[]]$Arguments)
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $Exe
    $psi.Arguments = (($Arguments | ForEach-Object { '"' + $_ + '"' }) -join ' ')
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8
    $p = [System.Diagnostics.Process]::Start($psi)
    $text = $p.StandardOutput.ReadToEnd()
    $p.WaitForExit(30000) | Out-Null
    return $text
}

function Test-CnConsoleOutput {
    param(
        [string]$Name,
        [string]$Source,
        [string[]]$Needles,          # 四语种各一段（读的是**屏幕缓冲**, 与代码页无关）
        [string]$Arch = ""
    )
    # vbp 分片: 本片不跑这例 (闸门在 run_tests.ps1 里定义)。
    if (-not (Enter-VbpShard)) { return }
    $script:total++
    Write-Host -NoNewline "  [ENC] $Name ... "

    $built = Build-CnSample -Source $Source -Arch $Arch
    if (-not $built.Exe) {
        $script:fail++
        Write-Host "FAIL (compile)" -ForegroundColor Red
        if ($Verbose) { Write-Host $built.Log }
        return
    }
    $probe = Get-ConCaptureExe
    if (-not $probe) { $script:fail++; Write-Host "FAIL (no con_capture)" -ForegroundColor Red; return }

    $text = Invoke-CapturedUtf8 -Exe $probe -Arguments @("-hold", $built.Exe)
    $flat = ($text -split "`n" | Where-Object { $_ -like "LINE|*" }) -join "`n"
    $missing = @($Needles | Where-Object { -not $flat.Contains($_) })
    if ($missing.Count -eq 0) {
        $script:pass++
        Write-Host "PASS (console renders every language, code page independent)" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        Write-Host ("        missing: " + ($missing -join ' / ')) -ForegroundColor DarkGray
        if ($Verbose) { Write-Host $flat }
    }
}

# 字节子序列查找（落盘断言用：先按指定编码把期望文本编成字节，再到文件里找）
function Test-ByteSeq {
    param([byte[]]$Hay, [byte[]]$Needle)
    if ($Needle.Length -eq 0 -or $Hay.Length -lt $Needle.Length) { return $false }
    for ($i = 0; $i -le $Hay.Length - $Needle.Length; $i++) {
        $ok = $true
        for ($j = 0; $j -lt $Needle.Length; $j++) {
            if ($Hay[$i + $j] -ne $Needle[$j]) { $ok = $false; break }
        }
        if ($ok) { return $true }
    }
    return $false
}

function Test-CnRedirectOutput {
    param(
        [string]$Name,
        [string]$Source,
        [int]$ConsoleCp,             # 先 chcp 到这个代码页, 再让程序把 stdout 落文件
        [string]$Arch = ""
    )
    # vbp 分片: 本片不跑这例 (闸门在 run_tests.ps1 里定义)。
    if (-not (Enter-VbpShard)) { return }
    $script:total++
    Write-Host -NoNewline "  [ENC] $Name ... "

    $built = Build-CnSample -Source $Source -Arch $Arch
    if (-not $built.Exe) {
        $script:fail++
        Write-Host "FAIL (compile)" -ForegroundColor Red
        if ($Verbose) { Write-Host $built.Log }
        return
    }
    $file = Join-Path $OutDir "enc_redirect_$ConsoleCp.txt"
    Remove-Item $file -ErrorAction SilentlyContinue
    # cmd 层的重定向: 程序看到的是**文件句柄**, 走“控制台代码页”那条字节路
    cmd /c "chcp $ConsoleCp >nul & `"$($built.Exe)`" > `"$file`""
    if (-not (Test-Path $file)) { $script:fail++; Write-Host "FAIL (no output file)" -ForegroundColor Red; return }
    $bytes = [System.IO.File]::ReadAllBytes($file)

    # 期望: 该代码页装得下的行按它落字节; 装不下的行**整行**退回 UTF-8
    $expect = @(
        @{ Text = "中文";        Enc = $ConsoleCp },
        @{ Text = "日本語 テスト"; Enc = $ConsoleCp },
        @{ Text = "한국어 테스트"; Enc = 65001 }
    )
    $detail = @()
    foreach ($e in $expect) {
        $needle = [System.Text.Encoding]::GetEncoding([int]$e.Enc).GetBytes([string]$e.Text)
        if (-not (Test-ByteSeq -Hay $bytes -Needle $needle)) {
            $detail += ("missing '" + $e.Text + "' as CP" + $e.Enc)
        }
    }
    if ($detail.Count -eq 0) {
        $script:pass++
        Write-Host "PASS (bytes follow the console code page, unrepresentable lines fall back to UTF-8)" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "FAIL" -ForegroundColor Red
        Write-Host ("        " + ($detail -join '; ') + "  file=$($bytes.Length)B") -ForegroundColor DarkGray
        if ($Verbose) { Write-Host ("        hex: " + (($bytes | ForEach-Object { "{0:X2}" -f $_ }) -join " ")) }
    }
}
