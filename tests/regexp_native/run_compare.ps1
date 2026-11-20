param(
    [string]$C3 = ".build\C3.exe",
    [ValidateSet('x86','x64')][string]$Arch = 'x86',
    [string]$OutDir = ""
)
# C30-10 正则引擎对照：cscript 跑真 VBScript.RegExp（基准） vs C3 跑内置实现，逐行 diff。
#
# 两侧用例清单必须一致：tests/regexp_native/ReCases.vbs（基准） 与 ReMain.bas（C3 侧）。
# 判定：两边输出行数一致且逐字相同 -> PASS；否则打印前 20 条差异。
# 退出码 0/1。

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent   # tests/
$proj = $PSScriptRoot
if (-not $OutDir) { $OutDir = Join-Path (Split-Path $root -Parent) (".build\re_native_" + $Arch) }
$work = Join-Path (Split-Path $root -Parent) ".build"

# 1) 基准：cscript 跑真组件
#    环境缺件（没有 cscript / 没装 VBScript 功能）→ exit 2 = SKIP，不许当 PASS 混过去。
if (-not (Get-Command cscript -ErrorAction SilentlyContinue)) {
    Write-Output "SKIP regexp_native 本机没有 cscript"; exit 2
}
try { $null = New-Object -ComObject VBScript.RegExp }
catch { Write-Output "SKIP regexp_native 本机没有 VBScript.RegExp（VBScript 功能未安装）"; exit 2 }
$vbs = Join-Path $proj 'ReCases.vbs'
$tmpVbs = Join-Path $work 're_cases_oracle.vbs'
Copy-Item -LiteralPath $vbs -Destination $tmpVbs -Force
$oracle = & cmd.exe /c "cscript //nologo `"$tmpVbs`"" 2>&1
$oracleLines = @($oracle | ForEach-Object { $_.ToString().TrimEnd() })
if ($LASTEXITCODE -ne 0) { Write-Output ("FAIL cscript 退出码 " + $LASTEXITCODE); exit 1 }

# 2) C3 侧
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$c3exe = if ([IO.Path]::IsPathRooted($C3)) { $C3 } else { (Resolve-Path (Join-Path (Split-Path $root -Parent) $C3)).Path }  # harness 传绝对路径也在
$vbp = Join-Path $proj 'ReApp.vbp'
$env:C3_PROJECT_DIR = (Split-Path $root -Parent)
$log = & cmd.exe /c "`"$c3exe`" `"$vbp`" --arch $Arch --output-dir `"$OutDir`" 2>&1"
$exe = Join-Path $OutDir 'ReApp.exe'
if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $exe)) {
    Write-Output "FAIL 编译失败（没有 $exe）"
    $log | Select-Object -Last 12 | ForEach-Object { "  " + $_ }
    exit 1
}
$txt = Join-Path $OutDir 're_c3.txt'
if (Test-Path -LiteralPath $txt) { Remove-Item -LiteralPath $txt -Force }
$p = Start-Process -FilePath $exe -WorkingDirectory $OutDir -PassThru
if (-not $p.WaitForExit(30000)) { try { $p.Kill() } catch {} ; Write-Output "FAIL 用例进程超时（30s）"; exit 1 }
if (-not (Test-Path -LiteralPath $txt)) { Write-Output "FAIL 没有产出 re_c3.txt（进程提前退出 code=" + $p.ExitCode + "）"; exit 1 }
$c3Lines = @(Get-Content -LiteralPath $txt -Encoding Default | ForEach-Object { $_.TrimEnd() })

# 3) diff
$bad = 0
$max = [Math]::Max($oracleLines.Count, $c3Lines.Count)
$report = New-Object System.Collections.Generic.List[string]
for ($i = 0; $i -lt $max; $i++) {
    $a = if ($i -lt $oracleLines.Count) { $oracleLines[$i] } else { "<缺行>" }
    $b = if ($i -lt $c3Lines.Count) { $c3Lines[$i] } else { "<缺行>" }
    if ($a -ne $b) {
        $bad++
        if ($report.Count -lt 20) {
            $report.Add(("第 {0} 行`n    ocx: {1}`n    c3 : {2}" -f ($i + 1), $a, $b))
        }
    }
}
if ($bad -eq 0) {
    Write-Output ("PASS regexp_native 用例={0} 行 两侧逐字一致（oracle=cscript VBScript.RegExp）" -f $oracleLines.Count)
    exit 0
}
Write-Output ("FAIL regexp_native 不一致 {0} 行 / 共 {1} 行" -f $bad, $max)
foreach ($r in $report) { Write-Output $r }
exit 1
