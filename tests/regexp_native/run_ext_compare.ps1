param(
    [string]$C3 = ".build\C3.exe",
    [ValidateSet('x86','x64')][string]$Arch = 'x86',
    [string]$OutDir = ""
)
# C30-10 扩展语法的对照：**Python re** 当判据（VB 的 VBScript.RegExp 不认这些写法）。
#   基准 = python ext_oracle.py
#   被测 = C3 编译 ReApp.vbp 后跑出来的 re_c3_ext.txt（ExtCases）
# 判定：两边行数一致且逐字相同 -> PASS；否则打印差异。退出码 0/1。
#
# 与 run_compare.ps1 的分工：那一份管"与真 VBScript.RegExp 逐字相同"（VB 兼容面），
# 这一份管"超集面"，判据换成 PCRE2/Python re —— 两条判据不许互相借。

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$proj = $PSScriptRoot
if (-not $OutDir) { $OutDir = Join-Path (Split-Path $root -Parent) (".build\re_native_" + $Arch) }

# 1) 基准：Python re
$py = (Get-Command python -ErrorAction SilentlyContinue)
if (-not $py) { Write-Output "FAIL 找不到 python（扩展语法判据依赖 Python re）"; exit 1 }
$oracle = & python (Join-Path $proj 'ext_oracle.py') 2>&1
$oracleLines = @($oracle | ForEach-Object { $_.ToString().TrimEnd() })
if ($LASTEXITCODE -ne 0) {
    Write-Output "FAIL ext_oracle.py 退出码 $LASTEXITCODE"
    $oracleLines | Select-Object -Last 10 | ForEach-Object { "  " + $_ }
    exit 1
}

# 2) C3 侧
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$c3exe = (Resolve-Path (Join-Path (Split-Path $root -Parent) $C3)).Path
$vbp = Join-Path $proj 'ReApp.vbp'
$env:C3_PROJECT_DIR = (Split-Path $root -Parent)
$log = & cmd.exe /c "`"$c3exe`" `"$vbp`" --arch $Arch --output-dir `"$OutDir`" 2>&1"
$exe = Join-Path $OutDir 'ReApp.exe'
if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $exe)) {
    Write-Output "FAIL 编译失败（没有 $exe）"
    $log | Select-Object -Last 12 | ForEach-Object { "  " + $_ }
    exit 1
}
$txt = Join-Path $OutDir 're_c3_ext.txt'
if (Test-Path -LiteralPath $txt) { Remove-Item -LiteralPath $txt -Force }
$p = Start-Process -FilePath $exe -WorkingDirectory $OutDir -PassThru
if (-not $p.WaitForExit(30000)) { try { $p.Kill() } catch {}; Write-Output "FAIL ReApp.exe 超时"; exit 1 }
if (-not (Test-Path -LiteralPath $txt)) { Write-Output "FAIL 没有产出 re_c3_ext.txt"; exit 1 }
$c3Lines = @(Get-Content -LiteralPath $txt -Encoding UTF8 | ForEach-Object { $_.TrimEnd() })

# 3) 逐行比
$n = [Math]::Max($oracleLines.Count, $c3Lines.Count)
$bad = 0
$report = New-Object System.Collections.Generic.List[string]
for ($i = 0; $i -lt $n; $i++) {
    $a = if ($i -lt $oracleLines.Count) { $oracleLines[$i] } else { "<缺行>" }
    $b = if ($i -lt $c3Lines.Count) { $c3Lines[$i] } else { "<缺行>" }
    if ($a -ne $b) {
        $bad++
        if ($bad -le 20) { $report.Add("第 " + ($i + 1) + " 行"); $report.Add("    py: " + $a); $report.Add("    c3: " + $b) }
    }
}
if ($bad -eq 0) {
    Write-Output ("PASS regexp_ext 用例=" + $oracleLines.Count + " 行 两侧逐字一致（oracle=python re）")
    exit 0
}
Write-Output ("FAIL regexp_ext 不一致 " + $bad + " 行 / 共 " + $n + " 行")
$report | ForEach-Object { Write-Output $_ }
exit 1
