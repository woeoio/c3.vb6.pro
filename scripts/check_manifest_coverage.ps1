# check_manifest_coverage.ps1 - 账 §B99① 的结构哨兵 (第 42 道；不起 cl)
#
# 为什么要这一道: 门 #434 唯一红就是 `Emit manifest (shape oracle)` —— 57b1f64e 往 tests/ 加了
# 一份新夹具 (tests/test_vbnet_ext.bas) 却没把它的哈希登记进 emit-manifest.expected.txt。
# 这不算产品缺陷, 但它是**只有跑完 CI 那二十分钟才报得出来**的一类红: 比较器把「新增 1 行」判红，
# 而本机在提交前没有任何一句话会说「你少登记了」。这一道就是那句话。
#
# 口径只有一份: 语料枚举走 `emit_manifest.ps1 -ListInputs` 自己那一步 (Get-ChildItem + 同一套排除 +
# 同一个 Sort)。哨兵里绝不另写一遍 glob —— 那是给「谁算语料」开第二份权威 (#166/#231/#245 那一族)。
#
# 规则 (改坏了会红):
#   K1 枚举出的输入集合 == emit-manifest.expected.txt 里的路径集合 (两个方向的差集都要空)。
#   K2 期望清单必须 BOM+CRLF、每一行数据都长成 `sha256=… ascii256=… rc=… bytes=… <relpath>`
#      (手工编辑时最容易把行拼坏或把 BOM/行尾弄丢)。
#   K3 枚举本身不许退化：至少 200 份、且 .vbp 与 .bas **两类都有**。没有这一条，K1 会在
#      「两边都空」上假绿 (「恰好 N 处」那类钉死同理，见 memory judgment-from-owner-expectation-table)。
#
# 用法: powershell -File scripts/check_manifest_coverage.ps1   (PASS => exit 0)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$exp = Join-Path $root 'emit-manifest.expected.txt'
$list = Join-Path $PSScriptRoot 'emit_manifest.ps1'
$fails = @()

if (-not (Test-Path -LiteralPath $list)) { Write-Host 'FAIL missing scripts/emit_manifest.ps1'; exit 1 }
if (-not (Test-Path -LiteralPath $exp)) { Write-Host "FAIL missing $exp"; exit 1 }

# ---- 取枚举 (唯一口径) ----
$out = & powershell -NoProfile -ExecutionPolicy Bypass -File $list -ListInputs 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host ('FAIL emit_manifest -ListInputs rc=' + $LASTEXITCODE + ' :: ' + (($out | Select-Object -First 3) -join ' | '))
    exit 1
}
$inputs = @($out | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ -ne '' -and $_ -notmatch '^#' })

# ---- K3: 枚举退化检查 ----
$nVbp = @($inputs | Where-Object { $_ -like '*.vbp' }).Count
$nBas = @($inputs | Where-Object { $_ -like '*.bas' }).Count
if ($inputs.Count -lt 200) {
    $fails += ('K3 enumerated only ' + $inputs.Count + ' corpus inputs (expect >=200) - a degenerate ' +
               'enumeration would make the set comparison below vacuously green')
}
if ($nVbp -lt 1 -or $nBas -lt 1) {
    $fails += ('K3 enumeration lost a whole kind (.vbp=' + $nVbp + ' .bas=' + $nBas + ') - the manifest covers both')
}

# ---- 读期望清单 ----
$bytes = [System.IO.File]::ReadAllBytes($exp)
if ($bytes.Length -lt 3 -or $bytes[0] -ne 0xEF -or $bytes[1] -ne 0xBB -or $bytes[2] -ne 0xBF) {
    $fails += 'K2 emit-manifest.expected.txt lost its BOM (the rest of this file is read as UTF-8 by PS 5.1)'
}
$text = [System.Text.Encoding]::UTF8.GetString($bytes, 3, $bytes.Length - 3)
$nl = ([regex]::Matches($text, "`r`n")).Count
if ($nl -eq 0 -or $nl -ne ([regex]::Matches($text, "`n")).Count) {
    $fails += 'K2 emit-manifest.expected.txt is not pure CRLF'
}
$row = [regex]'^sha256=([0-9a-f]{64}) ascii256=([0-9a-f]{64}) rc=([0-9]+) bytes=([0-9]+) (.+)$'
$registered = @()
$bad = 0
foreach ($line in ($text -split "`r`n")) {
    $l = $line.Trim()
    if ($l -eq '' -or -not $l.StartsWith('sha256=')) { continue }
    $m = $row.Match($l)
    if (-not $m.Success) { $bad++; continue }
    $registered += $m.Groups[5].Value
}
if ($bad -gt 0) { $fails += ('K2 ' + $bad + ' data row(s) do not match "sha256=<64hex> ascii256=<64hex> rc=<n> bytes=<n> <relpath>"') }

# ---- K1: 两个方向的差集 ----
$miss = @($inputs | Where-Object { $registered -notcontains $_ })
$extra = @($registered | Where-Object { $inputs -notcontains $_ })
if ($miss.Count -gt 0) {
    $fails += ('K1 ' + $miss.Count + ' corpus input(s) are NOT in emit-manifest.expected.txt (the shape ' +
               'oracle red-gates these as "added rows"): ' + (($miss | Select-Object -First 6) -join ' , ') +
               ' - add the row(s) from a manifest produced by the current C3.exe')
}
if ($extra.Count -gt 0) {
    $fails += ('K1 emit-manifest.expected.txt registers ' + $extra.Count + ' input(s) that no longer exist: ' +
               (($extra | Select-Object -First 6) -join ' , '))
}

if ($fails.Count -gt 0) {
    $fails | ForEach-Object { Write-Host "FAIL $_" }
    exit 1
}
Write-Host ('OK check_manifest_coverage: K1..K3 (inputs=' + $inputs.Count + ' [.vbp=' + $nVbp +
             ' .bas=' + $nBas + '], registered=' + $registered.Count + ', sets equal)')
exit 0
