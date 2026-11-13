# diff_c3error.ps1 - PD 整编错误面**逐码**比对。
#
# 这条判据的唯一权威口径。PD 整编（623 个项目 TU + 65 个 RTL .c）在 CI 上必然失败
# （PD 的源码本身带着真实错误，GA 的目的是"别把错误面改大了/改小了"，不是"编过"），
# 所以每次提速刀都必须证明：**错误/警告的码种与逐码计数**在新旧两次之间完全相等。
#
# 为什么不能用"字节数相同"当判据（§56.55 实测结论）：
#   c3-error.log 的字节数**不可比**。cl 的 /MP 调度顺序一变，同一批 'compiling source file'
#   行就换位置；再加上 session 目录名长度不同（16 vs 13 字符），字节数就会差上百 KB，
#   而内容逐行看是等价的。所以判据只认**逐码计数**，字节数只作参考信息打印。
#
# 口径只设一处：下面 Get-CodeCounts 的正则 `(fatal error|error|warning) C####:`。
# 别的脚本、别的台账要判"错误面变了没有"，都必须走这里，不要另起一份正则。
#
# 用法：
#   pwsh -NoProfile -File scripts\diff_c3error.ps1 <oldRun> <newRun>
#   pwsh -NoProfile -File scripts\diff_c3error.ps1 -OldLog a.log -NewLog b.log
#
# 退出码：
#   0 = 逐码相等（差异码种数 = 0）—— 判据通过
#   1 = 有差异码种 —— 判据不通过
#   2 = 日志找不到 / 参数不对 —— 没判（别把它当成"通过"）
#
# 注意：本脚本含中文，必须用 pwsh（PowerShell 7）跑；powershell 5.1 会按 GBK
# 误读 UTF-8 无 BOM 的源码，引号配对会错乱。
param(
    [Parameter(Position = 0)][string]$OldRun = "",
    [Parameter(Position = 1)][string]$NewRun = "",
    [string]$Root = "",
    [string]$ArtifactRoot = "",
    [string]$OldLog = "",
    [string]$NewLog = "",
    [switch]$Quiet
)
$ErrorActionPreference = 'Continue'
if (-not $Root) { $Root = (Split-Path -Parent $PSScriptRoot) }
$Root = (Resolve-Path $Root).Path
if (-not $ArtifactRoot) { $ArtifactRoot = Join-Path $Root ".temp" }

function Find-Log([string]$run) {
    if (-not $run) { return $null }
    $d = Join-Path $ArtifactRoot ("art_" + $run)
    if (-not (Test-Path $d)) { return $null }
    $cands = @()
    foreach ($p in @(
            (Join-Path $d 'pd_full\c3-error.log'),
            (Join-Path $d 'c3-error.log'))) {
        if (Test-Path $p) { $cands += $p }
    }
    if ($cands.Count -eq 0) {
        $hit = Get-ChildItem -Path $d -Recurse -Filter 'c3-error.log' -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($hit) { $cands += $hit.FullName }
    }
    if ($cands.Count -eq 0) { return $null }
    return $cands[0]
}

if (-not $OldLog) { $OldLog = Find-Log $OldRun }
if (-not $NewLog) { $NewLog = Find-Log $NewRun }
if (-not $OldLog -or -not $NewLog) {
    Write-Host ("NO LOG: old=" + $OldRun + " new=" + $NewRun +
                " (artifactRoot=" + $ArtifactRoot + ")")
    Write-Host "  ⇒ 没判，不是判过。exit 2。"
    exit 2
}

# ★ 口径唯一出处：码种 = (fatal error|error|warning) + C####。
function Get-CodeCounts([string]$p) {
    $m = @{}
    $sr = New-Object IO.StreamReader($p)
    try {
        while (($line = $sr.ReadLine()) -ne $null) {
            if ($line -match '(?<k>fatal error|error|warning)\s+(?<c>C\d{4})\s*:') {
                $key = $Matches['k'] + ' ' + $Matches['c']
                if (-not $m.ContainsKey($key)) { $m[$key] = 0 }
                $m[$key]++
            }
        }
    } finally { $sr.Close() }
    return $m
}

function Get-LineStats([string]$p) {
    $total = 0; $compiling = 0
    $sr = New-Object IO.StreamReader($p)
    try {
        while (($line = $sr.ReadLine()) -ne $null) {
            $total++
            # 实测形态：(compiling source file <path>)，**不带引号**（§56.55 工件核对过）。
            if ($line -like '*compiling source file*') { $compiling++ }
        }
    } finally { $sr.Close() }
    return @{ Total = $total; Compiling = $compiling }
}

$o = Get-CodeCounts $OldLog
$n = Get-CodeCounts $NewLog
$os = Get-LineStats $OldLog
$ns = Get-LineStats $NewLog

Write-Host ("旧 {0}" -f $OldLog)
Write-Host ("新 {0}" -f $NewLog)
Write-Host ("字节（旧 {0} / 新 {1}）—— ★ 仅供参考，不可当判据：cl 的 /MP 调度顺序与 session 目录名长度都会改它" -f `
            (Get-Item $OldLog).Length, (Get-Item $NewLog).Length)

Write-Host ''
Write-Host '=== 错误/警告码对照（判据：逐码计数） ==='
$keys = @(@($o.Keys) + @($n.Keys)) | Sort-Object -Unique
$diff = 0
foreach ($k in $keys) {
    $ov = if ($o.ContainsKey($k)) { $o[$k] } else { 0 }
    $nv = if ($n.ContainsKey($k)) { $n[$k] } else { 0 }
    $mark = if ($ov -eq $nv) { '  =' } else { '  *' }
    if ($ov -ne $nv) { $diff++ }
    Write-Host ("{0} {1,-16} 旧 {2,6}  新 {3,6}  差 {4}" -f $mark, $k, $ov, $nv, ($nv - $ov))
}
Write-Host ''
Write-Host ("码种总数 = " + $keys.Count)
Write-Host ("差异码种数 = " + $diff)
Write-Host ("参考（不参与判定）: 总行数 旧 {0} / 新 {1}；'compiling source file' 行 旧 {2} / 新 {3}" -f `
            $os.Total, $ns.Total, $os.Compiling, $ns.Compiling)

if ($diff -eq 0) {
    Write-Host 'VERDICT=PASS（逐码相等）'
    exit 0
}
Write-Host 'VERDICT=FAIL（错误面变了）'
exit 1
