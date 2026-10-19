# check_b424_datedpart_params.ps1 - 那两条形参的两档口径是**用户拍的**，不许被实现悄悄换掉
#   (账 #424 §B198 / 第一百零九刀)
#
# 为什么单开这一道: 账上明写「两种口径都合法，但**不能靠实现默默选一个**」。本仓的实现一度
#   两边都悄悄选了一个 —— `firstWeekOfYear >= 4` 静默按 1 档答（运行期报 err 5），
#   `firstDayOfWeek = 0` 静默 clamp 到周日（运行期交系统）。用户拍板的结果:
#   ① firstWeekOfYear >= 4 ⇒ **报 err 5**（对齐运行期，与 #436 越界 namedFormat 同一口径）
#   ② firstDayOfWeek = 0 ⇒ **继续 clamp 到 vbSunday**，作为一条**显式的已知偏离**
#      （跟随系统跨机器不可复现；clamp 换来读数稳定 = emit-manifest 逐字节对账的前提）
#   这一道就是把那两条决定**钉成形状**，免得下一个改这附近的人以为「clamp / 静默」只是随手写的。
#
# 规则（改坏了会红，不是装饰）:
#   R1  firstWeekOfYear >= 4 必须 vb6_ErrRaiseNumber(5) 并返回（不许静默当 1 档）
#   R2  合法档 0..3 的归一**不许动**：0 与 1 同档、2 / 3 另算（实测口径，改了会挪动读数）
#   R3  firstDayOfWeek 的 clamp 到 vbSunday **不许**被换成"跟随系统"（GetSystemDefaultLCID / 读注册表
#       那一族）—— 那会让同一份源码在不同机器上读数不同，破坏整套可复现对账
#   R4  两处 clamp/报错的位置**只在** vb6_WeekOfYear 一处（读者 = DatePart 的 ww 臂与 DateDiff 的 ww 臂）
#
# 用法:  powershell -File scripts\check_b424_datedpart_params.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$dateC = Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_date.c'
if (-not (Test-Path -LiteralPath $dateC)) { Write-Host ('FAIL missing ' + $dateC) -ForegroundColor Red; exit 1 }
$txt = [IO.File]::ReadAllText($dateC)
$code = [regex]::Replace($txt, '/\*[\s\S]*?\*/', ' ')
$code = [regex]::Replace($code, '(?m)//.*$', ' ')
$bad = @()

function Get-Body([string]$src, [string]$sig) {
    $start = $src.IndexOf($sig)
    if ($start -lt 0) { return '' }
    $open = $src.IndexOf('{', $start)
    if ($open -lt 0) { return '' }
    $depth = 0; $end = -1
    for ($k = $open; $k -lt $src.Length; $k++) {
        if ($src[$k] -eq '{') { $depth++ }
        elseif ($src[$k] -eq '}') { $depth--; if ($depth -eq 0) { $end = $k; break } }
    }
    if ($end -lt 0) { return '' }
    return $src.Substring($open, $end - $open + 1)
}
$body = Get-Body $code 'int32_t vb6_WeekOfYear('
if (-not $body) { Write-Host 'FAIL vb6_WeekOfYear body not found -> renamed/removed' -ForegroundColor Red; exit 1 }

# ---- R1: 越界必须报 err 5 ----
if ($body -notmatch 'fwoy\s*>=\s*4') {
    $bad += 'R1 the `fwoy >= 4` guard is gone -> out-of-range firstWeekOfYear silently answers again (the runtime ' +
             'raises err 5; measured fwy=4/5 both throw)'
} elseif ($body -notmatch 'vb6_ErrRaiseNumber\(5\)') {
    $bad += 'R1 the `fwoy >= 4` guard does not raise err 5 -> the guard is there but answers instead of reporting, ' +
             'which is the same silent-wrong-answer shape this knife closed'
}

# ---- R2: 合法档的归一不许动 ----
if ($body -notmatch 'fwoy\s*!=\s*2\s*&&\s*fwoy\s*!=\s*3\s*\)\s*fwoy\s*=\s*1') {
    $bad += 'R2 the 0..3 normalisation changed -> measured policy is "0 and 1 are the same tier; 2 and 3 are the other", ' +
             'and editing it moves readings on the fixtures that pin the ww tiers'
}

# ---- R3: fdow 必须 clamp 到周日，不许跟随系统 ----
if ($body -notmatch 'fdow\s*<\s*1\s*\|\|\s*fdow\s*>\s*7\s*\)\s*fdow\s*=\s*1') {
    $bad += 'R3 firstDayOfWeek no longer clamps to vbSunday(1) -> the deliberate, documented deviation was ' +
             'replaced (possibly by "follow the system"); that makes readings machine-dependent and breaks the ' +
             'byte-for-byte shape oracle this repo relies on'
}
foreach ($banned in @('GetSystemDefaultLCID', 'GetUserDefaultLCID', 'RegGetValueA', 'GetLocaleInfoW')) {
    if ($body -match [regex]::Escape($banned)) {
        $bad += ('R3 vb6_WeekOfYear now reaches for ' + $banned + ' -> that is the "follow the system" behaviour this ' +
                 'knife explicitly rejected; the clamp must stay deterministic')
    }
}

# ---- R4: 这两问只住一处 ----
$clampSites = ([regex]::Matches($code, 'fdow\s*<\s*1\s*\|\|\s*fdow\s*>\s*7\s*\)\s*fdow\s*=\s*1')).Count
if ($clampSites -gt 1) {
    $bad += ('R4 the firstDayOfWeek clamp appears ' + $clampSites + ' times -> DatePart("ww") and DateDiff("ww") must ' +
             'share ONE owner (vb6_WeekOfYear); two copies is how they drift apart')
}

if ($bad.Count -eq 0) {
    Write-Host 'PASS b424_datedpart_params: fwoy>=4 raises err5 | tiers 0..3 unchanged | fdow clamps to Sunday (no system lookup) | clamp in one place | violations 0' -ForegroundColor Green
    exit 0
}
foreach ($b in $bad) { Write-Host ('FAIL ' + $b) -ForegroundColor Red }
exit 1