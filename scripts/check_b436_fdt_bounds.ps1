# check_b436_fdt_bounds.ps1 - FormatDateTime 的「窗外 / 越界档不许装没事」（账 #436 §B210 / 第一百零九刀）
#
# 为什么单开这一道: 这一家原先有**两条**「拿今天兜底」的路 ——
#   ① `dateSerial == 0.0` → GetLocalTime()          ⇒ **序列号 0 本来是合法日期 1899-12/30**
#   ② `VariantTimeToSystemTime` 失败 → GetLocalTime() ⇒ 窗外静默答今天
#   外加一条更隐蔽的: 转换**成功**但落在 GetDateFormatW 的下限(年份 1601)之外 ⇒ 交**空串 + err=0**。
#   三条的共同形状 = **「不知道就装没事」**：读数看着像答上了，其实答的是「今天」或空串。
#   这类病单测抓不全(要么没测到，要么测到了被当成"合法空结果"放过)，所以钉形状。
#
# 规则（改坏了会红，不是装饰）:
#   R1  `GetLocalTime` **不许**再出现在这一家的兜底路上（两条兜底整个删掉）
#   R2  `dateSerial == 0.0` 那条特判不许回来 —— 0 是合法日期(1899-12-30)，不是哨兵
#   R3  转换失败必须 **vb6_ErrRaiseNumber(5)**（运行期 err 5 = 无效的过程调用或参数）
#   R4  年份 **1601** 那道下限不许被撤掉/改小 —— GetDateFormatW 的文档下限，实测窗逐格二分
#       (尺 .build/k429e|429f|429g|429h.vbs: -109200→1601/1/6 合法、-109300→ERR5)
#   R5  namedFormat 的 `default:` **不许**回落到 general date —— 越界档运行期报 err 5
#       (语料 44 处全用 vb* 常量 ⇒ 承面 0，但"越界静默照常交卷"正是本账的病)
#
# 用法:  powershell -File scripts\check_b436_fdt_bounds.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$sysC = Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_system.c'
if (-not (Test-Path -LiteralPath $sysC)) { Write-Host ('FAIL missing ' + $sysC) -ForegroundColor Red; exit 1 }
$txt = [IO.File]::ReadAllText($sysC)
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
$body = Get-Body $code 'BSTR vb6_FormatDateTime('
if (-not $body) { Write-Host 'FAIL vb6_FormatDateTime body not found -> renamed/removed' -ForegroundColor Red; exit 1 }

# ---- R1: 兜底路上不许再有 GetLocalTime ----
if ($body -match 'GetLocalTime') {
    $bad += 'R1 GetLocalTime is back in this function -> the "can''t tell, so answer with today" fallback is back; ' +
             'serial 0 is a real date (1899-12-30) and an out-of-window serial must raise err 5'
}

# ---- R2: 0 那一格特判不许回来 ----
if ($body -match 'dateSerial\s*==\s*0\.0') {
    $bad += 'R2 the `dateSerial == 0.0` special case is back -> serial 0 is a legitimate OLE date (1899-12-30), ' +
             'not a "no value given" sentinel'
}

# ---- R3: 失败必须报错 ----
$raises = ([regex]::Matches($body, 'vb6_ErrRaiseNumber\(5\)')).Count
if ($raises -lt 2) {
    $bad += ('R3 only ' + $raises + ' vb6_ErrRaiseNumber(5) site(s), want >= 2 (conversion failure AND out-of-window). ' +
             'Both paths used to return quietly instead of raising err 5')
}

# ---- R4: 年份 1601 那道下限 ----
if ($body -notmatch 'wYear\s*<\s*1601') {
    $bad += 'R4 the year-1601 lower bound is gone -> VariantTimeToSystemTime happily decodes pre-1601 dates, but ' +
             'GetDateFormatW cannot format them, so that band returns an EMPTY STRING with err=0 (measured window: ' +
             '-109200 -> 1601/1/6 formats, -109300 -> err 5)'
}

# ---- R5: 越界 namedFormat 不许回落 ----
# default 分支里必须**报错**而不是调 general date 那一档
$defIdx = $body.IndexOf('default:')
if ($defIdx -lt 0) {
    $bad += 'R5 no default: arm for namedFormat -> an out-of-range value has no owner at all'
} else {
    $seg = $body.Substring($defIdx)
    $seg = $seg.Substring(0, [Math]::Min(220, $seg.Length))
    if ($seg -notmatch 'vb6_ErrRaiseNumber') {
        $bad += 'R5 the default: arm does not raise -> an out-of-range namedFormat silently falls through to ' +
                 'vb6_fmtLocaleGeneralDate (runtime reports err 5; the corpus uses only vb* constants, so this is silent)'
    }
    if ($seg -match 'vb6_fmtLocaleGeneralDate') {
        $bad += 'R5 the default: arm still calls vb6_fmtLocaleGeneralDate -> out-of-range namedFormat answers as a ' +
                 'general date instead of raising err 5'
    }
}

if ($bad.Count -eq 0) {
    Write-Host 'PASS b436_fdt_bounds: no today-fallback | zero-not-sentinel | raises 2 | year-1601 bound | nf-default raises | violations 0' -ForegroundColor Green
    exit 0
}
foreach ($b in $bad) { Write-Host ('FAIL ' + $b) -ForegroundColor Red }
exit 1