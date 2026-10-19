# check_b442_dateadd_interval.ps1 - DateAdd 的 w / ww 那一格是**对的**，不许照着 DateDiff 去"修"
#   (账 #442 §B211，第一百一十一刀 = 本刀只钉行为、不改行为)
#
# 为什么单开这一道: 账 #423 (§B197) 给 `DateDiff` 的 w / ww 两条臂各补了一次「按 builtin 分岔」的
#   问句。那一格修完以后，隔壁 `DateAdd` 的同款分岔看起来就**可疑**了 —— 共享枚举把单枚 `w`
#   答成 `vb6_di_weekday`、双 `w` 答成 `vb6_di_week`，而 `DateAdd` 的 weekday 臂是 `datePart += n`
#   （加**天**）、week 臂是 `datePart += n * 7`（加**周**）⇒ 看上去像「DateAdd("w",1,·) 只加了
#   一天，是 bug」。
#
#   ⚠ **量完证明它是对的**：运行期口径就是
#       DateAdd("w",  1, 2026/1/5) = 2026/1/6   (w  = 加**天**)
#       DateAdd("ww", 1, 2026/1/5) = 2026/1/12  (ww = 加**周**)
#   本机逐格量了 **12 格**（w/ww 各含负数、0、三周、跨月、闰月，以及 d/m/y/h 四档不许动的对照），
#   **零失配**。所以那一格**一个字都不改**。
#
# 这一点同时**反证了 #423 的架构是对的**：那枚共享枚举对 **DateAdd 与 DatePart 两家**都恰好
#   是对的，只有 DateDiff 一家分岔 ⇒ #423 选择「不动枚举、让 DateDiff 自己认」正是唯一正解。
#   若当时「顺手把枚举改成 DateDiff 口径」，DateAdd 与 DatePart 会同时被弄坏。
#
# 规则（改坏了会红，不是装饰）:
#   R1  `vb6_parse_interval` 的 w / ww 分岔**不许**改（那是 DateAdd + DatePart 两家共用的正解）
#   R2  DateAdd 的 weekday 臂（`datePart += n`）与 week 臂（`datePart += n * 7`）都**不许**动
#   R3  夹具必须在，且必须同时钉住 w（加天）与 ww（加周）两档 —— 只钉一档等于没钉这格分岔
#   R4  d / m / y / h 四档不许动（它们与 w/ww 共用那张表，改 w/ww 时最容易顺手碰到它们）
#
# 用法:  powershell -File scripts\check_b442_dateadd_interval.ps1
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
$parse = Get-Body $code 'static vb6_date_interval vb6_parse_interval('
$add   = Get-Body $code 'double vb6_DateAdd('
if (-not $parse) { Write-Host 'FAIL vb6_parse_interval body not found' -ForegroundColor Red; exit 1 }
if (-not $add)   { Write-Host 'FAIL vb6_DateAdd body not found' -ForegroundColor Red; exit 1 }

# ---- R1: 共用枚举的 w/ww 分岔不许动 ----
$seg = $parse
$idx = $seg.IndexOf("case L'w'")
if ($idx -lt 0) {
    $bad += 'R1 the w/W case is gone from vb6_parse_interval -> single "w" and double "ww" must stay distinguishable'
} else {
    $seg = $seg.Substring($idx, [Math]::Min(260, $seg.Length - $idx))
    if ($seg -notmatch 'vb6_di_weekday') {
        $bad += 'R1 single "w" no longer maps to vb6_di_weekday -> that mapping is what DateAdd (adds days) and ' +
                 'DatePart (weekday number) both rely on; flipping it to DateDiff''s meaning breaks those two'
    }
    if ($seg -notmatch 'vb6_di_week\b') {
        $bad += 'R1 double "ww" no longer maps to vb6_di_week -> DateAdd("ww") adds weeks through this arm'
    }
}

# ---- R2: 两条臂的形状不许动 ----
if ($add -notmatch 'case\s+vb6_di_weekday:\s*datePart\s*\+=\s*n\s*;') {
    $bad += 'R2 the weekday arm no longer does `datePart += n` -> DateAdd("w") adds DAYS (measured: "w",1 -> +1 day)'
}
if ($add -notmatch 'case\s+vb6_di_week:\s*datePart\s*\+=\s*n\s*\*\s*7\s*;') {
    $bad += 'R2 the week arm no longer does `datePart += n * 7` -> DateAdd("ww") adds WEEKS (measured: "ww",1 -> +7 days)'
}

# ---- R3/R4: 夹具必须把两档一起钉住，外加四档对照 ----
$fixture = Join-Path $root 'tests\b442_dateadd_interval.bas'
if (-not (Test-Path -LiteralPath $fixture)) {
    $bad += 'R4 tests\b442_dateadd_interval.bas is missing -> nothing pins "w adds days, ww adds weeks"'
} else {
    $fx = [IO.File]::ReadAllText($fixture)
    foreach ($needle in @('w_p1', 'ww_p1', 'd_p1', 'm_p1', 'y_p1', 'h_p1')) {
        if ($fx -notmatch [regex]::Escape($needle)) {
            $bad += ('R3/R4 the fixture does not pin ' + $needle + ' -> pinning only one of w/ww leaves the split untested')
        }
    }
}

if ($bad.Count -eq 0) {
    Write-Host 'PASS b442_dateadd_interval: shared split intact | w adds days | ww adds weeks | fixture pins both + 4 controls | violations 0' -ForegroundColor Green
    exit 0
}
foreach ($b in $bad) { Write-Host ('FAIL ' + $b) -ForegroundColor Red }
exit 1