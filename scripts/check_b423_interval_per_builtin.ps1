# check_b423_interval_per_builtin.ps1 - 「w / ww 这一格按 builtin 分岔」不许被共用枚举抹平 (账 #423 §B197 / 第一百零八刀)
#
# 为什么单开这一道: `vb6_parse_interval` 是 **DateAdd / DateDiff / DatePart 三家共用**的一枚枚举，
#   而 `w` 与 `ww` 的含义**按 builtin 分岔**:
#     DatePart("w",  …) = 星期序号(与周首无关)     DateDiff("w",  …) = 整段七天(与周首无关)
#     DatePart("ww", …) = 一年里第几周             DateDiff("ww", …) = 半开区间里落在第 f 天那一格的条数
#   两边对 `w` 的答案**不同**(2 vs 整段天数)，而共用一枚枚举只有一个答案可发 ⇒ 改它就弄坏
#   另一家。本刀的选择是**不动那枚枚举**、DateDiff 自己认这一格(口径三件由 §B197 量死:
#   162400 格对 ww 公式零失配、81200 格确认 w 与周首无关)。
#
# 规则（改坏了会红，不是装饰）:
#   R1  `vb6_parse_interval` **不许**被改成 DateDiff 口径 —— 它三家共用，一改就是拆另外两家
#   R2  DateDiff 的 w / ww 两问必须**在 vb6_DateDiff 体内**各答一次，不许转问共用枚举那两格
#   R3  「w」那条公式**不许**含 firstDayOfWeek（运行期与周首无关，81200 格零失配）
#   R4  「ww」那条公式**必须**含 firstDayOfWeek（f 参与，162400 格零失配）
#   R5  「w」必须是 **trunc** 向零（(b-a)/7 直接整数除），不许改成 floor —— 反向区间会差一
#
# 用法:  powershell -File scripts\check_b423_interval_per_builtin.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$dateC = Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_date.c'
if (-not (Test-Path -LiteralPath $dateC)) { Write-Host ('FAIL missing ' + $dateC) -ForegroundColor Red; exit 1 }
$txt = [IO.File]::ReadAllText($dateC)
$code = [regex]::Replace($txt, '/\*[\s\S]*?\*/', ' ')
$code = [regex]::Replace($code, '(?m)//.*$', ' ')
$bad = @()

# 取 vb6_parse_interval 与 vb6_DateDiff 两段（括号配平，早先用 \{[^{}]*\} 只够单语句 = 死电）
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
$parseBody = Get-Body $code 'static vb6_date_interval vb6_parse_interval'
$diffBody  = Get-Body $code 'int64_t vb6_DateDiff('
if (-not $parseBody) { $bad += 'R1 vb6_parse_interval body not found -> renamed/removed; the shared enum just lost its one owner' }
if (-not $diffBody)  { $bad += 'R2 vb6_DateDiff body not found -> renamed/removed; codegen emits it by name' }

# ---- R1: 共用枚举不许被 DateDiff 口径污染 ----
# 它仍按「DatePart 口径」把单枚 w 答成 vb6_di_weekday、双 w 答成 vb6_di_week —— 这是**对的**，
# 因为另外两家( DatePart / DateAdd )靠它。若有人把它单 w 改成 vb6_di_week，
# DatePart("w") 就会从星期序号答成一年里第几周 —— 那是拿一家的对换另一家的错。
if ($parseBody -match 'case\s+L''[wW]''') {
    $seg = $parseBody.Substring($parseBody.IndexOf("case L'w'"))
    $seg = $seg.Substring(0, [Math]::Min(260, $seg.Length))
    if ($seg -notmatch 'vb6_di_weekday') {
        $bad += 'R1 vb6_parse_interval no longer answers single "w" as vb6_di_weekday -> that is DatePart''s weekday ' +
                 'answer; flipping it to vb6_di_week would make DatePart("w") report week-of-year instead'
    }
    if ($seg -notmatch 'vb6_di_week\b') {
        $bad += 'R1 vb6_parse_interval no longer answers double "ww" as vb6_di_week -> DatePart("ww") lost its answer'
    }
}

# ---- R2: DateDiff 自己认这两格 ----
if ($diffBody -notmatch 'ivIsW423') {
    $bad += 'R2 vb6_DateDiff no longer recognises the single-"w" interval -> that arm falls back to the shared enum, ' +
            'which answers DatePart''s question and yields the DAY COUNT (1/5->1/12 would answer 7, not 1)'
}
if ($diffBody -notmatch 'ivIsWw423') {
    $bad += 'R2 vb6_DateDiff no longer recognises the double-"ww" interval -> that arm has no owner'
}

# ---- R3/R4: 周首参不参与这一格，是两格各自的判据 ----
$wLine = ''; $wwLine = ''
if ($diffBody -match '(?s)if \(ivIsW423\)\s*\{(?<b>[^{}]*)\}') { $wLine = $Matches['b'] }
if ($diffBody -match '(?s)if \(ivIsWw423\)\s*\{(?<b>[^{}]*)\}') { $wwLine = $Matches['b'] }
if (-not $wLine) {
    $bad += 'R3 the ivIsW423 branch is gone -> cannot check that "w" ignores firstDayOfWeek'
} elseif ($wLine -match 'firstDayOfWeek') {
    $bad += 'R3 the "w" formula now uses firstDayOfWeek -> the runtime answer does NOT depend on the week-start ' +
             '(81200 cells, zero mismatch); adding it there makes the reading vary across machines'
}
if (-not $wwLine) {
    $bad += 'R4 the ivIsWw423 branch is gone -> cannot check that "ww" uses firstDayOfWeek'
} elseif ($wwLine -notmatch 'firstDayOfWeek') {
    $bad += 'R4 the "ww" formula no longer uses firstDayOfWeek -> the half-open-interval count is defined by where ' +
             'the f-th weekday falls; dropping f silently changes every answer'
}

# ---- R5: "w" 是 trunc 向零，不是 floor ----
if ($wLine -and $wLine -match 'floor') {
    $bad += 'R5 the "w" formula uses floor -> the runtime truncates toward zero, so a reversed interval ' +
             '(b < a) differs by one week between the two'
}

if ($bad.Count -eq 0) {
    Write-Host 'PASS b423_interval_per_builtin: shared-enum intact | DateDiff owns w+ww | w ignores f | ww uses f | w is trunc | violations 0' -ForegroundColor Green
    exit 0
}
foreach ($b in $bad) { Write-Host ('FAIL ' + $b) -ForegroundColor Red }
exit 1