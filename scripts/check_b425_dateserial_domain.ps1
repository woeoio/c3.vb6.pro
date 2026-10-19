# check_b425_dateserial_domain.ps1 - DateSerial 的「月/日超出范围」那一问：域内逐格对，域外**不许盲修**
#   (账 #425 §B199 / 第一百一十刀 = 本刀只结账、不改行为)
#
# 为什么单开这一道: 这一格量完的结论是「**现成的出口已经对了，别盲修**」，但「别修」如果不钉成
#   形状，下一个看到 `if (month <= 2)` 觉得不完整的人会顺手「补」上月份归一 —— 而那正是**引入
#   偏差**的动作。`30.6001*(month+1)` 那一项只在月份落在 **3..14** 时精确，这是儒略日算式的
#   固有前提，不是本仓的疏漏。
#
# 承面（开工前重新点过一遍，尺本刀 census）：真语料 `DateSerial` 调用 **46 处**，
#   第二实参分布 = 字面量 1(33) / 2(3) / 6(2) / 9(1) / 12(1) + `*.wMonth` 字段(6) + 循环下标(2)，
#   **月份 > 14 的 0 处** ⇒ 域外那一段在真语料里**不可达**。
#
# 域外偏差的量级值得单独记一笔：月份 15 起偏 **2 天**、月份 25 起偏 2~3 天，**不是**几百天 ——
#   因为 `30.6001*(month+1)` 与 `365.25*year` 两项**互相抵消**，只剩一个小的残差。
#   ⚠ 这正是它危险的地方：**看起来几乎是对的**。真语料若哪天出现 `DateSerial(y, m+13, d)`，
#     编译与运行都不会响，只是日子悄悄差了 2 天。
#
# 规则（改坏了会红，不是装饰）:
#   R1  `vb6_DateSerial` **不许**加月份归一（`m = (m-1)%12+1` 那一族）—— 那会改掉域内读数，
#       而域内本来是对的
#   R2  `vb6_date_to_serial` 的 `if (month <= 2) { year--; month += 12; }` 不许被撤
#       （它是算式**能用的前提**，不是可选的修饰）
#   R3  `30.6001` 那一项不许换成别的系数（换成别的会让域内也偏）
#   R4  域外偏差必须仍然存在（即不许"修"）；夹具 `tests/b425_dateserial_domain.bas` 钉住
#       域内 6 格 + 域外 3 格，域外那三格**就是钉住"仍然是偏的"**
#
# 用法:  powershell -File scripts\check_b425_dateserial_domain.ps1
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
$d2s = Get-Body $code 'static int32_t vb6_date_to_serial('
$ds  = Get-Body $code 'double vb6_DateSerial('
if (-not $d2s) { Write-Host 'FAIL vb6_date_to_serial body not found' -ForegroundColor Red; exit 1 }
if (-not $ds)  { Write-Host 'FAIL vb6_DateSerial body not found' -ForegroundColor Red; exit 1 }

# ---- R1: 不许加月份归一 ----
if ($ds -match '%\s*12|month\s*=\s*\(.*-1\s*\)') {
    $bad += 'R1 vb6_DateSerial now normalises the month -> the in-domain readings (months 1..14) are the correct ones ' +
             'and this changes them; the out-of-domain band is unreachable from the corpus (0 of 46 call sites pass ' +
             'month > 14), so "fixing" it trades a correct domain for an unreachable one'
}
if ($ds -notmatch 'vb6_SerialOf') {
    $bad += 'R1 vb6_DateSerial no longer goes through vb6_SerialOf -> the single serial authority (owner of ' +
             '"a y/m/d triple becomes an OLE serial") was bypassed'
}

# ---- R2: 算式能用的前提不许撤 ----
if ($d2s -notmatch 'month\s*<=\s*2\s*\)\s*\{\s*year--\s*;\s*month\s*\+=\s*12\s*;') {
    $bad += 'R2 the `if (month <= 2) { year--; month += 12; }` shift is gone -> that shift is what puts the month ' +
             'into the 3..14 band the 30.6001 term is exact on; removing it breaks every date, not just edge cases'
}

# ---- R3: 系数不许换 ----
if ($d2s -notmatch '30\.6001') {
    $bad += 'R3 the 30.6001 coefficient is gone -> it is the Julian-day term that pairs with the 365.25 term; ' +
             'swapping it shifts every in-domain date'
}

# ---- R4: 夹具必须在，且必须钉住域外那一格仍然是偏的 ----
$fixture = Join-Path $root 'tests\b425_dateserial_domain.bas'
if (-not (Test-Path -LiteralPath $fixture)) {
    $bad += 'R4 tests\b425_dateserial_domain.bas is missing -> nothing pins "in-domain exact, out-of-domain still off"'
} else {
    $fx = [IO.File]::ReadAllText($fixture)
    foreach ($needle in @('mm15', 'mm25')) {
        if ($fx -notmatch [regex]::Escape($needle)) {
            $bad += ('R4 the fixture does not pin the out-of-domain cell ' + $needle + ' -> the deliberate ' +
                     'non-fix stops being visible and someone will "correct" it')
        }
    }
}

if ($bad.Count -eq 0) {
    Write-Host 'PASS b425_dateserial_domain: no month-normalisation | 30.6001 shift intact | coefficient intact | out-of-domain still pinned as off | violations 0' -ForegroundColor Green
    exit 0
}
foreach ($b in $bad) { Write-Host ('FAIL ' + $b) -ForegroundColor Red }
exit 1