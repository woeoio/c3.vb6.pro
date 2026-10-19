# check_b440_serial_split_authority.ps1 - 「序列号拆成日期 + 日内秒」这一问只许一位主人 (账 #440 §B209 / 第一百零七刀)
#
# 为什么单开这一道: 这一问原先**两处各答一半**，两半的归一还不一样：
#   时间那三枚 (Hour/Minute/Second) 走 `vb6_SecondsOfDay` 的 `t<0 → t+=1.0`；
#   日期那三枚 (Year/Month/Day) 走裸 `(int32_t)date`  —— 一天末尾那一格两半**不一致**
#   (日内秒已归零、日期还没进)，epoch 之前零头还与运行期**反号**(运行期取绝对值)。
# 两条腿各自看着都对，合起来对同一枚瞬间给出两套拆法 —— 这类病靠单测抓不全，
#   所以这里钉「**一位主人**」这个形状。
#
# 规则（改坏了会红，不是装饰）:
#   R1  `vb6_split_serial` 必须存在，且是**唯一**做「零头 → 日内秒 + 日期进位」这件事的地方
#   R2  零头归一只许**一处**写 `t = -t`（取绝对值）；再抄第二处 ⇒ 红
#   R3  `t += 1.0` 那一格（旧的反号归一）不许回来 —— 它是 §B195 那笔账的病根
#   R4  日期进位 `day += 1` 只许在 vb6_split_serial 里出现**一次**；
#       Year/Month/Day 必须转问 vb6_serial_day_half，不得自己抄 `(int32_t)date`
#   R5  `+0.5` 那条进位口径（§B193）不许被顺手改掉 —— 它量过：+0.4/+0.5 不进位、+0.6/+0.99 进位
#
# 用法:  powershell -File scripts\check_b440_serial_split_authority.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$dateC = Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_date.c'
if (-not (Test-Path -LiteralPath $dateC)) { Write-Host ('FAIL missing ' + $dateC) -ForegroundColor Red; exit 1 }
$txt = [IO.File]::ReadAllText($dateC)
# 注释里引用的形状不是实现 —— 先剥注释再数，否则这句注释自己变成一个「站点」
$code = [regex]::Replace($txt, '/\*[\s\S]*?\*/', ' ')
$code = [regex]::Replace($code, '(?m)//.*$', ' ')
$bad = @()

# ---- R1 ----
if ($code -notmatch 'static\s+void\s+vb6_split_serial\s*\(') {
    $bad += 'R1 no vb6_split_serial -> the split question lost its single owner (regression to two half-answers)'
}

# ---- R2/R3: 零头归一两格 ----
$absWrites = ([regex]::Matches($code, 'if\s*\(\s*t\s*<\s*0\s*\)\s*t\s*=\s*-t\s*;')).Count
if ($absWrites -eq 0) {
    $bad += 'R2 no `t = -t` -> the pre-epoch fraction sign flip is back (§B195 / #421)'
} elseif ($absWrites -gt 1) {
    $bad += ('R2 the absolute-value normalization is written ' + $absWrites + ' times -> the split was copied again')
}
if ($code -match 'if\s*\(\s*t\s*<\s*0\s*\)\s*t\s*\+=\s*1\.0') {
    $bad += 'R3 `t += 1.0` is back -> that is the OLD normalization whose sign disagrees with the runtime (the #421 defect)'
}

# ---- R4: 日期进位只许一处 ----
$carry = ([regex]::Matches($code, 'day\s*\+=\s*1\s*;')).Count
if ($carry -ne 1) {
    $bad += ('R4 the date carry `day += 1` appears ' + $carry + ' times (want exactly 1) -> either the day half stopped ' +
             'carrying at end-of-day (§B194 / #420), or it is being hand-copied into a second split')
}
foreach ($fn in @('vb6_Year', 'vb6_Month', 'vb6_Day')) {
    if ($code -notmatch ('int32_t\s+' + $fn + '\s*\(')) {
        $bad += ('R4 ' + $fn + ' is missing from vb6rtl_date.c -> renamed/deleted; check its callers')
        continue
    }
    # 取这一枚的**函数头到下一个 `int32_t`(或 EOF)**: 体里含嵌套调用括号, 用括号配平数深度
    # (早先用 `\{[^{}]*\}` 只够单语句, 三枚 getter 一改形状就**静默不再匹配** = 死电)
    $start = $code.IndexOf('int32_t ' + $fn + '(')
    if ($start -lt 0) { continue }
    $open = $code.IndexOf('{', $start)
    if ($open -lt 0) {
        $bad += ('R4 ' + $fn + ' has no body -> check before deleting (codegen emits these by name)')
        continue
    }
    $depth = 0; $end = -1
    for ($k = $open; $k -lt $code.Length; $k++) {
        if ($code[$k] -eq '{') { $depth++ }
        elseif ($code[$k] -eq '}') { $depth--; if ($depth -eq 0) { $end = $k; break } }
    }
    if ($end -lt 0) {
        $bad += ('R4 ' + $fn + ' body does not close -> unbalanced braces in vb6rtl_date.c')
        continue
    }
    $body = $code.Substring($open, $end - $open + 1)
    if ($body -notmatch 'vb6_serial_day_half') {
        $bad += ('R4 ' + $fn + ' does not ask vb6_serial_day_half -> the date half is answering for itself again, ' +
                 'which is how the time half and the date half drifted apart in the first place')
    }
}

# ---- R5: +0.5 那条口径 ----
if ($code -notmatch 't\s*\*\s*86400\.0\s*\+\s*0\.5') {
    $bad += 'R5 the `+0.5` rounding is gone -> §B193''s measured rule (+0.4/+0.5 do NOT carry, +0.6/+0.99 do) ' +
            'was measured, not guessed; changing it silently moves every sub-second cell'
}

if ($bad.Count -eq 0) {
    Write-Host 'PASS b440_serial_split_authority: owner 1 | abs-norm 1 | old-norm 0 | day-carry 1 | round-rule 1 | violations 0' -ForegroundColor Green
    exit 0
}
foreach ($b in $bad) { Write-Host ('FAIL ' + $b) -ForegroundColor Red }
exit 1