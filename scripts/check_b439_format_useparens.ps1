# check_b439_format_useparens.ps1 - UseParentheses 这一问只有一位主人，且**省略档不许和显式 True 同码**
#   (账 #439 §B212 / 第一百一十二刀)
#
# 为什么单开这一道: FormatNumber / FormatCurrency / FormatPercent 三家共用同一条 5 位 C 签名，
#   而「UseParentheses 要什么」原先被三家各自的 `(void)useParens` **明确丢掉** —— 三份抄写、
#   三份都答错（负数永远交 `-1,234.57`，运行期在该位为真时交 `(1,234.57)`）。这一刀把它收成
#   一位主人，并把**真正难的那一半**钉住:
#
#   ⚠ **省略档与显式 True 在 C 侧同码** —— VB6 的 `True` 恒为 **-1**，而发码那一侧原先也给省略档填
#     -1 ⇒ 两边分不开:
#       把 -1 当默认 ⇒ 显式 True 被吃掉（实测 -1,234.57 vs 运行期 (1,234.57)）
#       把 -1 当真   ⇒ **省略档全部变括号** = 把绝大多数正常调用一起改坏
#     VB6 里 UseParentheses 的默认是 **False**，所以发码那一档必须是 **0**。这一格是最容易被
#     「顺手统一成 -1」改坏的，而改坏之后**默认读数**才会变，肉眼看不出来。
#
# 规则（改坏了会红，不是装饰）:
#   R1  三家都必须走 `vb6_fmtApplyUseParens` 这一位主人；任何一家再写 `(void)useParens` ⇒ 红
#   R2  `kOD_FormatNum` 位 3（UseParentheses）的省略档必须是 **"0"**，不许是 "-1"
#       （同码 ⇒ 显式 True 被吃 / 省略档全变括号，两头都是静默错）
#   R3  位 1/2/4 的默认（NumDigits / IncludeLeadingDigits / GroupDigits）**不许**被顺手改成 "0"
#       —— 那三位在 VB6 里的默认是真/照默认，改 0 会把 GroupDigits 关掉、少千分位
#   R4  `kOD_FormatNum` 必须同时登记 **三家**（只登一家 ⇒ 另两家省略实参时 C2198 整份编不过）
#   R5  主人里的判据是 `useParens == 0` 放行负号那一支（不是 `> 0`）：默认档已改成 0
#
# 用法:  powershell -File scripts\check_b439_format_useparens.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$compatC = Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_compat.c'
$fixupInc = Join-Path $root 'src\backend\detail\expr\cgen_expr_call_builtin_fixup.inc'
foreach ($p in @($compatC, $fixupInc)) {
    if (-not (Test-Path -LiteralPath $p)) { Write-Host ('FAIL missing ' + $p) -ForegroundColor Red; exit 1 }
}
$comp = [IO.File]::ReadAllText($compatC)
$compCode = [regex]::Replace([regex]::Replace($comp, '/\*[\s\S]*?\*/', ' '), '(?m)//.*$', ' ')
$fix = [IO.File]::ReadAllText($fixupInc)
$bad = @()

# ---- R1: 三家都走那一位主人 ----
$owners = @('vb6_FormatNumber', 'vb6_FormatCurrency', 'vb6_FormatPercent')
foreach ($fn in $owners) {
    $m = [regex]::Match($compCode, '(?:BSTR|int32_t|double)\s+' + [regex]::Escape($fn) + '\s*\([^)]*\)\s*\{(?<b>[\s\S]*?)\n\}')
    if (-not $m.Success) { $bad += ('R1 cannot locate the body of ' + $fn + ' -> renamed/removed?'); continue }
    $b = $m.Groups['b'].Value
    if ($b -notmatch 'vb6_fmtApplyUseParens') {
        $bad += ('R1 ' + $fn + ' does not go through vb6_fmtApplyUseParens -> UseParentheses is answered by this ' +
                 'function alone again (the original shape: three copies, all wrong)')
    }
    if ($b -match '\(void\)\s*useParens') {
        $bad += ('R1 ' + $fn + ' still discards useParens via `(void)` -> the parameter is dropped on the floor')
    }
}

# ---- R2/R3: 默认档那一格 ----
$dm = [regex]::Match($fix, 'kOD_FormatNum\[\]\s*=\s*\{(?<b>[^}]*)\}')
if (-not $dm.Success) {
    $bad += 'R2 kOD_FormatNum table not found -> renamed/removed; the omitted-argument defaults lost their owner'
} else {
    $cells = @([regex]::Matches($dm.Groups['b'].Value, '"([^"]*)"') | ForEach-Object { $_.Groups[1].Value })
    $nulls = ([regex]::Matches($dm.Groups['b'].Value, 'nullptr')).Count
    # cells[0] = 位1(numDigits), [1] = 位2(incLeading), [2] = 位3(useParens), [3] = 位4(groupDigits)
    if ($cells.Count -lt 4) {
        $bad += ('R2 kOD_FormatNum has only ' + $cells.Count + ' default cells (want 4 for positions 1..4)')
    } else {
        if ($cells[2] -ne '0') {
            $bad += ('R2 position 3 (UseParentheses) defaults to "' + $cells[2] + '" instead of "0" -> VB6''s True is ' +
                     'always -1, so an omitted argument and an explicit True collide in C. Either an explicit True ' +
                     'gets ignored, or every omitted argument turns into parentheses.')
        }
        foreach ($i in @(0, 1, 3)) {
            if ($cells[$i] -eq '0') {
                $bad += ('R3 position ' + ($i + 1) + ' default changed to "0" -> those three default to -1/True in VB6; ' +
                         'zeroing GroupDigits would drop the thousands separators')
            }
        }
    }
}

# ---- R4: 三家都要登记那张默认档表 ----
foreach ($fn in $owners) {
    $n = ([regex]::Matches($fix, [regex]::Escape('"' + $fn + '"') + ',\s*kOD_FormatNum')).Count
    if ($n -ne 1) {
        $bad += ('R4 ' + $fn + ' is registered ' + $n + ' time(s) in the omitted-default table (want exactly 1) -> ' +
                 'an unregistered one emits too few arguments into the 5-parameter RTL exit = cl C2198')
    }
}

# ---- R5: 主人那一支的判据 ----
if ($compCode -notmatch 'static\s+BSTR\s+vb6_fmtApplyUseParens') {
    $bad += 'R5 vb6_fmtApplyUseParens is gone -> the UseParentheses question lost its single owner'
} elseif ($compCode -notmatch 'useParens\s*==\s*0\s*\)\s*return') {
    $bad += 'R5 the owner no longer tests `useParens == 0` for the minus-sign branch -> the omitted default (0) must ' +
             'be the one that keeps the minus sign; `> 0` would reject the explicit True (-1)'
}

if ($bad.Count -eq 0) {
    Write-Host 'PASS b439_format_useparens: 3 functions share one owner | no (void) useParens | UseParentheses default=0 | other defaults untouched | all 3 registered | violations 0' -ForegroundColor Green
    exit 0
}
foreach ($b in $bad) { Write-Host ('FAIL ' + $b) -ForegroundColor Red }
exit 1