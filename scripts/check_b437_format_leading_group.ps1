# check_b437_format_leading_group.ps1 - 5 位签名上剩下那两枚位，每一枚都有主人，且数得对
#   (账 #437 §B213 / 第一百一十三刀)
#
# 为什么值得单开一道: §B212 收了 useParens，但同一条 5 位签名上还有 **incLeading** 与
#   **groupDigits** 两枚是 `(void)` 丢掉的 —— 丢法还各不相同（FormatNumber 只 void 了
#   incLeading 却手写了一遍千分位；FormatCurrency 两个一起 void；FormatPercent 的 groupDigits
#   **连 void 都没有**，是压根没接）。三份抄写、三种坏法，全部**静默答错**（不编译失败）。
#
# 这道门钉的是本刀真踩到的四个坑，每一个都是「形状对、格子错」那一族：
#   R1  两枚都各有**唯一**主人；三家都调得到，任何一家再 `(void)` 掉 ⇒ 红
#   R2  删前导 0 的 memmove 必须搬 **wcslen+1** 个元素（含 '\0'）。少搬一格 ⇒ 被删的 '0'
#       原地留在末尾 ⇒ "0.50" 变 ".500"、"0.50%" 变 ".50%%"、"0" 压根没删（实测全中）
#   R3  插完千分位必须把**负号接回最前面**：p 为了数数字跳过了符号位，out 是从数字段重写的，
#       不补符号 ⇒ `FormatNumber(-1234.5678,2)` 交 `1,234.57`（引擎 `-1,234.57`）
#   R4  FormatCurrency 的括号**不许补 ')'**：body+skip 的收尾 ')' 已经在 body 里
#       （`vb6_fmtApplyUseParens` 交的是完整 `(…)`），补一个 ⇒ `(¥1,234.57))`
#   R5  货币符号必须在 '(' **之后**（引擎 `(¥1,234.57)`）—— §B212 写成括号外，
#       而那一刀只钉了 `InStr(s,"(")>0` 的**形状**，两边都有括号 ⇒ 形状绿、位置错一刀没人看见
#   R6  percent 的 buf 带着 '%'，两位主人都必须按「数字跑」看，不能按「到 '.' 为止」：
#       判前导 0 时漏掉 '%' ⇒ "0%" 被当成「不止一位」⇒ 交 `0%` 而不是运行期的 `%`；
#       数整数段时把 '%' 当一位 ⇒ 逗号插错位置
#   R7  夹具必须**连符号位置一起钉**（k_cur_par=/k_cur_ng=），且不许把货币符号写进字面量
#
# 用法:  powershell -File scripts\check_b437_format_leading_group.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$compatC = Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_compat.c'
$fixupInc = Join-Path $root 'src\backend\detail\expr\cgen_expr_call_builtin_fixup.inc'
$fixture = Join-Path $root 'tests\b437_format_leading_group.bas'
foreach ($p in @($compatC, $fixupInc, $fixture)) {
    if (-not (Test-Path -LiteralPath $p)) { Write-Host ('FAIL missing ' + $p) -ForegroundColor Red; exit 1 }
}
$comp = [IO.File]::ReadAllText($compatC)
$code = [regex]::Replace([regex]::Replace($comp, '/\*[\s\S]*?\*/', ' '), '(?m)//.*$', ' ')
$fix = [IO.File]::ReadAllText($fixupInc)
$fixCode = [regex]::Replace([regex]::Replace($fix, '/\*[\s\S]*?\*/', ' '), '(?m)//.*$', ' ')
$fx = [IO.File]::ReadAllText($fixture)
$bad = @()

# ---- R1: 两枚各有唯一主人，三家都调 ----
foreach ($h in @('vb6_fmtDropLeadingZero', 'vb6_fmtGroupThousands')) {
    $defs = ([regex]::Matches($code, 'static\s+void\s+' + [regex]::Escape($h) + '\s*\(')).Count
    if ($defs -ne 1) { $bad += ('R1 ' + $h + ' is defined ' + $defs + ' time(s), want exactly 1 -> three copies all wrong is ' +
                                'exactly how §B212 happened') }
}
foreach ($fn in @('vb6_FormatNumber', 'vb6_FormatCurrency', 'vb6_FormatPercent')) {
    $m = [regex]::Match($code, 'BSTR\s+' + [regex]::Escape($fn) + '\s*\([^)]*\)\s*\{(?<b>[\s\S]*?)\n\}')
    if (-not $m.Success) { $bad += ('R1 cannot locate the body of ' + $fn); continue }
    $b = $m.Groups['b'].Value
    foreach ($h in @('vb6_fmtDropLeadingZero', 'vb6_fmtGroupThousands')) {
        if ($b -notmatch [regex]::Escape($h)) {
            $bad += ('R1 ' + $fn + ' never calls ' + $h + ' -> this position is answered by that one function alone again')
        }
    }
    if ($b -match '\(void\)\s*(incLeading|groupDigits)') {
        $bad += ('R1 ' + $fn + ' still voids an optional position -> it is dropped on the floor again')
    }
}
# 千分位的顺序：必须在**同一个函数体内**先于括号/符号定形，否则 '(' 或 '¥' 会被当成一位数字数错。
# ⚠ 这个顺序判据**必须逐个函数体查**：早先写成跨全文找 `ApplyUseParens(...)` 后面 400 字内有没有
#   `GroupThousands(`，结果跨过了函数边界 —— FormatCurrency 交完括号之后就是下一个函数的
#   千分位调用 —— 于是一份**顺序完全正确**的源码被判红。跨边界的「顺序」不是顺序。
foreach ($fn in @('vb6_FormatNumber', 'vb6_FormatCurrency', 'vb6_FormatPercent')) {
    $m = [regex]::Match($code, 'BSTR\s+' + [regex]::Escape($fn) + '\s*\([^)]*\)\s*\{(?<b>[\s\S]*?)\n\}')
    if (-not $m.Success) { continue }
    $b = $m.Groups['b'].Value
    $g = $b.IndexOf('vb6_fmtGroupThousands')
    $p = $b.IndexOf('vb6_fmtApplyUseParens')
    if ($g -ge 0 -and $p -ge 0 -and $g -gt $p) {
        $bad += ('R6 in ' + $fn + ' the parens owner runs before grouping -> the "(" would be counted as a digit ' +
                 'and the commas land one place off')
    }
}

# ---- R2: memmove 必须含 '\0' ----
$dzm = [regex]::Match($code, 'static\s+void\s+vb6_fmtDropLeadingZero\([^)]*\)\s*\{(?<b>[\s\S]*?)\n\}')
if ($dzm.Success) {
    $mm = [regex]::Match($dzm.Groups['b'].Value, 'memmove\([^;]*\)')
    if (-not $mm.Success) {
        $bad += 'R2 vb6_fmtDropLeadingZero no longer memmoves -> the leading 0 is not removed'
    } elseif ($mm.Value -notmatch 'wcslen\s*\([^)]*\)\s*\+\s*1\s*\)') {
        # ⚠⚠ 早先两版都栽在这里。第一版用 `[^)]*` 抓第三格，在 `wcslen(p + 1)` 的 ')' 就停 ⇒ 把
        #   **正确**的源码判红。第二版改成找 `memmove` 整段再搜 `+ 1`，可 `p + 1` 与
        #   `wcslen(p + 1)` **本身就含 `+ 1`** ⇒ 把缺 `+ 1` 的坏源码放过（实测 NC1 绿）。
        #   现在只认「wcslen(...) 之后再 +1 再收右括号」这一形状：`wcslen(p + 1) + 1)` 命中，
        #   坏写法 `(wcslen(p + 1))` 因为 `)` 后面不是 `+ 1)` 而不命中。
        $bad += ('R2 the shift is `' + $mm.Value.Trim() + '`, which does not move wcslen(...)+1 -> the deleted ''0'' ' +
                 'stays behind: "0.50"->".500", "0.50%"->".50%%", "0"->"0" (all measured)')
    }
}

# ---- R3: 负号接回 ----
$gtm = [regex]::Match($code, 'static\s+void\s+vb6_fmtGroupThousands\([^)]*\)\s*\{(?<b>[\s\S]*?)\n\}')
if ($gtm.Success -and $gtm.Groups['b'].Value -notmatch "out\s*\[\s*0\s*\]\s*=\s*L'-'") {
    # ⚠ 只搜 `L'-'` 是不够的：函数开头那句 `wchar_t* p = (buf[0] == L'-') ? buf + 1 : buf;`
    #   **本来就有** `L'-'`（跳过符号位），把整段补符号的代码删掉之后它照样在 ⇒ 坏源码照样绿。
    #   判据必须是「**写回**输出缓冲」那一枚，即 `out[0] = '-'`。
    $bad += 'R3 vb6_fmtGroupThousands never writes the sign back into the output buffer -> FormatNumber(-1234.5678,2) ' +
            'yields 1,234.57 instead of the engine''s -1,234.57'
}

# ---- R4/R5: 括号与符号的位置 ----
$fc = [regex]::Match($code, 'BSTR\s+vb6_FormatCurrency\s*\([^)]*\)\s*\{(?<b>[\s\S]*?)\n\}')
if (-not $fc.Success) {
    $bad += 'R4 cannot locate vb6_FormatCurrency'
} else {
    $b = $fc.Groups['b'].Value
    # R4: skip 分支里补的那个 ')' —— body 已经带收尾括号
    $skipBranch = [regex]::Match($b, '(?s)if\s*\(\s*skip\s*\)\s*\{(?<b>.*?)\n    \}')
    if ($skipBranch.Success) {
        $inner = $skipBranch.Groups['b'].Value
        if ($inner -match 'wcscat\s*\(\s*result\s*,\s*L"\)\s*"') {
            $bad += 'R4 the skip branch appends another ")" -> body+skip already ends with it, giving (Y1,234.57))'
        }
        if ($inner -notmatch 'L"\(\\x00a5"') {
            $bad += 'R5 the skip branch no longer puts the symbol right after "(" -> the engine emits (Y1,234.57), ' +
                    'symbol inside; §B212 had it outside and the shape-only fixture let it through'
        }
    }
}

# ---- R6: percent 缓冲带 '%'，两位主人都必须认 ----
if ($code -notmatch "nxt\s*!=\s*L'%'") {
    $bad += 'R6 the leading-zero rule does not know about ''%'' -> "0%" counts as multi-digit and stays "0%" ' +
            'instead of the engine''s "%"'
}
if ($gtm.Success -and $gtm.Groups['b'].Value -notmatch ">= L'0'\s*&&\s*p\[intLen\]\s*<=\s*L'9'") {
    $bad += 'R6 vb6_fmtGroupThousands does not count the digit run -> a ''%'' suffix would be counted as a digit ' +
            'and the commas land one place off'
}

# ---- R7: 夹具连符号位置一起钉，且不钉货币符号字面量 ----
# ---- R7: 断言连符号位置一起钉，且不钉货币符号字面量 ----
# ⚠ 断言住在 run_tests.ps1 的 needle 名单里，不在 .bas 里 —— 早先在 .bas 里找
#   `k_cur_par=(1,234.57)` 找了一整轮找不到，因为 .bas 只写 `Debug.Print "k_cur_par=" & …`。
# ⚠⚠ 必须在**去掉注释行**之后的文本里找 needle。§B438 实测过同一个坑：直接读整个
#   run_tests.ps1，上面说明文字里反引号包着的 needle 字面量**也算命中** ⇒ 真把 needle 改掉
#   之后这道门照样绿。needle 是**断言**，散文不是。
$harness = (@([IO.File]::ReadAllLines((Join-Path $root 'tests\run_tests.ps1')) | Where-Object { $_ -notmatch '^\s*#' }) -join "`n")
foreach ($n in @('k_cur_par=(1,234.57)', 'k_cur_ng=(1234.57)', 'k_pct_true=(123,456.78%)', 'k_pct_ng=(123456.78%)')) {
    if ($harness -notmatch [regex]::Escape($n)) { $bad += ('R7 the harness does not pin ' + $n + ' -> the symbol/paren placement can drift again') }
}
# 货币符号只许出现在**注释**里（那里是在解释运行期读数），不许进任何 Debug.Print 的字面量。
$printLines = @([IO.File]::ReadAllLines($fixture) | Where-Object { $_ -match '^\s*Debug\.Print' }) -join "`n"
foreach ($glyph in @([char]0x00A5, [char]0xFFE5, [char]0x00A4)) {
    if ($printLines.IndexOf($glyph) -ge 0) {
        $bad += ('R7 a Debug.Print pins currency glyph U+' + ('{0:X4}' -f [int]$glyph) + ' -> that glyph follows the locale (§B142); strip it with StripCur instead')
    }
}
if ($fx -notmatch 'StripCur') { $bad += 'R7 the fixture lost StripCur -> FormatCurrency readings would pin the symbol' }

# ---- R8: 这两位的省略档仍必须是 -1（默认 True），不许顺手改成 0 ----
$dm = [regex]::Match($fixCode, 'kOD_FormatNum\[\]\s*=\s*\{(?<b>[^}]*)\}')
if (-not $dm.Success) {
    $bad += 'R8 kOD_FormatNum table not found'
} else {
    $cells = @([regex]::Matches($dm.Groups['b'].Value, '"([^"]*)"') | ForEach-Object { $_.Groups[1].Value })
    if ($cells.Count -lt 4) {
        $bad += ('R8 kOD_FormatNum has only ' + $cells.Count + ' default cells (want 4)')
    } else {
        foreach ($i in @(0, 1, 3)) {
            if ($cells[$i] -eq '0') {
                $bad += ('R8 position ' + ($i + 1) + ' default changed to "0" -> NumDigits/IncludeLeadingDigits/GroupDigits ' +
                         'all default to -1 (True) in VB6; zeroing GroupThousands would drop the separators')
            }
        }
    }
}

if ($bad.Count -eq 0) {
    Write-Host 'PASS b437_format_leading_group: 2 owners, 3 callers each | shift keeps the NUL | sign restored | no doubled ")" | symbol inside "(" | "%" handled | defaults intact; violations 0' -ForegroundColor Green
    exit 0
}
foreach ($b in $bad) { Write-Host ('FAIL ' + $b) -ForegroundColor Red }
exit 1