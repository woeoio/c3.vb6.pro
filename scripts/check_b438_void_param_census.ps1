# check_b438_void_param_census.ps1 - 三条「(void) 形参」看着像丢了一位，其实都不是
#   (账 #438 §B214)
#
# 这道门不新增行为，它**阻止误修**。§B212 与 §B213 两刀都是从「某处写着 (void)x;」这句
# 源码里读出来的；普查顺着同一句去找，读到 121 枚 (void)，其中三枚最像「静默丢了一位」。
# 逐个量完之后，**三条全是误报**：
#
#   ① vb6_DateDiff 的 (void)firstWeekOfYear —— **实测惰性**，不是丢掉了一位。
#      尺 .build/kfwy.vbs：6 个间隔(w/ww/y/q/m/d) x 5 组**跨年边界**的日期对，fwy=1/2/3
#      **读数全同**。§B197 那两条公式只吃 f=firstDayOfWeek，压根不吃 firstWeekOfYear
#      —— 所以这里 (void) 是**对的**。⚠ 它形同「丢了一位」，任何普查都会重新举它一次，
#      然后有人去「修」，把一个对的实现改成错的 ⇒ 这条必须钉住。
#   ② vb6_IIfLong / vb6_IIfDouble 的 (void)falsepart —— **多余**的，下一行就用到了
#      （`return cond ? truepart : falsepart`）。⚠ 而且这三个出口发码**真的走得到**
#      （cgen_expr_call_builtin_pre.inc 会发 vb6_IIfLong/Double/BSTR），IIf 是 VB6 高频件，
#      普查前它**一处回归都没有**。
#   ③ vb6_IPmt 的 (void)nper / (void)fv —— 同样**多余**，两个都喂给了下一行的 vb6_Pmt。
#
# 规则（改坏了会红）：
#   R1  ①的 (void)firstWeekOfYear 必须**留在** DateDiff 里，且 DateDiff 的 w / ww 两条臂
#       **不许**去读 firstWeekOfYear（读了就与实测的全同读数矛盾）
#   R2  ②③ 的形参必须**仍然被用到** —— `(void)x;` 旁边那句若不再引用 x，说明有人把
#       它「清理」成了真丢弃，那才是把对的实现改成错的
#   R3  夹具必须钉住「fwy=1/2/3 三格读数相同」（惰性）与 IIf 的两个分支（含负数即真）
#   R4  诚实性：财务那一族（IPmt/PPmt/Pmt/FV/PV/RATE/NPER/IRR）**在本机没有引擎代理**
#       —— VBScript 压根没有财务函数（连 IIf 也没有）。所以本门**不许**假装钉了它们的
#       绝对读数，只钉「形参仍被引用」这一条静态事实
#
# 用法:  powershell -File scripts\check_b438_void_param_census.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$dateC = Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_date.c'
$miscC = Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_misc.c'
$finC = Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_financial.c'
$fixture = Join-Path $root 'tests\b438_void_param_census.bas'
$harness = Join-Path $root 'tests\run_tests.ps1'
foreach ($p in @($dateC, $miscC, $finC, $fixture, $harness)) {
    if (-not (Test-Path -LiteralPath $p)) { Write-Host ('FAIL missing ' + $p) -ForegroundColor Red; exit 1 }
}
function Get-Body([string]$path, [string]$fn) {
    $src = [IO.File]::ReadAllText($path)
    $m = [regex]::Match($src, '(?:^|\n)[^\n]*\b' + [regex]::Escape($fn) + '\s*\([^;]*\)\s*\{(?<b>[\s\S]*?)\n\}')
    if (-not $m.Success) { return $null }
    $b = $m.Groups['b'].Value
    return [regex]::Replace([regex]::Replace($b, '/\*[\s\S]*?\*/', ' '), '(?m)//.*$', ' ')
}
$bad = @()

# ---- R1: DateDiff 的 firstWeekOfYear 惰性 ----
$dd = Get-Body $dateC 'vb6_DateDiff'
if (-not $dd) {
    $bad += 'R1 cannot locate vb6_DateDiff -> renamed/removed; the verified-inert pin lost its anchor'
} else {
    if ($dd -notmatch '\(void\)\s*firstWeekOfYear') {
        $bad += 'R1 the (void)firstWeekOfYear in vb6_DateDiff is gone -> either it was "fixed" into a bogus reader, ' +
                'or the anchor moved. Measured: firstWeekOfYear changes nothing (6 intervals x 5 year-boundary pairs)'
    }
    # 若两条臂真的去读了 firstWeekOfYear，读数就会与「fwy=1/2/3 全同」矛盾。
    # ⚠ 判据钉的是**出现次数**：函数体（不含形参表）里正确状态应当**只有那一条 (void) 铸死**，
    #   也就是恰好 1 次；真去读它就变成 2 次。门槛必须写 `>= 2`。
    #   早先写成 `> 2` —— 正因为「形参表那一行也在同一段文本里」时正确值就是 2，
    #   门槛等于要求 3 次，**把最该逮的那次误修放过了**（负控是绿的才发现）。
    $fwyHits = ([regex]::Matches($dd, '\bfirstWeekOfYear\b')).Count
    if ($fwyHits -ge 2) {
        $bad += ('R1 vb6_DateDiff READS firstWeekOfYear (it occurs ' + $fwyHits + 'x in the body, not just the (void) cast) -> ' +
                 'the engine holds fwy=1/2/3 identical, so this reading must be wrong')
    }
}

# ---- R2: ②③ 的形参仍被引用 ----
foreach ($spec in @(
    @{ file = $miscC;    fn = 'vb6_IIfLong';   p = 'falsepart'; why = 'IIf would always take the true branch' },
    @{ file = $miscC;    fn = 'vb6_IIfDouble'; p = 'falsepart'; why = 'IIf would always take the true branch' },
    @{ file = $finC;     fn = 'vb6_IPmt';      p = 'nper';      why = 'IPmt could not know the term' },
    @{ file = $finC;     fn = 'vb6_IPmt';      p = 'fv';        why = 'IPmt could not know the residual' })) {
    $b = Get-Body $spec.file $spec.fn
    if (-not $b) { $bad += ('R2 cannot locate ' + $spec.fn); continue }
    $withoutVoid = [regex]::Replace($b, '\(void\)\s*[A-Za-z_][A-Za-z0-9_]*\s*;', '')
    if ($withoutVoid -notmatch ('\b' + [regex]::Escape($spec.p) + '\b')) {
        $bad += ('R2 ' + $spec.fn + ' no longer uses $' + $spec.p + ' once the (void) casts are removed -> ' +
                 'the cast was "cleaned up" into a real drop and ' + $spec.why)
    }
}

# ---- R3: 夹具钉住惰性与两个分支 ----
$fx = [IO.File]::ReadAllText($fixture)
$hh = [IO.File]::ReadAllText($harness)
# ⚠⚠ 必须在**去掉注释行**之后的文本里找 needle。早先直接读整个 run_tests.ps1，于是上面那段
#   说明文字里反引号包着的 `iifL_n=7` **也算命中** ⇒ 真把 needle 改成 iifL_n=8 之后，这道门
#   照样绿（负控是绿的才发现）。needle 是**断言**，散文不是 —— 这与 §B437 的教训同源：
#   判据要钉「那一份名单」，不能钉「这个文件里出现过这个字符串」。
$hh = [IO.File]::ReadAllText($harness)
# ⚠⚠ 必须在**去掉注释行**之后的文本里找 needle。早先直接读整个 run_tests.ps1，于是上面那段
#   说明文字里反引号包着的 `iifL_n=7` **也算命中** ⇒ 真把 needle 改成 iifL_n=8 之后，这道门
#   照样绿（负控是绿的才发现）。needle 是**断言**，散文不是 —— 这与 §B437 的教训同源：
#   判据要钉「那一份名单」，不能钉「这个文件里出现过这个字符串」。
$hh = (@([IO.File]::ReadAllLines($harness) | Where-Object { $_ -notmatch '^\s*#' }) -join "`n")
$needles = @(
    'w_f1=52', 'w_f2=52', 'w_f3=52',          # w  与 fwy 无关（实测 52/52/52）
    'ww_f1=53', 'ww_f2=53', 'ww_f3=53',       # ww 与 fwy 无关（实测 53/53/53）
    'wwb_1=1', 'wwb_2=1',                    # 跨年那组
    'iifL_t=11', 'iifL_f=22', 'iifD_t=1.5', 'iifD_f=2.5', 'iifS_t=yes', 'iifS_f=no',
    'iifL_0=8', 'iifL_n=7',                  # 0 是假；-5 非零即真
    'B438-DONE')
foreach ($n in $needles) {
    if ($hh -notmatch [regex]::Escape($n)) { $bad += ('R3 the harness does not pin ' + $n) }
}
# 三格读数必须**字面相同** —— 那正是「惰性」这条账的全部内容
foreach ($triple in @(@('w_f1', 'w_f2', 'w_f3'), @('ww_f1', 'ww_f2', 'ww_f3'), @('wwb_1', 'wwb_2'))) {
    $vals = @()
    foreach ($k in $triple) { if ($hh -match ([regex]::Escape($k) + '=(\S+)"')) { $vals += $Matches[1] } }
    if ($vals.Count -ge 2 -and (@($vals | Sort-Object -Unique)).Count -ne 1) {
        $bad += ('R3 ' + ($triple -join '/') + ' no longer agree (' + ($vals -join ',') + ') -> firstWeekOfYear is not inert any more')
    }
}

# ---- R4: 诚实性 —— 不许假装钉了财务族的绝对读数 ----
$printLines = @([IO.File]::ReadAllLines($fixture) | Where-Object { $_ -match '^\s*Debug\.Print' }) -join "`n"
foreach ($fn in @('IPmt', 'PPMT', 'Pmt(', 'FV(', 'RATE(', 'NPER(', 'IRR(')) {
    if ($printLines -match ('\b' + [regex]::Escape($fn.TrimEnd('(')) + '\s*\(')) {
        $bad += ('R4 the fixture pins an absolute reading for ' + $fn.TrimEnd('(') + ' -> VBScript has no financial ' +
                 'functions at all, so there is NO engine oracle behind that number; it would be a made-up expectation')
    }
}

if ($bad.Count -eq 0) {
    Write-Host 'PASS b438_void_param_census: DateDiff firstWeekOfYear inert (and unread) | IIf falsepart + IPmt nper/fv still referenced | fixture pins 52/52/52, 53/53/53, 1/1 and both IIf branches | no fabricated financial readings; violations 0' -ForegroundColor Green
    exit 0
}
foreach ($b in $bad) { Write-Host ('FAIL ' + $b) -ForegroundColor Red }
exit 1