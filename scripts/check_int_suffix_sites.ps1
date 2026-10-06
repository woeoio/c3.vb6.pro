# check_int_suffix_sites.ps1 - 账 #195 的结构性哨兵 (只扫编译器源码, 不起 cl)
#
# 为什么单开这一道: VB 的整数类型后缀 (% Integer / & Long / ^ LongPtr) 在词法器里被**抄了四份**
# —— 十进制扫描器一份、&H / &O / &B 各一份, 外加 radixDigits 的剥离表一份。`%` 那一支只把字符
# 吃进 rawText、不置任何标志, 于是 `3%` 落回「无后缀十进制按数值大小定档」那一段,
# parseIntLit 看见残留的 % 就报「十进制数字超出 64 位整数表示范围」—— 一条合法语句级联出
# 9-13 条诊断, 整个工程卡在最前面 (实测 .build/p195_*.bas: 十进制 / 三种进制 / 负数 / 表达式 /
# 参数默认值 六形, 修前全红, 修后全 0 条)。
#
# 规则 (改坏了会红, 不是装饰):
#   S1  十进制扫描器的 `%` 支必须置 isInteger (且不许留「只吃字符然后 break」的空支)
#   S2  radixDigits 的剥离表必须同时含 ^ & % 三个后缀字符
#   S3  四处消费后缀的地方都认 `%` (含 "== '%'" 的行数 >= 4)
#   S4  显式 `%` 的 16 位范围检查还在 (按 int32 收下再让 int16 截 = 把值改错)
#
# 用法:  pwsh -File scripts\check_int_suffix_sites.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$bad = @()
$lex = Join-Path $root "src\lexer\lexer_number.cpp"
if (-not (Test-Path -LiteralPath $lex)) {
    Write-Host "FAIL S0 src/lexer/lexer_number.cpp is missing" -ForegroundColor Red
    exit 1
}
$t = [System.IO.File]::ReadAllText($lex)
$rows = $t -split "\r?\n"

$set = 0
$bare = 0
$pctLines = 0
$ampLines = 0
foreach ($ln in $rows) {
    $tt = $ln.Trim()
    if ($tt.StartsWith("case '%':")) {
        if ($tt.Contains("isInteger = true")) { $set = $set + 1 }
        if ($tt.Contains("advance(); break;")) { $bare = $bare + 1 }
    }
    if ($ln.Contains("== '%'")) { $pctLines = $pctLines + 1 }
    if ($ln.Contains("== '&'")) { $ampLines = $ampLines + 1 }
}

if ($set -ne 1) {
    $bad += ("S1 the decimal scanner sets isInteger on the % branch " + $set + " times (want 1) -" +
             " an empty branch is exactly how 3% turned into a lex error")
}
if ($bare -ne 0) {
    $bad += ("S1 a bare % branch (consume the char, break, set nothing) is back: " + $bare)
}

# S2: 三个后缀都在剥离表里 (同一行里出现 ^ & % 三个字符比较)
$strip = $rows | Where-Object { $_.Contains("== '^'") -and $_.Contains("== '%'") }
if (@($strip).Count -ne 1) {
    $bad += ("S2 radixDigits' strip table does not carry ^ and % together (matched " +
             @($strip).Count + " lines, want 1) - the leftover % breaks parseIntLit")
}

# S3: 后缀消费点
if ($pctLines -lt 4) {
    $bad += ("S3 only " + $pctLines + " lines consume % (want >= 4: decimal + &H + &O + &B)")
}
if ($ampLines -lt 4) {
    $bad += ("S3 only " + $ampLines + " lines consume & (want >= 4) - the four copies drifted apart")
}

# S4: 范围守卫
if (-not $t.Contains("INT16_MIN") -or -not $t.Contains("INT16_MAX")) {
    $bad += "S4 the explicit-% range guard is gone - a too-big % would be silently wrapped, not reported"
}

if ($bad.Count -eq 0) {
    Write-Host ("PASS int-suffix sites: % sets isInteger once, strip table carries ^ and %, " +
                $pctLines + " consumers, 16-bit guard present") -ForegroundColor Green
    exit 0
}
foreach ($b in $bad) { Write-Host ("FAIL " + $b) -ForegroundColor Red }
exit 1
