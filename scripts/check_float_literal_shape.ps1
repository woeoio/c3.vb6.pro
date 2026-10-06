# check_float_literal_shape.ps1 - 账 #188 的结构性哨兵 (只扫编译器源码, 不起 cl)
#
# 为什么单开这一道: C 里的 `12f` 是**非法 token** (MSVC C2059 "bad suffix on number") ——
# 整数值的 Single 必须写成 `12.0f`。这条规则以前在三个地方各抄一遍, 其中「.frm 设计期字体块」
# 那处抄漏了补小数点那一步 (`snprintf("%.4g")` 之后直接拼 "f") ⇒ 凡设计期字号是整数的工程
# 直接编不过 (实测 Charts 2020 的 ucChartBar / ucProgressCircular 两家, 生成码里 `->Size = 12f`)。
# VB6 的默认字号 8.25 带小数点, 所以这条在门禁里一直隐身。
#
# 口径现在收在一处: src/common/float_literal.hpp 的 floatSingleLiteral / floatingLiteralText。
#
# 规则 (改坏了会红, 不是装饰):
#   R1  src/backend 与 src/semantics 里**手拼** f 后缀 = 0 处 (`+ "f"` / `<< "f"` / `to_string(...) + "f`)
#   R2  权威头存在, 三个函数各只定义一次, 且必须带「没有 . 或 eE 就补 .0」那一步
#   R3  调用点 >= 4 处 (只删调用者、把规则写回本地也算红)
#
# 用法:  pwsh -File scripts\check_float_literal_shape.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$bad = @()

$dirs = @("src\backend", "src\semantics") | ForEach-Object { Join-Path $root $_ }
$files = @()
foreach ($d in $dirs) {
    if (Test-Path -LiteralPath $d) {
        $files += Get-ChildItem -LiteralPath $d -Recurse -File |
                  Where-Object { $_.Extension -in ".cpp", ".inc", ".hpp" }
    } else { $bad += ("R0 missing dir: " + $d) }
}

# R1: 手拼后缀 ( (?![A-Za-z]) 是为了不把 "font..." 这类字符串误当后缀 )
$patHand = '(\+\s*"f(?![A-Za-z]))|(<<\s*"f(?![A-Za-z]))|(std::to_string\s*\([^)]*\)\s*\+\s*"f(?![A-Za-z]))'
$hand = 0
$handWhere = @()
foreach ($f in $files) {
    $t = [System.IO.File]::ReadAllText($f.FullName)
    foreach ($m in [regex]::Matches($t, $patHand)) {
        $hand++
        $handWhere += ($f.Name + ":" + $m.Value.Trim())
    }
}
if ($hand -ne 0) {
    $bad += ("R1 hand-concatenated float suffixes = " + $hand + " (must be 0; use floatSingleLiteral) -> " +
             (($handWhere | Select-Object -First 6) -join " | "))
}

# R2: 权威头
$hdr = Join-Path $root "src\common\float_literal.hpp"
if (-not (Test-Path -LiteralPath $hdr)) {
    $bad += "R2 src/common/float_literal.hpp is missing (the single authority)"
} else {
    $h = [System.IO.File]::ReadAllText($hdr)
    foreach ($fn in @("floatingLiteralText", "floatSingleLiteral", "floatFixed6Literal")) {
        $defs = @([regex]::Matches($h, 'inline\s+std::string\s+' + $fn + '\s*\('))
        if ($defs.Count -ne 1) { $bad += ("R2 " + $fn + " defined " + $defs.Count + " times in the header (want 1)") }
    }
    if ($h -notmatch 'find_first_of\("\.eE"\)') {
        $bad += 'R2 the header lost the dot/exponent guard (find_first_of(".eE")) - that is the whole bug'
    }
    if ($h -notmatch 'locale::classic') {
        $bad += "R2 the header lost the classic-locale imbue (a comma decimal separator would break the literal)"
    }
}

# R3: 调用点数量
$calls = 0
foreach ($f in @($files + @(Get-ChildItem -LiteralPath (Join-Path $root "src\common") -File -ErrorAction SilentlyContinue))) {
    if ($f.FullName -eq $hdr) { continue }
    $t = [System.IO.File]::ReadAllText($f.FullName)
    $calls += @([regex]::Matches($t, '(floatSingleLiteral|floatFixed6Literal)\s*\(')).Count
    $calls += @([regex]::Matches($t, 'floatingLiteralText\s*\(')).Count
}
if ($calls -lt 4) {
    $bad += ("R3 only " + $calls + " call sites into the authority (want >= 4) - the rule is being re-inlined somewhere")
}

if ($bad.Count -eq 0) {
    Write-Host ("PASS float-literal authority: 0 hand-concat sites, " + $calls + " call sites into common/float_literal.hpp") -ForegroundColor Green
    exit 0
}
foreach ($b in $bad) { Write-Host ("FAIL " + $b) -ForegroundColor Red }
exit 1
