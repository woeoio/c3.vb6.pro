# check_int_literal_shape.ps1 - 账 #194 的结构性哨兵 (只扫编译器源码, 不起 cl)
#
# 为什么单开这一道: VB 的整数类型后缀 (`12&` = Long、`3%` = Integer、`&H10&` = 十六进制 Long)
# 属于**词法**, token 的 rawText 里带着它。C 只认 `12L` / `16` —— 所以凡"按文本抄默认值"的地方
# 都会往生成 C 里塞一个非法 token (MSVC C2059 "bad suffix on number")。
# 实测: Charts 2020/ucTreeMaps 的 PropPagFMR.pag:740/886 `Optional ... As Long = 0&` 发成
# `if (!_has_FontIndex) (*FontIndex) = 0&;` ⇒ 两条 C2059, 而那两条压在 #192 的一堆 C2039 底下,
# 不做结构性哨兵就永远看不见它。
#
# 口径现在收在一处: src/common/int_literal.hpp 的 intLiteralText(数值, 是否 Long 档)。
#
# 规则 (改坏了会红, 不是装饰):
#   I1  src/backend 与 src/semantics 里**手拼**整数后缀 = 0 处 (`+ "L"` / `+ "LL"`)
#   I2  权威头存在, intLiteralText 只定义一次, 且保留两条: LL 那一步的 32 位界限判定
#       (INT32_MIN/INT32_MAX) 与 Integer 档的"无后缀"支路
#   I3  调用点 >= 3, 且发码侧 (cgen_expr.cpp) 与语义侧 (semantic_analyzer_util.cpp) 各至少一处
#   I4  语义侧的 Integer/Long 两支不许再 return rawText (那条抄文本的路就是这次的根因)
#
# 用法:  pwsh -File scripts\check_int_literal_shape.ps1
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
    } else { $bad += ("I0 missing dir: " + $d) }
}

# I1: 手拼后缀。"L" 必须是一个**完整**的 C 字符串字面量才算 (发射 wide 字面量那种 "L\"" + x 不算)。
$patHand = '\+\s*"L{1,2}"'
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
    $bad += ("I1 hand-concatenated integer suffixes = " + $hand + " (must be 0; use intLiteralText) -> " +
             (($handWhere | Select-Object -First 6) -join " | "))
}

# I2: 权威头
$hdr = Join-Path $root "src\common\int_literal.hpp"
if (-not (Test-Path -LiteralPath $hdr)) {
    $bad += "I2 src/common/int_literal.hpp is missing (the single authority)"
} else {
    $h = [System.IO.File]::ReadAllText($hdr)
    $defs = @([regex]::Matches($h, 'inline\s+std::string\s+intLiteralText\s*\('))
    if ($defs.Count -ne 1) { $bad += ("I2 intLiteralText defined " + $defs.Count + " times in the header (want 1)") }
    if ($h -notmatch 'INT32_MIN' -or $h -notmatch 'INT32_MAX') {
        $bad += "I2 the header lost the 32-bit bound check - that is the LL step (2147483648L is not a long)"
    }
    if ($h -notmatch 'static_cast<int>') {
        $bad += "I2 the header lost the Integer (suffix-free) branch - every Integer would get an L"
    }
}

# I3: 调用点
$calls = 0
$callBackend = 0
$callSemantics = 0
foreach ($f in $files) {
    $t = [System.IO.File]::ReadAllText($f.FullName)
    $n = @([regex]::Matches($t, 'intLiteralText\s*\(')).Count
    $calls += $n
    if ($f.Name -eq "cgen_expr.cpp") { $callBackend = $n }
    if ($f.Name -eq "semantic_analyzer_util.cpp") { $callSemantics = $n }
}
if ($calls -lt 3) {
    $bad += ("I3 only " + $calls + " call sites into the authority (want >= 3) - the rule is being re-inlined somewhere")
}
if ($callBackend -lt 1) { $bad += "I3 the codegen side (visit(LiteralExpr)) no longer calls the authority" }
if ($callSemantics -lt 1) { $bad += "I3 the semantics side (evalOptionalDefault) no longer calls the authority" }

# I4: 语义侧不许对 Integer/Long 抄 rawText
$sem = Join-Path $root "src\semantics\semantic_analyzer_util.cpp"
if (Test-Path -LiteralPath $sem) {
    $st = [System.IO.File]::ReadAllText($sem)
    $i = $st.IndexOf("case LiteralKind::Integer:")
    if ($i -lt 0) {
        $bad += "I4 the Integer branch vanished from evalOptionalDefault"
    } else {
        $j = $st.IndexOf("case LiteralKind::LongPtr:", $i)
        if ($j -lt 0) { $j = $i + 1600 }
        $blk = $st.Substring($i, $j - $i)
        if ($blk -match 'return\s+lit->rawText') {
            $bad += "I4 the semantics default reverts to copying rawText - VB's &/% suffix then lands in the generated C (C2059)"
        }
    }
}

if ($bad.Count -eq 0) {
    Write-Host ("PASS int-literal authority: 0 hand-concat sites, " + $calls + " call sites into common/int_literal.hpp") -ForegroundColor Green
    exit 0
}
foreach ($b in $bad) { Write-Host ("FAIL " + $b) -ForegroundColor Red }
exit 1
