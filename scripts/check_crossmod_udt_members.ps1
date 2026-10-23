# check_crossmod_udt_members.ps1 - Type 的成员表: 作者只许一个, 跨模块只许三处转抄, 换人只许一个方向
#
# 起因 (2026-10-22, 账 #305 第一百七十三刀 / 新账 #468): 一份 Type 声明在 A 模块、被 B 模块
# 用作变量类型时, B 看到的成员表是**空的** —— 不是查不到符号 (lookupModule 命中), 而是那张
# 按 storageKey 去重的工程级代表表 (globalPublicSyms) 先到先得, 而消费者自己解析 `As <跨模块
# Type>` 时先立的那枚**无成员占位符号**也进 getPublicSymbols(), 模块遍历顺序又不固定 (§B97),
# 于是占位版常当上代表 ⇒ 下面那条 Fix 047 的回填闸门 (`localSym 空 && srcSym 非空 → 抄`)
# 拿不到东西可抄, 每个消费者只抄到一枚空表。后果: 成员数组的下标没人折算秩,
# `me->Location.Coord(1)` 原样发进 C = cl C2064 (语料 Ray.vbp 真编 111 行 C2064 / 95 站里 87 站)。
#
# 规则 (改坏了会红, 不是装饰):
#   U1  成员表的作者全仓恰好 1 处 (push_back), 且它住在语义层那一格里。
#       发码侧再多一份"自己数 AST 成员"的表 = 第二份真相, 本条要红。
#   U2  跨模块转抄恰好 3 处 (赋值), 且全在 driver_crossmod.cpp 一个文件里; 别处再抄一份要红。
#   U3  代表表换人只许 空→非空 这一个方向: 那个形状恰好 1 处; 反向 (非空→空) 恒 0 处
#       —— 反向一旦出现, 空占位又能当上代表, 本账原样复发。
#   U4  Fix 047 那条回填闸门 (localSym 空 && srcSym 非空) 恰好 1 处。它从本刀起才有活干,
#       不许被改成"双向都比"或删掉。
#
# 用法: powershell -File scripts\check_crossmod_udt_members.ps1 [-Root <副本目录>]  (PASS = exit 0)

param([string]$Root = "")

$ErrorActionPreference = "Stop"
if ($Root -eq "") { $Root = Split-Path -Parent $PSScriptRoot }

$crossmod = Join-Path $Root "src\driver\driver_crossmod.cpp"
$authored = Join-Path $Root "src\semantics\semantic_analyzer_decl_type.cpp"
foreach ($p in @($crossmod, $authored)) {
    if (-not (Test-Path -LiteralPath $p)) { Write-Host ("FAIL missing path: " + $p); exit 1 }
}

function Count-Match([string]$text, [string]$pattern) {
    return ([regex]::Matches($text, $pattern)).Count
}

$allFiles = Get-ChildItem -LiteralPath (Join-Path $Root "src") -Recurse -File |
    Where-Object { $_.Extension -in @(".cpp", ".inc", ".h") }

$authorTotal = 0
$authorFiles = @()
$copyTotal = 0
$copyFiles = @()
foreach ($f in $allFiles) {
    $t = [IO.File]::ReadAllText($f.FullName)
    $a = Count-Match $t "udtMembers\.push_back"
    $c = Count-Match $t "udtMembers\s*=\s*[^=]"
    if ($a -gt 0) { $authorTotal += $a; $authorFiles += $f.Name }
    if ($c -gt 0) { $copyTotal += $c; $copyFiles += $f.Name }
}

$crossText = [IO.File]::ReadAllText($crossmod)
$authText = [IO.File]::ReadAllText($authored)
$fwd = Count-Match $crossText "existing->udtMembers\.empty\(\)\s*&&\s*!sym->udtMembers\.empty\(\)"
$rev = Count-Match $crossText "!existing->udtMembers\.empty\(\)\s*&&\s*sym->udtMembers\.empty\(\)"
$backfill = Count-Match $crossText "localSym->udtMembers\.empty\(\)\s*&&\s*!srcSym->udtMembers\.empty\(\)"

$fail = @()
if ($authorTotal -ne 1) { $fail += ("U1 作者数=" + $authorTotal + " (要恰好 1; 住 " + ($authorFiles -join ",") + ")") }
if ($authorFiles.Count -ne 1 -or $authorFiles[0] -ne "semantic_analyzer_decl_type.cpp") {
    $fail += ("U1 作者不在语义层那一格: " + ($authorFiles -join ","))
}
if ((Count-Match $authText "udtMembers\.push_back") -ne 1) {
    $fail += "U1 语义层那一格里作者不是恰好 1 处"
}
if ($copyTotal -ne 3) { $fail += ("U2 转抄数=" + $copyTotal + " (要恰好 3; 住 " + ($copyFiles -join ",") + ")") }
if ($copyFiles.Count -ne 1 -or $copyFiles[0] -ne "driver_crossmod.cpp") {
    $fail += ("U2 转抄不在一个文件里: " + ($copyFiles -join ","))
}
if ($fwd -ne 1) { $fail += ("U3 空→非空那一支=" + $fwd + " 处 (要恰好 1)") }
if ($rev -ne 0) { $fail += ("U3 反向 非空→空=" + $rev + " 处 (要 0 —— 空占位不许再当代表)") }
if ($backfill -ne 1) { $fail += ("U4 回填闸门=" + $backfill + " 处 (要恰好 1)") }

if ($fail.Count -gt 0) {
    foreach ($m in $fail) { Write-Host ("FAIL " + $m) }
    exit 1
}
Write-Host ("PASS 作者=1(semantic_analyzer_decl_type.cpp) 转抄=3(driver_crossmod.cpp) 换人=1向(空→非空) 反向=0 回填=1")
