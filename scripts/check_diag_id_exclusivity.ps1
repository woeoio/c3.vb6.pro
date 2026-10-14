# check_diag_id_exclusivity.ps1 - census: VB3001 只报「名字解析不出来」，Implements 那两条有自己的号（账 #278 / §B102 ③）
#
# 起因（2026-10-09）：要谈「Option Explicit 在场时把未声明标识符升 error」，先得让这个号**只表示一件事**。
# 实测 VB3001 被四种互不相干的意思共用（未声明标识符 / For 循环变量 / Implements 的接口与槽 / Gosub 标签 /
# Delegate 绑定目标 / 未使用的模块级符号），而严重级是在**调用点**选的 —— 只要 ID 还共用，
# 任何按号做的动作（升严重级、--suppress-warning 3001 那条抑制通道、按号数条的 census）都会连带误伤。
# 本批把 Implements 那一族分家出去（接口找不到 = VB3044，缺槽 = 沿用既有的 VB3012），
# 这一道钉的就是「分家之后不许再借回去」。
#
# 规则（改坏了会红，不是装饰）：
#   S1  src/ 里 `DiagnosticID::SemUndeclaredIdentifier` 的使用点恰好 7 处，且逐处点名 ——
#       多了 = 有新地方借这个号；少了 = 有报点被删而这里没跟着改（census 会跟着骗人）。
#   S2  semantic_analyzer.cpp（Implements 校验住的那个文件）里这个号**一次都不许出现**。
#   S3  接口找不到那条用自己的号：`SemImplementsInterfaceNotFound` 恰好 1 处声明 + 1 处使用。
#   S4  缺槽那条沿用 `SemInterfaceNotImplemented`：semantic_analyzer.cpp 里恰好 1 处。
#
# 用法: powershell -File scripts\check_diag_id_exclusivity.ps1   (PASS = exit 0)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$srcDir = Join-Path $root "src"
$diagHpp = Join-Path $srcDir "common\diagnostics.hpp"
$analyzer = Join-Path $srcDir "semantics\semantic_analyzer.cpp"

foreach ($p in @($srcDir, $diagHpp, $analyzer)) {
    if (-not (Test-Path -LiteralPath $p)) { Write-Host ("FAIL missing path: " + $p); exit 1 }
}

$bad = @()

# ---- 收集 src/ 里这三个号的全部出现点（声明在 diagnostics.hpp，使用在别处）----
$occ = @()
$nameRx = [regex]'(SemUndeclaredIdentifier|SemImplementsInterfaceNotFound|SemInterfaceNotImplemented)'
foreach ($f in (Get-ChildItem -Path $srcDir -File -Recurse -Include *.cpp, *.hpp)) {
    $isDeclFile = ($f.FullName -ieq $diagHpp)
    $n = 0
    foreach ($ln in (Get-Content -LiteralPath $f.FullName -Encoding UTF8)) {
        $n++
        foreach ($m in $nameRx.Matches($ln)) {
            $occ += [pscustomobject]@{
                Id = $m.Groups[1].Value
                Where = ($f.FullName.Substring($root.Length + 1) + ":" + $n)
                Kind = $(if ($isDeclFile) { 'decl' } elseif ($ln -match 'DiagnosticID::') { 'use' } else { 'other' })
            }
        }
    }
}

$undecl = @($occ | Where-Object { $_.Id -eq 'SemUndeclaredIdentifier' -and $_.Kind -eq 'use' })
$impl44 = @($occ | Where-Object { $_.Id -eq 'SemImplementsInterfaceNotFound' })
$slot12 = @($occ | Where-Object { $_.Id -eq 'SemInterfaceNotImplemented' })

# ---- S1: 恰好 7 处使用点 ----
if ($undecl.Count -ne 7) {
    $bad += ("S1 SemUndeclaredIdentifier has " + $undecl.Count + " use sites (want exactly 7): " +
             (($undecl | ForEach-Object { $_.Where }) -join ", ") +
             " - this id must mean one thing: a name that resolves to nothing")
}

# ---- S2: Implements 那个文件里一次都不许借这个号 ----
$inAnalyzer = @($undecl | Where-Object { $_.Where -like "*semantic_analyzer.cpp:*" })
if ($inAnalyzer.Count -ne 0) {
    $bad += ("S2 Implements checks live in semantic_analyzer.cpp but it borrows SemUndeclaredIdentifier " +
             $inAnalyzer.Count + " time(s): " + (($inAnalyzer | ForEach-Object { $_.Where }) -join ", ") +
             " - escalate-by-id or suppress-by-id would hit Implements too")
}

# ---- S3: 新号 = 1 处声明 + 1 处使用 ----
$s3decl = @($impl44 | Where-Object { $_.Kind -eq 'decl' })
$s3use = @($impl44 | Where-Object { $_.Kind -eq 'use' })
if ($s3decl.Count -ne 1 -or $s3use.Count -ne 1) {
    $bad += ("S3 SemImplementsInterfaceNotFound declared " + $s3decl.Count + " (want 1) / used " +
             $s3use.Count + " (want 1): " + (($impl44 | ForEach-Object { $_.Where }) -join ", "))
}

# ---- S4: 缺槽那条走 VB3012，且在 legacy 路径里恰好一处 ----
$s4use = @($slot12 | Where-Object { $_.Kind -eq 'use' })
$s4analyzer = @($s4use | Where-Object { $_.Where -like "*semantic_analyzer.cpp:*" })
if ($s4analyzer.Count -ne 1) {
    $bad += ("S4 legacy missing-slot Implements diagnostic should use SemInterfaceNotImplemented " +
             "exactly once in semantic_analyzer.cpp, found " + $s4analyzer.Count +
             " (all uses: " + (($s4use | ForEach-Object { $_.Where }) -join ", ") + ")")
}

if ($bad.Count -gt 0) {
    foreach ($b in $bad) { Write-Host ("FAIL " + $b) }
    exit 1
}

Write-Host ("PASS diag_id_exclusivity (S1 7 sites, S2 0 in the Implements file, S3 1+1, S4 1; " +
            "VB3012 family uses=" + $s4use.Count + ")")
exit 0
