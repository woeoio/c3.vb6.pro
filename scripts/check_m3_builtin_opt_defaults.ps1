# check_m3_builtin_opt_defaults.ps1 - M3 (§39) 的结构性哨兵 (只扫源码, 不起 cl)
#
# 为什么单开这一道: `MonthName(month[, abbreviate])` 与 `WeekdayName(weekday[, abbreviate[,
# firstdayofweek]])` 的省略档**只有一个补齐点** (cgen_expr_call_builtin_fixup.inc 的
# kBuiltinOptDefaults 表)。这张表的档位必须与 RTL 出口的 C 形参表逐档对齐:
#   · 表里少一档 ⇒ `vb6_MonthName(Month(Now))` 1 实参进 2 形参 = cl C2198 (PhotoDemon 7 处);
#   · 表里多一档 ⇒ 多发一枚实参进 2 参签名 = cl C2197;
#   · 第一档 (month/weekday) 是**必填档**: 写成可补值就把"少写实参"从编译错误变成静默错值。
# 故本哨兵钉住四件事:
#   S1  补齐点唯一 (C3-SENTINEL(m3-builtin-opt-defaults) 恰 1 处, 且在该文件)。
#   S2  表里两家各登记一次, 指向各自的 kOD 数组。
#   S3  两份 kOD 数组的档数 = 2 / 3, 且**首档是 nullptr** (必填档不可补)。
#   S4  RTL 声明仍是 2 / 3 参 —— 表与 RTL 双向对账, 任一侧改形即红。
#
# 用法:  powershell -NoProfile -ExecutionPolicy Bypass -File scripts\check_m3_builtin_opt_defaults.ps1
#        -SrcRoot <dir> 可指向一棵只含 src/ 的副本树 (负测试用; 默认 = 仓库根)
# 退出码: 0 = 全绿; 1 = 红

param(
    [string]$SrcRoot = ""
)

$ErrorActionPreference = "Stop"
$root = if ($SrcRoot -ne "") { $SrcRoot } else { Split-Path -Parent $PSScriptRoot }
$bad = @()

function Read-Text([string]$rel) {
    $p = Join-Path $root $rel
    if (-not (Test-Path -LiteralPath $p)) { return $null }
    return [System.IO.File]::ReadAllText($p)
}

$fixRel = "src\backend\detail\expr\cgen_expr_call_builtin_fixup.inc"
$fixText = Read-Text $fixRel
if ($null -eq $fixText) {
    Write-Host ("FAIL S0 file missing: " + $fixRel) -ForegroundColor Red
    exit 1
}

# ---- S1: 补齐点唯一 ----
$markRx = [regex]::Escape('C3-SENTINEL(m3-builtin-opt-defaults)')
$markCount = @([regex]::Matches($fixText, $markRx)).Count
if ($markCount -ne 1) {
    $bad += ("S1 C3-SENTINEL(m3-builtin-opt-defaults) occurrences in " + $fixRel + " = " + $markCount + " (want 1)")
}
$srcDirs = @("src\backend", "src\semantics", "src\parser", "src\common") |
           ForEach-Object { Join-Path $root $_ }
$markAll = 0
foreach ($d in $srcDirs) {
    if (-not (Test-Path -LiteralPath $d)) { continue }
    foreach ($f in (Get-ChildItem -LiteralPath $d -Recurse -File |
                    Where-Object { $_.Extension -in ".cpp", ".inc", ".hpp" })) {
        $markAll += @([regex]::Matches([System.IO.File]::ReadAllText($f.FullName), $markRx)).Count
    }
}
if ($markAll -ne 1) {
    $bad += ("S1 sentinel occurrences across src = " + $markAll + " (want 1: one padding point)")
}

# ---- S2: 表里两家各登记一次, 且指向各自的 kOD 数组 ----
foreach ($pair in @(@("vb6_MonthName", "kOD_MonthName"), @("vb6_WeekdayName", "kOD_WeekdayName"))) {
    $callee = $pair[0]; $arr = $pair[1]
    $rowRx = [regex]("\{\s*""" + [regex]::Escape($callee) + """\s*,\s*" + [regex]::Escape($arr) + "\s*,")
    $n = @([regex]::Matches($fixText, $rowRx)).Count
    if ($n -ne 1) {
        $bad += ("S2 kBuiltinOptDefaults row for " + $callee + " -> " + $arr + " = " + $n + " (want 1)")
    }
}

# ---- S3: 档数 2 / 3, 首档 nullptr ----
$slotChecks = @(
    @{ Arr = "kOD_MonthName";   Want = 2 },
    @{ Arr = "kOD_WeekdayName"; Want = 3 }
)
foreach ($sc in $slotChecks) {
    $arrRx = [regex]("static\s+const\s+char\*\s+const\s+" + [regex]::Escape($sc.Arr) + "\[\]\s*=\s*\{([^}]*)\}")
    $m = $arrRx.Match($fixText)
    if (-not $m.Success) {
        $bad += ("S3 " + $sc.Arr + " definition not found")
        continue
    }
    $slots = @($m.Groups[1].Value -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })
    if ($slots.Count -ne $sc.Want) {
        $bad += ("S3 " + $sc.Arr + " slot count = " + $slots.Count + " (want " + $sc.Want + ", must match the RTL C signature)")
    }
    if ($slots.Count -ge 1 -and $slots[0] -ne "nullptr") {
        $bad += ("S3 " + $sc.Arr + " first slot = '" + $slots[0] + "' (must stay nullptr: the required slot is never padded)")
    }
}

# ---- S4: 与 RTL 出口双向对账 ----
$rtlRel = "src\rtl\core\vb6rtl\vb6rtl_builtin.h"
$rtlText = Read-Text $rtlRel
if ($null -eq $rtlText) {
    $bad += ("S4 file missing: " + $rtlRel)
} else {
    $rtlChecks = @(
        @{ Name = "vb6_MonthName";   Want = 2 },
        @{ Name = "vb6_WeekdayName"; Want = 3 }
    )
    foreach ($rc in $rtlChecks) {
        $declRx = [regex]("BSTR\s+" + [regex]::Escape($rc.Name) + "\s*\(([^)]*)\)\s*;")
        $m = $declRx.Match($rtlText)
        if (-not $m.Success) {
            $bad += ("S4 RTL declaration for " + $rc.Name + " not found in " + $rtlRel)
            continue
        }
        $params = @($m.Groups[1].Value -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })
        if ($params.Count -ne $rc.Want) {
            $bad += ("S4 RTL " + $rc.Name + " arity = " + $params.Count + " (want " + $rc.Want + "); table slots and C signature diverged")
        }
    }
}

if ($bad.Count -gt 0) {
    Write-Host "FAIL m3_builtin_opt_defaults" -ForegroundColor Red
    $bad | ForEach-Object { Write-Host ("  " + $_) -ForegroundColor Red }
    exit 1
}
Write-Host ("PASS m3_builtin_opt_defaults (authority=" + $fixRel + ", rtl=" + $rtlRel + ")") -ForegroundColor Green
exit 0
