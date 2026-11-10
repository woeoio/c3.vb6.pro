# check_g4_udt_member_nd.ps1 - 账 §38 的结构性哨兵 (只扫编译器源码, 不起 cl)
#
# 为什么单开这一道: UDT 的**动态数组字段**被下标访问 (`t.Items(i, j)`) 曾经在两个发码站
# 各自写死"实参正好 1 枚"这一条 —— cgen_expr_call_callee_member.inc (obj.arr(i)) 与
# cgen_expr_call_callee_withm.inc (With 块 .arr(i))。秩 2 于是整条落空, 退回通用调用路,
# 发出 `field(0, 0)` ⇒ cl C2064 (term does not evaluate to a function) 以及同站 C2197。
# PhotoDemon 实测: C2064 66 条里 48 条是这一族 (pdGIF 12 + pdXCF 36), 同站 C2197 32 条。
#
# 规则 (改坏了会红, 不是装饰):
#   B1  CCodeGen::udtDynMemberIndex 在 src 全域**只定义一次**, 且在权威文件
#       src/backend/cgen_util_classtype.cpp。
#   B2  哨兵标记 C3-SENTINEL(g4-udt-member-nd) 在 src 全域**恰好 1 处**, 且在 B1 的函数体内。
#   B3  出口形状:
#         a) 秩 1 发 VB6_SA_AT (存量形状: 秩 1 的索引走裸表达式, 不再包 toLongIfVariant);
#         b) 秩 2/3 发 VB6_SA_ND_AT2/AT3, 且数组实参带 `(vb6_SafeArrayND*)` 强转;
#         c) 秩 >=4 发 VB6_SA_ND_ATN, 秩数与全部下标一次交全 (最外层那对括号不是装饰);
#         d) 秩 >=2 的索引过 toLongIfVariant (ND 出口的形参是 int32_t)。
#   B4  调用面: 成员站 1 处、With 站 2 处 (UDT 臂 + 类成员臂), 其余 src 文件 0 处;
#       声明在 src/backend/detail/util/cgen_helpers.inc 恰 1 处。
#   B5  两个站**都不许**再出现写死一维的判据 (`...DynArr && node.positional.size() == 1`)
#       —— 那正是这一族的病根; 有它就是把秩 2 又推回通用调用路。
#
# 用法:  powershell -NoProfile -ExecutionPolicy Bypass -File scripts\check_g4_udt_member_nd.ps1
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

$authRel  = "src\backend\cgen_util_classtype.cpp"
$hdrRel   = "src\backend\detail\util\cgen_helpers.inc"
$memRel   = "src\backend\detail\expr\cgen_expr_call_callee_member.inc"
$withRel  = "src\backend\detail\expr\cgen_expr_call_callee_withm.inc"

foreach ($r in @($authRel, $hdrRel, $memRel, $withRel)) {
    if ($null -eq (Read-Text $r)) {
        Write-Host ("FAIL B0 file missing: " + $r) -ForegroundColor Red
        exit 1
    }
}

$srcDirs = @("src\backend", "src\semantics", "src\ast", "src\driver", "src\common") |
           ForEach-Object { Join-Path $root $_ }
$srcFiles = @()
foreach ($d in $srcDirs) {
    if (-not (Test-Path -LiteralPath $d)) { continue }
    $srcFiles += @(Get-ChildItem -LiteralPath $d -Recurse -File |
                   Where-Object { $_.Extension -in ".cpp", ".inc", ".hpp" })
}

# ---- B1: 定义恰一处, 且在权威文件 ----
$defRx = 'std::string\s+CCodeGen::udtDynMemberIndex\s*\('
$authText = Read-Text $authRel
$defsInAuth = @([regex]::Matches($authText, $defRx)).Count
if ($defsInAuth -ne 1) {
    $bad += ("B1 udtDynMemberIndex is defined " + $defsInAuth + " time(s) in " + $authRel + " (want exactly 1)")
}
$defAll = 0
foreach ($f in $srcFiles) {
    $defAll += @([regex]::Matches([System.IO.File]::ReadAllText($f.FullName), $defRx)).Count
}
if ($defAll -ne 1) {
    $bad += ("B1 udtDynMemberIndex definition sites across src = " + $defAll + " (want 1: one adoption point, not two)")
}

# ---- B2 / B3: 函数体形状 ----
if ($defsInAuth -eq 1) {
    $m = [regex]::Match($authText, $defRx)
    $tail = $authText.Substring($m.Index)
    $endIdx = $tail.IndexOf("`n}")
    $body = if ($endIdx -ge 0) { $tail.Substring(0, $endIdx) } else { $tail }

    $markRx = [regex]::Escape('C3-SENTINEL(g4-udt-member-nd)')
    $markAll = 0
    foreach ($f in $srcFiles) {
        $markAll += @([regex]::Matches([System.IO.File]::ReadAllText($f.FullName), $markRx)).Count
    }
    if ($markAll -ne 1) {
        $bad += ("B2 C3-SENTINEL(g4-udt-member-nd) occurrences across src = " + $markAll + " (want 1, inside the adoption point)")
    }
    if ($body -notmatch $markRx) {
        $bad += "B2 the sentinel marker is not inside the udtDynMemberIndex body (it must pin the adoption point itself)"
    }

    if ($body -notmatch '"VB6_SA_AT\("') {
        $bad += "B3a rank-1 branch no longer emits VB6_SA_AT (existing 1-D shapes must stay byte-identical)"
    }
    if ($body -notmatch 'args\.size\(\)\s*==\s*1') {
        $bad += "B3a the rank-1 legacy guard is gone (args.size() == 1 must keep the bare index expression)"
    }
    if ($body -notmatch '"VB6_SA_ND_AT2\("') {
        $bad += "B3b rank-2 branch is gone (VB6_SA_ND_AT2)"
    }
    if ($body -notmatch '"VB6_SA_ND_AT3\("') {
        $bad += "B3b rank-3 branch is gone (VB6_SA_ND_AT3)"
    }
    if ($body -notmatch '"VB6_SA_ND_ATN\("') {
        $bad += "B3c rank>=4 branch is gone (VB6_SA_ND_ATN)"
    }
    if ($body -notmatch '\(vb6_SafeArrayND\*\)') {
        $bad += "B3b the ND accessors lost the (vb6_SafeArrayND*) cast (dynamic array members are declared vb6_SafeArray1D*)"
    }
    if ($body -notmatch '\(int32_t\)\(' -or $body -notmatch 'const int32_t\[\]') {
        $bad += "B3c VB6_SA_ND_ATN no longer passes all indices as ((const int32_t[]){...})"
    }
    if ($body -notmatch 'toLongIfVariant') {
        $bad += "B3d rank>=2 indices no longer go through toLongIfVariant (ND accessors take int32_t)"
    }
}

# ---- B4: 调用面 ----
$cMem  = @([regex]::Matches((Read-Text $memRel),  'lastExpr_\s*=\s*udtDynMemberIndex\s*\(')).Count
if ($cMem -ne 1) {
    $bad += ("B4 " + $memRel + ": udtDynMemberIndex calls = " + $cMem + " (want 1: obj.arr(i,j) station)")
}
$cWith = @([regex]::Matches((Read-Text $withRel), 'lastExpr_\s*=\s*udtDynMemberIndex\s*\(')).Count
if ($cWith -ne 2) {
    $bad += ("B4 " + $withRel + ": udtDynMemberIndex calls = " + $cWith + " (want 2: With UDT arm + With class-member arm)")
}
$cHdr = @([regex]::Matches((Read-Text $hdrRel), 'udtDynMemberIndex\s*\(')).Count
if ($cHdr -ne 1) {
    $bad += ("B4 " + $hdrRel + ": udtDynMemberIndex declaration = " + $cHdr + " (want exactly 1)")
}
$others = 0
foreach ($f in $srcFiles) {
    if ($f.FullName -eq (Join-Path $root $authRel)) { continue }
    if ($f.FullName -eq (Join-Path $root $memRel))  { continue }
    if ($f.FullName -eq (Join-Path $root $withRel)) { continue }
    if ($f.FullName -eq (Join-Path $root $hdrRel))  { continue }
    $others += @([regex]::Matches([System.IO.File]::ReadAllText($f.FullName), 'udtDynMemberIndex\s*\(')).Count
}
if ($others -ne 0) {
    $bad += ("B4 udtDynMemberIndex referenced " + $others + " time(s) outside the four sanctioned files (a second adoption path)")
}

# ---- B5: 两个站都不许再写死一维 ----
$memText = Read-Text $memRel
$withText = Read-Text $withRel
if ($memText -match 'memIsDynArr\s*&&\s*node\.positional\.size\(\)\s*==\s*1') {
    $bad += "B5 member station still hardcodes the 1-argument dynamic-array branch (rank 2 falls back to a call => C2064)"
}
if ($withText -match 'wIsDynArr\s*&&\s*node\.positional\.size\(\)\s*==\s*1') {
    $bad += "B5 With UDT arm still hardcodes the 1-argument dynamic-array branch (rank 2 => C2064)"
}
if ($withText -match 'wClsIsDynArr\s*&&\s*node\.positional\.size\(\)\s*==\s*1') {
    $bad += "B5 With class-member arm still hardcodes the 1-argument dynamic-array branch (rank 2 => C2064)"
}
if ($memText -notmatch 'memIsDynArr\s*&&\s*!node\.positional\.empty\(\)') {
    $bad += "B5 member station no longer routes every non-empty argument list through udtDynMemberIndex"
}
if ($withText -notmatch 'wIsDynArr\s*&&\s*!node\.positional\.empty\(\)') {
    $bad += "B5 With UDT arm no longer routes every non-empty argument list through udtDynMemberIndex"
}

if ($bad.Count -gt 0) {
    Write-Host "FAIL g4_udt_member_nd" -ForegroundColor Red
    $bad | ForEach-Object { Write-Host ("  " + $_) -ForegroundColor Red }
    exit 1
}
Write-Host ("PASS g4_udt_member_nd (authority=" + $authRel + ", member calls=1, with calls=2)") -ForegroundColor Green
exit 0
