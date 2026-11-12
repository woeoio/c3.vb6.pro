# check_g4_receiver_class_gate.ps1 - 账 G4/M1 的结构性哨兵 (只扫编译器源码, 不起 cl)
#
# 为什么单开这一道: G4 这一刀把"接收者类解析"从 inferClassTypeOfExpr 扩到
# 控件/控件数组 → 工程 UserControl 类 (knownUserControlCtrlVars_), 好让
# `ddWhichFrame(1).AddItem ...` 这种 IndexOrCallExpr 接收者也能拿到**该类自己的**形参表。
# 代价是: 一旦类判错, 本已正确的调用会被改坏。所以这一刀立了一道闸 ——
#   "类感知形参表只在 findClassMemberCallParams 命中时才采用, 答不出逐字回落原兜底",
# 这一道哨兵就是把那句口径变成可判红的句子 (口径说明不算数)。
#
# 规则 (改坏了会红, 不是装饰):
#   A1  CCodeGen::tryAdoptReceiverClassParams 在 src/backend 全域**只定义一次**,
#       且落在权威文件 src/backend/cgen_util_classtype.cpp。
#   A2  采纳点函数体形状:
#         a) 先问 inferCallReceiverClassForParams, 且只问一次;
#         b) 类答不出 (空串) 必须**先** return false —— 这一条在 findClassMemberCallParams 之前;
#         c) findClassMemberCallParams 只用一次, 且 `!...(...)` 未命中就 return false;
#         d) 最后才赋值 outParams / return true (命中才算采纳)。
#   A3  全 src 里 tryAdoptReceiverClassParams 的**调用点恰 2 处** (cgen_call.cpp 的 Fix 254
#       与 callee_params.inc 的 Fix 033/042b), 且每处都在 `if (tryAdoptReceiverClassParams(`
#       守卫之内 —— 不许有绕开闸直接写 calleeParams 的第二条路。
#   A4  两处 class-unaware 兜底**都还在**(闸的"回落"半边不许被删):
#         a) cgen_call.cpp 的 Fix 254: `if (!byClassCS254) { … }` 那一块**里面**仍有
#            lookupModule(maExpr.memberName)（取块的办法 = 花括号配平取 span，不是定宽窗口 ——
#            理由与归因见 A4a 那段注释）;
#         b) callee_params.inc 仍存在 lookupModule(maExpr.memberName) 兜底。
#   A5  哨兵标记 C3-SENTINEL(g4-m1-gate) 在权威文件里恰好 1 处 (A2 的采纳点)。
#
# 用法:  pwsh -File scripts\check_g4_receiver_class_gate.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$bad = @()

function Read-Text([string]$rel) {
    $p = Join-Path $root $rel
    if (-not (Test-Path -LiteralPath $p)) { return $null }
    return [System.IO.File]::ReadAllText($p)
}

# 取一段花括号平衡的块体 (start 指向块首 '{' 的位置)，返回**块内**文本。
# 认字符串字面量：生成器源码里有把 '{' 写进串里的行，不认串就会把深度数错。
function Get-BlockBody([string]$t, [int]$start) {
    if ($start -lt 0 -or $start -ge $t.Length) { return $null }
    $depth = 0
    $inStr = $false
    for ($j = $start; $j -lt $t.Length; $j++) {
        $ch = $t[$j]
        if ($inStr) {
            if ($ch -eq '\') { $j++; continue }
            if ($ch -eq '"') { $inStr = $false }
            continue
        }
        if ($ch -eq '"') { $inStr = $true; continue }
        if ($ch -eq '{') { $depth++ }
        elseif ($ch -eq '}') {
            $depth--
            if ($depth -eq 0) { return $t.Substring($start + 1, $j - $start - 1) }
        }
    }
    return $null
}

$authRel = "src\backend\cgen_util_classtype.cpp"
$callRel = "src\backend\stmt\cgen_call.cpp"
$paramRel = "src\backend\detail\expr\cgen_expr_call_callee_params.inc"

$authText = Read-Text $authRel
if ($null -eq $authText) { Write-Host ("FAIL A0 authority file missing: " + $authRel) -ForegroundColor Red; exit 1 }

# ---- A1: 定义恰一处, 且在权威文件 ----
$defRx = 'bool\s+CCodeGen::tryAdoptReceiverClassParams\s*\('
$defsInAuth = @([regex]::Matches($authText, $defRx)).Count
if ($defsInAuth -ne 1) {
    $bad += ("A1 tryAdoptReceiverClassParams is defined " + $defsInAuth + " time(s) in " + $authRel + " (want exactly 1)")
}
$dirs = @("src\backend", "src\semantics", "src\ast", "src\driver") | ForEach-Object { Join-Path $root $_ }
$srcFiles = @()
foreach ($d in $dirs) {
    if (Test-Path -LiteralPath $d) {
        $srcFiles += Get-ChildItem -LiteralPath $d -Recurse -File |
                     Where-Object { $_.Extension -in ".cpp", ".inc", ".hpp" }
    }
}
$defAll = 0
foreach ($f in $srcFiles) {
    $t = [System.IO.File]::ReadAllText($f.FullName)
    $defAll += @([regex]::Matches($t, $defRx)).Count
}
if ($defAll -ne 1) {
    $bad += ("A1 tryAdoptReceiverClassParams definition sites across src = " + $defAll + " (want 1: the gate has one body, not two)")
}

# ---- A2: 采纳点函数体形状 ----
if ($defsInAuth -eq 1) {
    $m = [regex]::Match($authText, $defRx)
    # 函数体 = 定义行到下一个独占一行的 '}' 为止
    $tail = $authText.Substring($m.Index)
    $endIdx = $tail.IndexOf("`n}")
    $body = if ($endIdx -ge 0) { $tail.Substring(0, $endIdx) } else { $tail }

    $cInfer = @([regex]::Matches($body, 'inferCallReceiverClassForParams\s*\(')).Count
    if ($cInfer -ne 1) { $bad += ("A2a inferCallReceiverClassForParams calls inside the gate body = " + $cInfer + " (want 1)") }

    $emptyGuard = [regex]::Match($body, 'if\s*\(\s*\w+\.empty\(\)\s*\)\s*return\s+false\s*;')
    if (-not $emptyGuard.Success) {
        $bad += "A2b no early `if (<recvCls>.empty()) return false;` inside the gate body (unresolved receiver must fall back)"
    }

    $fcMatch = [regex]::Match($body, 'if\s*\(\s*!\s*findClassMemberCallParams\s*\(')
    if (-not $fcMatch.Success) {
        $bad += "A2c gate body has no `if (!findClassMemberCallParams(...)) return false;` (adoption must require a hit)"
    }
    $cFc = @([regex]::Matches($body, 'findClassMemberCallParams\s*\(')).Count
    if ($cFc -ne 1) { $bad += ("A2c findClassMemberCallParams calls inside the gate body = " + $cFc + " (want 1)") }

    if ($emptyGuard.Success -and $fcMatch.Success -and $emptyGuard.Index -gt $fcMatch.Index) {
        $bad += "A2d the empty-receiver early exit comes AFTER findClassMemberCallParams (order inverted: class must be resolved first)"
    }

    $cTrue = @([regex]::Matches($body, 'return\s+true\s*;')).Count
    if ($cTrue -lt 1) { $bad += "A2e gate body never returns true (nothing can ever be adopted)" }

    # ---- A5 ----
    $cMark = @([regex]::Matches($authText, [regex]::Escape('C3-SENTINEL(g4-m1-gate)'))).Count
    if ($cMark -ne 1) { $bad += ("A5 C3-SENTINEL(g4-m1-gate) occurrences in " + $authRel + " = " + $cMark + " (want 1)") }
}

# ---- A3: 调用点恰 2 处, 且都在 if 守卫里 ----
$callRx = 'if\s*\(\s*tryAdoptReceiverClassParams\s*\('
$callCount = 0
foreach ($rel in @($callRel, $paramRel)) {
    $t = Read-Text $rel
    if ($null -eq $t) { $bad += ("A0 missing file: " + $rel); continue }
    $guarded = @([regex]::Matches($t, $callRx)).Count
    $anyCall = @([regex]::Matches($t, 'tryAdoptReceiverClassParams\s*\(')).Count
    if ($guarded -ne 1) {
        $bad += ("A3 " + $rel + ": guarded `if (tryAdoptReceiverClassParams(` sites = " + $guarded + " (want 1)")
    }
    if ($anyCall -ne $guarded) {
        $bad += ("A3 " + $rel + ": tryAdoptReceiverClassParams mentions = " + $anyCall + " but guarded calls = " + $guarded +
                 " (a call that bypasses the `if` is a bypassed gate)")
    }
    $callCount += $guarded

    # 每个采纳点必须把结果写进 calleeParams (否则闸形同虚设)
    if ($guarded -eq 1) {
        $idx = [regex]::Match($t, $callRx).Index
        $seg = $t.Substring($idx, [Math]::Min(400, $t.Length - $idx))
        if ($seg -notmatch 'calleeParams\s*=\s*std::move\(') {
            $bad += ("A3 " + $rel + ": the guarded gate call does not assign calleeParams (dead gate)")
        }
    }
}
if ($callCount -ne 2) { $bad += ("A3 total guarded gate call sites = " + $callCount + " (want 2: Fix 254 + Fix 033/042b)") }

# ---- A4: 两处 class-unaware 兜底仍在 ----
# A4a 的口径 = 「`if (!byClassCS254) { … }` 那一块**里面**仍有 lookupModule(maExpr.memberName)」。
#   取块的办法从「定宽窗口 `.{0,200}?`」改成「花括号配平取 span」，**判据形状只放宽取窗口的方式，
#   不变量 (class-unaware 兜底不许被删) 一字未动** —— 兜底块里少了 lookupModule(maExpr.memberName)
#   照样红。
#   为什么必须改（归因已定，不重新调查）：§56.29（提交 6f923ac6）把「跨窗体语句调用先问模块限定
#   二级键」那一块（cgen_call.cpp:595-614）插进了兜底块**里面**，于是两个关键点相距 **1338 字符**
#   > 定宽窗口的 200 ⇒ 窗口永远跨不到 ⇒ 判据误红。代码比判据**更严**：兜底仍在（cgen_call.cpp:615
#   `if (!sym) sym = symTab_.lookupModule(maExpr.memberName);` 就在那一块里）。
$callText = Read-Text $callRel
$paramText = Read-Text $paramRel
if ($callText) {
    $gateM = [regex]::Match($callText, 'if\s*\(\s*!\s*byClassCS254\s*\)\s*\{')
    if (-not $gateM.Success) {
        $bad += "A4a Fix 254 fallback gone: `if (!byClassCS254) {` not found in cgen_call.cpp"
    } else {
        $openAt = $callText.IndexOf('{', $gateM.Index)
        $blk = Get-BlockBody $callText $openAt
        if ($null -eq $blk) {
            $bad += "A4a the `if (!byClassCS254) {` block is not brace-balanced in cgen_call.cpp"
        } elseif ($blk -notmatch 'lookupModule\(maExpr\.memberName\)') {
            $bad += "A4a Fix 254 fallback gone: `if (!byClassCS254) { ... lookupModule(maExpr.memberName) ... }` not found in cgen_call.cpp"
        }
    }
}
if ($paramText -and $paramText -notmatch 'lookupModule\(maExpr\.memberName\)') {
    $bad += "A4b callee_params.inc fallback gone: lookupModule(maExpr.memberName) not found (the gate's fall back half was deleted)"
}

if ($bad.Count -gt 0) {
    Write-Host "FAIL g4_receiver_class_gate" -ForegroundColor Red
    $bad | ForEach-Object { Write-Host ("  " + $_) -ForegroundColor Red }
    exit 1
}
Write-Host ("PASS g4_receiver_class_gate (authority=" + $authRel + ", guarded call sites=" + $callCount + ")" ) -ForegroundColor Green
exit 0
