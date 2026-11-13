# check_g4_event_index_forward.ps1 - 账 G4/M2 的结构性哨兵 (只扫编译器源码, 不起 cl)
#
# 为什么单开这一道: 事件转发站曾经把调用写死成 `{ extern void fn(); fn(); }`。
# 控件数组 / 菜单数组的处理器首形参是元素下标 ⇒ 0 实参打 1 形参 = C2198;
# 修好之后还要保证"非数组成员 / 无参处理器"那半边**逐字不变**(否则一条修法换来一片形状改动)。
# 本刀只落两处转发站: 菜单派发 (cgen_form_menu.cpp) 与 Slider 的 Change/Scroll 臂
# (cgen_form_wndproc_dispatch.inc)。两句口径都收进唯一采纳点 CCodeGen::eventIndexCallText。
#
# 规则 (改坏了会红, 不是装饰):
#   B1  CCodeGen::eventIndexCallText 在 src 全域**只定义一次**, 且落在权威文件
#       src/backend/module/cgen_form_menu.cpp。
#   B2  哨兵标记 C3-SENTINEL(g4-m2-event-index) 在 src 全域**恰好 1 处**, 且在 B1 的函数体内。
#   B3  采纳点函数体形状:
#         a) 回落分支仍是**逐字**的 0 参文本 `{ extern void " + fn + "(); " + fn + "(); }`;
#         b) 回落条件同时认 `params.empty()` (无参处理器) 与 `elemIndex < 0` (非数组成员);
#         c) 有 `isByVal` 分支 —— ByVal 直接传值, ByRef 必须借址 (`(&` 形态);
#         d) 形参类型取自 `mapType(p0.type)`, 不许写死 int16_t。
#   B4  调用面: 权威文件 3 次 (1 定义 + 2 调用)、wndproc 片段 2 次、helpers 声明 1 次,
#       其余 src 文件 0 次; 菜单两处必须传 `menuCtrl.index` / `child.index`,
#       且任何调用里都不许出现 `menuId` (它是叶菜单项顺序计数器, 与 VB6 Index 无关)。
#   B5  菜单派发不许再出现写死的 0 参文本 (`extern void " + clickFn + "();`)。
#   B6  Slider 臂不许再出现写死的 0 参文本; 必须保留 UC 跳过门
#       (`findUserControlSpec(ctrl.controlTypeName)`), 且数组元素句柄按
#       `vb6_CtrlArr_GetAt(&vb6_arr_` 取 (与 ScrollBar 臂同一口径, 全文件 2 处)。
#
# 用法:  powershell -NoProfile -ExecutionPolicy Bypass -File scripts\check_g4_event_index_forward.ps1
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

$authRel = "src\backend\module\cgen_form_menu.cpp"
$wndRel  = "src\backend\detail\module\cgen_form_wndproc_dispatch.inc"
$hdrRel  = "src\backend\detail\util\cgen_helpers.inc"

$authText = Read-Text $authRel
if ($null -eq $authText) { Write-Host ("FAIL B0 authority file missing: " + $authRel) -ForegroundColor Red; exit 1 }
$wndText = Read-Text $wndRel
if ($null -eq $wndText) { Write-Host ("FAIL B0 file missing: " + $wndRel) -ForegroundColor Red; exit 1 }
$hdrText = Read-Text $hdrRel
if ($null -eq $hdrText) { Write-Host ("FAIL B0 file missing: " + $hdrRel) -ForegroundColor Red; exit 1 }

# ---- B1: 定义恰一处, 且在权威文件 ----
$defRx = 'std::string\s+CCodeGen::eventIndexCallText\s*\('
$defsInAuth = @([regex]::Matches($authText, $defRx)).Count
if ($defsInAuth -ne 1) {
    $bad += ("B1 eventIndexCallText is defined " + $defsInAuth + " time(s) in " + $authRel + " (want exactly 1)")
}
$defAll = 0
$srcDirs = @("src\backend", "src\semantics", "src\ast", "src\driver", "src\common") |
           ForEach-Object { Join-Path $root $_ }
foreach ($d in $srcDirs) {
    if (-not (Test-Path -LiteralPath $d)) { continue }
    foreach ($f in (Get-ChildItem -LiteralPath $d -Recurse -File |
                    Where-Object { $_.Extension -in ".cpp", ".inc", ".hpp" })) {
        $t = [System.IO.File]::ReadAllText($f.FullName)
        $defAll += @([regex]::Matches($t, $defRx)).Count
    }
}
if ($defAll -ne 1) {
    $bad += ("B1 eventIndexCallText definition sites across src = " + $defAll + " (want 1: one adoption point, not two)")
}

# ---- B2 / B3: 函数体形状 ----
if ($defsInAuth -eq 1) {
    $m = [regex]::Match($authText, $defRx)
    $tail = $authText.Substring($m.Index)
    $endIdx = $tail.IndexOf("`n}")
    $body = if ($endIdx -ge 0) { $tail.Substring(0, $endIdx) } else { $tail }

    $markRx = [regex]::Escape('C3-SENTINEL(g4-m2-event-index)')
    $markAll = 0
    foreach ($d in $srcDirs) {
        if (-not (Test-Path -LiteralPath $d)) { continue }
        foreach ($f in (Get-ChildItem -LiteralPath $d -Recurse -File |
                        Where-Object { $_.Extension -in ".cpp", ".inc", ".hpp" })) {
            $t = [System.IO.File]::ReadAllText($f.FullName)
            $markAll += @([regex]::Matches($t, $markRx)).Count
        }
    }
    if ($markAll -ne 1) {
        $bad += ("B2 C3-SENTINEL(g4-m2-event-index) occurrences across src = " + $markAll + " (want 1, inside the adoption point)")
    }
    if ($body -notmatch $markRx) {
        $bad += "B2 the sentinel marker is not inside the eventIndexCallText body (it must pin the adoption point itself)"
    }

    if ($body -notmatch 'return\s+"\{\s*extern void "\s*\+\s*fn\s*\+\s*"\(\);\s*"\s*\+\s*fn\s*\+\s*"\(\);\s*\}\s*";') {
        $bad += "B3a fallback branch is no longer the verbatim 0-arg text '{ extern void fn(); fn(); }'"
    }
    if ($body -notmatch 'params\.empty\(\)') {
        $bad += "B3b fallback guard lost params.empty() (no-param handlers must stay 0-arg)"
    }
    if ($body -notmatch 'elemIndex\s*<\s*0') {
        $bad += "B3b fallback guard lost 'elemIndex < 0' (non-array members must stay 0-arg)"
    }
    if ($body -notmatch 'isByVal') {
        $bad += "B3c no isByVal branch (ByVal and ByRef event ABIs are not the same call)"
    }
    if ($body -notmatch 'fn\s*\+\s*"\(&"\s*\+\s*var') {
        $bad += "B3c ByRef branch no longer passes an address (fn(&var))"
    }
    if ($body -notmatch 'mapType\(\s*p0\.type\s*\)') {
        $bad += "B3d the element-index type does not come from mapType(p0.type) (hardcoded type = wrong for non-Integer Index)"
    }
    if ($body -match 'int16_t') {
        $bad += "B3d the adoption point hardcodes int16_t (type must follow the handler's first parameter)"
    }
}

# ---- B4: 调用面 ----
$cAuth = @([regex]::Matches($authText, 'eventIndexCallText\s*\(')).Count
if ($cAuth -ne 3) {
    $bad += ("B4 " + $authRel + ": eventIndexCallText mentions = " + $cAuth + " (want 3: 1 definition + 2 menu leaves)")
}
if ($authText -notmatch 'eventIndexCallText\(\s*clickFn\s*,\s*clickSym\s*,\s*menuCtrl\.index\s*\)') {
    $bad += "B4 top-level menu leaf does not pass menuCtrl.index (the element's own VB6 Index)"
}
if ($authText -notmatch 'eventIndexCallText\(\s*clickFn\s*,\s*leafSym\s*,\s*child\.index\s*\)') {
    $bad += "B4 child menu leaf does not pass child.index (the element's own VB6 Index)"
}
if ($authText -match 'eventIndexCallText\s*\([^)]*menuId') {
    $bad += "B4 a menu dispatch passes menuId into eventIndexCallText (menuId is a sequential counter, not the VB6 Index)"
}
$cWnd = @([regex]::Matches($wndText, 'eventIndexCallText\s*\(')).Count
if ($cWnd -ne 2) {
    $bad += ("B4 " + $wndRel + ": eventIndexCallText mentions = " + $cWnd + " (want 2: Slider Change + Scroll)")
}
if ($wndText -notmatch 'eventIndexCallText\(\s*slChangeFn\s*,\s*slChangeSym\s*,\s*slIsArr \? ctrl\.index : -1\s*\)') {
    $bad += "B4 Slider Change arm does not pass the element index (ctrl.index for arrays, -1 otherwise)"
}
if ($wndText -notmatch 'eventIndexCallText\(\s*slScrollFn\s*,\s*slScrollSym\s*,\s*slIsArr \? ctrl\.index : -1\s*\)') {
    $bad += "B4 Slider Scroll arm does not pass the element index (ctrl.index for arrays, -1 otherwise)"
}
if ($hdrText -notmatch 'std::string\s+eventIndexCallText\s*\(\s*const std::string& fn,\s*const Symbol\* handler,\s*int elemIndex\s*\)\s*const\s*;') {
    $bad += ("B4 declaration missing/shape changed in " + $hdrRel + " (one shared signature for all stations)")
}
$others = 0
foreach ($d in $srcDirs) {
    if (-not (Test-Path -LiteralPath $d)) { continue }
    foreach ($f in (Get-ChildItem -LiteralPath $d -Recurse -File |
                    Where-Object { $_.Extension -in ".cpp", ".inc", ".hpp" })) {
        if ($f.FullName -eq (Join-Path $root $authRel)) { continue }
        if ($f.FullName -eq (Join-Path $root $wndRel)) { continue }
        if ($f.FullName -eq (Join-Path $root $hdrRel)) { continue }
        $t = [System.IO.File]::ReadAllText($f.FullName)
        $others += @([regex]::Matches($t, 'eventIndexCallText\s*\(')).Count
    }
}
if ($others -ne 0) {
    $bad += ("B4 eventIndexCallText referenced " + $others + " time(s) outside the three sanctioned files (a second adoption path)")
}

# ---- B5: 菜单派发不再写死 0 参 ----
if ($authText -match 'extern void "\s*\+\s*clickFn\s*\+\s*"\(\)') {
    $bad += "B5 menu dispatch still emits the hardcoded 0-arg call text (bypasses the adoption point)"
}

# ---- B6: Slider 臂 ----
# 只看 Slider 那一臂 (同一文件里还有 ScrollBar 等别的臂, 各自有自己的句柄取法)
$slAnchor = 'if (ctrl.controlType != FrmControlType::Slider) continue;'
$slPos = $wndText.IndexOf($slAnchor)
if ($slPos -lt 0) {
    $bad += "B6 the Slider arm is gone (no 'if (ctrl.controlType != FrmControlType::Slider) continue;')"
} else {
    $slArm = $wndText.Substring($slPos, [Math]::Min(2000, $wndText.Length - $slPos))
    $cArrGet = @([regex]::Matches($slArm, [regex]::Escape('vb6_CtrlArr_GetAt(&vb6_arr_'))).Count
    if ($cArrGet -ne 1) {
        $bad += ("B6 control-array handle lookup inside the Slider arm = " + $cArrGet + " (want 1: array elements have no vb6_hwnd_<name>)")
    }
    if (@([regex]::Matches($slArm, 'extern void "\s*\+\s*sl(Change|Scroll)Fn\s*\+\s*"\(\)')).Count -gt 0) {
        $bad += "B6 Slider arm still emits the hardcoded 0-arg call text (bypasses the adoption point)"
    }
    if ($slArm -notmatch 'if\s*\(\s*findUserControlSpec\(ctrl\.controlTypeName\)\s*\)\s*continue\s*;') {
        $bad += "B6 Slider arm lost the UserControl skip gate (a project UC whose name contains 'slider' is dispatched by its sink thunk; a second dispatch is both C2198 and a double fire)"
    }
}

if ($bad.Count -gt 0) {
    Write-Host "FAIL g4_event_index_forward" -ForegroundColor Red
    $bad | ForEach-Object { Write-Host ("  " + $_) -ForegroundColor Red }
    exit 1
}
Write-Host ("PASS g4_event_index_forward (authority=" + $authRel + ", menu calls=2, slider calls=2)") -ForegroundColor Green
exit 0
