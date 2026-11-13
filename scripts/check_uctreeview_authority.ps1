# ucTreeView 家族静态哨兵 —— 事件分派 / 模态出口 / TreeView 槽位 / 多窗体 (只扫源码)
#
# 为什么单开这一道 (ucTreeView 夹具实测, 2026-11-06):
#   用户实测四条: 展开丢 icon、子项不能独立切换、三点按钮不弹色选、XML 加载无效果、
#   五个按钮不弹窗体。逐条取证后全部落在**四条"唯一出口"没被守住**:
#     ① 事件 sink 未挂/处理器未接时 RaiseEvent 整条静默吞掉 —— 与"点了没反应"同现象;
#     ② Form.Show / 模态窗口各站点手拼 —— 模态框会阻塞宿主, 没有唯一收口就没人能关;
#     ③ `With ... .Add(...).Image = 1` 的目标地址被求值两次 (多插一行); `Is` 比较
#        COM 结果未拆身份 (单选永远全 false); 省略 Optional Variant 实参填 Empty 而非
#        Missing (icon 被当成"有 SelectedImage"画丢);
#     ④ 控件数组容器里的子控件 WM_COMMAND 不转发 —— 三点按钮的 Click 永远到不了窗体。
#
# 规则 (改坏了会红, 不是装饰):
#   T1  RaiseEvent 的分发必须有可观测 miss: cgen 只许在 cgen_com_events.cpp 一处
#       发 `if (me->events && me->events->onX)`, 且紧跟 `else { vb6_UCEventMiss(...) }`;
#       vb6_UCEventMiss 在 src/rtl 恰好定义一次, 体内必须 `vb6_ucEventMiss++` +
#       C3_OCX_TRACE 分支。禁止"事件名解析失败就静默返回"。
#   T2  模态/对话框唯一出口: vb6_ShowForm 在 src/rtl 恰好定义一次 (vb6forms.c);
#       ShowWindow(SW_SHOWDEFAULT) 全仓恰好 1 站, 且在同一文件; ChooseColorA/W 在
#       src/rtl 只许出现在 DI 桩 (vb6_di_*.c); src/backend 不许自己发 ShowWindow。
#   T3  TreeView 槽位唯一出口 (Node.Image / SelectedImage / Expanded / Selected /
#       Nodes.Add 的 With 链): (a) Optional As Variant 省略位必须发 vb6_VariantMissing;
#       (b) `Is` 两侧 COM 结果必须过身份拆箱 (vb6_ComResultIdentity / vb6_ComObjectUnbox);
#       (c) Variant 槽赋值的目标地址必须提成临时量 (`_vtg_`), 目标只求值一次。
#   T4  多窗体统一入口: Form.Show 只走 vb6_form_show_<Form> -> vb6_ShowForm;
#       控件数组容器里的子控件 WM_COMMAND 必须逐元素转发 (vb6_fw 循环),
#       禁止对控件数组容器整个跳过。
#
# 白名单 (带理由, 改动时同步复查):
#   W1  tests\ucTreeView\Form1.frm 里的 `Private Declare Function ChooseColorA ...`
#       是**夹具用户代码**的 Declare, 不归 src/rtl 管; 本哨兵只扫 src/rtl, 天然不命中。
#       (用户实测要的就是它弹出 —— 生成侧只需保证 WM_COMMAND 到得了 Form1。)
#   W2  src\rtl\core\vb6forms\vb6forms_ctrl.c 的 vb6_CdShowColor: 色选对话框在 RTL 的
#       **唯一出口** (comdlg32!ChooseColorW + CHOOSECOLORW 组装只此一站)。放行, 但断言
#       它仍在; 别处再出现一次才算散点。
#   W3  无其他白名单 —— 当前四类散点均为 0。
#
# 用法:  powershell -NoProfile -File scripts\check_uctreeview_authority.ps1 [-Root <repo>]
# 退出码: 0 = 全绿; 1 = 红

param([string]$Root = "")

$ErrorActionPreference = "Stop"
if ($Root -ne "") { $root = (Resolve-Path -LiteralPath $Root).Path }
else { $root = Split-Path -Parent $PSScriptRoot }
$bad = @()

function Read-Text([string]$rel) {
    $p = Join-Path $root $rel
    if (-not (Test-Path -LiteralPath $p)) { return "" }
    return [System.IO.File]::ReadAllText($p)
}

function Get-SrcLines([string]$subdir) {
    $dir = Join-Path $root $subdir
    if (-not (Test-Path -LiteralPath $dir)) { return @() }
    # 单次 List 追加; 早前用 `+=` 是 O(n²) —— 整仓几万行会把哨兵拖成"挂着不动"。
    $out = New-Object System.Collections.Generic.List[object]
    foreach ($f in (Get-ChildItem -LiteralPath $dir -Recurse -File | Where-Object { $_.Extension -in ".c", ".h", ".inc", ".cpp", ".hpp" })) {
        $ln = 0
        foreach ($line in ([System.IO.File]::ReadAllLines($f.FullName))) {
            $ln++
            $t = $line.Trim()
            if ($t.StartsWith("//") -or $t.StartsWith("/*") -or $t.StartsWith("*") -or $t.StartsWith("*/")) { continue }
            $out.Add([pscustomobject]@{ File = $f.Name; Path = $f.FullName; Line = $ln; Text = $t })
        }
    }
    return $out
}

$rtl = @(Get-SrcLines "src\rtl")
$backend = @(Get-SrcLines "src\backend")

# ---- T1: RaiseEvent 的 miss 必须可观测, 且只有一处分发发码 ----
$evtSink = @($backend | Where-Object { $_.Text -match 'me->events\s*&&\s*me->events->' })
if ($evtSink.Count -ne 1 -or $evtSink[0].File -ne "cgen_com_events.cpp") {
    $bad += ("T1 RaiseEvent sink dispatch emitted " + $evtSink.Count + " times (" +
             (($evtSink | ForEach-Object { $_.File + ":" + $_.Line }) -join " | ") +
             ") want exactly 1 in cgen_com_events.cpp")
}
$evtCpp = Read-Text "src\backend\module\cgen_com_events.cpp"
if ($evtCpp -notmatch 'emitLine\("} else \{"\);[\s\S]{0,600}vb6_UCEventMiss\(') {
    $bad += "T1 RaiseEvent no longer emits an else-branch vb6_UCEventMiss (事件又被静默吞掉)"
}
$missDef = @($rtl | Where-Object { $_.Text -match 'void\s*vb6_UCEventMiss\s*\(' -and $_.Text -notmatch '\)\s*;' })
if ($missDef.Count -ne 1) {
    $bad += ("T1 vb6_UCEventMiss defined " + $missDef.Count + " times in src/rtl (want exactly 1)")
} else {
    $full = [System.IO.File]::ReadAllText($missDef[0].Path)
    $blk = [regex]::Match($full, 'void\s*vb6_UCEventMiss\s*\([^)]*\)\s*\{[\s\S]{0,500}?\r?\n\}')
    if (-not $blk.Success) { $bad += "T1 vb6_UCEventMiss body not found" }
    else {
        if ($blk.Value -notmatch 'vb6_ucEventMiss\+\+') { $bad += "T1 vb6_UCEventMiss does not increment the miss counter" }
        if ($blk.Value -notmatch 'C3_OCX_TRACE') { $bad += "T1 vb6_UCEventMiss has no C3_OCX_TRACE log gate" }
    }
}

# ---- T2: 模态/对话框唯一出口 ----
$showDef = @($rtl | Where-Object { $_.Text -match 'void\s*vb6_ShowForm\s*\(' -and $_.Text -notmatch '\)\s*;' })
if ($showDef.Count -ne 1 -or $showDef[0].File -ne "vb6forms.c") {
    $bad += ("T2 vb6_ShowForm defined " + $showDef.Count + " times (" +
             (($showDef | ForEach-Object { $_.File + ":" + $_.Line }) -join " | ") + ") want exactly 1 in vb6forms.c")
}
$showDefault = @($rtl | Where-Object { $_.Text -match 'ShowWindow\s*\([^;]*SW_SHOWDEFAULT' })
if ($showDefault.Count -ne 1 -or $showDefault[0].File -ne "vb6forms.c") {
    $bad += ("T2 ShowWindow(SW_SHOWDEFAULT) sites = " + $showDefault.Count + " (" +
             (($showDefault | ForEach-Object { $_.File + ":" + $_.Line }) -join " | ") + ") want exactly 1 (vb6_ShowForm)")
}
$ccSites = @($rtl | Where-Object { $_.Text -match 'ChooseColorA|ChooseColorW' })
# W1: vb6forms_ctrl.c 的 vb6_CdShowColor 是 ColorDialog/CommonDialog.ShowColor 的**唯一出口**
#     (LoadLibrary comdlg32 + GetProcAddress ChooseColorW + CHOOSECOLORW 组装只此一站)。
#     只有"别处再手拼一枚 ChooseColor"才算散点 ⇒ 这里放行该文件, 并断言它确实只有这一站。
$ccBad = @($ccSites | Where-Object { $_.File -notmatch '^vb6_di_' -and $_.File -ne 'vb6forms_ctrl.c' })
$ccOk = @($ccSites | Where-Object { $_.File -eq 'vb6forms_ctrl.c' })
if ($ccOk.Count -lt 1) { $bad += "T2 vb6forms_ctrl.c 的 ChooseColorW 唯一出口不见了 (色选对话框没有 RTL 落点)" }
if ($ccBad.Count -gt 0) {
    $bad += ("T2 naked ChooseColor call at " + (($ccBad | ForEach-Object { $_.File + ":" + $_.Line }) -join " | ") +
             " -> W1 只放行 vb6forms_ctrl.c 的唯一出口 + DI 桩; 新增站点必须并进那一站")
}
$beShowWin = @($backend | Where-Object { $_.Text -match 'ShowWindow\s*\(' })
if ($beShowWin.Count -gt 0) {
    $bad += ("T2 cgen emits ShowWindow directly at " +
             (($beShowWin | ForEach-Object { $_.File + ":" + $_.Line }) -join " | ") +
             " -> Form 显示只许经 vb6_ShowForm")
}

# ---- T3: TreeView 槽位三件套 (Missing / Is 身份 / 目标只求值一次) ----
$missingDef = @($rtl | Where-Object { $_.Text -match 'vb6_VARIANT\s*vb6_VariantMissing\s*\(' -and $_.Text -notmatch '\)\s*;' })
if ($missingDef.Count -ne 1) {
    $bad += ("T3 vb6_VariantMissing defined " + $missingDef.Count + " times in src/rtl (want exactly 1)")
} else {
    $full = [System.IO.File]::ReadAllText($missingDef[0].Path)
    $blk = [regex]::Match($full, 'vb6_VariantMissing\s*\([^)]*\)\s*\{[\s\S]{0,400}?\r?\n\}')
    if (-not $blk.Success) { $bad += "T3 vb6_VariantMissing body not found" }
    else {
        if ($blk.Value -notmatch 'vb6_vtError') { $bad += "T3 vb6_VariantMissing is not VT_ERROR (VarType=vbError)" }
        if ($blk.Value -notmatch '0x80020004') { $bad += "T3 vb6_VariantMissing lacks DISP_E_PARAMNOTFOUND" }
    }
}
$naming = Read-Text "src\backend\cgen_base_naming.cpp"
if ($naming -notmatch 'p\.type == Vb6Type::Variant && p\.isOptional && !p\.hasDefaultValue') {
    $bad += "T3 defaultValueForParam no longer routes omitted Optional Variant to vb6_VariantMissing"
} elseif ($naming -notmatch 'return "vb6_VariantMissing\(\)";') {
    $bad += "T3 defaultValueForParam branch exists but does not return vb6_VariantMissing()"
}
$identDef = @($rtl | Where-Object { $_.Text -match 'void\s*\*\s*vb6_ComResultIdentity\s*\(' -and $_.Text -notmatch '\)\s*;' })
$unboxDef = @($rtl | Where-Object { $_.Text -match 'void\s*\*\s*vb6_ComObjectUnbox\s*\(' -and $_.Text -notmatch '\)\s*;' })
if ($identDef.Count -ne 1) { $bad += ("T3 vb6_ComResultIdentity defined " + $identDef.Count + " times (want 1)") }
if ($unboxDef.Count -ne 1) { $bad += ("T3 vb6_ComObjectUnbox defined " + $unboxDef.Count + " times (want 1)") }
$bin = Read-Text "src\backend\expr\cgen_expr_binary.cpp"
if ($bin -notmatch '"vb6_ComResultIdentity\("') { $bad += "T3 Is operands no longer unwrap vb6_ComCall results (单选身份又会被比对成 false)" }
if ($bin -notmatch '"vb6_ComObjectUnbox\("') { $bad += "T3 Is operands no longer unbox vb6_ComGetObjectProp results" }
$util = Read-Text "src\backend\cgen_util_ctrl.cpp"
if ($util -notmatch 'emitVariantSlotStore') { $bad += "T3 emitVariantSlotStore vanished" }
else {
    $fn = [regex]::Match($util, 'void\s*CCodeGen::emitVariantSlotStore[\s\S]{0,2200}?\r?\n\}')
    if (-not $fn.Success) { $bad += "T3 emitVariantSlotStore body not found" }
    else {
        if ($fn.Value -notmatch 'needsAddrTemp') { $bad += "T3 emitVariantSlotStore lacks the side-effecting-target test" }
        if ($fn.Value -notmatch '_vtg_') { $bad += "T3 emitVariantSlotStore no longer hoists the target address (With 链会多插一行)" }
    }
}

# ---- T4: 多窗体入口 + 控件数组容器命令转发 ----
$frame = Read-Text "src\backend\detail\module\cgen_form_frame_menu.inc"
if ($frame -notmatch 'vb6_ForwardChildCommands') { $bad += "T4 ForwardChildCommands emission vanished" }
elseif ($frame -notmatch 'vb6_CtrlArr_GetAt\(&vb6_arr_" \+ cIdent\(ctrl\.controlName\) \+ ", vb6_fw\)') {
    $bad += "T4 control-array container forwarding loop missing (数组容器里的子按钮 Click 又会被吞)"
} elseif ($frame -notmatch 'if \(vb6_fw_hwnd\) vb6_ForwardChildCommands\(vb6_fw_hwnd\)') {
    $bad += "T4 control-array container forwarding does not call ForwardChildCommands per element"
}
$formShowSites = @($backend | Where-Object { $_.Text -match 'vb6_form_show_"' })
if ($formShowSites.Count -lt 1) { $bad += "T4 no vb6_form_show_<Form> emission site found" }

if ($bad.Count -eq 0) {
    Write-Host ("PASS uctreeview authority: RaiseEvent 分发 1 / ChooseColor DI 站 " + $ccSites.Count +
                " / 后端 ShowWindow 散点 " + $beShowWin.Count + " / Missing+Is+槽位出口齐 / 数组容器转发齐") -ForegroundColor Green
    exit 0
}
foreach ($b in $bad) { Write-Host ("FAIL " + $b) -ForegroundColor Red }
exit 1
