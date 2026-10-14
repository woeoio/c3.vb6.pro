# check_ctrl_geom_cache.ps1 - 账 #230 的结构性哨兵 (只扫源码, 不起 cl)
#
# 为什么单开这一道: 控件几何 (Left/Top/Width/Height) 的缺陷形状是**读法**而不是值 ——
# getter 现问 GetWindowRect 再折算回容器单位，于是写进去的数与读出来的数天生差一个像素格
# (写 5000 读回 4995)。这类缺陷编译不响、链接不响，语料里也没有一条针在量它 (存量夹具读的
# 都是 15 的倍数)。更要紧的是它有三个写侧来路 (属性赋值 / Move / 建窗设计值) 与四个读侧出口，
# 少接一个就又是一次"只接了一半"。
#
# 规则 (改坏了会红, 不是装饰):
#   S1  "VB6_Geom*" 那批窗口属性名只许住在 vb6forms_ctrl.c —— 别处出现一对 Get/SetPropW
#       就是又开了一份几何存储 (账 #234/#235/#239 同一族)
#   S2  读写成对: 四个档 (LEFT/TOP/WIDTH/HEIGHT) 每一档都要既有读路线又有写路线，
#       写侧还须凑齐属性 setter + Move (ctrl.c) 与建窗 (vb6forms.c) 三个来路
#   S3  来路计数: vb6forms_ctrl.c 恰好 8 次 Write (4 setter + 4 条 Move 腿)，
#       vb6forms.c 恰好 4 次 Write (建窗四档) —— 少一条腿就是本刀回归的那个形状
#   S4  建窗那一档交出的单位必须是缇 (字面量 1)：.frm 的几何恒按缇写，与容器声明的 ScaleMode
#       无关 (实证 tests/czUI-main/frmDemo.frm：窗体 ScaleMode=3，仍写 ClientWidth=6600、
#       子控件 Width=6240 —— 按像素读它就是 6240 像素，放不下)
#   S5  旧形不回潮: 四个 getter 各含且仅含一次 vb6_GeomCacheRead(；把投影直接 return
#       的那一形 (return (int)vb6_ScalePxToUser) 全文件 0 次
#   S6  读侧那道像素闸不许被删: vb6_GeomCacheRead 体内必须问 vb6_ScaleUserToPx —— 存着的数
#       换算出的像素与窗口现在的像素对不上，就说明**别人**挪过这枚窗口 (ComboBox 建窗时自己
#       补下拉高度就是这一格)，此时必须回投影而不是回缓存
#   S7  账 #247: 宿主模型 (IDispatch 那一路) 的四个几何名两头都只许问 vb6forms_ctrl.c 的
#       出口；它自己那份 MoveWindow / vb6_ho_ctrlRect 投影一出现就是又开了第二份实现
#   S8  账 #251: UC 里 `UserControl.Parent.Move l,t,w,h` 那条 call 路 (vb6_UC_ParentMove)
#       也只许转调同一个收口一次，四档全交 (mask 15)；函数体里出现 MoveWindow(/SetWindowPos(
#       就是把缇当像素摆位的第二份实现回潮
#
# 用法:  pwsh -File scripts\check_ctrl_geom_cache.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$bad = @()

$ctrlRel = "src\rtl\core\vb6forms\vb6forms_ctrl.c"
$formsRel = "src\rtl\core\vb6forms\vb6forms.c"
$propRel = "src\rtl\core\vb6forms\vb6forms_prop.h"
$ctrl = Join-Path $root $ctrlRel
$forms = Join-Path $root $formsRel
$prop = Join-Path $root $propRel
foreach ($need in @($ctrl, $forms, $prop)) {
    if (-not (Test-Path -LiteralPath $need)) {
        Write-Host ("FAIL S0 missing " + $need) -ForegroundColor Red
        exit 1
    }
}

$ctrlText = [System.IO.File]::ReadAllText($ctrl)
$formsText = [System.IO.File]::ReadAllText($forms)
$propText = [System.IO.File]::ReadAllText($prop)
$slots = @("LEFT", "TOP", "WIDTH", "HEIGHT")

function Get-SrcFiles($dir) {
    Get-ChildItem -LiteralPath $dir -Recurse -File |
        Where-Object { $_.Extension -in ".c", ".h", ".cpp", ".inc", ".hpp" }
}

# ---- S1: 几何存储只许有一份 ----
$patStore = '(GetPropW|SetPropW)\s*\([^;]*L"VB6_Geom'
$leak = @()
foreach ($f in Get-SrcFiles (Join-Path $root "src\rtl")) {
    if ($f.FullName -eq $ctrl) { continue }
    foreach ($m in [regex]::Matches([System.IO.File]::ReadAllText($f.FullName), $patStore)) {
        $leak += ($f.Name + ":" + $m.Value.Substring(0, [Math]::Min(48, $m.Value.Length)))
    }
}
if ($leak.Count -ne 0) {
    $bad += ("S1 geometry store opened a second home = " + $leak.Count + " -> " +
             (($leak | Select-Object -First 5) -join " | "))
}
foreach ($s in $slots) {
    if (@([regex]::Matches($ctrlText, ('"VB6_Geom' + $s.Substring(0, 1) + '"'))).Count -lt 1) {
        $bad += ("S1 vb6forms_ctrl.c no longer names its own VB6_Geom" + $s)
    }
}

# ---- S2: 四档读写成对 ----
foreach ($s in $slots) {
    $nr = @([regex]::Matches($ctrlText, 'vb6_GeomCacheRead\([^;]*VB6_GEOM_' + $s)).Count
    $nwCtrl = @([regex]::Matches($ctrlText, 'vb6_GeomCacheWrite\([^;]*VB6_GEOM_' + $s)).Count
    $nwForms = @([regex]::Matches($formsText, 'vb6_GeomCacheWrite\([^;]*VB6_GEOM_' + $s)).Count
    if ($nr -lt 1) { $bad += ("S2 slot " + $s + " has no read route (reads=" + $nr + ")") }
    if ($nwCtrl -lt 2 -or $nwForms -lt 1) {
        $bad += ("S2 slot " + $s + " write routes incomplete: setter+Move=" + $nwCtrl +
                 " creation=" + $nwForms + " (need >=2 and >=1)")
    }
}

# ---- S3: 三个来路的计数 ----
$wCtrl = @([regex]::Matches($ctrlText, 'vb6_GeomCacheWrite\(hwnd, ')).Count
if ($wCtrl -ne 8) {
    $bad += ("S3 vb6forms_ctrl.c writes=" + $wCtrl +
             " (expected 8 = 4 property setters + 4 Move legs)")
}
$wForms = @([regex]::Matches($formsText, 'vb6_GeomCacheWrite\(hwnd, ')).Count
if ($wForms -ne 4) {
    $bad += ("S3 vb6forms.c (creation route) writes=" + $wForms + " (expected 4)")
}

# ---- S4: 建窗那档按缇记 ----
$nTwill = @([regex]::Matches($formsText, 'vb6_GeomCacheWrite\(hwnd, VB6_GEOM_[A-Z]+, (x|y|width|height), 1\);')).Count
if ($nTwill -ne 4) {
    $bad += ("S4 creation-route writes that record the unit as twips=" + $nTwill +
             " (expected 4; .frm geometry is twips regardless of the container ScaleMode -- czUI frmDemo proves it)")
}
if (@([regex]::Matches($formsText, 'vb6_GeomCacheWrite\([^;]*ContainerScaleMode')).Count -ne 0) {
    $bad += "S4 the creation route started asking the container ScaleMode -- design geometry must stay recorded as twips"
}

# ---- S5: getter 都过缓存, 旧形不回潮 ----
foreach ($fn in @("vb6_GetControlLeft", "vb6_GetControlTop", "vb6_GetControlWidth", "vb6_GetControlHeight")) {
    $mFn = [regex]::Match($ctrlText, $fn + '\s*\(void\*\s+hwnd\)\s*\{[\s\S]*?\r?\n\}')
    if (-not $mFn.Success) { $bad += ("S5 getter body not found: " + $fn); continue }
    $nr = @([regex]::Matches($mFn.Value, 'vb6_GeomCacheRead\(')).Count
    if ($nr -ne 1) { $bad += ($fn + " reads the geometry cache " + $nr + " times (need exactly 1)") }
}
if (@([regex]::Matches($ctrlText, 'return\s+\(int\)vb6_ScalePxToUser')).Count -ne 0) {
    $bad += "S5 a getter returned the projection directly again (the pre-230 shape)"
}

# ---- S6: 像素闸还在 ----
$mRead = [regex]::Match($ctrlText, 'int vb6_GeomCacheRead\([\s\S]*?\r?\n\}')
if (-not $mRead.Success) {
    $bad += "S6 vb6_GeomCacheRead definition not found"
} elseif (@([regex]::Matches($mRead.Value, 'vb6_ScaleUserToPx\(')).Count -lt 1) {
    $bad += "S6 vb6_GeomCacheRead lost its pixel gate -- a window somebody else resized would answer the stale cached number"
}
if (@([regex]::Matches($propText, 'enum \{ VB6_GEOM_LEFT')).Count -ne 1) {
    $bad += "S6 the slot enum is not declared exactly once in vb6forms_prop.h"
}

# ---- S7: 账 #247 —— 宿主模型 (IDispatch 那一路) 不许自带第二份几何实现 ----
# 晚绑定的 `With obj : .Width = ...`、UserControl 里的 `Parent.Width` 都落在
# vb6_Host_GetProp / vb6_Host_SetProp。这两支以前自己 GetWindowRect + 写死缇、
# 写侧四档一起 MoveWindow 且不存 VB 侧读数 —— 同一句话两条路给两个数 (实测宿主
# 那路 `f.Width = 7222` 读回 7215)。现在四档两头都必须问 vb6forms_ctrl.c 的出口。
$hoGetRel = "src\rtl\core\vb6forms\uc\detail\uc_hostmodel_getprop.inc"
$hoSetRel = "src\rtl\core\vb6forms\uc\detail\uc_hostmodel_setprop.inc"
$hoGet = Join-Path $root $hoGetRel
$hoSet = Join-Path $root $hoSetRel
foreach ($need in @($hoGet, $hoSet)) {
    if (-not (Test-Path -LiteralPath $need)) {
        Write-Host ("FAIL S7 missing " + $need) -ForegroundColor Red
        exit 1
    }
}
$hoGetText = [System.IO.File]::ReadAllText($hoGet)
$hoSetText = [System.IO.File]::ReadAllText($hoSet)
foreach ($s in @("Left", "Top", "Width", "Height")) {
    $n = @([regex]::Matches($hoGetText, 'vb6_ho_setVariantLong\(out,\s*vb6_GetControl' + $s + '\(obj\)\)')).Count
    if ($n -ne 1) { $bad += ("S7 host-model read of " + $s + " asks vb6_GetControl" + $s + " " + $n + " times (need exactly 1)") }
    $n2 = @([regex]::Matches($hoSetText, 'vb6_SetControl' + $s + '\(obj,\s*vb6_ho_variantToLong\(v\)\)')).Count
    if ($n2 -ne 1) { $bad += ("S7 host-model write of " + $s + " asks vb6_SetControl" + $s + " " + $n2 + " times (need exactly 1)") }
}
if (@([regex]::Matches($hoSetText, 'MoveWindow\(')).Count -ne 0) {
    $bad += "S7 the host model moved a window itself again -- geometry writes must go through the four setters (they own the unit + the VB-side cache)"
}
if (@([regex]::Matches($hoGetText + $hoSetText, 'vb6_ho_ctrlRect')).Count -ne 0) {
    $bad += "S7 the host model's own geometry projection helper came back"
}

# ---- S8: 账 #251 —— UC 里 `Parent.Move` 那条 call 路也只许问同一个收口 ----
# vb6_UC_ParentMove 以前把四个 VB 侧的数 (窗体容器 = 缇) 原样交给 MoveWindow, 于是
# 607 缇落在 607 像素上, 而且四档都不进 VB 侧缓存 (Move 之后读 Me.Width 是投影)。
# 它是 §B75/§B76 那一族的第三条来路: 属性赋值 / 宿主模型之外, 还有语言级 Move 的
# call 那一路。规则: 函数体内恰好一次那条转调语句, 且不许自带摆位 API。
$ucHostRel = "src\rtl\core\vb6forms\uc\uc_host.c"
$ucHost = Join-Path $root $ucHostRel
if (-not (Test-Path -LiteralPath $ucHost)) {
    Write-Host ("FAIL S8 missing " + $ucHost) -ForegroundColor Red
    exit 1
}
$ucHostText = [System.IO.File]::ReadAllText($ucHost)
$mPM = [regex]::Match($ucHostText, '(?s)void vb6_UC_ParentMove\([^\)]*\)\s*\{.*?\r?\n\}')
if (-not $mPM.Success) {
    $bad += "S8 vb6_UC_ParentMove is gone from uc_host.c -- the Parent.Move call path needs its single placement forwarder"
} else {
    $body = $mPM.Value
    if (@([regex]::Matches($body, 'vb6_ControlMove\(\(void\*\)fw, \(double\)left, \(double\)top, \(double\)width, \(double\)height, 15\);')).Count -ne 1) {
        $bad += "S8 Parent.Move no longer forwards all four axes to vb6_ControlMove (mask 15) exactly once"
    }
    foreach ($pat in @('MoveWindow\(', 'SetWindowPos\(')) {
        $n = @([regex]::Matches($body, $pat)).Count
        if ($n -ne 0) { $bad += ("S8 Parent.Move placed a window by itself again (" + $pat + " x" + $n + ") -- twips/units/cache only come from the one authority") }
    }
}

if ($bad.Count -ne 0) {
    foreach ($b in $bad) { Write-Host ("FAIL " + $b) -ForegroundColor Red }
    exit 1
}
Write-Host "PASS ctrl_geom_cache (S1 one store, S2 four slots read+write, S3 8+4 write routes, S4 creation records twips, S5 every getter cached, S6 pixel gate alive, S7 host model asks the same exit, S8 Parent.Move forwards the four axes too)"
