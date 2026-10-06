# check_subclass_slots.ps1 - 账 #185 的结构性哨兵 (只扫源码, 不起 cl)
#
# 为什么单开这一道: 控件子类化在 RTL 里是**分层**的 —— 建窗时自绘那一层 (PictureBox/Image
# = STATIC+SS_BITMAP) 与发码为事件装的那一层, 都要存"我的下面那层的窗口过程"。两层用同一个
# 属性名时, 配上两边都写的"属性已存在就不装", 症状是后装那层**静默不生效**: 事件臂生成得好好的、
# 编得过、一次也不响 (实测 PictureBox 的 Paint/MouseDown/MouseUp/Click 与 Image 的 Click 全灭),
# 而同窗体上没人抢的 Label 一直正常。这类"装了但没装"只有运行期才看得见 ⇒ 静态钉住。
#
# 规则 (改坏了会红, 不是装饰):
#   S1  存原始窗口过程的属性名 (VB6_*OrigProc) 必须由**唯一**一个文件写入 —— 一个名字 = 一层
#   S2  通用事件层 VB6_OrigProc 归 vb6forms_widget.c; 自绘层必须用自己的名字, 不得再出现裸的
#       VB6_OrigProc (那是 #185 之前的形状)
#   S3  共享 DC 的画序两头都在: 自绘层读 VB6_PaintDC (有就只画不 BeginPaint),
#       发码那一臂既 SetProp 又向下转调
#   S4  发码那一臂里"向下转调"必须在"抬用户 _Paint"**之前** (顺序反了就是表面盖住用户笔画)
#
# 用法:  pwsh -File scripts\check_subclass_slots.ps1
# 退出码: 0 = 全绿; 1 = 红 (每条失败单独打一行)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$bad = @()

function Get-Text([string]$path) {
    if (-not (Test-Path -LiteralPath $path)) { return $null }
    return [System.IO.File]::ReadAllText($path)
}

$widget = Join-Path $root "src\rtl\core\vb6forms\vb6forms_widget.c"
$picture = Join-Path $root "src\rtl\core\vb6forms\vb6forms_picture_prop.c"
$arm = Join-Path $root "src\backend\detail\module\cgen_form_wndproc_subclass.inc"
foreach ($p in @($widget, $picture, $arm)) {
    if (-not (Test-Path -LiteralPath $p)) { $bad += ("S0 missing file: " + $p) }
}
if ($bad.Count -gt 0) {
    foreach ($b in $bad) { Write-Host ("FAIL " + $b) -ForegroundColor Red }
    exit 1
}

# --- S1: 一个 OrigProc 属性名只允许一个文件写 ---
$rtlFiles = Get-ChildItem -LiteralPath (Join-Path $root "src\rtl") -Recurse -File -Include *.c
$owners = @{}
foreach ($f in $rtlFiles) {
    $t = [System.IO.File]::ReadAllText($f.FullName)
    foreach ($m in [regex]::Matches($t, 'SetPropW\s*\([^,]+,\s*(L"(VB6_\w*OrigProc)")')) {
        $name = $m.Groups[2].Value
        if (-not $owners.ContainsKey($name)) { $owners[$name] = @{} }
        $owners[$name][$f.Name] = $true
    }
}
foreach ($name in ($owners.Keys | Sort-Object)) {
    $files = @($owners[$name].Keys)
    if ($files.Count -gt 1) {
        $bad += ("S1 slot " + $name + " is written by " + $files.Count + " files: " + ($files -join ", "))
    }
}
# 一条负控口径: 至少要认到 4 个分层槽位, 认出 0 个说明扫描本身失效(假绿)
if ($owners.Count -lt 4) {
    $bad += ("S1 census found only " + $owners.Count + " OrigProc slots (<4) - scan is not seeing the layers")
}

# --- S2: 自绘层不得再用裸的 VB6_OrigProc ---
$tw = Get-Text $widget
$tp = Get-Text $picture
if (([regex]::Matches($tw, '"VB6_OrigProc"')).Count -lt 3) {
    $bad += "S2 vb6forms_widget.c no longer owns VB6_OrigProc (expected >=3 uses)"
}
if (([regex]::Matches($tp, '"VB6_OrigProc"')).Count -ne 0) {
    $bad += ("S2 vb6forms_picture_prop.c still references bare VB6_OrigProc x" +
             ([regex]::Matches($tp, '"VB6_OrigProc"')).Count + " (the collision #185 removed)")
}
if (([regex]::Matches($tp, '"VB6_ImageOrigProc"')).Count -lt 2) {
    $bad += "S2 vb6forms_picture_prop.c does not carry its own slot VB6_ImageOrigProc"
}

# --- S3: 共享 DC 两头都在 ---
if (([regex]::Matches($tp, 'GetPropW\s*\(\s*hwnd,\s*L"VB6_PaintDC"')).Count -lt 1) {
    $bad += "S3 self-draw layer never reads VB6_PaintDC (it would BeginPaint a second time)"
}
$ta = Get-Text $arm
if (([regex]::Matches($ta, 'VB6_PaintDC')).Count -lt 2) {
    $bad += "S3 the emitted paint arm no longer hands the DC over via VB6_PaintDC"
}
if (([regex]::Matches($ta, 'CallWindowProcW\(vb6_below')).Count -lt 1) {
    $bad += "S3 the emitted paint arm no longer forwards WM_PAINT to the layer below"
}

# --- S4: 向下转调必须在抬处理器之前 (只在 hasPaint 那一块里比 —— 焦点臂也有 raise) ---
$iStart = $ta.IndexOf("if (info.hasPaint) {")
if ($iStart -lt 0) {
    $bad += "S4 cannot find the hasPaint arm block at all"
} else {
    $iEnd = $ta.IndexOf("// WM_DESTROY", $iStart)
    if ($iEnd -lt 0) { $iEnd = $ta.Length }
    $blk = $ta.Substring($iStart, $iEnd - $iStart)
    $iBelow = $blk.IndexOf("CallWindowProcW(vb6_below")
    $iRaise = $blk.IndexOf('{ extern void " + fn + "(); " + fn + "(); }')
    if ($iBelow -lt 0 -or $iRaise -lt 0) {
        $bad += ("S4 paint arm lost one of the two lines (forward=" + $iBelow + " raise=" + $iRaise + ")")
    } elseif ($iBelow -gt $iRaise) {
        $bad += "S4 paint arm raises the handler BEFORE forwarding down (surface would cover the strokes)"
    }
}

if ($bad.Count -eq 0) {
    $names = ($owners.Keys | Sort-Object) -join ", "
    Write-Host ("PASS subclass slots: " + $owners.Count + " names, one file each -> " + $names) -ForegroundColor Green
    exit 0
}
foreach ($b in $bad) { Write-Host ("FAIL " + $b) -ForegroundColor Red }
exit 1
