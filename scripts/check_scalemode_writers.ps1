# W 账 #197 的结构性哨兵 (只扫源码, 不起 cl)
#
# 为什么单开这一道: `VB6_ScaleMode` 这个窗口属性在 RTL 里**读点齐全、写者为零** ——
# vb6_GetScaleMode / vb6_WindowScaleModeSelf / vb6_ContainerScaleMode 全在读它,
# 缺省 1=缇, 而生成代码里从来没有人调过 vb6_SetScaleMode。后果不是某一枚控件长歪,
# 是**量出来的数全是错的**: ScaleWidth/ScaleHeight、控件几何 (#175 那一族换算)、
# 文字量纲 (#177) 都建立在一个恒为 1 的数上。
# "没人写"这件事在源码层面看跟"有人写"一模一样 (读点一条不少), 所以判据必须是结构性的:
#   设计块写了 ScaleMode ⇒ 产物里必须有一句 vb6_SetScaleMode。
#
# 口径现在收在一处: CCodeGen::emitDesignerScaleModeProp (src/backend/cgen_util_ctrl.cpp),
# 三条落点同调 —— 顶层控件创建路 / 容器子控件创建路 (#83 那条"两条创建路都要打") / 窗体自己的 WM_CREATE。
# 读的那一头也收在一处: 表里给的是 vb6_WindowScaleModeSelf (#175 的单位权威), 不是再拼一份换算。
#
# 规则 (改坏了会红, 不是装饰):
#   W1  src/backend 里手拼发码 `"vb6_SetScaleMode(` = 恰好 1 处 (在权威里; 表里那是返回**名字**, 不算)
#   W2  权威恰好定义一次, 函数体里 `properties.find("ScaleMode")` 与 `design ScaleMode` 都还在
#   W3  三条落点各有一处调用 (顶层 / 容器 / 窗体 WM_CREATE) —— 少一条就是"只接一头"那个老症状
#   W4  读写两张表成对: 读侧 `vb6_WindowScaleModeSelf` 的行数 >= 2 且写侧 `vb6_SetScaleMode` >= 2
#       (账 #278 §B109: Form 那一行搬进了 src\common\form_pseudo.hpp 的表 ⇒ 读侧跨两处数)
#   W5  RTL 里写 `VB6_ScaleMode` 这个属性名的地方 = 1 (只有 setter 自己), 读的地方 >= 2
#
# 用法:  pwsh -File scripts\check_scalemode_writers.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$bad = @()
# 只认非注释部分（§B108 那条负控实测过一次：纯文本正则会连 `// addObj("X")` 一起算进名单）。
function Strip-RowComments([string]$text) {
    $kept = @()
    foreach ($line in ($text -split "`r?`n")) {
        $cut = $line.IndexOf('//')
        $kept += $(if ($cut -ge 0) { $line.Substring(0, $cut) } else { $line })
    }
    return ($kept -join "`n")
}


$authRel = "src\backend\cgen_util_ctrl.cpp"
$auth = Join-Path $root $authRel
if (-not (Test-Path -LiteralPath $auth)) {
    Write-Host ("FAIL W0 authority file missing: " + $authRel) -ForegroundColor Red
    exit 1
}

# W1: 后端手拼发码点只许在权威里一处。表里那条是 `return "vb6_SetScaleMode";` —— 交的是**名字**,
# 由调用点拼实参, 所以这里只数带左括号的那种。
$hand = 0
$handWhere = @()
$beDir = Join-Path $root "src\backend"
foreach ($f in (Get-ChildItem -LiteralPath $beDir -Recurse -File | Where-Object { $_.Extension -in ".cpp", ".inc", ".hpp" })) {
    $ln = 0
    foreach ($line in ([System.IO.File]::ReadAllText($f.FullName) -split "`r?`n")) {
        $ln++
        $t = $line.Trim()
        if ($t.StartsWith("//")) { continue }
        if ($line.Contains('"vb6_SetScaleMode(')) {
            $hand++
            $handWhere += ($f.Name + ":" + $ln)
        }
    }
}
if ($hand -ne 1) {
    $bad += ("W1 hand-composed emit sites of vb6_SetScaleMode( = " + $hand +
             " (want exactly 1, inside the authority) -> " + ($handWhere -join " | "))
}

# W2: 权威的形状
$h = [System.IO.File]::ReadAllText($auth)
$defs = @([regex]::Matches($h, 'void\s+CCodeGen::emitDesignerScaleModeProp\s*\('))
if ($defs.Count -ne 1) {
    $bad += ("W2 emitDesignerScaleModeProp defined " + $defs.Count + " times (want 1)")
}
$blk = [regex]::Match($h, 'emitDesignerScaleModeProp[\s\S]{0,1400}?\r?\n\}')
if (-not $blk.Success) {
    $bad += "W2 authority body not found"
} else {
    if ($blk.Value -notmatch 'properties\.find\("ScaleMode"\)') {
        $bad += "W2 the authority no longer reads the design-block ScaleMode property"
    }
    if ($blk.Value -notmatch 'design ScaleMode') {
        $bad += "W2 the authority lost its emit line (or the tag that identifies it)"
    }
}

# W3: 三条落点
$spots = @(
    @("顶层创建路",   "src\backend\detail\module\cgen_form_ctrl_style_apply.inc"),
    @("容器子控件路", "src\backend\detail\module\cgen_form_frame_menu.inc"),
    @("窗体 WM_CREATE", "src\backend\detail\module\cgen_form_wndproc_create.inc")
)
$n3 = 0
foreach ($pair in $spots) {
    $fp = Join-Path $root $pair[1]
    if (-not (Test-Path -LiteralPath $fp)) { $bad += ("W3 missing " + $pair[0] + " file"); continue }
    # 注释行不算 —— 第一版这里是整篇正则, 把调用点**注释掉** W3 照样绿, 那条假 needle 白验了。
    $n = 0
    foreach ($line in ([System.IO.File]::ReadAllText($fp) -split "`r?`n")) {
        if ($line.Trim().StartsWith("//")) { continue }
        if ($line.Contains("emitDesignerScaleModeProp(")) { $n++ }
    }
    if ($n -lt 1) {
        $bad += ("W3 " + $pair[0] + " 那条路不再调用出口 (设计期 ScaleMode 又没人写了)")
    }
    $n3 += $n
}

# W4: 读写成对
# 账 #278 §B109: Form 那一行现在住在 src\common\form_pseudo.hpp 的表里（两层的唯一出口），
# backend 只剩 PictureBox 那条 => 读侧要跨两处数, 少任何一处都算没登记。
$fpRawA = [System.IO.File]::ReadAllText((Join-Path $root "src\common\form_pseudo.hpp"))
$fpTxtA = Strip-RowComments $fpRawA
$nR = @([regex]::Matches($h, 'return\s+"vb6_WindowScaleModeSelf";')).Count +
      @([regex]::Matches($fpTxtA, '\{\s*"scalemode"\s*,\s*"vb6_WindowScaleModeSelf"')).Count
$nW = @([regex]::Matches($h, 'return\s+"vb6_SetScaleMode";')).Count
if ($nR -lt 2) { $bad += ("W4 read-table rows for scalemode = " + $nR + " (want >= 2: backend 的 PictureBox 一条 + common 表里的 Form 一条)") }
if ($nW -lt 2) { $bad += ("W4 write-table rows for scalemode = " + $nW + " (want >= 2, 与读侧成对)") }

# W5: RTL 那头, 属性名只许 setter 一个人写; 读点不许哑
$rtlDir = Join-Path $root "src\rtl"
$wr = 0; $rd = 0
foreach ($f in (Get-ChildItem -LiteralPath $rtlDir -Recurse -File | Where-Object { $_.Extension -in ".c", ".h" })) {
    foreach ($line in ([System.IO.File]::ReadAllText($f.FullName) -split "`r?`n")) {
        $t = $line.Trim()
        if ($t.StartsWith("//")) { continue }
        if ($t.Contains("SetPropW") -and $t.Contains('L"VB6_ScaleMode"')) { $wr++ }
        if ($t.Contains("GetPropW") -and $t.Contains('L"VB6_ScaleMode"')) { $rd++ }
    }
}
if ($wr -ne 1) { $bad += ("W5 RTL writers of the VB6_ScaleMode window property = " + $wr + " (want exactly 1)") }
if ($rd -lt 1) { $bad += ("W5 RTL readers of VB6_ScaleMode = " + $rd + " (读点哑了说明单位表又分家了)") }

if ($bad.Count -eq 0) {
    Write-Host ("PASS ScaleMode writers: 手拼发码 " + $hand + " 处 / 落点调用 " + $n3 +
                " 处 / 读表 " + $nR + " 写表 " + $nW + " / RTL 写 " + $wr + " 读 " + $rd) -ForegroundColor Green
    exit 0
}
foreach ($b in $bad) { Write-Host ("FAIL " + $b) -ForegroundColor Red }
exit 1
