# check_form_draw_state.ps1 - 账 #233 的结构性哨兵 (只扫源码, 不起 cl)
#
# 为什么单开这一道: Form 的绘图状态属性 (CurrentX / CurrentY / DrawWidth) 这一族, 缺陷形状是
# "只接了一半" —— 读侧在 cgen 的一份侧表里硬编码四条, 写侧压根没登记进 cgen 的写表, 于是
# `Me.DrawWidth = 3` 发成 `vb6_Form_DrawGetWidth(hwnd) = 3;` (C2106, 两架构都编不出来);
# 同时 RTL 在**同一个窗口属性名** `VB6_CurrentX` 上挂了两套编码 (绘图家族 int32、Form Print
# float 位图案), 互读必错。这些都是"看着像做完了"的形状, 编译与运行各只响一半, 所以要结构钉。
#
# 规则 (改坏了会红, 不是装饰):
#   S1  笔位那份存储唯一: `VB6_CurrentX` / `VB6_CurrentY` 这两个窗口属性名的 GetPropW/SetPropW
#       在 src/rtl 里只许出现在 vb6forms_widget_prop.c (float 那一户) —— 再开第二处 = 又一份编码
#   S2  cgen 侧表不回潮: 四个笔位导出名 (vb6_Form_DrawGet/SetCurrentX/Y) 在 src/backend 里 0 次
#   S3  读写成对: currentx / currenty / drawwidth 三个名在窗体那一档读表(账 #278 §B109 起在
#       src\common\form_pseudo.hpp)与写表里都各有一行
#       getControlPropWriteFn 的 Form 档**都要**有 (少一边就是这一刀回归)
#   S4  编码对称: vb6_DrawSetI 那条 SetPropW 必须写裸值 (只看写行 —— 注释里提到 "v+1" 是在讲历史)
#   S5  画笔色也只有一份存储 (账 #235): 属性名 VB6_DrawForeColor 在 src/rtl 里不许出现在任何
#       GetPropW/SetPropW 行 —— Form 绘图家族的画笔色就是控件那个 ForeColor (唯一出口
#       vb6_GetControlForeColor)。另存一枚的现场就是 `Me.ForeColor = vbRed` 之后 PSet 画出来是黑。
#   S6  Print 是绘图家族的一员 (账 #237): `vb6_Form_Print` 在 src/rtl 里恰好定义一次, 而那一条
#       必须住在 vb6forms_draw.c; 它体内四件权威一个都不许自己答 —— DC(vb6_DrawAcquire)、
#       字体(vb6_ControlFont()、色(vb6_DrawForeColor、笔位(vb6_GetCurrentY/vb6_SetCurrentY +
#       两条换算码头(vb6_DrawUserToPx / vb6_DrawPxToUser)。以前它住在 vb6forms.c 里自带一份
#       GetDC/ReleaseDC、不选字体、把用户单位当像素落笔、把像素行高存回用户单位那份笔位里。
#   S7  换算只许是码头 (账 #237): 两条码头分别交回 vb6_ScaleUserToPx / vb6_ScalePxToUser,
#       每个调用点都必须**显式写出纵/横那一档**(0 或 1), 且本文件不许再出现自己的 dpi
#       (LOGPIXELSX / v * dpi) —— 那份 `v * dpi / 1440` 只认缇, Point/Inch/cm 差 20 倍。
#   S8  Print/Cls 全仓只有一份实现 (账 #239): 控件那户 vb6_ControlPrint / vb6_ControlCls 只许
#       转调 vb6_Form_Print / vb6_Form_Cls; 窗口属性 VB6_PrintX/Y 在 src/rtl 里 0 次 (那份私有
#       像素光标一回来, pic.Print 就又不读也不动 pic.CurrentX/Y); Cls 的背景色必须问
#       vb6_GetControlBackColor —— 黑色按值存就是 NULL, 自己判空等于把 vbBlack 读成"没设过"。
#
# 用法:  pwsh -File scripts\check_form_draw_state.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
# 只认非注释部分（§B108 那条负控实测过一次：纯文本正则会连 `// addObj("X")` 一起算进名单）。
function Strip-RowComments([string]$text) {
    $kept = @()
    foreach ($line in ($text -split "`r?`n")) {
        $cut = $line.IndexOf('//')
        $kept += $(if ($cut -ge 0) { $line.Substring(0, $cut) } else { $line })
    }
    return ($kept -join "`n")
}

$root = Split-Path -Parent $PSScriptRoot
$bad = @()

$rtlDir = Join-Path $root "src\rtl"
$beDir = Join-Path $root "src\backend"
if (-not (Test-Path -LiteralPath $rtlDir)) { Write-Host "FAIL S0 missing src\rtl" -ForegroundColor Red; exit 1 }
if (-not (Test-Path -LiteralPath $beDir)) { Write-Host "FAIL S0 missing src\backend" -ForegroundColor Red; exit 1 }

function Get-SrcFiles($dir) {
    Get-ChildItem -LiteralPath $dir -Recurse -File |
        Where-Object { $_.Extension -in ".c", ".h", ".cpp", ".inc", ".hpp" }
}

# ---- S1: 笔位存储唯一 ----
$keepRel = "src\rtl\core\vb6forms\vb6forms_widget_prop.c"
$keep = Join-Path $root $keepRel
$patProp = '(GetPropW|SetPropW)\s*\([^;]*L"VB6_Current[XY]"'
$leak = @()
foreach ($f in Get-SrcFiles $rtlDir) {
    if ($f.FullName -eq $keep) { continue }
    foreach ($m in [regex]::Matches([System.IO.File]::ReadAllText($f.FullName), $patProp)) {
        $leak += ($f.Name + ":" + $m.Value.Substring(0, [Math]::Min(40, $m.Value.Length)))
    }
}
if ($leak.Count -ne 0) {
    $bad += ("S1 pen-position store opened a second home = " + $leak.Count + " -> " +
             (($leak | Select-Object -First 5) -join " | "))
}
$keepText = [System.IO.File]::ReadAllText($keep)
foreach ($n in @("VB6_CurrentX", "VB6_CurrentY")) {
    if (@([regex]::Matches($keepText, '"' + $n + '"')).Count -lt 2) {
        $bad += ("S1 " + $keepRel + " no longer owns both ends of " + $n)
    }
}

# ---- S2: cgen 侧表不回潮 ----
$patSide = 'vb6_Form_Draw(Set)?Current[XY]'
$side = @()
foreach ($f in Get-SrcFiles $beDir) {
    foreach ($m in [regex]::Matches([System.IO.File]::ReadAllText($f.FullName), $patSide)) {
        $side += ($f.Name + ":" + $m.Value)
    }
}
if ($side.Count -ne 0) {
    $bad += ("S2 the deleted side-list is back in cgen = " + (($side | Select-Object -First 5) -join " | "))
}

# ---- S3: 读写成对 (Form 档) ----
$util = Join-Path $root "src\backend\cgen_util_ctrl.cpp"
$u = [System.IO.File]::ReadAllText($util)
function Get-FormCase([string]$text, [string]$fnPat, [string]$tag) {
    $mFn = [regex]::Match($text, $fnPat + '[\s\S]*?\r?\n\}')
    if (-not $mFn.Success) { return $null }
    $mCase = [regex]::Match($mFn.Value, 'case\s+FrmControlType::Form\s*:[\s\S]*?break\s*;')
    if (-not $mCase.Success) { return $null }
    return $mCase.Value
}
$rd = Get-FormCase $u 'std::string\s+CCodeGen::getControlPropReadFn' 'read'
# 账 #278 §B109: Form 那一臂整张搬进了 common 的表 ⇒ 读侧的名字在那里；两处合起来算，
# 两边都查不到才算没登记（写表仍住 cgen_util_ctrl.cpp，没搬）。
if ($null -ne $rd) {
    $fpRawC = [System.IO.File]::ReadAllText((Join-Path $root "src\common\form_pseudo.hpp"))
    $rd = $rd + (Strip-RowComments $fpRawC)
}
$wr = Get-FormCase $u 'std::string\s+CCodeGen::getControlPropWriteFn' 'write'
if ($null -eq $rd) { $bad += "S3 getControlPropReadFn Form case not found" }
if ($null -eq $wr) { $bad += "S3 getControlPropWriteFn Form case not found" }
if ($null -ne $rd -and $null -ne $wr) {
    foreach ($p in @("currentx", "currenty", "drawwidth")) {
        $nr = @([regex]::Matches($rd, '"' + $p + '"')).Count
        $nw = @([regex]::Matches($wr, '"' + $p + '"')).Count
        if ($nr -lt 1 -or $nw -lt 1) {
            $bad += ("S3 property '" + $p + "' is registered read=" + $nr + " write=" + $nw +
                     " (both ends are required; the missing half is what made 233 compile-broken)")
        }
    }
}

# ---- S4: 编码对称 ----
$draw = Join-Path $root "src\rtl\core\vb6forms\vb6forms_draw.c"
$d = [System.IO.File]::ReadAllText($draw)
$mSet = [regex]::Match($d, 'static void vb6_DrawSetI\([\s\S]*?\r?\n\}')
if (-not $mSet.Success) {
    $bad += "S4 vb6_DrawSetI not found in vb6forms_draw.c"
} elseif (@([regex]::Matches($mSet.Value, 'SetPropW\s*\(\s*hw\s*,\s*name\s*,[^;]*')).Count -ne 1) {
    $bad += "S4 vb6_DrawSetI no longer has exactly one SetPropW(hw, name, ...) write"
} else {
    # 只看那一行写进去的表达式 —— 注释里出现 "v+1" 是在讲历史, 不算数
    $w = [regex]::Match($mSet.Value, 'SetPropW\s*\(\s*hw\s*,\s*name\s*,[^;]*').Value
    if ($w -match '\+\s*1') {
        $bad += ("S4 vb6_DrawSetI stores v+1 again while vb6_DrawGetI does not subtract 1 -> " + $w.Trim())
    }
}
# ---- S5: 画笔色只有一份存储 (账 #235) ----
$patFg = '(GetPropW|SetPropW)\s*\([^;]*L"VB6_DrawForeColor"'
$fghost = @()
foreach ($f in Get-SrcFiles $rtlDir) {
    foreach ($m in [regex]::Matches([System.IO.File]::ReadAllText($f.FullName), $patFg)) {
        $fghost += ($f.Name + ":" + $m.Value.Substring(0, [Math]::Min(44, $m.Value.Length)))
    }
}
if ($fghost.Count -ne 0) {
    $bad += ("S5 pen color opened a second store = " + $fghost.Count + " -> " + (($fghost | Select-Object -First 4) -join " | "))
}




# ---- S6: Print 住在家族里, 且四件都问权威 (账 #237) ----
$printDefs = @()
foreach ($f in Get-SrcFiles $rtlDir) {
    $txt = [System.IO.File]::ReadAllText($f.FullName)
    foreach ($m in [regex]::Matches($txt, 'void\s+vb6_Form_Print\s*\([^)]*\)\s*\{')) {
        $printDefs += $f.Name
    }
}
if ($printDefs.Count -ne 1) {
    $bad += ("S6 vb6_Form_Print defined " + $printDefs.Count + " times in src\rtl (want exactly 1) -> " +
             ($printDefs -join ", "))
} elseif ($printDefs[0] -ne "vb6forms_draw.c") {
    $bad += ("S6 vb6_Form_Print lives in " + $printDefs[0] + " again -> it left the drawing family" +
             " and grew its own DC/units answer back")
} else {
    $mPrint = [regex]::Match($d, 'void\s+vb6_Form_Print\s*\([\s\S]*?\r?\n\}')
    if (-not $mPrint.Success) {
        $bad += "S6 vb6_Form_Print body not found in vb6forms_draw.c"
    } else {
        foreach ($need in @("vb6_DrawAcquire", "vb6_ControlFont(", "vb6_DrawForeColor",
                            "vb6_GetCurrentY", "vb6_SetCurrentY", "vb6_DrawUserToPx",
                            "vb6_DrawPxToUser")) {
            if ($mPrint.Value.IndexOf($need) -lt 0) {
                $bad += ("S6 vb6_Form_Print no longer asks the authority " + $need)
            }
        }
    }
}

# ---- S7: 单位换算只许是码头, 而且必须点名纵/横 (账 #237) ----
$mU = [regex]::Match($d, 'static double vb6_DrawUserToPx\([\s\S]*?\r?\n\}')
$mP = [regex]::Match($d, 'static double vb6_DrawPxToUser\([\s\S]*?\r?\n\}')
if (-not $mU.Success -or ($mU.Value.IndexOf("vb6_ScaleUserToPx(") -lt 0)) {
    $bad += "S7 vb6_DrawUserToPx no longer delegates to vb6_ScaleUserToPx"
}
if (-not $mP.Success -or ($mP.Value.IndexOf("vb6_ScalePxToUser(") -lt 0)) {
    $bad += "S7 vb6_DrawPxToUser no longer delegates to vb6_ScalePxToUser"
}
# Every call of the two docks must name its axis (0 = x, 1 = y): the pre-237 helper had
# no axis at all, so the vertical leg was converted with the horizontal DPI.
$callTot = @([regex]::Matches($d, 'vb6_Draw(UserToPx|PxToUser)\s*\(')).Count
$axisTot = @([regex]::Matches($d, 'vb6_Draw(?:UserToPx|PxToUser)\s*\([^;{]*?,\s*[01]\s*\)')).Count
if ($callTot - 2 -ne $axisTot) {
    $bad += ("S7 " + ($callTot - 2 - $axisTot) + " conversion call site(s) without an explicit" +
             " 0/1 axis flag (calls " + ($callTot - 2) + ", flagged " + $axisTot + ")")
}
$own = @()
foreach ($ln in ($d -split "`r?`n")) {
    $t = $ln.Trim()
    if ($t.StartsWith("//")) { continue }
    foreach ($tok in @("LOGPIXELSX", "1440", "dpi")) {
        if ($t.Contains($tok)) {
            $own += ($tok + " :: " + $t.Substring(0, [Math]::Min(58, $t.Length)))
        }
    }
}
if ($own.Count -ne 0) {
    $bad += ("S7 vb6forms_draw.c carries its own DPI math again -> " +
             (($own | Select-Object -First 3) -join " | ") +
             " (Fix 184: every scale conversion shares the one real-DPI pair in vb6forms.c)")
}

# ---- S8: Print/Cls 全仓只许一份实现, 控件那户只许转调 (账 #239) ----
# 改前 vb6forms_ctrl.c 的 Print/Cls 是本族另写的一份: 笔位存在窗口属性 VB6_PrintX/Y 上、
# 按像素推进, 既不读 pic.CurrentX/CurrentY 也不动它们; Cls 也只复位那份私有存储。
$patPrintProp = '(GetPropW|SetPropW|RemovePropW)\s*\([^;]*L"VB6_Print[XY]"'
$leakPen = @()
foreach ($f in Get-SrcFiles $rtlDir) {
    foreach ($m in [regex]::Matches([System.IO.File]::ReadAllText($f.FullName), $patPrintProp)) {
        $leakPen += ($f.Name + ":" + $m.Value.Substring(0, [Math]::Min(44, $m.Value.Length)))
    }
}
if ($leakPen.Count -ne 0) {
    $bad += ("S8 the control-side pixel cursor is back = " + $leakPen.Count + " -> " +
             (($leakPen | Select-Object -First 4) -join " | "))
}
$ctrlFile = Join-Path $root "src\rtl\core\vb6forms\vb6forms_ctrl.c"
$ct = [System.IO.File]::ReadAllText($ctrlFile)
foreach ($pair in @(@("vb6_ControlPrint", "vb6_Form_Print"), @("vb6_ControlCls", "vb6_Form_Cls"))) {
    $mDock = [regex]::Match($ct, 'void\s+' + $pair[0] + '\s*\([\s\S]*?\r?\n\}')
    if (-not $mDock.Success) {
        $bad += ("S8 " + $pair[0] + " definition not found in vb6forms_ctrl.c")
    } elseif ($mDock.Value.IndexOf($pair[1] + "(") -lt 0) {
        $bad += ("S8 " + $pair[0] + " no longer forwards to " + $pair[1] + " -> the control side" +
                " grew its own Print/Cls implementation back (pen store, units, background)")
    }
}
foreach ($fn in @("vb6_Form_Print", "vb6_Form_Cls")) {
    $defs = @()
    foreach ($f in Get-SrcFiles $rtlDir) {
        $t2 = [System.IO.File]::ReadAllText($f.FullName)
        foreach ($m in [regex]::Matches($t2, 'void\s+' + $fn + '\s*\([^)]*\)\s*\{')) { $defs += $f.Name }
    }
    if ($defs.Count -ne 1 -or $defs[0] -ne "vb6forms_draw.c") {
        $bad += ("S8 " + $fn + " must be defined exactly once and inside vb6forms_draw.c -> found " +
                 $defs.Count + " in " + ($defs -join ", "))
    }
}
# Cls 不许自己答背景色: 黑色存进窗口属性就是 NULL, "0 与没设过同构"那颗哨兵住在
# vb6_GetControlBackColor (Fix 187) —— 自己按值判空就把 BackColor = vbBlack 读成未设置。
$mCls = [regex]::Match($d, 'void\s+vb6_Form_Cls\s*\([\s\S]*?\r?\n\}')
if (-not $mCls.Success) {
    $bad += "S8 vb6_Form_Cls body not found in vb6forms_draw.c"
} else {
    # 只看代码行: 讲历史的注释里出现这两个名字不算数 —— 负控就是这么发现的 (注释写着
    # vb6_GetControlBackColor 而代码答的是 VB6_BackColor, 只查体内文本的判据是哑的)。
    $clsLines = @()
    foreach ($ln in ($mCls.Value -split "`r?`n")) {
        $t = $ln.Trim()
        if ($t.StartsWith("//")) { continue }
        $clsLines += $t
    }
    $clsCode = $clsLines -join " "
    if ($clsCode.IndexOf("vb6_GetControlBackColor(") -lt 0) {
        $bad += "S8 vb6_Form_Cls no longer asks vb6_GetControlBackColor (the Fix 187 set-sentinel lives there)"
    }
    if ($clsCode -match 'L"VB6_BackColor"') {
        $bad += 'S8 vb6_Form_Cls answers the background from the raw property name again (black = 0 is indistinguishable from unset)'
    }
}

# ---- S9: 画布家族 (Cls / Print / Line) 的名字只许出自一张表 (账 #232②) ----
# 这一族以前把同一个决定答了四遍: 成员侧打标记时硬编码名字名单, 表达式码头与语句码头
# 又各抄一份"接收者是 PictureBox 才答"。窗体自己那枚接收者不在任何一份里 ——
# 实测 `Me.Cls` 发成 vb6_ComCall(vb6_hwnd_<窗体>, L"Cls", NULL, 0)、
# `Me.Print "AB"` 发成 vb6_ComCallObject(vb6_ComGetObjectProp(同一 HWND, L"Print"), L"Item", ...),
# 两形都编得过、跑得起、一笔不画 (本线第四次栽在"落 COM 兜底 = 静默空转"这一味上)。
# 现在: 名字与 C 出口出自 controlCanvasMethod, 接收者出自 formCtrlSlot, 两条码头只做翻译。
# 只看代码行 —— 讲历史的注释里出现这些名字不算数 (S8 那条负控就是这么变哑的)。

function Get-CodeText($text) {
    $kept = @()
    foreach ($ln in ($text -split "`r?`n")) {
        $t = $ln.Trim()
        if ($t.StartsWith("//")) { continue }
        $kept += $t
    }
    return ($kept -join "`n")
}

$tableFile = Join-Path $root "src\backend\cgen_util_ctrl.cpp"
$declFile = Join-Path $root "src\backend\detail\util\cgen_helpers.inc"

# S9.1 C 出口名只许住在表里 —— 表的家已搬到 src/common/canvas_drawing.hpp（账 #232①-c：
# 语义层也要问同一张表，而 common 是两层唯一都能问的家）。backend 下任何一处再拼一次名字 = 又一份答案。
# 名单**从表里读**（以前这里硬编码三枚名字 —— 表加一档这条判据就看不见那一档，正是本账那一族）。
$canvasTable = Join-Path $root "src\common\canvas_drawing.hpp"
if (-not (Test-Path -LiteralPath $canvasTable)) {
    $bad += ("S9 canvas table missing -> " + $canvasTable)
}
$tableText = ""
if (Test-Path -LiteralPath $canvasTable) {
    $tableText = [System.IO.File]::ReadAllText($canvasTable)
    $mRows = [regex]::Match($tableText, 'kCanvasVerbRows\[\] = \{[\s\S]*?\};')
    if (-not $mRows.Success) {
        $bad += "S9 the canvas verb table body not found in canvas_drawing.hpp"
        $canvasNames = @()
    } else {
        $canvasNames = @([regex]::Matches($mRows.Value, '"(vb6_[A-Za-z0-9_]+)"') |
                         ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
        # census 地板: 表被清空 / 正则退化到只认一行时, 这条判据不许"绿着的空转"
        if ($canvasNames.Count -lt 3) {
            $bad += ("S9 canvas table yields only " + $canvasNames.Count +
                     " entry names (want >= 3) -> the guard went vacuous")
        }
    }
}
foreach ($name in $canvasNames) {
    foreach ($f in Get-SrcFiles $beDir) {
        $code = Get-CodeText ([System.IO.File]::ReadAllText($f.FullName))
        if ($code.IndexOf($name) -lt 0) { continue }
        $bad += ("S9 " + $name + " is spelled outside the canvas table -> " +
                 (Split-Path -Leaf $f.FullName) +
                 " (only src/common/canvas_drawing.hpp may name a canvas entry)")
    }
}

# S9.2 覆盖面矩阵钉死：每一对（动词, 接收者）都得在表里占一行，而且总行数不许悄悄变。
# 上一版这里只查"每个名字在不在"，于是删掉 {"point", CR_CANVAS_FORM} 那一行它一声不响
# （负控实测 rc=0）—— 而"只接了一半接收者"恰恰是本账 ② 那一轮的定义，判据必须换成对。
$wantPairs = @(
    @('cls',    'FORM'),    @('cls',    'PBOX'),    @('cls',    'PRINTER'),
    @('enddoc', 'PRINTER'),
    @('print',  'FORM'),    @('print',  'PBOX'),
    @('line',   'PBOX'),    @('line',   'FORM'),    @('line',   'PRINTER'),
    @('pset',   'FORM'),    @('pset',   'PRINTER'),
    @('circle', 'FORM'),    @('circle', 'PRINTER'),
    @('point',  'FORM'),    @('point',  'PRINTER'),
    @('scale',  'FORM'))
if (Test-Path -LiteralPath $canvasTable) {
    $mTbl = [regex]::Match($tableText, 'kCanvasVerbRows\[\] = \{[\s\S]*?\};')
    if (-not $mTbl.Success) {
        $bad += "S9 the canvas verb table body not found in canvas_drawing.hpp"
    } else {
        $tbl = Get-CodeText $mTbl.Value
        $rowTotal = [regex]::Matches($tbl, '\{"[a-z]+",\s*CR_CANVAS_[A-Z]+').Count
        if ($rowTotal -ne $wantPairs.Count) {
            $bad += ("S9 the canvas table has " + $rowTotal + " (verb, receiver) rows, want exactly " +
                    $wantPairs.Count + " -> add the pair to this matrix when a verb gains a receiver")
        }
        foreach ($pr in $wantPairs) {
            $rx = '\{"' + $pr[0] + '",\s*CR_CANVAS_' + $pr[1] + '\b'
            if (-not [regex]::IsMatch($tbl, $rx)) {
                $bad += ("S9 the canvas table no longer pairs " + $pr[0] + " with CR_CANVAS_" + $pr[1] +
                         " -> a receiver silently loses its drawing entry")
            }
        }
        foreach ($owner in @('CANVAS_OWNER_METHOD', 'CANVAS_OWNER_DRAW')) {
            if ($tbl.IndexOf($owner) -lt 0) {
                $bad += ("S9 the canvas table no longer carries " + $owner +
                         " -> the two owner kinds (member call / drawing dock) collapsed into one")
            }
        }
    }
}

# S9.2b 表不许是死代码：backend 那一条出口必须 include 它并问它要答案，
# 而且 AST 侧的接收者映射两处都得在（Form / PictureBox 少一条就是"表在、没人接"回到本账那一形）。
$askCode = Get-CodeText ([System.IO.File]::ReadAllText($tableFile))
foreach ($want in @('canvas_drawing.hpp', 'canvasMethodEntry(', 'FrmControlType::Form',
                    'FrmControlType::PictureBox')) {
    if ($askCode.IndexOf($want) -lt 0) {
        $bad += ("S9 cgen_util_ctrl.cpp no longer asks the canvas table for " + $want +
                 " -> the table exists but a dock answers from somewhere else again")
    }
}

# S9.3 调用点名单钉死: 定义 1 + 声明 1 + 码头 6
# (表达式形 3 / 打标记 1 / 语句形 1 / **With 控件臂 1**)
# 第七个码头 = Fix 263 在 cgen_expr_with.cpp 的 FormControl 臂补上的 `.Cls`
# (`With PictureDraw : .Cls`)。为什么不能走 `controlZeroArgMethod`: 那一张表的
# 行数被 rtl_proto_arity 的 R3 钉死 (恰好 23), 加一行就红; 而把出口名直接拼进
# 发码侧又会踩 S9.1 (名字只许出自 canvas_drawing.hpp)。
$callTotal = 0
foreach ($f in Get-SrcFiles $beDir) {
    $code = Get-CodeText ([System.IO.File]::ReadAllText($f.FullName))
    $callTotal += [regex]::Matches($code, 'controlCanvasMethod\s*\(').Count
}
if ($callTotal -ne 8) {
    $bad += ("S9 call sites of the canvas table = " + $callTotal + " (want exactly 8:" +
             " 1 definition + 1 declaration + 6 docks)")
}

# S9.4 打标记那一路只许问表: 那张 PictureBox 判据起 400 字符内必须调用 controlCanvasMethod，
# 而且不许把成员名抄成名单 (以前写的是 memLower == "print" || memLower == "cls" || 表 ——
# 表加一档它就漏一档，本账那两形正是这么漏掉的)。
$mbFile = Join-Path $root "src\backend\detail\expr\cgen_expr_member_form_builtin.inc"
$mbCode = Get-CodeText ([System.IO.File]::ReadAllText($mbFile))
$mbHits = @([regex]::Matches($mbCode, 'FrmControlType::PictureBox'))
if ($mbHits.Count -ne 1) {
    $bad += ("S9 the member side asks PictureBox " + $mbHits.Count +
             " times (want exactly 1) -> the canvas receiver test grew a second copy")
} elseif ($mbCode.Substring($mbHits[0].Index, [Math]::Min(400, $mbCode.Length - $mbHits[0].Index)) -notlike "*controlCanvasMethod(*") {
    $bad += "S9 the member side no longer asks controlCanvasMethod for that receiver"
}
if ($mbCode -match '==\s*"print"\s*\|\|\s*\w+\s*==\s*"cls"') {
    $bad += 'S9 the member side hardcodes the canvas member names again (print / cls list)'
}
# S9.5 绘图码头也必须问表 (名字与"这一档有没有出口"都来自表, 不再自带一份动词名单)
$dwFile = Join-Path $root "src\backend\detail\expr\cgen_expr_call_callee_withm.inc"
$dwCode = Get-CodeText ([System.IO.File]::ReadAllText($dwFile))
if ($dwCode.IndexOf('canvasDrawEntry(') -lt 0) {
    $bad += "S9 the drawing dock no longer asks the canvas table (canvasDrawEntry gone)"
}
$mFormP = Get-CodeText ([System.IO.File]::ReadAllText($mbFile))
if ($mFormP.IndexOf('canvasDrawEntry("enddoc"') -lt 0) {
    $bad += "S9 the Printer member route spells its own entry again (enddoc)"
}

# S10 尾巴吸收的名字也只许问那张表（账 #232③）。这一族以前在 parser 里是**第四份**动词名单
# (`toLower(memberName) == "circle"` / `== "pset"` / `== "line"` 三条硬写)，而且只认带接收者那一形
# ⇒ 裸写 `Circle (20,21), 22` 的半径与 `Line (0,0)-(10,10)` 的第二点都收不进来。
# 现在 parser 两趟都问 canvasVerbTailKind(名字)，名单只剩表那一处。
$parserFile = Join-Path $root "src\parser\parser_expr_postfix.cpp"
$parserCode = Get-CodeText ([System.IO.File]::ReadAllText($parserFile))
foreach ($want in @('common/canvas_drawing.hpp', 'canvasVerbTailKind(', 'CANVAS_TAIL_COORD', 'CANVAS_TAIL_TWO_POINT')) {
    if ($parserCode.IndexOf($want) -lt 0) {
        $bad += ("S10 parser_expr_postfix.cpp no longer asks the canvas table for " + $want +
                 " -> a canvas verb's tail is being decided by a hand-written list again")
    }
}
# 一问只许问一次（尾巴档的名字归一全在本文件这一处），而"两坐标那一档"只许**判**一次:
# 吸收点与 `case TokenKind::Minus` 的分发点都问 isTwoPointCanvasVerb —— 两个语法位置
# 各问一次是必然的（要不要消费 token 只能在动手之前决定），但答案只有一位主人。
# 账 #232④ 把这两点钉死: 吸收动作一份都不许抄第二遍 (抄一遍 = 旗标折法分家),
# "哪一档是这一形"也不许在第二处重新判。
if ([regex]::Matches($parserCode, 'canvasVerbTailKind\s*\(').Count -ne 1) {
    $bad += "S10 the parser asks canvasVerbTailKind a number of times other than 1 (one question, both absorptions read the answer)"
}
$twoPointJudges = [regex]::Matches($parserCode, 'CANVAS_TAIL_TWO_POINT').Count
if ($twoPointJudges -ne 1) {
    $bad += ("S10 the parser judges CANVAS_TAIL_TWO_POINT " + $twoPointJudges +
             " times, want exactly 1 (the two-point tail has one owner: isTwoPointCanvasVerb)")
}
$absorbSites = [regex]::Matches($parserCode, 'absorbCanvasLineTail\s*\(').Count
if ($absorbSites -ne 3) {
    $bad += ("S10 absorbCanvasLineTail appears " + $absorbSites +
             " times, want exactly 3 (1 definition + the '(' branch + the '-' branch) -> " +
             "the second-point tail got copied instead of shared")
}
# 坐标尾巴那一格 (Circle 的半径 / PSet 的颜色) 仍由左括号那一支自己判，位置没挪。
if ([regex]::Matches($parserCode, 'tailKind\s*==\s*CANVAS_TAIL_').Count -ne 1) {
    $bad += ("S10 the parser gates on tailKind " +
             [regex]::Matches($parserCode, 'tailKind\s*==\s*CANVAS_TAIL_').Count +
             " times, want exactly 1 (the coordinate tail; the second-point tail moved " +
             "into isTwoPointCanvasVerb, which is the one place that decides it)")
}
foreach ($lit in @('== "circle"', '== "pset"', '== "line"')) {
    if ($parserCode.IndexOf($lit) -ge 0) {
        $bad += ("S10 the parser spells the verb list itself again (" + $lit +
                 ") -> that is the fourth copy; canvas_drawing.hpp answers this question")
    }
}
# 表那一头也得真的答得出三档尾巴（撤掉一处 = 裸写的尾巴又没人收了）
$tailBody = [regex]::Match($tableText, 'canvasVerbTailKind[\s\S]*?\r\n\}')
if (-not $tailBody.Success) {
    $bad += "S10 canvasVerbTailKind is gone from canvas_drawing.hpp -> the tail question has no answer left"
} else {
    foreach ($want in @('"pset"', '"circle"', '"line"', 'CANVAS_TAIL_COORD', 'CANVAS_TAIL_TWO_POINT')) {
        if ($tailBody.Value.IndexOf($want) -lt 0) {
            $bad += ("S10 canvasVerbTailKind no longer maps " + $want + " -> a tail stops being absorbed")
        }
    }
}

# S11 = `[Step]` 前缀这一族 (账 #232⑥)。三条各钉**一件事**, 每件都有"坏掉长什么样"的实物:
#   (a) 名单只许住表里 (canvasVerbHasStep), parser 恰好问一次 —— Step 是不是软关键字 + 左边
#       是不是画布动词, 这两问都在同一枚 file-static 里, 不许在第二处重判;
#   (b) parser 必须在**消费掉 Step 那一枚 token 的那一刻**把两枚旗标写上 (canvasStep1/2);
#   (c) 发码侧三处码头 (Form.PSet / Form.Line / PictureBox.Line) 都得**读**旗标, 一个都不许
#       把 step 位写死 —— 写死的那一格编得过、跑得起、画到别处去 (本线栽过四次的"静默"那一味)。
$withmTxt = [System.IO.File]::ReadAllText((Join-Path $root "src\backend\detail\expr\cgen_expr_call_callee_withm.inc"))
$stmtTxt = [System.IO.File]::ReadAllText((Join-Path $root "src\parser\stmt\parser_stmt.cpp"))
$astTxt = [System.IO.File]::ReadAllText((Join-Path $root "src\ast\detail\ast_expr.hpp"))

$askStep = @([regex]::Matches($parserCode, 'canvasVerbHasStep\s*\(')).Count
if ($askStep -ne 1) {
    $bad += ("S11 the parser asks canvasVerbHasStep " + $askStep +
             " times (want exactly 1 - the verb list lives in the table, not in a second reader)")
}
$stepBody = [regex]::Match($tableText, 'canvasVerbHasStep[\s\S]*?\r\n\}')
if (-not $stepBody.Success) {
    $bad += "S11 canvasVerbHasStep is gone from canvas_drawing.hpp -> the Step question has no answer left"
} else {
    # Point 刻意**不在**名单里: 本仓 vb6_Form_Point(hwnd,x,y) 没有 step 位 (台账 §B225 那一格)。
    foreach ($want in @('"line"', '"pset"', '"circle"')) {
        if ($stepBody.Value.IndexOf($want) -lt 0) {
            $bad += ("S11 canvasVerbHasStep no longer maps " + $want + " -> that verb's [Step] stops parsing")
        }
    }
    if ($stepBody.Value -match '"point"') {
        $bad += "S11 canvasVerbHasStep maps point -> its RTL exit has no step bit; that is a signature change (ledger 232 open grid), not a parser answer"
    }
}
foreach ($flag in @('canvasStep1', 'canvasStep2')) {
    $sets = @([regex]::Matches($parserCode, $flag + '\s*=')).Count
    $has  = @([regex]::Matches($astTxt, 'bool\s+' + $flag + '\s*=')).Count
    if ($sets -lt 1) { $bad += ("S11 the parser never assigns " + $flag + " -> a [Step] prefix would be swallowed and the coordinates read as absolute") }
    if ($sets -gt 1) { $bad += ("S11 the parser assigns " + $flag + " " + $sets + " times (want 1 - the hand-off from the Step branch / the absorb point)") }
    if ($has -ne 1) { $bad += ("S11 IndexOrCallExpr carries " + $flag + " " + $has + " times (want 1)") }
}
# 发码侧: 三处码头读旗标的次数 —— PSet/Circle 共用 st1, 两条 Line 支各读两个位。
if (@([regex]::Matches($withmTxt, 'node\.canvasStep1')).Count -ne 3) {
    $bad += ("S11 only " + @([regex]::Matches($withmTxt, 'node\.canvasStep1')).Count +
             " docks read node.canvasStep1 (want 3: PictureBox.Line / Form-Printer Line / the shared PSet+Circle st1) -> some dock writes the step bit as a literal 0")
}
if (@([regex]::Matches($withmTxt, 'node\.canvasStep2')).Count -ne 2) {
    $bad += ("S11 only " + @([regex]::Matches($withmTxt, 'node\.canvasStep2')).Count +
             " docks read node.canvasStep2 (want 2: PictureBox.Line / Form-Printer Line)")
}
if ($stmtTxt -notmatch 'next_\.kind\s*==\s*TokenKind::Step') {
    $bad += "S11 the Line statement no longer routes a Step successor -> `Line Step (x,y)-(x2,y2)` goes back to `Line Input` (VB2001)"
}

if ($bad.Count -eq 0) {
    Write-Host ("PASS form draw state: pen store unique, side-list gone, " +
                "read/write paired for 3 props, encoding symmetric, pen color store unique, " +
                "Print inside the family asking 7 authorities, conversions dock-only with an axis, " +
                "Print/Cls one implementation that the control side only forwards to, canvas verbs one common table that the backend asks with both receivers, the parser's coordinate tails asked of that same table, " +
                "and the [Step] prefix asked once, recorded by the parser and read by all three drawing docks")
    exit 0
}
foreach ($b in $bad) { Write-Host ("FAIL " + $b) -ForegroundColor Red }
exit 1
