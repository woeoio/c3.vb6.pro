# VB6 颜色值 → COLORREF 的唯一出口哨兵 (只扫源码, 不起 cl)
#
# 为什么单开这一道: VB6 的系统色是 32 位 Long 的负值 `&H80000000& Or n`
# (n = COLOR_SCROLLBAR(0) .. COLOR_MENUBAR(30)), 必须按 GetSysColor(n) 解析;
# 非系统色 (含负 COLORREF) 按 0x00BBGGRR 语义。这两件事此前在 RTL 里有 4 份
# 各写各的 inline 判据, 绘制层还直接把原始负值裸转 (COLORREF) 造刷 ——
# VBFlexGridDemo 的 Picture2 (`BackColor = &H8000000D&`) 因此被画成 RGB(13,0,0)
# 而不是 GetSysColor(13)。现在翻译只许发生在 vb6_ColorToColorRef 一处。
#
# 规则 (改坏了会红, 不是装饰):
#   C1  vb6_ColorToColorRef 在 src/rtl 里恰好定义一次, 函数体同时有
#       GetSysColor(...) / 0x80000000 判据 / 0x00FFFFFF 兜底 (三档缺一就是口径变形)
#   C2  src/rtl 里不许再出现 `(COLORREF)vb6_GetControlBackColor` / `...ForeColor` /
#       `(COLORREF)vb6_Ambient_BackColor` / `(COLORREF)vb6_UserControl_ForeColor`
#       (裸强转 = 把系统色负值当 COLORREF, 本缺陷的直接形状)
#   C3  src/rtl 里除唯一出口函数体那一行外, 任何 `GetSysColor(<表达式含 &>)` 都算
#       第二份解码 (允许的只有 `GetSysColor(COLOR_XXX)` 常量与出口那一行)
#   C4  src/backend / src/project / src/driver 里不许出现 `(int)(unsigned int)` ——
#       设计期颜色发射曾靠这一对截断"碰巧"折回负数, 现在统一走 frmValueToVbLong
#   C5  设计期 BackColor/ForeColor 的发射行 (三个 .inc, 共 5 站) 必须经
#       frmValueToVbLong 折符号, 不许直接 to_string(intValue) 或裸强转
#
# 白名单 (带保号/合法理由, 改动时同步复查):
#   W1  vb6forms_ctrl.c 的 `return GetSysColor((int)(v & 0xFFu));`
#       —— 这是**唯一出口**的函数体, 全仓唯一允许的 VB 系统色解码行。
#   W2  frm_parser.hpp 的 frmValueToVbLong(): 代码生成侧唯一的符号折回出口,
#       纯算术实现 (unsigned long long 掩码 + 减 2^32), 不含任何有符号窄化强转。
#
# 用法:  powershell -NoProfile -File scripts\check_color_ref_authority.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$bad = @()

# 去掉块注释与整行 // 注释; 保留行号 (块注释按等长空格抹掉, 换行不动)。
function Get-CodeLines([string]$path, [string]$exts) {
    $f = Get-Item -LiteralPath $path
    $txt = [System.IO.File]::ReadAllText($f.FullName)
    $txt = [regex]::Replace($txt, '/\*[\s\S]*?\*/', { param($m) [regex]::Replace($m.Value, '[^\r\n]', ' ') })
    $ln = 0
    foreach ($line in ($txt -split "`r?`n")) {
        $ln++
        $t = $line.Trim()
        if ($t.StartsWith("//")) { continue }
        [pscustomobject]@{ File = $f.Name; Path = $f.FullName; Line = $ln; Text = $t }
    }
}

function Get-SrcLines([string]$subdir) {
    $dir = Join-Path $root $subdir
    if (-not (Test-Path -LiteralPath $dir)) { return @() }
    $out = @()
    foreach ($f in (Get-ChildItem -LiteralPath $dir -Recurse -File | Where-Object { $_.Extension -in ".c", ".h", ".inc", ".cpp", ".hpp" })) {
        $out += @(Get-CodeLines $f.FullName "")
    }
    return $out
}

$rtl = @(Get-SrcLines "src\rtl")

# ---- C1: 唯一出口恰好一处, 三档判据都在 ----
$authDef = @($rtl | Where-Object { $_.Text -match '\bCOLORREF\s+vb6_ColorToColorRef\s*\(' -and -not $_.Text.EndsWith(";") })
if ($authDef.Count -ne 1) {
    $bad += ("C1 vb6_ColorToColorRef defined " + $authDef.Count + " times in src/rtl (want exactly 1; .h 里以 ; 结尾的声明不算) -> " +
             (($authDef | ForEach-Object { $_.File + ":" + $_.Line }) -join " | "))
} elseif ($authDef[0].File -ne "vb6forms_ctrl.c") {
    $bad += ("C1 the authority moved to " + $authDef[0].File + " (want vb6forms_ctrl.c; 测地表按这里写)")
} else {
    $blk = [regex]::Match([System.IO.File]::ReadAllText($authDef[0].Path),
                          'COLORREF\s+vb6_ColorToColorRef\s*\([^)]*\)\s*\{[\s\S]{0,600}?\r?\n\}')
    if (-not $blk.Success) {
        $bad += "C1 authority body not found"
    } else {
        if ($blk.Value -notmatch 'GetSysColor\(') {
            $bad += "C1 the authority no longer resolves system colors through GetSysColor"
        }
        if ($blk.Value -notmatch '0x80000000u') {
            $bad += "C1 the authority lost the &H80000000& bit test"
        }
        if ($blk.Value -notmatch '<= 30u') {
            $bad += "C1 the authority no longer bounds the system-index to [0..30]"
        }
        if ($blk.Value -notmatch '0x00FFFFFF') {
            $bad += "C1 the authority lost the non-system COLORREF fallback (low 24 bits)"
        }
    }
}

# ---- C2: 裸强转 VB 颜色值 = 本缺陷直形 ----
foreach ($ln in $rtl) {
    if ($ln.Text -match '\(COLORREF\)\s*vb6_(GetControl(Back|Fore)Color|Ambient_BackColor|UserControl_ForeColor)') {
        $bad += ("C2 bare (COLORREF) cast of a VB color value at " + $ln.File + ":" + $ln.Line +
                 " -> 一律改问 vb6_ColorToColorRef 唯一出口 (" + $ln.Text.Substring(0, [Math]::Min(90, $ln.Text.Length)) + ")")
    }
}

# ---- C3: 第二份系统色解码 ----
# W1: 唯一出口函数体那一行是合法的。逐行精确匹配 (不是按文件豁免 —— 同文件里再抄一份也要红)。
$w1 = 'return GetSysColor((int)(v & 0xFFu));'
$w1Hits = 0
foreach ($ln in $rtl) {
    if (-not $ln.Text.Contains("GetSysColor(")) { continue }
    if (-not $ln.Text.Contains("&")) { continue }        # GetSysColor(COLOR_XXX) 常量不算解码
    if ($ln.File -eq "vb6forms_ctrl.c" -and $ln.Text -eq $w1) { $w1Hits++; continue }   # W1 白名单
    $bad += ("C3 second VB system-color decoder at " + $ln.File + ":" + $ln.Line +
             " -> 只许 vb6_ColorToColorRef 一处 (" + $ln.Text.Substring(0, [Math]::Min(90, $ln.Text.Length)) + ")")
}
if ($w1Hits -ne 1) {
    $bad += ("C3 whitelisted authority line not found exactly once (found " + $w1Hits +
             "; 白名单失效说明唯一出口的函数体被改写/搬走, 需同步复查本哨兵)")
}

# ---- C4: 代码生成侧不许靠 (int)(unsigned int) 截断保号 ----
# W2: frm_parser.hpp 的 frmValueToVbLong 是唯一折回出口 (它本身不含这对强转)。
$gen = @()
$gen += @(Get-SrcLines "src\backend")
$gen += @(Get-SrcLines "src\project")
$gen += @(Get-SrcLines "src\driver")
foreach ($ln in $gen) {
    if ($ln.Text -match '\(int\)\s*\(unsigned\s+int\)') {
        $bad += ("C4 (int)(unsigned int) sign-truncation at " + $ln.File + ":" + $ln.Line +
                 " -> 设计期整数发射一律走 frmValueToVbLong (" + $ln.Text.Substring(0, [Math]::Min(90, $ln.Text.Length)) + ")")
    }
}

# ---- C5: 设计期 Back/ForeColor 发射必须经唯一折回出口 ----
# 三个 .inc 里 emission 字符串是跨行的 (值在下一行), 所以按"本站起 3 行窗口"检查。
$emitFiles = @("cgen_form_frame_menu.inc", "cgen_form_ctrl_style_apply.inc", "cgen_form_wndproc_create.inc")
$emitSites = 0
foreach ($name in $emitFiles) {
    $p = Join-Path $root ("src\backend\detail\module\" + $name)
    if (-not (Test-Path -LiteralPath $p)) { $bad += ("C5 missing " + $name); continue }
    $lines = @([System.IO.File]::ReadAllText($p) -split "`r?`n")
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $t = $lines[$i].Trim()
        if ($t.StartsWith("//")) { continue }
        if ($t -notmatch '"vb6_SetControl(Back|Fore)Color\(') { continue }
        $emitSites++
        $win = ($lines[$i..([Math]::Min($i + 2, $lines.Count - 1))] -join " ")
        if ($win -notmatch 'frmValueToVbLong\(') {
            $bad += ("C5 design-time color emit without frmValueToVbLong at " + $name + ":" + ($i + 1) +
                     " -> 裸 to_string(intValue) 对 &H8000000D& 会发出 2147483661 (或靠截断碰巧变负)")
        }
    }
}
if ($emitSites -ne 5) {
    $bad += ("C5 design-time color emit sites = " + $emitSites + " (want 5: 顶层 2 + 容器子控件 2 + 窗体自身 1)")
}

if ($bad.Count -eq 0) {
    Write-Host ("PASS VB color authority: 出口 " + $authDef.Count + " / 白名单行 " + $w1Hits +
                " / 设计期发射 " + $emitSites + " 站全部经 frmValueToVbLong") -ForegroundColor Green
    exit 0
}
foreach ($b in $bad) { Write-Host ("FAIL " + $b) -ForegroundColor Red }
exit 1
