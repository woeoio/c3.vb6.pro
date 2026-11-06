# Tooltip 视觉属性/空串指针的唯一出口哨兵 (只扫源码, 不起 cl)
#
# 为什么单开这一道 (BalloonTooltips 夹具实测):
#   cTT.cls 建 tooltip 后有一句
#       SetWindowTheme hWndTT, StrPtr(sEmpty), StrPtr(sEmpty)
#   sEmpty 是 `Const sEmpty As String = ""` —— VB6 里这是**非 NULL 的零长 BSTR**,
#   StrPtr 答数据地址 (非 0), uxtheme 因此把 tooltip 的视觉样式关掉, 气泡按经典程序
#   画: 底色就是 TTM_SETTIPBKCOLOR 设的那枚 (#FFFFE1/#00FFFF/#FFFF00/#00FF00)。
#   C3 的 vb6_StrPtr 却按**长度**判空: 零长 -> NULL => 实际调用是
#   SetWindowTheme(hTT, NULL, NULL), comctl32 v6 保留视觉样式 => 白底、尺寸 +8px、
#   lfQuality=5(ClearType)。参考 VB6 与 C3 的 TTM_GETTIPBKCOLOR/TTM_GETTIPTEXTCOLOR/
#   WM_GETFONT(LOGFONT) 读数**完全相同**, 唯一差异就是这张"主题皮" —— 所以本家族的
#   出口不在绘制层, 在 BSTR→LPWSTR 这一条指针语义上。
#
# 规则 (改坏了会红, 不是装饰):
#   T1  vb6_StrPtr 在 src/rtl 里恰好定义一次 (vb6rtl_compat.c), 函数体必须是
#       `return (void*)s;` —— 只有 NULL(即 vbNullString) 答 0; 不许出现任何长度判据
#       (vb6_BSTR_Len / == 0 / Len(...))。这就是本缺陷的直接形状。
#   T2  vb6_StrPtrVar 的 BSTR 档必须转调 T1 (Variant 里那枚空串同一口径), 不许自带
#       第二份长度判空。
#   T3  tooltip 底色/字色 (TTM_SETTIPBKCOLOR / TTM_SETTIPTEXTCOLOR) 在 src/rtl 里
#       只许出现在唯一出口 vb6_ToolTipApplyVisual, 且同一条语句必须经
#       vb6_ColorToColorRef 解色 (系统色 &H80000000& 也走同一条解析)。
#       现有散点: 0 处 (台账).
#   T4  对 tooltip 宿主发 WM_SETFONT 同 T3 口径 (字体的唯一出口, 现有 0 处).
#   T5  src/rtl 不许替用户决定 tooltip 主题: `SetWindowTheme` 只许出现在 DI 桩
#       (vb6_di_stubs.c) 且必须**原样透传**两个指针实参 —— 不许把 0 换空串、也不许
#       把空串换 0 (账 #305 第三百五刀: 主题与否由进程激活上下文/用户代码决定).
#   T6  src/rtl 里 tooltip 宿主创建 (`TOOLTIPS_CLASS`) 恰好 2 站: vb6forms_style.c
#       共享宿主 + vb6forms_slider.c 滑块气泡。新增宿主必须同步改台账与本哨兵.
#
# 白名单 (带理由, 改动时同步复查):
#   W1  无 —— T3/T4 的现有散点是 0, 这正是"唯一出口"还没被撬开的证据。
#   W2  SetWindowTheme 的 DI 桩 (vb6_di_stubs.c) —— T5 的合法落点: 它只做
#       Declare 转发, 不改形参。
#
# 用法:  powershell -NoProfile -File scripts\check_tooltip_authority.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$bad = @()

# 去掉块注释与整行 // 注释; 保留行号 (块注释按等长空格抹掉, 换行不动)。
function Get-CodeLines([string]$path) {
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
        $out += @(Get-CodeLines $f.FullName)
    }
    return $out
}

$rtl = @(Get-SrcLines "src\rtl")

# ---- T1: BSTR -> LPWSTR 唯一出口, 零长非 NULL 必须给指针 ----
$spDef = @($rtl | Where-Object { $_.Text -match 'void\s*\*\s*vb6_StrPtr\s*\(' -and $_.Text -notmatch '\)\s*;' })
if ($spDef.Count -ne 1) {
    $bad += ("T1 vb6_StrPtr defined " + $spDef.Count + " times in src/rtl (want exactly 1; .h 声明不算) -> " +
             (($spDef | ForEach-Object { $_.File + ":" + $_.Line }) -join " | "))
} elseif ($spDef[0].File -ne "vb6rtl_compat.c") {
    $bad += ("T1 the BSTR pointer authority moved to " + $spDef[0].File + " (want vb6rtl_compat.c; 测地表按这里写)")
} else {
    $blk = [regex]::Match([System.IO.File]::ReadAllText($spDef[0].Path),
                          'void\s*\*\s*vb6_StrPtr\s*\([^)]*\)\s*\{[\s\S]{0,400}?\r?\n\}')
    if (-not $blk.Success) {
        $bad += "T1 vb6_StrPtr body not found"
    } else {
        if ($blk.Value -notmatch 'return\s*\(void\s*\*\)\s*s\s*;') {
            $bad += "T1 vb6_StrPtr no longer returns the BSTR data pointer directly (want `return (void*)s;`)"
        }
        if ($blk.Value -match 'vb6_BSTR_Len|==\s*0|!=\s*0') {
            $bad += "T1 vb6_StrPtr reintroduced a length test -> 零长 BSTR 又会被折成 NULL " +
                    "(SetWindowTheme(hTT, StrPtr(""""), ...) 收 NULL, comctl32 保留视觉样式 => 白底/+8px/q5)"
        }
    }
}

# ---- T2: Variant 档必须复用 T1 ----
$spvDef = @($rtl | Where-Object { $_.Text -match 'void\s*\*\s*vb6_StrPtrVar\s*\(' -and $_.Text -notmatch '\)\s*;' })
if ($spvDef.Count -ne 1) {
    $bad += ("T2 vb6_StrPtrVar defined " + $spvDef.Count + " times in src/rtl (want exactly 1)")
} else {
    $blk2 = [regex]::Match([System.IO.File]::ReadAllText($spvDef[0].Path),
                           'void\s*\*\s*vb6_StrPtrVar\s*\([^)]*\)\s*\{[\s\S]{0,600}?\r?\n\}')
    if (-not $blk2.Success) {
        $bad += "T2 vb6_StrPtrVar body not found"
    } else {
        if ($blk2.Value -notmatch 'vb6_StrPtr\s*\(') {
            $bad += "T2 vb6_StrPtrVar no longer delegates the BSTR case to vb6_StrPtr"
        }
        if ($blk2.Value -match 'vb6_BSTR_Len') {
            $bad += "T2 vb6_StrPtrVar has its own length test (第二份判空)"
        }
    }
}

# ---- T3: tooltip 底色/字色只许经唯一出口, 且必须经颜色唯一出口解色 ----
$authName = 'vb6_ToolTipApplyVisual'
$tooltipColorHits = 0
foreach ($ln in $rtl) {
    if ($ln.Text -notmatch 'TTM_SETTIPBKCOLOR|TTM_SETTIPTEXTCOLOR') { continue }
    $tooltipColorHits++
    $ok = $ln.File -eq "vb6forms_style.c" -and $ln.Text -match ([regex]::Escape($authName) + '\s*\(') -and
          $ln.Text -match 'vb6_ColorToColorRef\s*\('
    if (-not $ok) {
        $bad += ("T3 naked tooltip color call at " + $ln.File + ":" + $ln.Line +
                 " -> 只许 " + $authName + " 唯一出口, 且颜色经 vb6_ColorToColorRef 解析 (" +
                 $ln.Text.Substring(0, [Math]::Min(90, $ln.Text.Length)) + ")")
    }
}

# ---- T4: tooltip 字体的唯一出口 ----
$tooltipFontHits = 0
foreach ($ln in $rtl) {
    if ($ln.Text -notmatch 'WM_SETFONT') { continue }
    if ($ln.Text -notmatch 'hTT|hwndTT|ToolTip|hToolTip') { continue }
    $tooltipFontHits++
    $ok = $ln.File -eq "vb6forms_style.c" -and $ln.Text -match ([regex]::Escape($authName) + '\s*\(')
    if (-not $ok) {
        $bad += ("T4 naked WM_SETFONT on a tooltip host at " + $ln.File + ":" + $ln.Line +
                 " -> 只许 " + $authName + " 唯一出口 (" + $ln.Text.Substring(0, [Math]::Min(90, $ln.Text.Length)) + ")")
    }
}

# ---- T5: RTL 不替用户决定 tooltip 主题, DI 桩原样透传 ----
$themeHits = 0
foreach ($ln in $rtl) {
    if ($ln.Text -notmatch 'SetWindowTheme\s*\(') { continue }
    $themeHits++
    $ok = $ln.File -eq "vb6_di_stubs.c"
    if ($ok) {
        $tail = $ln.Text.Substring($ln.Text.IndexOf('SetWindowTheme'))
        if ($tail -match '\bNULL\b|\b0\s*,') {
            $ok = $false
        }
    }
    if (-not $ok) {
        $bad += ("T5 SetWindowTheme outside the DI passthrough at " + $ln.File + ":" + $ln.Line +
                 " -> RTL 不许替用户关/开 tooltip 主题, 且桩不许把 0 换空串 (" +
                 $ln.Text.Substring(0, [Math]::Min(90, $ln.Text.Length)) + ")")
    }
}

# ---- T6: tooltip 宿主台账恰好 2 站 ----
$hosts = @($rtl | Where-Object { $_.Text -match 'TOOLTIPS_CLASS' -and $_.Text -match 'CreateWindowEx' })
$hostFiles = @($hosts | ForEach-Object { $_.File } | Sort-Object -Unique)
if ($hosts.Count -ne 2 -or ($hostFiles -join ",") -ne "vb6forms_slider.c,vb6forms_style.c") {
    $bad += ("T6 tooltip hosts = " + $hosts.Count + " (" + ($hostFiles -join ",") +
             ") want 2 (vb6forms_style.c 共享宿主 + vb6forms_slider.c 滑块气泡) -> " +
             (($hosts | ForEach-Object { $_.File + ":" + $_.Line }) -join " | "))
}

if ($bad.Count -eq 0) {
    Write-Host ("PASS tooltip authority: StrPtr 唯一出口 1 / tooltip 颜色散点 " + $tooltipColorHits +
                " / tooltip 字体散点 " + $tooltipFontHits + " / SetWindowTheme 站 " + $themeHits +
                " / 宿主 " + $hosts.Count) -ForegroundColor Green
    exit 0
}
foreach ($b in $bad) { Write-Host ("FAIL " + $b) -ForegroundColor Red }
exit 1
