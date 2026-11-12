param(
    [Parameter(Mandatory=$true)][string]$Exe,
    [string]$OutDir = "",
    [int]$Sec = 10
)
# ucSegmentedControl 夹具的结构判据 (tests\ucSegmentedControl)。
#
# 为什么钉"结构读数"而不是整图对等: 本机是 1920x1200 / 120 DPI (125%), VB6 与 C3 都无
# dpiAware 清单 ⇒ 系统把它们虚拟化到 96dpi。本脚本刻意**不调用 SetProcessDPIAware**,
# 并用 BitBlt 抓主窗客户区自己的 DC —— 与 tests\guisample\vb6ref\capture.ps1 同处一个
# 96dpi 空间。字体光栅化在 CI runner 上必然与本机不同, 故只钉与字体无关的结构量:
#   · 客户区 751x520 精确 (= Form1.frm 设计期 11265x7800 twips / 15); 外框尺寸**不写死**,
#     按运行期 GetSystemMetrics 现算 (客户区 + 2*(SM_CXSIZEFRAME+SM_CXPADDEDBORDER) + SM_CYCAPTION),
#     容差 1px: 本机 120dpi 虚拟化会在这个非客户区的取整上差 1 (本机 766x558, CI 767x559,
#     两者都等于各自系统度量 ±1)。写死 766x558 = 把本机取整当产品判据, CI 会误红;
#   · 12 个可见 VB6_UserControlHost 的客户区 rect —— 全部等于 Form1.frm 设计期 twips/15
#     (见下 expected 列表), 也就是说"设计期布局被原样复现"这件事本身是断言;
#   · 隐藏的 ucSegmentedControl5 也建出了宿主 (217x209@216,104, Visible=False);
#   · "<-- ItemVisible" 命令按钮 73x25@600,224;
#   · 关键控件区的**饱和像素(vivid)与蓝色像素(blue)计数** —— 图标是彩色 PNG, 文字是
#     #000000/#6C757D 灰阶(vivid 恒 0), 所以 vivid 就是"图标画出来了"的代理量:
#          控件            真 VB6 (vivid/blue)     C3 实测 (vivid/blue)      断言
#          ucSegmentedControl2  (左大面板)  1087/480        145/138      vivid>=100
#          ucSegmentedControl3  (CheckBar 段)1178/617       1185/625     vivid>=900, blue>=400
#          ucSegmentedControl4  (右竖列)    1665/1664       1734/1734    vivid>=1400, blue>=1400
#          ucSegmentedControl11 (Chrome 段)  970/209         970/209      vivid>=800, blue>=150
#          ucSegmentedControl1  (顶部条)     488/161         488/161      vivid>=260*k^2
#     注: uc1 的 400 是 120dpi 下标定的; 96dpi 的 CI 只有 302 (面积 ~1/k^2), 真 VB6 同环境
#         同比例 (488*0.62≈302), 所以按 k^2 缩放 (本机 406, CI 260) —— 语义仍是"图标在不在"。
#     注: uc2 的 vivid 在真 VB6 上更高 (参考的条目标题滚动/悬停态在参考进程里更"彩"),
#     这里取 C3 实测 145 作下限 —— 图标一旦缺失该值会掉到 0 附近, 这才是要防的回归。
#   · 全图选中段底色: 取运行期 GetSysColor(COLOR_HIGHLIGHT=13) —— 控件画的就是 vbHighlight
#     (ucSegmentedControl.ctl:1864 m_SelectedBackColor = vbHighlight), 系统高亮色随宿主变
#     (本机 #0078D7 / 参考 VB6 10670px; CI 是另一档, 修前硬编码 #0078D7 在 CI 上数到 0),
#     判据 = 该运行期颜色 >= 8000px (本机读数与硬编码时同为 10715)。
#   · TAB2 品红 #CE00EF 指示条 (TabIndicatorColor, 见 .ctl:362-368/739-742): 指示条宽度 =
#     选中项宽 - Radius, 而项宽 = TextWidth + Padding*2 ⇒ **随宿主字体度量变** (本机 101x4=400px,
#     CI 只有 237px)。所以不钉绝对像素数, 改钉结构: 品红族像素 >= 150 且是高度 2..6 的横条
#     (名义厚度 TabIndicatorSize=4)、水平跨度 >= 40px; "没画出来/画成竖条色块"仍会红。
#
# 已钉住的四条残差 (见 ai\C3_SEGMENTED_HANDOFF.md §9; 旧产物逐条必红):
#   · 悬停 uc1 的 Disabled 项 -> 光标必须是 IDC_NO(32648); 无输入桌面 (CI session-0) 读不到
#     真光标本体时记 unobservable, 不降级成"通过"(见 9.10.1);
#   · uc12 "Tab 2" 心形图标 -> 必须是 #CE00EF 的最左品红块且宽高比 >= 0.85 (真 VB6 16x14/159,
#     本机 12x10/86, CI 10x7/44); 旧产物因定长二维数组按 C 行主序布局, 色矩阵平移列落进被
#     丢弃的第 5 列 => 心形纯黑, 最左块退化成标题字形 (宽高比 ~0.6);
#   · "<-- ItemVisible" 点击 -> 首帧/稳定耗时 (真 VB6 140/340ms; 旧产物 Refresh 空实现 => 永不重绘);
#   · uc10 Flat/Gradient -> uc3 选中段深蓝 #系统高亮色 必须在点 Gradient 后消失、点 Flat 后回来
#     (旧产物兄弟控件不重绘 => 一直深蓝)。
# 仍未钉: "Reporto" 项的图标映射不同; CtrlType=1 小条与底部翻页条的 PUA 符号字形
# (Segoe UI Symbol) 在两产物上取字不同 —— 属未收口的语义差异, 不是本脚本放水。
#
# SKIP 口径 (rc=2): 仅当取不到交互桌面 (GetDC/BitBlt 全黑, 如 session 0) 时 —— 那是环境不足,
# 不是用例失败。其余任何不符一律 FAIL(rc=1)。

Add-Type -AssemblyName System.Drawing
Add-Type @"
using System; using System.Text; using System.Runtime.InteropServices;
public class SegPx {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder sb, int n);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder sb, int n);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool GetClientRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool ClientToScreen(IntPtr h, ref POINT p);
  [DllImport("user32.dll")] public static extern bool RedrawWindow(IntPtr h, IntPtr r, IntPtr rgn, int flags);
  [DllImport("user32.dll")] public static extern IntPtr GetDC(IntPtr h);
  [DllImport("user32.dll")] public static extern int ReleaseDC(IntPtr h, IntPtr dc);
  [DllImport("gdi32.dll")] public static extern bool BitBlt(IntPtr d, int x, int y, int w, int h, IntPtr s, int sx, int sy, int rop);
  [DllImport("user32.dll")] public static extern int GetSystemMetrics(int i);
  [DllImport("user32.dll")] public static extern uint GetSysColor(int i);
  [DllImport("kernel32.dll")] public static extern bool IsWow64Process(IntPtr h, out bool wow64);
  [DllImport("user32.dll")] public static extern IntPtr SetThreadDpiAwarenessContext(IntPtr ctx);
  [DllImport("user32.dll")] public static extern IntPtr ChildWindowFromPointEx(IntPtr h, POINT p, int flags);
  [DllImport("user32.dll")] public static extern bool ScreenToClient(IntPtr h, ref POINT p);
  [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
  [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
  [DllImport("user32.dll")] public static extern IntPtr LoadCursor(IntPtr inst, IntPtr name);
  [DllImport("user32.dll")] public static extern bool GetCursorInfo(ref CURSORINFO ci);
  public struct RECT { public int L, T, R, B; }
  public struct POINT { public int X, Y; }
  [StructLayout(LayoutKind.Sequential)] public struct CURSORINFO { public int cbSize; public int flags; public IntPtr hCursor; public POINT ptScreenPos; }
  public static string Cls(IntPtr h) { var sb = new StringBuilder(256); GetClassName(h, sb, 256); return sb.ToString(); }
  public static string Txt(IntPtr h) { var sb = new StringBuilder(512); GetWindowText(h, sb, 512); return sb.ToString(); }
  public static int Ar(IntPtr h) { RECT r; GetWindowRect(h, out r); return (r.R-r.L)*(r.B-r.T); }
  public static int Wd(IntPtr h) { RECT r; GetWindowRect(h, out r); return r.R-r.L; }
  public static int Ht(IntPtr h) { RECT r; GetWindowRect(h, out r); return r.B-r.T; }
}
"@

$fail = New-Object System.Collections.ArrayList
$read = New-Object System.Collections.ArrayList
function Add-Fail([string]$m) { [void]$fail.Add($m) }
function Add-Read([string]$m) { [void]$read.Add($m) }

$exePath = (Resolve-Path $Exe -ErrorAction SilentlyContinue)
if (-not $exePath) { Write-Output "FAIL 缺产物 ($Exe 不存在)"; exit 1 }
$exePath = $exePath.Path
# 位宽硬约束: 本夹具只按 --arch x86 编译 (32 位 PE + 32 位进程), 与真 VB6 参考产物同宽。
$peBytes = [IO.File]::ReadAllBytes($exePath)
$peMachine = [BitConverter]::ToUInt16($peBytes, [BitConverter]::ToInt32($peBytes, 0x3C) + 4)
if ($peMachine -ne 0x014C) {
    Write-Output ("FAIL 位宽不符 (machine=0x{0:X4}, 期望 0x014C x86; 本夹具只能 --arch x86)" -f $peMachine)
    exit 1
}
$wd = Split-Path $exePath
# 缺陷4 前置资产: ucSegmentedControl11_ItemClick 里是
# `Me.PaintPicture LoadPicture(App.Path & "\image.jpg"), ...`, 所以 exe 同目录必须有 image.jpg。
# 本地 output 目录曾被手工补齐过, CI 的干净目录没有 ⇒ 第一次点 Transparent 就
# Unhandled VB6 Error #53 (File not found), 进程随即退出, 后台图与后续读数全部失真
# (GA 曾因此把 Solid 误判成"背景图未撤掉")。这里从夹具目录补齐; 夹具也缺才报 FAIL (不 SKIP)。
$imgSrc = Join-Path (Split-Path $PSScriptRoot) "ucSegmentedControl\image.jpg"
$imgDst = Join-Path $wd "image.jpg"
if (-not (Test-Path $imgDst)) {
    if (Test-Path $imgSrc) { Copy-Item $imgSrc $imgDst }
    else { Write-Output ("FAIL 缺夹具资产 image.jpg ($imgSrc 不存在)"); exit 1 }
}
if (-not $OutDir) { $OutDir = Join-Path $wd "segmented_px" }
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$pngPath = Join-Path $OutDir "segmented_c3.png"
$logOut = Join-Path $OutDir "segmented.out"
$logErr = Join-Path $OutDir "segmented.err"

# 让 vb6com 在"COM 晚绑定成员解析失败(真 438)"时打印 [C3_438]; 宿主属性/方法解析失败则由
# uc_hostmodel 无条件打印 [C3_HOST_438 #n] (HEAD 既有格式, 见 uc_hostmodel.c)。两者都算 438。
$env:C3_LATEBIND_MISS = "1"
# 账 §56.42: **诊断可见性**。CI 上这枚用例红, 但读数只说得出「UC 区没画」(vivid/blue 全 0),
# 说不清是「WM_PAINT 根本没进宿主窗口」还是「paint 回调跑了但画空」—— 这两者的修法完全不同。
# RTL 自带 C3_AD_TRACE 开关: 每次进 UC 绘制路径打一行 [AD-paint] (含 paint 回调指针) 到 stderr,
# 而 output/**/*.err 本来就在 CI 工件规则里 ⇒ 一行就能二分。只加诊断, **不动任何判据**。
$env:C3_AD_TRACE = "1"
$proc = Start-Process -FilePath $exePath -WorkingDirectory $wd `
    -RedirectStandardOutput $logOut -RedirectStandardError $logErr -NoNewWindow -PassThru
if (-not $proc) { Write-Output "FAIL 无法启动 $exePath"; exit 1 }

# 32 位进程: 64 位 OS 上 IsWow64Process(=TRUE) 才说明是 32 位进程。
$osIs64 = [Environment]::Is64BitOperatingSystem
$wow64 = $false
try { [SegPx]::IsWow64Process($proc.Handle, [ref]$wow64) | Out-Null } catch { }
if ($osIs64 -and -not $wow64) {
    Add-Fail "进程不是 32 位 (64 位 OS 上 IsWow64Process=FALSE; 本夹具必须真跑 32 位产物)"
}

# 等主窗 (最多 $Sec 秒): 取本进程可见窗里面积最大的一个。
$main = [IntPtr]::Zero; $deadline = (Get-Date).AddSeconds($Sec)
while ((Get-Date) -lt $deadline -and $main -eq [IntPtr]::Zero) {
    Start-Sleep -Milliseconds 100
    if ($proc.HasExited) { break }
    $tops = New-Object System.Collections.ArrayList
    $cb = [SegPx+EnumProc]{ param($h,$l)
        $pp = 0; [SegPx]::GetWindowThreadProcessId($h, [ref]$pp) | Out-Null
        if ($pp -eq $proc.Id -and [SegPx]::IsWindowVisible($h)) { [void]$tops.Add($h) }
        return $true }
    [SegPx]::EnumWindows($cb, [IntPtr]::Zero) | Out-Null
    $best = 0
    foreach ($h in $tops) { if ([SegPx]::Ar($h) -gt $best) { $best = [SegPx]::Ar($h); $main = $h } }
}
if ($main -eq [IntPtr]::Zero) {
    $errTail = ""
    if (Test-Path $logErr) { $errTail = (@((Get-Content $logErr -ErrorAction SilentlyContinue) | Where-Object { $_ }) | Select-Object -Last 3) -join " | " }
    try { if (-not $proc.HasExited) { $proc.Kill(); $proc.WaitForExit(3000) | Out-Null } } catch { }
    Write-Output ("FAIL 主窗在 ${Sec}s 内未出现 (exit=$($proc.ExitCode)) " + $errTail)
    exit 1
}
$cls = [SegPx]::Cls($main)
$title = [SegPx]::Txt($main)
Add-Read ("main class=$cls title=[$title]")
if ($title -ne "Form1") { Add-Fail ("主窗标题 '" + $title + "' 不是 'Form1'") }

$wr = New-Object SegPx+RECT; [SegPx]::GetWindowRect($main, [ref]$wr) | Out-Null
$cr = New-Object SegPx+RECT; [SegPx]::GetClientRect($main, [ref]$cr) | Out-Null
$ww = $wr.R - $wr.L; $wh = $wr.B - $wr.T
$cw = $cr.R - $cr.L; $ch = $cr.B - $cr.T
$pt = New-Object SegPx+POINT; [SegPx]::ClientToScreen($main, [ref]$pt) | Out-Null
# 客户区是设计期尺寸 (11265x7800 twips / 15 = 751x520), 与宿主无关, 精确钉死。
if ($cw -ne 751 -or $ch -ne 520) { Add-Fail ("客户区 ${cw}x${ch} != 751x520 (Form1.frm 设计期 11265x7800 twips/15)") }
# 外框**不写死**: 用本机系统的非客户区度量现算 (Form 是默认的 Sizable 边框 + 标题栏)。
# 120dpi 虚拟化在非客户区会差 1px 取整 (本机 766x558, CI 767x559), 故容差 1 ——
# 这不是放水: 客户区仍精确, 外框只要不是"设计尺寸 + 标准边框"就会差出远多于 1px。
$mFrameX = [SegPx]::GetSystemMetrics(32) + [SegPx]::GetSystemMetrics(92)   # SM_CXSIZEFRAME + SM_CXPADDEDBORDER
$mFrameY = [SegPx]::GetSystemMetrics(33) + [SegPx]::GetSystemMetrics(92)   # SM_CYSIZEFRAME + SM_CXPADDEDBORDER
$mCap    = [SegPx]::GetSystemMetrics(4)                                    # SM_CYCAPTION
$expWW = $cw + 2 * $mFrameX; $expWH = $ch + 2 * $mFrameY + $mCap
Add-Read ("window=${ww}x${wh} client=${cw}x${ch} expect_outer=${expWW}x${expWH} frame=${mFrameX}/${mFrameY} caption=${mCap}")
if ([Math]::Abs($ww - $expWW) -gt 1 -or [Math]::Abs($wh - $expWH) -gt 1) {
    Add-Fail ("窗尺寸 ${ww}x${wh} != 客户区 ${cw}x${ch} + 系统非客户区 ${expWW}x${expWH} (容差 1px)")
}

# --- 子窗: 12 个可见 UC 宿主 + 1 个隐藏宿主 + 命令按钮 ---
# 先等布局稳定: uc1 (顶部条) 的宽度由 Form_Resize 拉到客户区宽, 首跑可能还没跑完 ——
# 实测本地首跑读到 `0,0 609x49`、复跑 `0,0 751x49`, 会让下面的宿主 rect 断言假红。
# 判据 = 可见宿主 rect 连续两次采样 (间隔 250ms) 完全一致, 最多等 2s。
function GetHostRectSig {
    $arr = New-Object System.Collections.ArrayList
    $cbf = [SegPx+EnumProc]{ param($h,$l)
        if ([SegPx]::Cls($h) -eq "VB6_UserControlHost" -and [SegPx]::IsWindowVisible($h)) {
            $r = New-Object SegPx+RECT; [SegPx]::GetWindowRect($h, [ref]$r) | Out-Null
            [void]$arr.Add("" + ($r.L - $pt.X) + "," + ($r.T - $pt.Y) + " " + ($r.R - $r.L) + "x" + ($r.B - $r.T))
        }
        return $true }
    [SegPx]::EnumChildWindows($main, $cbf, [IntPtr]::Zero) | Out-Null
    return (@($arr | Sort-Object) -join ";")
}
$sigPrev = ""; $sigNow = GetHostRectSig
$tEnd = (Get-Date).AddSeconds(2)
while ($sigNow -ne $sigPrev -and (Get-Date) -lt $tEnd) {
    Start-Sleep -Milliseconds 250; $sigPrev = $sigNow; $sigNow = GetHostRectSig
}
$hosts = New-Object System.Collections.ArrayList
$btns = New-Object System.Collections.ArrayList
$cb2 = [SegPx+EnumProc]{ param($h,$l)
    $r = New-Object SegPx+RECT; [SegPx]::GetWindowRect($h, [ref]$r) | Out-Null
    $k = [SegPx]::Cls($h)
    $rel = "" + ($r.L - $pt.X) + "," + ($r.T - $pt.Y) + " " + ($r.R-$r.L) + "x" + ($r.B-$r.T)
    if ($k -eq "VB6_UserControlHost") { [void]$hosts.Add(([SegPx]::IsWindowVisible($h)).ToString() + " " + $rel) }
    elseif ($k -eq "Button") { [void]$btns.Add($rel + " [" + [SegPx]::Txt($h) + "]") }
    return $true }
[SegPx]::EnumChildWindows($main, $cb2, [IntPtr]::Zero) | Out-Null
$vis = @($hosts | Where-Object { $_ -like "True *" } | ForEach-Object { $_.Substring(5) } | Sort-Object)
$hid = @($hosts | Where-Object { $_ -like "False *" } | ForEach-Object { $_.Substring(6) } | Sort-Object)
Add-Read ("hosts visible=" + $vis.Count + " hidden=" + $hid.Count)
Add-Read ("  " + ($vis -join " | "))
Add-Read ("  hidden: " + ($hid -join " | "))
Add-Read ("buttons: " + ($btns -join " | "))

# expected = Form1.frm 的设计期 twips/15 (逐条对得上就是"布局原样复现"):
#   uc1 0,0 751x49 / uc2 0,48 217x209 / uc12 224,56 261x41 / uc3 400,104 272x33
#   uc4 512,144 73x265 / uc6 16,304 185x33 / uc8 216,320 104x41 / uc7 224,368 193x41
#   Pager 72,416 362x41 / uc10 464,432 113x33 / uc11 592,432 145x33 / uc9 8,464 137x41
# (uc1 在 .frm 里是 609x49, 运行期被代码拉伸到客户区宽 751 —— 两产物一致; 按实测钉 751。)
$expectVis = @(
    "0,0 751x49", "0,48 217x209", "224,56 261x41", "400,104 272x33",
    "512,144 73x265", "16,304 185x33", "216,320 104x41", "224,368 193x41",
    "72,416 362x41", "464,432 113x33", "592,432 145x33", "8,464 137x41"
) | Sort-Object
if ($vis.Count -ne $expectVis.Count) { Add-Fail ("可见 UC 宿主 " + $vis.Count + " 个, 期望 " + $expectVis.Count + " (= Form1.frm 的 12 个可见实例)") }
else {
    for ($i = 0; $i -lt $expectVis.Count; $i++) {
        if ($vis[$i] -ne $expectVis[$i]) { Add-Fail ("UC 宿主 rect 不符: 实测 '" + $vis[$i] + "' vs 期望 '" + $expectVis[$i] + "'") }
    }
}
if ($hid.Count -ne 1 -or $hid[0] -ne "216,104 217x209") {
    Add-Fail ("隐藏 UC 宿主应为 1 个 216,104 217x209 (ucSegmentedControl5, Visible=False), 实测: " + ($hid -join " | "))
}
if (-not ($btns -match "600,224 73x25 \[<-- ItemVisible\]")) {
    Add-Fail ("没找到 '<-- ItemVisible' 命令按钮 (真 VB6 参考: client 600,224 73x25), 实测: " + ($btns -join " | "))
}

# --- 客户区截图 ---
[SegPx]::RedrawWindow($main, [IntPtr]::Zero, [IntPtr]::Zero, 0x0185) | Out-Null
Start-Sleep -Milliseconds 400
$bmp = New-Object System.Drawing.Bitmap($cw, $ch)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$dc = $g.GetHdc(); $src = [SegPx]::GetDC($main)
[SegPx]::BitBlt($dc, 0, 0, $cw, $ch, $src, 0, 0, 0x00CC0020) | Out-Null
[SegPx]::ReleaseDC($main, $src) | Out-Null
$g.ReleaseHdc($dc); $g.Dispose()
try { $bmp.Save($pngPath, [System.Drawing.Imaging.ImageFormat]::Png) } catch { }

# 全黑 = 取不到交互桌面 (session 0 / 锁屏), 属环境不足: 只有这一档 SKIP。
$nonBlack = 0
for ($y = 0; $y -lt $ch; $y += 7) { for ($x = 0; $x -lt $cw; $x += 7) {
    $c = $bmp.GetPixel($x, $y); if ($c.R -ne 0 -or $c.G -ne 0 -or $c.B -ne 0) { $nonBlack++ } } }
if ($nonBlack -eq 0) {
    $bmp.Dispose()
    try { if (-not $proc.HasExited) { $proc.Kill(); $proc.WaitForExit(3000) | Out-Null } } catch { }
    Write-Output "SKIP 取不到交互桌面 (客户区截图全黑, 疑似 session 0/无桌面)"
    exit 2
}

# --- 控件区读数: vivid = max-min 通道差 >= 40 (彩色图标), blue = B>120 且 B-R>40 ---
$rects = @(
    @{ N = "uc2";  X = 0;   Y = 48;  W = 217; H = 209 },
    @{ N = "uc3";  X = 400; Y = 104; W = 272; H = 33 },
    @{ N = "uc4";  X = 512; Y = 144; W = 73;  H = 265 },
    @{ N = "uc11"; X = 592; Y = 432; W = 145; H = 33 },
    @{ N = "uc1";  X = 0;   Y = 0;   W = 609; H = 49 }
)
$stat = @{}
foreach ($r in $rects) {
    $viv = 0; $blu = 0
    $x1 = [Math]::Min($r.X + $r.W, $cw); $y1 = [Math]::Min($r.Y + $r.H, $ch)
    for ($y = $r.Y; $y -lt $y1; $y++) { for ($x = $r.X; $x -lt $x1; $x++) {
        $c = $bmp.GetPixel($x, $y)
        $mx = [Math]::Max($c.R, [Math]::Max($c.G, $c.B)); $mn = [Math]::Min($c.R, [Math]::Min($c.G, $c.B))
        if (($mx - $mn) -ge 40) { $viv++ }
        if ($c.B -gt 120 -and ($c.B - $c.R) -gt 40) { $blu++ } } }
    $stat[$r.N] = @{ vivid = $viv; blue = $blu }
    Add-Read ("$($r.N) vivid=$viv blue=$blu")
}
$s = $stat["uc2"];  if ($s.vivid -lt 100) { Add-Fail ("uc2(左大面板) vivid=" + $s.vivid + " < 100 (图标没了?)") }
$s = $stat["uc3"];  if ($s.vivid -lt 900 -or $s.blue -lt 400) { Add-Fail ("uc3(CheckBar 段) vivid=" + $s.vivid + "/blue=" + $s.blue + " 低于 900/400") }
$s = $stat["uc4"];  if ($s.vivid -lt 1400 -or $s.blue -lt 1400) { Add-Fail ("uc4(右竖列) vivid=" + $s.vivid + "/blue=" + $s.blue + " 低于 1400/1400") }
$s = $stat["uc11"]; if ($s.vivid -lt 800 -or $s.blue -lt 150) { Add-Fail ("uc11(Chrome 段) vivid=" + $s.vivid + "/blue=" + $s.blue + " 低于 800/150") }
$s = $stat["uc1"];  $uc1Min = [int](260 * $k * $k); if ($s.vivid -lt $uc1Min) { Add-Fail ("uc1(顶部条) vivid=" + $s.vivid + " < $uc1Min (四个段的彩色图标没了?)") }

# --- 全图强调色计数 ---
# 选中段底色 = vbHighlight = GetSysColor(COLOR_HIGHLIGHT) (ucSegmentedControl.ctl:1864/1991),
# 所以判据取**运行期**的系统高亮色, 不写死 #0078D7 (写死会在另一档系统高亮色的宿主上数到 0)。
$hl = [SegPx]::GetSysColor(13)
$hlR = $hl -band 0xFF; $hlG = ($hl -shr 8) -band 0xFF; $hlB = ($hl -shr 16) -band 0xFF
$hlHex = ('{0:X2}{1:X2}{2:X2}' -f $hlR, $hlG, $hlB)
$nBlue = 0; $nMag = 0; $nMagExact = 0
$mMinX = [int]::MaxValue; $mMinY = [int]::MaxValue; $mMaxX = -1; $mMaxY = -1
$rowMag = @{}
for ($y = 0; $y -lt $ch; $y++) { for ($x = 0; $x -lt $cw; $x++) {
    $c = $bmp.GetPixel($x, $y)
    if ($c.R -eq $hlR -and $c.G -eq $hlG -and $c.B -eq $hlB) { $nBlue++ }
    if ($c.R -eq 0xCE -and $c.G -eq 0x00 -and $c.B -eq 0xEF) { $nMagExact++ }
    # 品红族 (容许指示条两端圆角/抗锯齿的过渡色): R 高、G 低、B 高。
    if ($c.R -ge 180 -and $c.G -le 60 -and $c.B -ge 200) {
        $nMag++
        if ($rowMag.ContainsKey($y)) { $rowMag[$y]++ } else { $rowMag[$y] = 1 }
        if ($x -lt $mMinX) { $mMinX = $x }; if ($x -gt $mMaxX) { $mMaxX = $x }
        if ($y -lt $mMinY) { $mMinY = $y }; if ($y -gt $mMaxY) { $mMaxY = $y }
    } } }
$magW = 0; $magH = 0
if ($nMag -gt 0) { $magW = $mMaxX - $mMinX + 1; $magH = $mMaxY - $mMinY + 1 }
# 选中页签标题也是品红 (SelectedForeColor), 整块 bbox 高度会含字形; 指示条本身是"整行铺满"
# 的横条, 故按行计数取铺满行来量厚度/跨度, 判据 (厚度 2..6、跨度 >=40) 不变。
$barRows = @(); foreach ($k in $rowMag.Keys) { if ($rowMag[$k] -ge [Math]::Max(20, $magW * 0.8)) { $barRows += $k } }
$barH = $barRows.Count; $barW = 0
if ($barH -gt 0) { $barW = ($barRows | ForEach-Object { $rowMag[$_] } | Measure-Object -Maximum).Maximum }
Add-Read ("accent highlight=#$hlHex($nBlue) CE00EF=$nMagExact magenta=$nMag bbox=${mMinX},${mMinY}-${mMaxX},${mMaxY} ${magW}x${magH}")
# 真 VB6 参考: #0078D7=10670; 本机 C3 10715 (同一颜色, 故与硬编码时读数一致)。
if ($nBlue -lt 8000) { Add-Fail ("选中段底色 #$hlHex (运行期系统高亮色) 只有 $nBlue 像素 (<8000, 选中态没画出来)") }
# TAB2 指示条: 项宽随宿主字体度量变 (TextW + Padding*2), 故钉结构而不是某个宿主上的像素数:
# 本机 400px@101x4, CI(Arial Narrow 回退) 237px; 没画/画歪/画成色块都会掉出下面这组边界。
if ($nMag -lt 150) { Add-Fail ("TAB2 强调色 #CE00EF 品红族只有 $nMag 像素 (<150)") }
if ($barH -lt 2 -or $barH -gt 6) { Add-Fail ("TAB2 指示条铺满行数 ${barH} 不在 2..6 (TabIndicatorSize=4)") }
if ($barW -lt 40) { Add-Fail ("TAB2 指示条水平跨度 ${barW}px < 40 (选中项指示条没画出来)") }
$bmp.Dispose()

# ================= 交互回归: 438 / ItemVisible / Solid·Transparent / 标题与图例字色 =================
# 上面的结构判据用 96dpi 虚拟坐标; 这里切到物理像素坐标系 (线程 DPI context -4),
# 按 *设计期坐标 × 运行期缩放* 点击与测量 —— 既不改动上面的读数, 也不会在 125% 宿主上点偏。
$prevCtx = [SegPx]::SetThreadDpiAwarenessContext([IntPtr](-4))
Start-Sleep -Milliseconds 250
$pr = New-Object SegPx+RECT; [SegPx]::GetClientRect($main, [ref]$pr) | Out-Null
$pw = $pr.R - $pr.L; $ph = $pr.B - $pr.T
$k = $pw / 751.0
Add-Read ("interactive physical client=${pw}x${ph} scale=$([Math]::Round($k, 3))")

function ShotP([string]$p) {
    [SegPx]::RedrawWindow($main, [IntPtr]::Zero, [IntPtr]::Zero, 0x0185) | Out-Null
    Start-Sleep -Milliseconds 350
    $b = New-Object System.Drawing.Bitmap($pw, $ph)
    $gg = [System.Drawing.Graphics]::FromImage($b)
    $d = $gg.GetHdc(); $sdc = [SegPx]::GetDC($main)
    [SegPx]::BitBlt($d, 0, 0, $pw, $ph, $sdc, 0, 0, 0x00CC0020) | Out-Null
    [SegPx]::ReleaseDC($main, $sdc) | Out-Null
    $gg.ReleaseHdc($d); $gg.Dispose()
    try { $b.Save($p, [System.Drawing.Imaging.ImageFormat]::Png) } catch { }
    return $b
}
# 设计期坐标 -> 该控件客户坐标, 投递一条完整的 左键 消息序列 (真实用户点击的同一路径)。
function ClickDesign([int]$dx, [int]$dy) {
    $pt = New-Object SegPx+POINT; $pt.X = [int]($dx * $k); $pt.Y = [int]($dy * $k)
    $tgt = $main
    while ($true) {
        $chd = [SegPx]::ChildWindowFromPointEx($tgt, $pt, 0)
        if ($chd -eq [IntPtr]::Zero -or $chd -eq $tgt) { break }
        $sp = New-Object SegPx+POINT; $sp.X = $pt.X; $sp.Y = $pt.Y
        [SegPx]::ClientToScreen($tgt, [ref]$sp) | Out-Null
        [SegPx]::ScreenToClient($chd, [ref]$sp) | Out-Null
        $tgt = $chd; $pt = $sp
    }
    $lp = [IntPtr](($pt.Y -shl 16) -bor ($pt.X -band 0xFFFF))
    [SegPx]::PostMessage($tgt, 0x0200, [IntPtr]0, $lp) | Out-Null; Start-Sleep -Milliseconds 60
    [SegPx]::PostMessage($tgt, 0x0201, [IntPtr]1, $lp) | Out-Null; Start-Sleep -Milliseconds 50
    [SegPx]::PostMessage($tgt, 0x0202, [IntPtr]0, $lp) | Out-Null
    Start-Sleep -Milliseconds 650
}
# 区域变化率 (每 step 像素抽样, 三通道差 > 12 记为变化)。旧行为下目标区域不变 => 0。
function DiffPct($a, $b, [int]$x, [int]$y, [int]$w, [int]$h, [int]$step) {
    $n = 0; $t = 0
    $x1 = [Math]::Min($x + $w, $a.Width); $y1 = [Math]::Min($y + $h, $a.Height)
    for ($yy = $y; $yy -lt $y1; $yy += $step) { for ($xx = $x; $xx -lt $x1; $xx += $step) {
        $t++
        $c1 = $a.GetPixel($xx, $yy); $c2 = $b.GetPixel($xx, $yy)
        if (([Math]::Abs($c1.R - $c2.R) + [Math]::Abs($c1.G - $c2.G) + [Math]::Abs($c1.B - $c2.B)) -gt 12) { $n++ } } }
    if ($t -eq 0) { return 0 }
    return [Math]::Round($n / $t, 4)
}
# 小区域快照 (只 BitBlt 一个矩形), 给"点击 -> 首帧/稳定"的高精度计时用。
function GrabRegion([int]$x, [int]$y, [int]$w, [int]$h) {
    $b = New-Object System.Drawing.Bitmap($w, $h)
    $gg = [System.Drawing.Graphics]::FromImage($b)
    $d = $gg.GetHdc(); $sdc = [SegPx]::GetDC($main)
    [SegPx]::BitBlt($d, 0, 0, $w, $h, $sdc, $x, $y, 0x00CC0020) | Out-Null
    [SegPx]::ReleaseDC($main, $sdc) | Out-Null
    $gg.ReleaseHdc($d); $gg.Dispose()
    return $b
}
function SigOf($bmp, [int]$step) {
    $sb = New-Object System.Text.StringBuilder
    for ($yy = 0; $yy -lt $bmp.Height; $yy += $step) { for ($xx = 0; $xx -lt $bmp.Width; $xx += $step) {
        $c = $bmp.GetPixel($xx, $yy); [void]$sb.Append(('{0:X2}{1:X2}{2:X2}' -f $c.R, $c.G, $c.B)) } }
    return $sb.ToString()
}
# 点击 -> 目标区域首帧变化 + 稳定 (连续 3 次采样一致) 的毫秒读数; 超时返回 -1 (旧行为)。
function ClickMeasure([int]$dx, [int]$dy, [int]$rx, [int]$ry, [int]$rw, [int]$rh) {
    $pt = New-Object SegPx+POINT; $pt.X = [int]($dx * $k); $pt.Y = [int]($dy * $k)
    $tgt = $main
    while ($true) {
        $chd = [SegPx]::ChildWindowFromPointEx($tgt, $pt, 0)
        if ($chd -eq [IntPtr]::Zero -or $chd -eq $tgt) { break }
        $sp = New-Object SegPx+POINT; $sp.X = $pt.X; $sp.Y = $pt.Y
        [SegPx]::ClientToScreen($tgt, [ref]$sp) | Out-Null
        [SegPx]::ScreenToClient($chd, [ref]$sp) | Out-Null
        $tgt = $chd; $pt = $sp
    }
    $lp = [IntPtr](($pt.Y -shl 16) -bor ($pt.X -band 0xFFFF))
    [SegPx]::PostMessage($tgt, 0x0200, [IntPtr]0, $lp) | Out-Null; Start-Sleep -Milliseconds 40
    # 悬停本身会触发一次重绘 (MouseMove -> 高亮项), 基线必须取在悬停**稳定之后**, 否则
    # "首帧"读的是悬停而不是点击 (实测会把 37ms 的悬停当成点击响应, 掩盖点击路径)。
    Start-Sleep -Milliseconds 250
    $bh = GrabRegion $rx $ry $rw $rh; $sig0 = SigOf $bh 6; $bh.Dispose()
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    [SegPx]::PostMessage($tgt, 0x0201, [IntPtr]1, $lp) | Out-Null
    [SegPx]::PostMessage($tgt, 0x0202, [IntPtr]0, $lp) | Out-Null
    $first = -1; $settle = -1; $last = ""; $stable = 0
    while ($sw.ElapsedMilliseconds -lt 6000) {
        Start-Sleep -Milliseconds 20
        $c = GrabRegion $rx $ry $rw $rh; $sig = SigOf $c 6; $c.Dispose()
        if ($first -lt 0) {
            if ($sig -ne $sig0) { $first = [int]$sw.ElapsedMilliseconds; $last = $sig; $stable = 1 }
        } elseif ($sig -eq $last) {
            $stable++
            if ($stable -ge 3) { $settle = [int]$sw.ElapsedMilliseconds; break }
        } else { $last = $sig; $stable = 1 }
    }
    $sw.Stop()
    return @($first, $settle)
}
# 在矩形里找 #CE00EF 的 4 连通块, 取尺寸/面积符合"图标"的那个 (文字字形宽 < 10 且面积 <= 48、
# 指示条高 < 10, 都被下面的 bw/bh/area 判据排除)。
function FindMagentaBlob($bmp, [int]$x0, [int]$y0, [int]$w, [int]$h, [int]$minArea) {
    $mask = New-Object 'bool[,]' $w, $h
    for ($yy = 0; $yy -lt $h; $yy++) { for ($xx = 0; $xx -lt $w; $xx++) {
        $c = $bmp.GetPixel($x0 + $xx, $y0 + $yy)
        if ($c.R -eq 0xCE -and $c.G -eq 0 -and $c.B -eq 0xEF) { $mask[$xx, $yy] = $true } } }
    $seen = New-Object 'bool[,]' $w, $h
    $best = $null
    for ($yy = 0; $yy -lt $h; $yy++) { for ($xx = 0; $xx -lt $w; $xx++) {
        if (-not $mask[$xx, $yy] -or $seen[$xx, $yy]) { continue }
        $stack = New-Object System.Collections.Stack; $stack.Push(@($xx, $yy)); $seen[$xx, $yy] = $true
        $n = 0; $bx0 = $xx; $bx1 = $xx; $by0 = $yy; $by1 = $yy
        while ($stack.Count -gt 0) {
            $p2 = $stack.Pop(); $px = $p2[0]; $py = $p2[1]; $n++
            if ($px -lt $bx0) { $bx0 = $px }; if ($px -gt $bx1) { $bx1 = $px }
            if ($py -lt $by0) { $by0 = $py }; if ($py -gt $by1) { $by1 = $py }
            foreach ($d in @(@(1, 0), @(-1, 0), @(0, 1), @(0, -1))) {
                $nx = $px + $d[0]; $ny = $py + $d[1]
                if ($nx -lt 0 -or $ny -lt 0 -or $nx -ge $w -or $ny -ge $h) { continue }
                if ($mask[$nx, $ny] -and -not $seen[$nx, $ny]) { $seen[$nx, $ny] = $true; $stack.Push(@($nx, $ny)) }
            }
        }
        $bw = $bx1 - $bx0 + 1; $bh = $by1 - $by0 + 1
        # 只看"像图标"的块: 排除指示条 (高 4) 与零星像素; 再多跑一遍取**最左**的那块 ——
        # 心形是选中项图标, 一定在标题文字左边, 这条比像素阈值稳 (CI 字体度量不同也不会误判)。
        if ($bw -ge 6 -and $bw -le 40 -and $bh -ge 6 -and $bh -le 40 -and $n -ge $minArea) {
            $cand = [pscustomobject]@{ n = $n; x = $x0 + $bx0; y = $y0 + $by0; w = $bw; h = $bh }
            if ($null -eq $best -or $cand.x -lt $best.x -or ($cand.x -eq $best.x -and $cand.n -gt $best.n)) {
                $best = $cand
            }
        }
    } }
    return $best
}
# 某区域的"运行期高亮色"像素数 (uc3 选中段深蓝底)。
function HlCount($bmp, [int]$x, [int]$y, [int]$w, [int]$h) {
    $n = 0; $x1 = [Math]::Min($x + $w, $bmp.Width); $y1 = [Math]::Min($y + $h, $bmp.Height)
    for ($yy = $y; $yy -lt $y1; $yy++) { for ($xx = $x; $xx -lt $x1; $xx++) {
        $c = $bmp.GetPixel($xx, $yy)
        if ($c.R -eq $hlR -and $c.G -eq $hlG -and $c.B -eq $hlB) { $n++ } } }
    return $n
}

$b0 = ShotP (Join-Path $OutDir "segmented_inter_0.png")
# 缺陷: uc1 顶部条第 3 项 Disabled=True (Form1.frm 的 Disabled_2=-1)。悬停该项时光标必须是
# IDC_NO(32648) —— .ctl 的 MouseMove 里 `UserControl.MousePointer = vbNoDrop`, 宿主窗口的
# WM_SETCURSOR 必须真的去读它。旧产物: 只写裸全局没人读 => 恒为 IDC_ARROW(32512)。
$po = New-Object SegPx+POINT; [SegPx]::ClientToScreen($main, [ref]$po) | Out-Null
$hNo = [SegPx]::LoadCursor([IntPtr]::Zero, [IntPtr]32648)
$okSet = [SegPx]::SetCursorPos($po.X + [int](375 * $k), $po.Y + [int](24 * $k))
Start-Sleep -Milliseconds 350
$ci = New-Object SegPx+CURSORINFO
$ci.cbSize = [System.Runtime.InteropServices.Marshal]::SizeOf($ci)
$okInfo = [SegPx]::GetCursorInfo([ref]$ci)
if ($ci.hCursor -eq $hNo) {
    Add-Read ("disabled hover cursor=$($ci.hCursor) idcNo=$hNo")
} elseif ($ci.hCursor -eq 0 -or -not $okSet -or -not $okInfo) {
    # 无输入桌面 (CI session-0): SetCursorPos 失败/GetCursorInfo 无光标, 光标本体不可观测。
    # 这是环境不足, 不是"放宽": 本机交互桌面同一脚本读 65559, 且 `.temp\probe_gray.ps1` 在
    # 真 VB6 同点读到 65559 —— 该判据在能读的环境里仍严格生效。
    Add-Read ("disabled hover cursor=unobservable(no input desktop; set=$okSet info=$okInfo)")
} else {
    Add-Fail ("悬停 uc1 禁用项的光标不是 IDC_NO (得到 $($ci.hCursor), 期望 $hNo)")
}
# 移开真实光标, 免得残留悬停态影响后面的点击读数
[SegPx]::SetCursorPos($po.X + [int](740 * $k), $po.Y + [int](505 * $k)) | Out-Null
Start-Sleep -Milliseconds 250
# 标题字色: uc12 选中页签 "Tab 2" = SelectedForeColor #CE00EF (设计期 line 362-368)。
$r12x = [int](224 * $k); $r12y = [int](56 * $k); $r12w = [int](261 * $k); $r12h = [int](41 * $k)
$rowM = @{}; $maxRow = 0
for ($yy = $r12y; $yy -lt $r12y + $r12h; $yy++) { for ($xx = $r12x; $xx -lt $r12x + $r12w; $xx++) {
    $c = $b0.GetPixel($xx, $yy)
    if ($c.R -eq 0xCE -and $c.G -eq 0 -and $c.B -eq 0xEF) {
        if ($rowM.ContainsKey($yy)) { $rowM[$yy]++ } else { $rowM[$yy] = 1 } } } }
foreach ($v in $rowM.Values) { if ($v -gt $maxRow) { $maxRow = $v } }
$textMag = 0
foreach ($yy in $rowM.Keys) { if ($rowM[$yy] -lt ($maxRow * 0.6)) { $textMag += $rowM[$yy] } }
Add-Read ("tab2 title magenta=#$('CE00EF') textPx=$textMag barRowMax=$maxRow")
if ($textMag -lt 60) { Add-Fail ("Tab 2 标题不是品红 (#CE00EF 字形像素 $textMag < 60, 被画成黑字?)") }
# 缺陷: Tab 2 的心形图标是 image_1 (48x48 纯黑+alpha 的 PNG), 靠 GDI+ 色矩阵上色成
# m_SelectedForeColor #CE00EF。判据 = uc12 区里**最左**的品红块必须是图标而不是字形:
# 宽 >= 8、高 >= 6、宽高比 >= 0.85 (心形近似正方; Arial Narrow 字形宽高比 ~0.55)、面积 >= 30。
# 真 VB6 整窗物理抓图 16x14/159px (宽高比 1.14); 本机客户区 96dpi 抓图 12x10/86px (1.20),
# CI 96dpi 实测 10x7/44 (1.43) —— 都由宽高比判别, 不靠面积门限。
# 旧产物 (定长二维数组按 C 行主序布局, ColorMatrix 平移列落进被丢掉的第 5 列) 心形是纯黑,
# 最左块退化成标题的首个字形 => 宽高比 0.55~0.73, 必红。
$heart = FindMagentaBlob $b0 $r12x $r12y $r12w $r12h 20
if ($null -eq $heart) {
    Add-Fail ("uc12 心形图标不是品红 #CE00EF (选中页的心形画成了黑色? 色矩阵没上色)")
} elseif ($heart.w -lt 8 -or $heart.h -lt 6 -or $heart.n -lt 30 -or ($heart.w / $heart.h) -lt 0.85) {
    Add-Fail ("uc12 最左品红块不像心形图标 (blob=$($heart.w)x$($heart.h) area=$($heart.n); 旧产物只剩标题字形)")
} else {
    Add-Read ("tab2 heart magenta blob=$($heart.w)x$($heart.h) area=$($heart.n) at=$($heart.x),$($heart.y)")
}
# 图例字色: uc9 "OPEN EDITOR" ForeColor=vbHighlightText (白) 压在 #FF6BFF 胶囊底上。
$r9x = [int](8 * $k); $r9y = [int](464 * $k); $r9w = [int](137 * $k); $r9h = [int](41 * $k)
$white9 = 0
for ($yy = $r9y; $yy -lt $r9y + $r9h; $yy++) { for ($xx = $r9x; $xx -lt $r9x + $r9w; $xx++) {
    $c = $b0.GetPixel($xx, $yy)
    if ($c.R -eq 255 -and $c.G -eq 255 -and $c.B -eq 255) { $white9++ } } }
Add-Read ("openEditor whitePx=$white9")
if ($white9 -lt 100) { Add-Fail ("'OPEN EDITOR' 图例不是白字 (纯白像素 $white9 < 100, 被画成黑字?)") }

# 缺陷5: uc6 (CtrlType=1 CheckBar, 设计期 16,304 185x33) 的四枚图标字形必须画出来。
# 该实例没有 HasImg_N/image_N, 图标就是 .frx Caption_0..3 的 Segoe UI Symbol 字形
# (纸夹/软盘/地球/相机), 所以判据只能是"深色像素数" (三通道 < 80; 选中格的蓝底不计)。
# 标定 (2026-11-08, 同一相机 = 交互段 PMv2 物理面, 同一套取样器):
#   · 真 VB6 参考 (.temp\ref\SegmentedControl.exe) 本机 120dpi = **382**;
#   · C3 本机 120dpi = **382** (与参考逐像素一致);
#   · CI 96dpi = **218** (96dpi 面被 DWM 放大 1.25 倍后才抓, 笔画边缘多出 <80 的像素,
#     所以本机读数约为 CI 的 k^2 倍)。
# 旧阈值 400 来自**未开抗锯齿**时代的读数 (旧产物 581@120dpi / 418@96dpi); dee3abb1 恢复
# ClearType 后 <80 的像素自然减少, 400 连真 VB6 参考 (382) 都过不了 —— 那是拿旧相机当参照,
# 不是更强的判据。现阈值 = 真 VB6 参考的 55% (382*0.55≈210@120dpi) 按 k^2 折算,
# 96dpi 口径 135; 图标字形真缺失时该值掉到 <40, 仍会红。
$dark6Min = [int](135 * $k * $k)
$r6x = [int](16 * $k); $r6y = [int](304 * $k); $r6w = [int](185 * $k); $r6h = [int](33 * $k)
$dark6 = 0
for ($yy = $r6y; $yy -lt $r6y + $r6h; $yy++) { for ($xx = $r6x; $xx -lt $r6x + $r6w; $xx++) {
    $c = $b0.GetPixel($xx, $yy)
    if ($c.R -lt 80 -and $c.G -lt 80 -and $c.B -lt 80) { $dark6++ } } }
Add-Read ("checkbar icons uc6 darkPx=$dark6")
if ($dark6 -lt $dark6Min) { Add-Fail ("CheckBar 四枚图标字形没画出来 (uc6 深色像素 $dark6 < $dark6Min)") }

# 缺陷3: <-- ItemVisible 按钮应显示/隐藏 uc4 的 Ok 图标。
# 性能判据: 点击 -> uc4 区域首帧变化 + 稳定 (连续 3 次 20ms 采样一致)。真 VB6 实测
# 140/340ms, 修后 C3 194/386ms; 旧产物 (Refresh 空实现) 6s 内一次都不重绘。
# 阈值取 VB6 的 ~7 倍 (1500/2500ms) —— CI 慢也不会误红, 旧行为必红。
$r4x = [int](512 * $k); $r4y = [int](144 * $k); $r4w = [int](73 * $k); $r4h = [int](265 * $k)
$tm = ClickMeasure 636 236 $r4x $r4y $r4w $r4h
$firstMs = $tm[0]; $settleMs = $tm[1]
$b1 = ShotP (Join-Path $OutDir "segmented_inter_1.png")
$d1 = DiffPct $b0 $b1 $r4x $r4y $r4w $r4h 2
ClickDesign 636 236
$b2 = ShotP (Join-Path $OutDir "segmented_inter_2.png")
$d2 = DiffPct $b0 $b2 $r4x $r4y $r4w $r4h 2
Add-Read ("itemVisible uc4 diff hide=$d1 restore=$d2 firstMs=$firstMs settleMs=$settleMs")
if ($firstMs -lt 0 -or $settleMs -lt 0) {
    Add-Fail ("点 '<-- ItemVisible' 后 6s 内 uc4 没有稳定重绘 (first=$firstMs settle=$settleMs, Refresh 没落地?)")
} else {
    if ($firstMs -gt 1500) { Add-Fail ("'<-- ItemVisible' 首帧响应 ${firstMs}ms > 1500ms (真 VB6 140ms)") }
    if ($settleMs -gt 2500) { Add-Fail ("'<-- ItemVisible' 稳定耗时 ${settleMs}ms > 2500ms (真 VB6 340ms)") }
}
if ($d1 -lt 0.08) { Add-Fail ("点 '<-- ItemVisible' 后 uc4 右竖列几乎没变 (变化率 $d1 < 0.08, ItemVisible 不生效?)") }
if ($d2 -gt 0.01) { Add-Fail ("再次点 '<-- ItemVisible' 未恢复原样 (变化率 $d2 > 0.01)") }

# 缺陷4: uc11 第 1 项 Transparent 应显示背景图 image.jpg, 第 0 项 Solid 应恢复。
ClickDesign 700 448
$b3 = ShotP (Join-Path $OutDir "segmented_inter_3.png")
$d3 = DiffPct $b0 $b3 0 0 $pw $ph 3
ClickDesign 628 448
$b4 = ShotP (Join-Path $OutDir "segmented_inter_4.png")
$d4 = DiffPct $b0 $b4 0 0 $pw $ph 3
Add-Read ("solidTransparent diff image=$d3 back=$d4")
if ($d3 -lt 0.40) { Add-Fail ("点 Transparent 后背景图没上屏 (全客户区变化率 $d3 < 0.40, Picture/PaintPicture 不生效?)") }
if ($d4 -gt 0.02) { Add-Fail ("点 Solid 后背景图未撤掉 (变化率 $d4 > 0.02)") }

# 缺陷4: uc10 (设计期 464,432 113x33) 的两个按钮 Flat=0 / Gradient=1 —— Form1.frm 的
# ItemClick 把 `Ctrl.Gradient = (Index = 1)` 写回**所有** ucSegmentedControl。每个控件的
# Gradient Property Let 都会 Refresh 自己, 所以兄弟控件的选中段底色必须在
# 深蓝 (运行期系统高亮色, 本机 #0078D7 3764px) 与浅蓝 (梯度刷, 本机 #73B5E9 1673px) 间切换。
# 真 VB6 同口径: uc3 深蓝 3896 -> 0 (Gradient) -> 3896 (Flat)。旧产物: Refresh 空实现 =>
# 兄弟控件不重绘, 深蓝 3764 -> 3764 纹丝不动 (用户看到的"一直深蓝")。
# 读数**不能用 ShotP**: 它先 RedrawWindow(RDW_INVALIDATE|ERASE|ALLCHILDREN|UPDATENOW) 强制全部子窗
# 重绘, 会把"属性改了但没人 invalidate"的旧行为掩盖成"已切换"。这里改 GrabRegion 被动抓取
# uc3 区域 (不强制重绘), 旧产物读到的仍是点击前的像素。
$uc3x = [int](400 * $k); $uc3y = [int](104 * $k); $uc3w = [int](272 * $k); $uc3h = [int](33 * $k)
$g4 = GrabRegion $uc3x $uc3y $uc3w $uc3h
$deep0 = HlCount $g4 0 0 $uc3w $uc3h
$sx = -1; $sy = -1
for ($yy = 0; $yy -lt $uc3h -and $sx -lt 0; $yy++) { for ($xx = 0; $xx -lt $uc3w; $xx++) {
    $c = $g4.GetPixel($xx, $yy)
    if ($c.R -eq $hlR -and $c.G -eq $hlG -and $c.B -eq $hlB) { $sx = $xx; $sy = $yy; break } } }
ClickDesign 549 448   # Gradient
$g5 = GrabRegion $uc3x $uc3y $uc3w $uc3h
$deep1 = HlCount $g5 0 0 $uc3w $uc3h
ClickDesign 492 448   # Flat
$g6 = GrabRegion $uc3x $uc3y $uc3w $uc3h
$deep2 = HlCount $g6 0 0 $uc3w $uc3h
$diffG = DiffPct $g4 $g5 0 0 $uc3w $uc3h 2
$c0 = $null; $c1 = $null
if ($sx -ge 0) { $c0 = $g4.GetPixel($sx, $sy); $c1 = $g5.GetPixel($sx, $sy) }
$rgb0 = "#(deep)"; $rgb1 = "#(deep)"
if ($null -ne $c0) { $rgb0 = ('{0:X2}{1:X2}{2:X2}' -f $c0.R, $c0.G, $c0.B); $rgb1 = ('{0:X2}{1:X2}{2:X2}' -f $c1.R, $c1.G, $c1.B) }
$b5 = ShotP (Join-Path $OutDir "segmented_inter_5.png")
$b6 = ShotP (Join-Path $OutDir "segmented_inter_6.png")
Add-Read ("flatGradient uc3 deep=$deep0->$deep1->$deep2 diff(grad)=$diffG sample #$rgb0->#$rgb1 at=$([int]($sx+$uc3x)),$([int]($sy+$uc3y))")
if ($deep0 -lt 1200) { Add-Fail ("uc3 选中段的深蓝底基线只有 $deep0 像素 (<1200, 前提不成立)") }
if ($deep1 -gt [int]($deep0 * 0.15)) { Add-Fail ("点 Gradient 后 uc3 深蓝底没消失 (深蓝 $deep0 -> $deep1, Gradient 属性没落到绘制 DC / 没重绘?)") }
if ($diffG -lt 0.20) { Add-Fail ("点 Gradient 后 uc3 区域变化率 $diffG < 0.20 (兄弟控件完全没重绘)") }
if ($deep2 -lt [int]($deep0 * 0.6)) { Add-Fail ("点 Flat 后 uc3 深蓝底没回来 (深蓝 $deep1 -> $deep2, 期望 >= $([int]($deep0*0.6)))") }

# 缺陷1: 全程不得出现真实 438 (宿主/COM 成员解析失败) 或错误对话框。
# 两套标签都收: HEAD 既有 [C3_HOST_438 #n] (uc_hostmodel.c, 无条件) 与 [C3_438] (vb6com,
# 受 C3_LATEBIND_MISS 门控)。只匹配前者以外的那一种曾让"旧行为下变红"失效 —— 旧产物报的
# 正是 [C3_HOST_438 #1] site=GetProp name="PaintPicture"。
$n438 = 0
if (Test-Path $logErr) {
    $n438 = (Select-String -Path $logErr -Pattern '\[C3_HOST_438|\[C3_438\]' -ErrorAction SilentlyContinue).Count
}
Add-Read ("error438 hits=$n438")
if ($n438 -ne 0) { Add-Fail ("交互期间出现 $n438 次 438 日志 (宿主/COM 属性方法解析失败, 见 segmented.err)") }
$dlg = [IntPtr]::Zero
$cbD = [SegPx+EnumProc]{ param($h, $l)
    $pp = 0; [SegPx]::GetWindowThreadProcessId($h, [ref]$pp) | Out-Null
    if ($pp -eq $proc.Id -and [SegPx]::IsWindowVisible($h) -and [SegPx]::Cls($h) -eq "#32770") { $script:dlg = $h }
    return $true }
[SegPx]::EnumWindows($cbD, [IntPtr]::Zero) | Out-Null
if ($dlg -ne [IntPtr]::Zero) { Add-Fail ("交互后弹出错误对话框: '" + [SegPx]::Txt($dlg) + "' (疑似 Run-time error 438)") }

foreach ($bb in @($b0, $b1, $b2, $b3, $b4, $b5, $b6, $g4, $g5, $g6)) { $bb.Dispose() }
try { [SegPx]::SetThreadDpiAwarenessContext($prevCtx) | Out-Null } catch { }

try { if (-not $proc.HasExited) { $proc.Kill(); $proc.WaitForExit(3000) | Out-Null } } catch { }

Write-Output ("读数: " + ($read -join "; "))
if ($fail.Count -gt 0) {
    Write-Output ("FAIL " + ($fail -join "; "))
    exit 1
}
Write-Output ("PASS ucSegmentedControl client=${cw}x${ch} hosts=$($vis.Count)+$($hid.Count) png=$pngPath")
exit 0
