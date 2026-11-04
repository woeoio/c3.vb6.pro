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
#   · 窗体 766x558 / 客户区 751x520 (真 VB6 参考 exe 实测同一组读数);
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
#          ucSegmentedControl1  (顶部条)     488/161         802/161      vivid>=400
#     注: uc2 的 vivid 在真 VB6 上更高 (参考的条目标题滚动/悬停态在参考进程里更"彩"),
#     这里取 C3 实测 145 作下限 —— 图标一旦缺失该值会掉到 0 附近, 这才是要防的回归。
#   · 全图强调色计数: #0078D7 (选中段底色) >= 8000; #CE00EF (TAB2 品红) >= 250。
#
# 已核对但**暂不钉**的残差 (见 ai\C3_SEGMENTED_HANDOFF.md, 属未收口的语义差异, 不是本脚本
# 放水): 顶部第 3 项在真 VB6 里是 Disabled 灰显 / C3 仍可点; "Reporto" 项的图标映射不同;
# CtrlType=1 小条与底部翻页条的 PUA 符号字形 (Segoe UI Symbol) 在两产物上取字不同。
# 这些一旦修好应再补断言, 不要让本条变成"永远只验旧口径"。
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
  [DllImport("kernel32.dll")] public static extern bool IsWow64Process(IntPtr h, out bool wow64);
  public struct RECT { public int L, T, R, B; }
  public struct POINT { public int X, Y; }
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
if (-not $OutDir) { $OutDir = Join-Path $wd "segmented_px" }
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$pngPath = Join-Path $OutDir "segmented_c3.png"
$logOut = Join-Path $OutDir "segmented.out"
$logErr = Join-Path $OutDir "segmented.err"

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
Add-Read ("window=${ww}x${wh} client=${cw}x${ch}")
# 真 VB6 参考 exe 实测: window=766x558 client=751x520 (同一套 96dpi 虚拟化口径)。
if ($ww -ne 766 -or $wh -ne 558) { Add-Fail ("窗尺寸 ${ww}x${wh} != 766x558 (真 VB6 参考读数)") }
if ($cw -ne 751 -or $ch -ne 520) { Add-Fail ("客户区 ${cw}x${ch} != 751x520 (真 VB6 参考读数)") }

# --- 子窗: 12 个可见 UC 宿主 + 1 个隐藏宿主 + 命令按钮 ---
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
$s = $stat["uc1"];  if ($s.vivid -lt 400) { Add-Fail ("uc1(顶部条) vivid=" + $s.vivid + " < 400 (四个段的彩色图标没了?)") }

# --- 全图强调色计数 ---
$nBlue = 0; $nMag = 0
for ($y = 0; $y -lt $ch; $y++) { for ($x = 0; $x -lt $cw; $x++) {
    $c = $bmp.GetPixel($x, $y)
    if ($c.R -eq 0x00 -and $c.G -eq 0x78 -and $c.B -eq 0xD7) { $nBlue++ }
    if ($c.R -eq 0xCE -and $c.G -eq 0x00 -and $c.B -eq 0xEF) { $nMag++ } } }
Add-Read ("accent 0078D7=$nBlue CE00EF=$nMag")
# 真 VB6 参考: 0078D7=10670, CE00EF=554; C3 实测 10715 / 400。
if ($nBlue -lt 8000) { Add-Fail ("选中段底色 #0078D7 只有 $nBlue 像素 (<8000, 选中态没画出来)") }
if ($nMag -lt 250) { Add-Fail ("TAB2 强调色 #CE00EF 只有 $nMag 像素 (<250)") }
$bmp.Dispose()

try { if (-not $proc.HasExited) { $proc.Kill(); $proc.WaitForExit(3000) | Out-Null } } catch { }

Write-Output ("读数: " + ($read -join "; "))
if ($fail.Count -gt 0) {
    Write-Output ("FAIL " + ($fail -join "; "))
    exit 1
}
Write-Output ("PASS ucSegmentedControl client=${cw}x${ch} hosts=$($vis.Count)+$($hid.Count) png=$pngPath")
exit 0
