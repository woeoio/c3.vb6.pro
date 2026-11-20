param(
    [Parameter(Mandatory=$true)][string]$Exe,
    [string]$OutDir = "",
    [int]$Sec = 4
)
# 内置 TChart 的结构像素判据（tests\ctrlteechart\TcPxApp.vbp）。
#
# 判据与 tests\guisample\teechart_display_px.ps1 同一条口径：真 VB6 那份录的
# 是 401x257 图表上的结构化读数，这里拿**同几何、同标题文字**的夹具比同一组数：
#   title bbox centerX=155.5（= 图例左边/2，标题居中于“客户区 - 图例”）
#   legend 横线 y=26 / x=310..371
#   topaxis y=53..54 / botaxis y=220..221
# 差异只在字体：真控件用 TeeChart 自带字体设置，内置版用系统 GUI 字体，
# 所以 title bbox 只判中心与纵向带（文字宽度随字体变，不钉左右端点）。
#
# 抓像素走 BitBlt 子窗自己的 DC（不 CopyFromScreen）：窗口被别的窗盖住也能量，
# 且与真 VB6 参照同处系统虚拟化后的 96dpi 空间（本脚本刻意不声明 DPI aware）。

Add-Type -AssemblyName System.Drawing
Add-Type @"
using System; using System.Text; using System.Runtime.InteropServices;
public class TcPx {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder sb, int n);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool GetClientRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool RedrawWindow(IntPtr h, IntPtr r, IntPtr rgn, int flags);
  [DllImport("user32.dll")] public static extern uint GetDpiForWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern IntPtr GetDC(IntPtr h);
  [DllImport("user32.dll")] public static extern int ReleaseDC(IntPtr h, IntPtr dc);
  [DllImport("gdi32.dll")] public static extern bool BitBlt(IntPtr d, int x, int y, int w, int h, IntPtr s, int sx, int sy, int rop);
  public struct RECT { public int L, T, R, B; }
  public static string Cls(IntPtr h) { var sb = new StringBuilder(256); GetClassName(h, sb, 256); return sb.ToString(); }
  public static IntPtr FindMain(uint pid) {
    IntPtr best = IntPtr.Zero; long bestA = 0;
    EnumWindows((h, l) => {
      uint p; GetWindowThreadProcessId(h, out p);
      if (p == pid && IsWindowVisible(h) && Cls(h).StartsWith("VB6_Form_")) {
        RECT r; GetWindowRect(h, out r);
        long a = (long)(r.R - r.L) * (r.B - r.T);
        if (a > bestA) { bestA = a; best = h; }
      }
      return true;
    }, IntPtr.Zero);
    return best;
  }
  public static IntPtr FindChart(IntPtr main) {
    IntPtr best = IntPtr.Zero; long bestA = 0;
    EnumChildWindows(main, (h, l) => {
      if (IsWindowVisible(h) && Cls(h) == "VB6_TCHART") {
        RECT r; GetWindowRect(h, out r);
        long a = (long)(r.R - r.L) * (r.B - r.T);
        if (a > bestA) { bestA = a; best = h; }
      }
      return true;
    }, IntPtr.Zero);
    return best;
  }
}
"@

$exePath = (Resolve-Path $Exe -ErrorAction SilentlyContinue)
if (-not $exePath) { Write-Output "FAIL 缺产物 ($Exe 不存在)"; exit 1 }
$exePath = $exePath.Path
$wd = Split-Path $exePath
if (-not $OutDir) { $OutDir = Join-Path $wd "tcpx" }
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$pngPath = Join-Path $OutDir "tcpx_chart.png"
$logOut = Join-Path $OutDir "tcpx.out"
$logErr = Join-Path $OutDir "tcpx.err"

$p = Start-Process -FilePath $exePath -WorkingDirectory $wd -PassThru `
     -RedirectStandardOutput $logOut -RedirectStandardError $logErr
try { $p.WaitForInputIdle(15000) | Out-Null } catch {}
Start-Sleep -Seconds $Sec

$main = [TcPx]::FindMain([uint32]$p.Id)
if ($main -eq [IntPtr]::Zero) {
    Write-Output ("FAIL 主窗未起 (pid=" + $p.Id + ")")
    try { $p.Kill() } catch {}
    exit 1
}
[TcPx]::ShowWindow($main, 9) | Out-Null
[TcPx]::SetForegroundWindow($main) | Out-Null
Start-Sleep -Milliseconds 600

$chart = [TcPx]::FindChart($main)
if ($chart -eq [IntPtr]::Zero) {
    Write-Output ("FAIL 主窗可见但无 VB6_TCHART 子窗 (pid=" + $p.Id + ")")
    try { $p.Kill() } catch {}
    exit 1
}
[TcPx]::RedrawWindow($chart, [IntPtr]::Zero, [IntPtr]::Zero, 0x0185) | Out-Null
Start-Sleep -Milliseconds 400

$chr = New-Object TcPx+RECT; [TcPx]::GetClientRect($chart, [ref]$chr) | Out-Null
$tw = $chr.R; $th = $chr.B
$dpi = [int][TcPx]::GetDpiForWindow($main)
if ($tw -le 0 -or $th -le 0) {
    Write-Output ("FAIL 图表客户区为空 ${tw}x${th}")
    try { $p.Kill() } catch {}
    exit 1
}

$bmp = New-Object System.Drawing.Bitmap($tw, $th)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$hdc = $g.GetHdc()
$srcDc = [TcPx]::GetDC($chart)
[TcPx]::BitBlt($hdc, 0, 0, $tw, $th, $srcDc, 0, 0, 0x00CC0020) | Out-Null
[TcPx]::ReleaseDC($chart, $srcDc) | Out-Null
$g.ReleaseHdc($hdc)
$g.Dispose()
$bmp.Save($pngPath, [System.Drawing.Imaging.ImageFormat]::Png)

$rect = New-Object System.Drawing.Rectangle(0, 0, $tw, $th)
$bd = $bmp.LockBits($rect, [System.Drawing.Imaging.ImageLockMode]::ReadOnly, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
$stride = $bd.Stride
$buf = New-Object byte[] ($stride * $th)
[System.Runtime.InteropServices.Marshal]::Copy($bd.Scan0, $buf, 0, $buf.Length)
$bmp.UnlockBits($bd)
$bmp.Dispose()

function Get-Px([int]$x, [int]$y) {
    $o = $y * $stride + $x * 4
    return @([int]$buf[$o + 2], [int]$buf[$o + 1], [int]$buf[$o])
}
function Test-Dark([int]$x, [int]$y, [int]$lim) {
    $c = Get-Px $x $y
    return ($c[0] -lt $lim -and $c[1] -lt $lim -and $c[2] -lt $lim)
}

$first = Get-Px 0 0
$blank = $true
for ($y = 0; $y -lt $th -and $blank; $y += 4) {
    for ($x = 0; $x -lt $tw -and $blank; $x += 4) {
        $c = Get-Px $x $y
        if ($c[0] -ne $first[0] -or $c[1] -ne $first[1] -or $c[2] -ne $first[2]) { $blank = $false }
    }
}
if ($blank) {
    Write-Output ("FAIL 捕获为空图 (单色 " + $first[0] + "," + $first[1] + "," + $first[2] + " 于 ${tw}x${th})")
    try { $p.Kill() } catch {}
    exit 1
}

$issues = New-Object System.Collections.Generic.List[string]
$read = New-Object System.Collections.Generic.List[string]

if ($tw -ne 401 -or $th -ne 257) { $issues.Add("图表 ${tw}x${th} != 401x257（夹具几何变了，判据的期望值要重录）") }
if ($dpi -ne 96) { $issues.Add("DPI=$dpi != 96 (判据按 96dpi 录制)") }
$read.Add("chart=${tw}x${th} dpi=$dpi")

# ① 标题：纵向必须落在 header 带 12..26 内；水平中心期望 155.5（图例左边/2）
# 只扫 y=6..24：y=26 是图例顶边，再往下是图例文字，混进来会把 bbox 拉歪（踩过）。
$minx = 9999; $maxx = -1; $miny = 9999; $maxy = -1
for ($y = 6; $y -le 24; $y++) {
    for ($x = 20; $x -le 380; $x++) {
        if (Test-Dark $x $y 110) {
            if ($x -lt $minx) { $minx = $x }
            if ($x -gt $maxx) { $maxx = $x }
            if ($y -lt $miny) { $miny = $y }
            if ($y -gt $maxy) { $maxy = $y }
        }
    }
}
if ($maxx -lt 0) {
    $issues.Add("标题区没扫到任何深色像素 (y=4..30)")
    $read.Add("title=none")
} else {
    $cx = [double]($minx + $maxx) / 2
    $read.Add("title=($minx,$miny)-($maxx,$maxy) centerX=$cx")
    if ($miny -lt 11 -or $maxy -gt 27) { $issues.Add("title y=$miny..$maxy 不在 header 带 12..26 内") }
    if ([Math]::Abs($cx - 155.5) -gt 2.0) { $issues.Add("title centerX=$cx (期望 155.5±2)") }
}

# ② 图例顶边：自上而下第一条 n>=40 的长横线（期望 y=26, x=310..371）
$topY = -1; $topN = 0; $topX0 = -1; $topX1 = -1
for ($y = 20; $y -le 70; $y++) {
    $n = 0; $x0 = -1; $x1 = -1
    for ($x = 280; $x -lt $tw; $x++) {
        if (Test-Dark $x $y 70) {
            $n++
            if ($x0 -lt 0) { $x0 = $x }
            $x1 = $x
        }
    }
    if ($n -ge 40) { $topY = $y; $topN = $n; $topX0 = $x0; $topX1 = $x1; break }
}
$read.Add("legend top y=$topY x=$topX0..$topX1 n=$topN")
if ($topN -lt 40) { $issues.Add("图例横线没找到 n=$topN (<40)") }
else {
    if ([Math]::Abs($topY - 26) -gt 1) { $issues.Add("图例横线 y=$topY (期望 26±1)") }
    if ([Math]::Abs($topX0 - 310) -gt 3) { $issues.Add("图例左边 x=$topX0 (期望 310±3)") }
    if ([Math]::Abs($topX1 - 371) -gt 3) { $issues.Add("图例右边 x=$topX1 (期望 371±3)") }
}

# ③ 上/下轴线所在行
$taHit = -1; $taN = 0
for ($y = 46; $y -le 60; $y++) {
    $n = 0
    for ($x = 40; $x -lt 380; $x++) { if (Test-Dark $x $y 90) { $n++ } }
    if ($n -gt $taN) { $taN = $n; $taHit = $y }
}
$read.Add("topaxis y=$taHit n=$taN")
if ($taN -lt 100) { $issues.Add("顶部轴线 n=$taN (<100) 于 y=46..60") }
elseif ([Math]::Abs($taHit - 53) -gt 2) { $issues.Add("顶部轴线 y=$taHit (期望 53±2)") }

$baHit = -1; $baN = 0
for ($y = 210; $y -le 232; $y++) {
    $n = 0
    for ($x = 40; $x -lt 380; $x++) { if (Test-Dark $x $y 90) { $n++ } }
    if ($n -gt $baN) { $baN = $n; $baHit = $y }
}
$read.Add("botaxis y=$baHit n=$baN")
if ($baN -lt 100) { $issues.Add("底部轴线 n=$baN (<100) 于 y=210..232") }
elseif ([Math]::Abs($baHit - 220) -gt 2) { $issues.Add("底部轴线 y=$baHit (期望 220±2)") }

if (-not $p.HasExited) { try { $p.Kill() } catch {} }

if ($issues.Count -gt 0) {
    Write-Output ("FAIL tc_px " + ($read -join " "))
    foreach ($it in $issues) { Write-Output ("  " + $it) }
    Write-Output ("  png=" + $pngPath)
    exit 1
}
Write-Output ("PASS tc_px " + ($read -join " ") + " png=" + $pngPath)
exit 0
