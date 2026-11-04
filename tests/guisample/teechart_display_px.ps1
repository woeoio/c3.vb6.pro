param(
    [Parameter(Mandatory=$true)][string]$Exe,
    [string]$OutDir = "",
    [int]$Sec = 10
)
# TeeChart Chart display 夹具的结构判据 (tests\TeeChartChartDisplay)。
#
# 为什么不能拿数据/整图对等: 示例 Form_Load 里 `Series(0).FillSampleValues 10` 用的是
# Delphi 的 Random ⇒ 每次运行条形与 Y 轴刻度都不同（真 VB6 两次运行同样不同）。所以只钉
# **结构性元素**，它们与随机数据无关，且正是本轮修出来的差异所在：
#   · 标题文字 bbox 的水平中心 —— 少了 Panel.MarginRight 的 Change 时整图右移 6px；
#   · 图例框的左右竖边 —— 同一原因右移 12px (322..383 → 310..371)；
#   · 顶部/底部轴线所在行 —— 少了 Panel.MarginBottom 时下移 5px (225 → 220)。
#
# 读数来源 (2026-11-04，**DPI 虚拟化 96dpi 空间**，x86 产物，TChartChart 子窗客户区 401x257):
#   本机物理屏 1920x1200 / 120 DPI (125%)；VB6 与 C3 都没有 dpiAware 清单 ⇒ 系统把它们
#   虚拟化到 96dpi（unaware 进程读到的窗口/客户区/子窗尺寸 = 514x539 / 498x500 / 401x257）。
#   本脚本刻意**不调用 SetProcessDPIAware**，并用 BitBlt 抓子窗自己的 DC —— 与
#   tests\guisample\vb6ref\capture.ps1 和 harness 的采样口径同一空间，读数才可对等。
#   真 VB6 (C:\...\VB6Mini /make 自同一份厂商源码)  C3 (修好后)
#   title bbox=(91,14)-(220,24) centerX=155.5      title bbox=(91,14)-(220,24) centerX=155.5
#   legend hline y=26 x=310..371 n=62              legend hline y=26 x=310..371 n=62
#   legend vline x=310 y=26..55 n=30               legend vline x=310 y=26..55 n=30
#   legend vline x=371 y=26..58 n=33               legend vline x=371 y=26..58 n=33
#   topaxis y=53..54 / botaxis y=220..221          topaxis y=53..54 / botaxis y=220..221
#   (轴线左端 x 随随机 Y 轴刻度宽度浮动: VB6 58/48, C3 49/40 ⇒ 只判"该行有长横线", 不判左端。)
#   C3 修前读数 (事件语义/设计期 Min-Max 都缺): title centerX=161.5, legend x=322..383,
#   botaxis y=225 —— 本判据在修前会红。
#
# SKIP 口径 (与 harness 的 RequiresFiles 同一条): 只在**产物目录缺 TeeChart5.ocx** 时 SKIP
# (rc=2)。资产在、脚本却 SKIP 就是"没验冒充通过"，故本脚本其余失败一律 FAIL(rc=1)。
#
# 厂商目录的 DisplayProject.png 是 633x664 的展示图 (与 514x539 窗口的客户区比例不一致,
# 非 1:1 基准), 因此不以整图像素做判据; 官方对齐以本脚本这些结构读数为准。

Add-Type -AssemblyName System.Drawing
Add-Type @"
using System; using System.Text; using System.Runtime.InteropServices;
public class TeePx {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder sb, int n);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool GetClientRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool ClientToScreen(IntPtr h, ref POINT p);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, int flags);
  [DllImport("user32.dll")] public static extern bool RedrawWindow(IntPtr h, IntPtr r, IntPtr rgn, int flags);
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  [DllImport("user32.dll")] public static extern uint GetDpiForWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern IntPtr GetDC(IntPtr h);
  [DllImport("user32.dll")] public static extern int ReleaseDC(IntPtr h, IntPtr dc);
  [DllImport("gdi32.dll")] public static extern bool BitBlt(IntPtr d, int x, int y, int w, int h, IntPtr s, int sx, int sy, int rop);
  public struct POINT { public int X, Y; }
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
      if (IsWindowVisible(h) && Cls(h) == "TChartChart") {
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
# 刻意不声明 DPI aware: 与真 VB6 参照 / cap.ps1 / harness 同处系统虚拟化后的 96dpi 空间。
# (在 125% 缩放的机器上, aware 进程会把 401x257 逻辑布局看成 501x321 物理图, 结构判据全错位。)

$exePath = (Resolve-Path $Exe -ErrorAction SilentlyContinue)
if (-not $exePath) { Write-Output "FAIL 缺产物 ($Exe 不存在)"; exit 1 }
$exePath = $exePath.Path
# 位宽硬约束: TeeChart5.ocx 是 32 位, 本工程只能 --arch x86。拿 x64 产物来跑
# 既加载不了 OCX 也量不出 96dpi 下的 514x539 几何, 那**不是跳过而是明确失败**
# (位宽不符 = 用例没按夹具约定编译), 不许静默放过。
$peBytes = [IO.File]::ReadAllBytes($exePath)
$peMachine = [BitConverter]::ToUInt16($peBytes, [BitConverter]::ToInt32($peBytes, 0x3C) + 4)
if ($peMachine -ne 0x014C) {
    Write-Output ("FAIL 位宽不符 (machine=0x{0:X4}, 期望 0x014C x86; 本夹具只能 --arch x86)" -f $peMachine)
    exit 1
}
$wd = Split-Path $exePath
if (-not $OutDir) { $OutDir = Join-Path $wd "teechart_display_px" }
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$pngPath = Join-Path $OutDir "teechart_chart_c3.png"
$logOut = Join-Path $OutDir "teechart.out"
$logErr = Join-Path $OutDir "teechart.err"

# 缺资产才允许 SKIP —— 与 harness 的 RequiresFiles 同一条口径。
if (-not (Test-Path -LiteralPath (Join-Path $wd 'TeeChart5.ocx'))) {
    Write-Output ("SKIP 缺资产 (" + $wd + " 下没有 TeeChart5.ocx)"); exit 2
}

function Get-LaunchDiag {
    param($Proc, [string]$OutLog, [string]$ErrLog)
    $state = if ($Proc -and -not $Proc.HasExited) { 'alive' } else { 'exited' }
    $code = '?'
    if ($state -eq 'exited') {
        try { $Proc.WaitForExit(1000) | Out-Null } catch {}
        try { $code = ('0x{0:X8}' -f [uint32]$Proc.ExitCode) } catch { try { $code = [string]$Proc.ExitCode } catch {} }
    }
    $e = @()
    try { $e = @(Get-Content -LiteralPath $ErrLog -ErrorAction SilentlyContinue | Where-Object { $_ }) } catch {}
    $crash = @($e | Where-Object { $_ -match '\[C3_CRASH\]' } | Select-Object -First 3) -join ' | '
    $el = if ($e.Count -gt 0) { @($e | Select-Object -Last 3) -join ' | ' } else { '(empty)' }
    $s = "state=$state exitCode=$code errTail='$el'"
    if ($crash) { $s += " crash='$crash'" }
    return $s
}

# 必须走 CreateProcess: 带 -RedirectStandard* 时 5.1 的 Start-Process 自动 UseShellExecute=$false,
# 不经 ShellExecute/附件管理器 (仓库根曾带 Low 完整性标签, ShellExecute 会弹安全警告并卡住)。
# 注意 5.1 下 -NoNewWindow 与 -RedirectStandard* 互斥, 不能同时给。
$env:C3_CRASH_TRACE = '1'
$p = Start-Process -FilePath $exePath -WorkingDirectory $wd -PassThru `
     -RedirectStandardOutput $logOut -RedirectStandardError $logErr
try { $p.WaitForInputIdle(15000) | Out-Null } catch {}
Start-Sleep -Seconds $Sec

$main = [TeePx]::FindMain([uint32]$p.Id)
if ($main -eq [IntPtr]::Zero) {
    Write-Output ("FAIL 主窗未起 (pid=" + $p.Id + " " + (Get-LaunchDiag $p $logOut $logErr) + ")")
    try { $p.Kill() } catch {}
    exit 1
}
[TeePx]::ShowWindow($main, 9) | Out-Null
[TeePx]::SetForegroundWindow($main) | Out-Null
[TeePx]::SetWindowPos($main, [IntPtr](-1), 0, 0, 0, 0, 0x0003) | Out-Null
Start-Sleep -Milliseconds 800

$chart = [TeePx]::FindChart($main)
if ($chart -eq [IntPtr]::Zero) {
    Write-Output ("FAIL 主窗可见但无 TChartChart 子窗 (pid=" + $p.Id + " " + (Get-LaunchDiag $p $logOut $logErr) + ")")
    try { $p.Kill() } catch {}
    exit 1
}
[TeePx]::RedrawWindow($chart, [IntPtr]::Zero, [IntPtr]::Zero, 0x0185) | Out-Null
Start-Sleep -Milliseconds 500

# --- 几何读数 (96dpi 物理像素; 与真 VB6 同: 窗口 514x539 / 客户区 498x500 / 图表 401x257) ---
$wr = New-Object TeePx+RECT; [TeePx]::GetWindowRect($main, [ref]$wr) | Out-Null
$cr = New-Object TeePx+RECT; [TeePx]::GetClientRect($main, [ref]$cr) | Out-Null
$chr = New-Object TeePx+RECT; [TeePx]::GetClientRect($chart, [ref]$chr) | Out-Null
$mw = $wr.R - $wr.L; $mh = $wr.B - $wr.T
$cw = $cr.R; $chh = $cr.B
$tw = $chr.R; $th = $chr.B
$dpi = [int][TeePx]::GetDpiForWindow($main)

# --- 图表客户区捕获: 直接 BitBlt 子窗自己的 DC ---
# 不用 CopyFromScreen: 顶层窗可能被别的窗盖住 (cap.ps1 当初就是被 IDE 窗遮住才换的),
# 而 TChartChart 自己有窗口/DC, 抓到的是它的真实像素 (PrintWindow 抓不到 Codejock 一类,
# 这里也不需要)。
if ($tw -le 0 -or $th -le 0) {
    Write-Output ("FAIL 图表客户区为空 ${tw}x${th}")
    try { $p.Kill() } catch {}
    exit 1
}
$bmp = New-Object System.Drawing.Bitmap($tw, $th)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$hdc = $g.GetHdc()
$srcDc = [TeePx]::GetDC($chart)
[TeePx]::BitBlt($hdc, 0, 0, $tw, $th, $srcDc, 0, 0, 0x00CC0020) | Out-Null
[TeePx]::ReleaseDC($chart, $srcDc) | Out-Null
$g.ReleaseHdc($hdc)
$g.Dispose()
$bmp.Save($pngPath, [System.Drawing.Imaging.ImageFormat]::Png)

# LockBits 读整幅 (GetPixel 在 400x260 上太慢且容易把"没抓到"变成超时)
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

# 空图/单色图 = 桌面没抓到, 这是失败不是跳过 (资产已确认在)
$first = Get-Px 0 0
$blank = $true
for ($y = 0; $y -lt $th -and $blank; $y += 4) {
    for ($x = 0; $x -lt $tw -and $blank; $x += 4) {
        $c = Get-Px $x $y
        if ($c[0] -ne $first[0] -or $c[1] -ne $first[1] -or $c[2] -ne $first[2]) { $blank = $false }
    }
}
if ($blank) {
    Write-Output ("FAIL 桌面捕获为空图 (单色 " + $first[0] + "," + $first[1] + "," + $first[2] + " 于 ${tw}x${th})")
    try { $p.Kill() } catch {}
    exit 1
}

$issues = New-Object System.Collections.Generic.List[string]
$read = New-Object System.Collections.Generic.List[string]

# ① 几何: 与真 VB6 读数逐值对等 (96dpi)
if ($mw -ne 514 -or $mh -ne 539) { $issues.Add("主窗 $($mw)x$($mh) != 514x539") }
if ($cw -ne 498 -or $chh -ne 500) { $issues.Add("客户区 $($cw)x$($chh) != 498x500") }
if ($tw -ne 401 -or $th -ne 257) { $issues.Add("图表 $($tw)x$($th) != 401x257") }
if ($dpi -ne 96) { $issues.Add("DPI=$dpi != 96 (像素判据按 96dpi 录制)") }
$read.Add("main=${mw}x${mh} client=${cw}x${chh} chart=${tw}x${th} dpi=$dpi")

# ② 标题蓝字 bbox: 期望 (91,14)-(220,24), centerX=155.5
$minx = 9999; $maxx = -1; $miny = 9999; $maxy = -1
for ($y = 0; $y -le 32; $y++) {
    for ($x = 80; $x -le 330; $x++) {
        $c = Get-Px $x $y
        if ($c[2] -gt 120 -and $c[0] -lt 110 -and $c[1] -lt 130) {
            if ($x -lt $minx) { $minx = $x }
            if ($x -gt $maxx) { $maxx = $x }
            if ($y -lt $miny) { $miny = $y }
            if ($y -gt $maxy) { $maxy = $y }
        }
    }
}
$cx = [double]($minx + $maxx) / 2
$read.Add("title=($minx,$miny)-($maxx,$maxy) centerX=$cx")
if ($minx -lt 90 -or $minx -gt 92) { $issues.Add("title minx=$minx (期望 91±1)") }
if ($maxx -lt 219 -or $maxx -gt 221) { $issues.Add("title maxx=$maxx (期望 220±1)") }
if ($miny -lt 13 -or $miny -gt 15) { $issues.Add("title miny=$miny (期望 14±1)") }
if ($maxy -lt 23 -or $maxy -gt 25) { $issues.Add("title maxy=$maxy (期望 24±1)") }
if ([Math]::Abs($cx - 155.5) -gt 1.0) { $issues.Add("title centerX=$cx (期望 155.5±1; 修前 161.5)") }

# ③ 图例: 顶边横线 y=26 (x=310..371, n=62) + 左右竖边 x=310 / x=371
# 自上而下取**第一条**满足 n>=55 的长横线 —— 底边框那条更长 (x=310..374, n=65),
# 用"最大 n"会挑到 y=55 (实测踩过)。
$topY = -1; $topN = 0; $topX0 = -1; $topX1 = -1
for ($y = 0; $y -le 70; $y++) {
    $n = 0; $x0 = -1; $x1 = -1
    for ($x = 300; $x -lt $tw; $x++) {
        $c = Get-Px $x $y
        if ($c[0] -lt 70 -and $c[1] -lt 70 -and $c[2] -lt 70) {
            $n++
            if ($x0 -lt 0) { $x0 = $x }
            $x1 = $x
        }
    }
    if ($n -ge 55) { $topY = $y; $topN = $n; $topX0 = $x0; $topX1 = $x1; break }
}
$read.Add("legend top y=$topY x=$topX0..$topX1 n=$topN")
if ($topN -lt 55) { $issues.Add("legend 横线 n=$topN (<55)") }
if ($topY -lt 25 -or $topY -gt 27) { $issues.Add("legend 横线 y=$topY (期望 26±1)") }
if ($topX0 -lt 309 -or $topX0 -gt 311) { $issues.Add("legend 左边 x=$topX0 (期望 310±1; 修前 322)") }
if ($topX1 -lt 370 -or $topX1 -gt 372) { $issues.Add("legend 右边 x=$topX1 (期望 371±1)") }

# 竖边是 3~4px 的粗带: 判"带的最左列" (310 / 371), 不判最右列 (374 也满足长游程)。
$vx1 = -1; $vn1 = 0; $vy0 = -1; $vy1 = -1
for ($x = 305; $x -le 315; $x++) {
    $n = 0; $y0 = -1; $y1 = -1
    for ($y = 0; $y -le 70; $y++) {
        $c = Get-Px $x $y
        if ($c[0] -lt 70 -and $c[1] -lt 70 -and $c[2] -lt 70) {
            $n++
            if ($y0 -lt 0) { $y0 = $y }
            $y1 = $y
        }
    }
    if ($n -ge 28 -and $vx1 -lt 0) { $vx1 = $x; $vn1 = $n; $vy0 = $y0; $vy1 = $y1 }
}
$read.Add("legend left x=$vx1 y=$vy0..$vy1 n=$vn1")
if ($vn1 -lt 28) { $issues.Add("legend 左竖边 n=$vn1 (<28)") }
if ($vx1 -lt 309 -or $vx1 -gt 311) { $issues.Add("legend 左竖边 x=$vx1 (期望 310±1; 修前 322)") }

$vx2 = -1; $vn2 = 0; $wy0 = -1; $wy1 = -1
for ($x = 366; $x -le 376; $x++) {
    $n = 0; $y0 = -1; $y1 = -1
    for ($y = 0; $y -le 70; $y++) {
        $c = Get-Px $x $y
        if ($c[0] -lt 70 -and $c[1] -lt 70 -and $c[2] -lt 70) {
            $n++
            if ($y0 -lt 0) { $y0 = $y }
            $y1 = $y
        }
    }
    if ($n -ge 25 -and $vx2 -lt 0) { $vx2 = $x; $vn2 = $n; $wy0 = $y0; $wy1 = $y1 }
}
$read.Add("legend right x=$vx2 y=$wy0..$wy1 n=$vn2")
if ($vn2 -lt 25) { $issues.Add("legend 右竖边 n=$vn2 (<25)") }
if ($vx2 -lt 370 -or $vx2 -gt 372) { $issues.Add("legend 右竖边 x=$vx2 (期望 371±1)") }

# ④ 顶部轴线 (y=53..54) / 底部轴线 (y=220..221): 只判"该行有长横线", 左端随随机刻度浮动
$topHit = -1; $topN = 0
for ($y = 52; $y -le 55; $y++) {
    $n = 0
    for ($x = 40; $x -lt 380; $x++) {
        $c = Get-Px $x $y
        if ($c[0] -lt 90 -and $c[1] -lt 90 -and $c[2] -lt 90) { $n++ }
    }
    if ($n -gt $topN) { $topN = $n; $topHit = $y }
}
$read.Add("topaxis y=$topHit n=$topN")
if ($topN -lt 150) { $issues.Add("顶部轴线 n=$topN (<150) 于 y=52..55") }

$botHit = -1; $botN = 0
for ($y = 219; $y -le 222; $y++) {
    $n = 0
    for ($x = 40; $x -lt 380; $x++) {
        $c = Get-Px $x $y
        if ($c[0] -lt 90 -and $c[1] -lt 90 -and $c[2] -lt 90) { $n++ }
    }
    if ($n -gt $botN) { $botN = $n; $botHit = $y }
}
$read.Add("botaxis y=$botHit n=$botN")
if ($botN -lt 150) { $issues.Add("底部轴线 n=$botN (<150) 于 y=219..222; 修前 y=225 (MarginBottom 丢失)") }

if (-not $p.HasExited) { try { $p.Kill() } catch {} }

if ($issues.Count -gt 0) {
    Write-Output ("FAIL teechart_display_px " + ($read -join " "))
    foreach ($it in $issues) { Write-Output ("  " + $it) }
    Write-Output ("  png=" + $pngPath)
    exit 1
}
Write-Output ("PASS teechart_display_px " + ($read -join " ") + " png=" + $pngPath)
exit 0
