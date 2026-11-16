param(
    [Parameter(Mandatory=$true)][string]$Exe,
    [string]$OutDir = "",
    [int]$Sec = 4
)
# 内置 TChart 的透视判据（tests\ctrlteechart\TcPerspApp.vbp）。
#
# 判据的物理含义：3D 柱的侧面色带宽 = 该柱的深度。平行投影下三根柱子一样宽；
# 透视打开后，越靠绘图区左边（远端）越厚、右边（近端）越薄。
# 所以量三根柱子各自的侧面色带宽，断言“左 > 右”且有梯度 —— 只看一根柱子
# 是分不出“透视没做”和“深度常量改了”的。
#
# 几何（客户区 401x257、Panel 边距 12、图例在右 ⇒ plot=(28,27)-(302,220)）：
#   3 根柱，bw=45、gap=45 ⇒ 左边 x=50/140/230，右缘 95/185/275
# 颜色同 Tc3dApp：正面 (0,128,255)、侧面暗 45% → (0,70,140)

Add-Type -AssemblyName System.Drawing
Add-Type @"
using System; using System.Text; using System.Runtime.InteropServices;
public class PsPx {
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
if (-not $OutDir) { $OutDir = Join-Path $wd "tcpersp" }
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$pngPath = Join-Path $OutDir "tcpersp_chart.png"

$p = Start-Process -FilePath $exePath -WorkingDirectory $wd -PassThru `
     -RedirectStandardOutput (Join-Path $OutDir "tcpersp.out") `
     -RedirectStandardError (Join-Path $OutDir "tcpersp.err")
try { $p.WaitForInputIdle(15000) | Out-Null } catch {}
Start-Sleep -Seconds $Sec

$main = [PsPx]::FindMain([uint32]$p.Id)
if ($main -eq [IntPtr]::Zero) {
    Write-Output ("FAIL 主窗未起 (pid=" + $p.Id + ")")
    try { $p.Kill() } catch {}
    exit 1
}
[PsPx]::ShowWindow($main, 9) | Out-Null
[PsPx]::SetForegroundWindow($main) | Out-Null
Start-Sleep -Milliseconds 600

$chart = [PsPx]::FindChart($main)
if ($chart -eq [IntPtr]::Zero) {
    Write-Output ("FAIL 主窗可见但无 VB6_TCHART 子窗 (pid=" + $p.Id + ")")
    try { $p.Kill() } catch {}
    exit 1
}
[PsPx]::RedrawWindow($chart, [IntPtr]::Zero, [IntPtr]::Zero, 0x0185) | Out-Null
Start-Sleep -Milliseconds 400

$chr = New-Object PsPx+RECT; [PsPx]::GetClientRect($chart, [ref]$chr) | Out-Null
$tw = $chr.R; $th = $chr.B
$dpi = [int][PsPx]::GetDpiForWindow($main)

$bmp = New-Object System.Drawing.Bitmap($tw, $th)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$hdc = $g.GetHdc()
$srcDc = [PsPx]::GetDC($chart)
[PsPx]::BitBlt($hdc, 0, 0, $tw, $th, $srcDc, 0, 0, 0x00CC0020) | Out-Null
[PsPx]::ReleaseDC($chart, $srcDc) | Out-Null
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

$issues = New-Object System.Collections.Generic.List[string]
$read = New-Object System.Collections.Generic.List[string]
if ($tw -ne 401 -or $th -ne 257) { $issues.Add("图表 ${tw}x${th} != 401x257") }
if ($dpi -ne 96) { $issues.Add("DPI=$dpi != 96") }

$sideC = @(0, 70, 140)
$bars = @(50, 140, 230)      # 三根柱子的左缘
$bw = 45

$widths = @()
foreach ($bx in $bars) {
    $right = $bx + $bw
    $maxX = -1
    for ($x = $right - 2; $x -le [Math]::Min($right + 24, $tw - 1); $x++) {
        for ($y = 20; $y -le 224; $y++) {
            $o = $y * $stride + $x * 4
            $b = [int]$buf[$o]; $g2 = [int]$buf[$o + 1]; $r2 = [int]$buf[$o + 2]
            if ([Math]::Abs($r2 - $sideC[0]) -le 6 -and [Math]::Abs($g2 - $sideC[1]) -le 6 -and [Math]::Abs($b - $sideC[2]) -le 6) {
                if ($x -gt $maxX) { $maxX = $x }
            }
        }
    }
    if ($maxX -lt 0) { $widths += -1 } else { $widths += ($maxX - $right) }
}
$read.Add("sideWidths=" + ($widths -join ","))

foreach ($w in $widths) {
    if ($w -lt 4) { $issues.Add("有一根柱子测不到侧面带宽 (w=$w) —— 3D 没画或颜色变了"); break }
}
if ($widths[0] -ge 0 -and $widths[2] -ge 0) {
    if ($widths[0] -le $widths[2]) {
        $issues.Add("左边（远端）侧面色带宽 $($widths[0]) 不大于右边（近端）$($widths[2]) —— 透视没生效")
    }
    if (($widths[0] - $widths[2]) -lt 3) {
        $issues.Add("左右带宽只差 $($widths[0] - $widths[2])px（期望 ≥3，这是渐变而不是常数）")
    }
    foreach ($w in $widths) {
        if ($w -lt 8 -or $w -gt 20) { $issues.Add("带宽 w=$w 超出 8..20（深度常量被改坏）"); break }
    }
}

if (-not $p.HasExited) { try { $p.Kill() } catch {} }

if ($issues.Count -gt 0) {
    Write-Output ("FAIL tc_persp_px chart=${tw}x${th} dpi=$dpi " + ($read -join " "))
    foreach ($it in $issues) { Write-Output ("  " + $it) }
    Write-Output ("  png=" + $pngPath)
    exit 1
}
Write-Output ("PASS tc_persp_px chart=${tw}x${th} dpi=$dpi " + ($read -join " ") + " png=" + $pngPath)
exit 0
