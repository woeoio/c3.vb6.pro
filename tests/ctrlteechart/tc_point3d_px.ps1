param(
    [Parameter(Mandatory=$true)][string]$Exe,
    [string]$OutDir = "",
    [int]$Sec = 4
)
# 内置 TChart 的点系列 3D 判据（tests\ctrlteechart\TcPoint3dApp.vbp）。
#
# 3D 的点是每个点画成“扁墩”：先画暗色偏移圆（+4,-4），再压上正面圆。
# 判据量两枚圆的相对位置，不是面积：
#   正面圆上缘 - 后墩上缘 ≈ 4（厚度的竖直分量）
#   两者最高那一行里的起始 x：后墩 - 正面 ≈ 4（厚度的水平分量）
# 点的深度按点径定（r=3 → 4px），远小于柱形的 12px：柱子有高度，点只有大小。
# 扫描跳过 x>=300 的图例区（图例色块就是正面色）。
#
# 颜色来自算法：正面 RGB(0,128,255)，暗 45% → (0,70,140)。
# 扫描跳过 x>=300 的图例区（图例色块就是正面色）。

Add-Type -AssemblyName System.Drawing
Add-Type @"
using System; using System.Text; using System.Runtime.InteropServices;
public class ArPx {
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
if (-not $OutDir) { $OutDir = Join-Path $wd "tcpoint3d" }
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$pngPath = Join-Path $OutDir "tcpoint3d_chart.png"

$p = Start-Process -FilePath $exePath -WorkingDirectory $wd -PassThru `
     -RedirectStandardOutput (Join-Path $OutDir "tcpoint3d.out") `
     -RedirectStandardError (Join-Path $OutDir "tcpoint3d.err")
try { $p.WaitForInputIdle(15000) | Out-Null } catch {}
Start-Sleep -Seconds $Sec

$main = [ArPx]::FindMain([uint32]$p.Id)
if ($main -eq [IntPtr]::Zero) {
    Write-Output ("FAIL 主窗未起 (pid=" + $p.Id + ")")
    try { $p.Kill() } catch {}
    exit 1
}
[ArPx]::ShowWindow($main, 9) | Out-Null
[ArPx]::SetForegroundWindow($main) | Out-Null
Start-Sleep -Milliseconds 600

$chart = [ArPx]::FindChart($main)
if ($chart -eq [IntPtr]::Zero) {
    Write-Output ("FAIL 主窗可见但无 VB6_TCHART 子窗 (pid=" + $p.Id + ")")
    try { $p.Kill() } catch {}
    exit 1
}
[ArPx]::RedrawWindow($chart, [IntPtr]::Zero, [IntPtr]::Zero, 0x0185) | Out-Null
Start-Sleep -Milliseconds 400

$chr = New-Object ArPx+RECT; [ArPx]::GetClientRect($chart, [ref]$chr) | Out-Null
$tw = $chr.R; $th = $chr.B
$dpi = [int][ArPx]::GetDpiForWindow($main)

$bmp = New-Object System.Drawing.Bitmap($tw, $th)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$hdc = $g.GetHdc()
$srcDc = [ArPx]::GetDC($chart)
[ArPx]::BitBlt($hdc, 0, 0, $tw, $th, $srcDc, 0, 0, 0x00CC0020) | Out-Null
[ArPx]::ReleaseDC($chart, $srcDc) | Out-Null
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

$frontC = @(0, 128, 255)
$backC  = @(0, 70, 140)
$cF = 0; $fMinY = 9999; $fMinX = 9999
$cB = 0; $bMinY = 9999; $bMinX = 9999
for ($y = 0; $y -lt $th; $y++) {
    $row = $y * $stride
    for ($x = 0; $x -lt 300; $x++) {
        $o = $row + $x * 4
        $b = [int]$buf[$o]; $g2 = [int]$buf[$o + 1]; $r2 = [int]$buf[$o + 2]
        if ([Math]::Abs($r2 - $frontC[0]) -le 6 -and [Math]::Abs($g2 - $frontC[1]) -le 6 -and [Math]::Abs($b - $frontC[2]) -le 6) {
            $cF++; if ($y -lt $fMinY) { $fMinY = $y; $fMinX = $x } elseif ($y -eq $fMinY -and $x -lt $fMinX) { $fMinX = $x }
        } elseif ([Math]::Abs($r2 - $backC[0]) -le 6 -and [Math]::Abs($g2 - $backC[1]) -le 6 -and [Math]::Abs($b - $backC[2]) -le 6) {
            $cB++; if ($y -lt $bMinY) { $bMinY = $y; $bMinX = $x } elseif ($y -eq $bMinY -and $x -lt $bMinX) { $bMinX = $x }
        }
    }
}
$read.Add("dot n=$cF minY=$fMinY topX=$fMinX | 3d n=$cB minY=$bMinY topX=$bMinX")

if ($cF -lt 40) { $issues.Add("正面圆像素太少 n=$cF (<40)") }
if ($cB -lt 20) { $issues.Add("后墩（暗色）像素太少 n=$cB (<20) —— 点的 3D 没画出来") }
if ($cF -ge 40 -and $cB -ge 20) {
    if ([Math]::Abs(($fMinY - $bMinY) - 4) -gt 2) { $issues.Add("后墩比正面圆高 $($fMinY - $bMinY)px（期望 4±2）") }
    if ([Math]::Abs(($bMinX - $fMinX) - 4) -gt 2) { $issues.Add("后墩顶行起点比正面圆右移 $($bMinX - $fMinX)px（期望 4±2）") }
}

if (-not $p.HasExited) { try { $p.Kill() } catch {} }

if ($issues.Count -gt 0) {
    Write-Output ("FAIL tc_point3d_px chart=${tw}x${th} dpi=$dpi " + ($read -join " "))
    foreach ($it in $issues) { Write-Output ("  " + $it) }
    Write-Output ("  png=" + $pngPath)
    exit 1
}
Write-Output ("PASS tc_point3d_px chart=${tw}x${th} dpi=$dpi " + ($read -join " ") + " png=" + $pngPath)
exit 0
