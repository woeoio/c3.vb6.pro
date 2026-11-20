param(
    [Parameter(Mandatory=$true)][string]$Exe,
    [string]$OutDir = "",
    [int]$Sec = 4
)
# 内置 TChart 的 3D 柱几何判据（tests\ctrlteechart\Tc3dApp.vbp）。
#
# 判据只量两件事，都是“深度常量 12”这一把尺子：
#   ① 侧面（正面色再暗 45%）的右缘比正面的右缘多出 12px  —— 深度真的横着伸出去
#   ② 顶面（正面色再亮 55%）的上缘比柱顶上缘高 12px       —— 深度真的朝上
# 这样“3D 打开了但画成 2D”“顶面画到下面去”“深度随手变”都会红。
#
# 颜色是算出来的，不是猜的：正面 RGB(0,128,255) ⇒
#   亮 55%: r=0+255*55/100=140, g=128+(255-128)*55/100=197, b=255  → (140,197,255)
#   暗 45%: r=0, g=128*55/100=70,   b=255*55/100=140              → (0,70,140)

Add-Type -AssemblyName System.Drawing
Add-Type @"
using System; using System.Text; using System.Runtime.InteropServices;
public class D3Px {
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
if (-not $OutDir) { $OutDir = Join-Path $wd "tc3d" }
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$pngPath = Join-Path $OutDir "tc3d_chart.png"

$p = Start-Process -FilePath $exePath -WorkingDirectory $wd -PassThru `
     -RedirectStandardOutput (Join-Path $OutDir "tc3d.out") `
     -RedirectStandardError (Join-Path $OutDir "tc3d.err")
try { $p.WaitForInputIdle(15000) | Out-Null } catch {}
Start-Sleep -Seconds $Sec

$main = [D3Px]::FindMain([uint32]$p.Id)
if ($main -eq [IntPtr]::Zero) {
    Write-Output ("FAIL 主窗未起 (pid=" + $p.Id + ")")
    try { $p.Kill() } catch {}
    exit 1
}
[D3Px]::ShowWindow($main, 9) | Out-Null
[D3Px]::SetForegroundWindow($main) | Out-Null
Start-Sleep -Milliseconds 600

$chart = [D3Px]::FindChart($main)
if ($chart -eq [IntPtr]::Zero) {
    Write-Output ("FAIL 主窗可见但无 VB6_TCHART 子窗 (pid=" + $p.Id + ")")
    try { $p.Kill() } catch {}
    exit 1
}
[D3Px]::RedrawWindow($chart, [IntPtr]::Zero, [IntPtr]::Zero, 0x0185) | Out-Null
Start-Sleep -Milliseconds 400

$chr = New-Object D3Px+RECT; [D3Px]::GetClientRect($chart, [ref]$chr) | Out-Null
$tw = $chr.R; $th = $chr.B
$dpi = [int][D3Px]::GetDpiForWindow($main)

$bmp = New-Object System.Drawing.Bitmap($tw, $th)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$hdc = $g.GetHdc()
$srcDc = [D3Px]::GetDC($chart)
[D3Px]::BitBlt($hdc, 0, 0, $tw, $th, $srcDc, 0, 0, 0x00CC0020) | Out-Null
[D3Px]::ReleaseDC($chart, $srcDc) | Out-Null
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

$front = @(0, 128, 255)
$topC  = @(140, 197, 255)
$sideC = @(0, 70, 140)

$cntF = 0; $cntT = 0; $cntS = 0
$fMaxX = -1; $fMinY = 9999
$tMinY = 9999
$sMaxX = -1
for ($y = 0; $y -lt $th; $y++) {
    $row = $y * $stride
    # 只扫 x<300：图例色块用的也是系列色（这就是正面色），追进来会把
    # “正面右缘”拉到图例上（实测 maxX=322、minY=3，差值全变负）。
    for ($x = 0; $x -lt 300; $x++) {
        $o = $row + $x * 4
        $b = [int]$buf[$o]; $g2 = [int]$buf[$o + 1]; $r2 = [int]$buf[$o + 2]
        if ([Math]::Abs($r2 - $front[0]) -le 6 -and [Math]::Abs($g2 - $front[1]) -le 6 -and [Math]::Abs($b - $front[2]) -le 6) {
            $cntF++; if ($x -gt $fMaxX) { $fMaxX = $x }; if ($y -lt $fMinY) { $fMinY = $y }
        } elseif ([Math]::Abs($r2 - $topC[0]) -le 6 -and [Math]::Abs($g2 - $topC[1]) -le 6 -and [Math]::Abs($b - $topC[2]) -le 6) {
            $cntT++; if ($y -lt $tMinY) { $tMinY = $y }
        } elseif ([Math]::Abs($r2 - $sideC[0]) -le 6 -and [Math]::Abs($g2 - $sideC[1]) -le 6 -and [Math]::Abs($b - $sideC[2]) -le 6) {
            $cntS++; if ($x -gt $sMaxX) { $sMaxX = $x }
        }
    }
}
$read.Add("front n=$cntF maxX=$fMaxX minY=$fMinY | top n=$cntT minY=$tMinY | side n=$cntS maxX=$sMaxX")

if ($cntF -lt 500) { $issues.Add("正面像素太少 n=$cntF (<500) —— 柱没画出来或颜色变了") }
if ($cntS -lt 150) { $issues.Add("侧面像素太少 n=$cntS (<150) —— 没有 3D 深度（View3D 没生效？）") }
elseif ([Math]::Abs(($sMaxX - $fMaxX) - 12) -gt 3) { $issues.Add("侧面右缘比正面右缘多 $($sMaxX - $fMaxX)px（期望 12±3）") }
if ($cntT -lt 100) { $issues.Add("顶面像素太少 n=$cntT (<100) —— 顶面没画出来") }
elseif ([Math]::Abs(($fMinY - $tMinY) - 12) -gt 3) { $issues.Add("顶面比柱顶高 $($fMinY - $tMinY)px（期望 12±3，负数=画到下面去了）") }

if (-not $p.HasExited) { try { $p.Kill() } catch {} }

if ($issues.Count -gt 0) {
    Write-Output ("FAIL tc_3d_px chart=${tw}x${th} dpi=$dpi " + ($read -join " "))
    foreach ($it in $issues) { Write-Output ("  " + $it) }
    Write-Output ("  png=" + $pngPath)
    exit 1
}
Write-Output ("PASS tc_3d_px chart=${tw}x${th} dpi=$dpi " + ($read -join " ") + " png=" + $pngPath)
exit 0
