param(
    [Parameter(Mandatory=$true)][string]$Exe,
    [string]$OutDir = "",
    [int]$Sec = 4
)
# 内置 TChart 的饼图掰开判据（tests\ctrlteechart\TcPieApp.vbp）。
#
# 为什么用“可见半径”当判据：扇区一顿乱掰的时候，颜色面积/像素总数都不敏感，
# 而“这块颜色的最远点到饼心的距离”直接等于 r + 掰开量。夹具是四等份：
#   第 0 块 -90°..0°（右上），中线 -45°，掰 18px ⇒ 最远点 ≈ r + 18；
#   其余三块不动 ⇒ 最远点 ≈ r。
# 饼心/半径由布局常量算（客户区 401x257、Panel 边距 12、无 header、图例在右）：
#   plot = (28,27)-(302,220) ⇒ cx=165 cy=123 r=94
# 这三条一旦被布局改动位移，下面会同时列读数，不会只报一句“半径不对”。

Add-Type -AssemblyName System.Drawing
Add-Type @"
using System; using System.Text; using System.Runtime.InteropServices;
public class PiePx {
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
if (-not $OutDir) { $OutDir = Join-Path $wd "tcpie" }
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$pngPath = Join-Path $OutDir "tcpie_chart.png"
$logOut = Join-Path $OutDir "tcpie.out"
$logErr = Join-Path $OutDir "tcpie.err"

$p = Start-Process -FilePath $exePath -WorkingDirectory $wd -PassThru `
     -RedirectStandardOutput $logOut -RedirectStandardError $logErr
try { $p.WaitForInputIdle(15000) | Out-Null } catch {}
Start-Sleep -Seconds $Sec

$main = [PiePx]::FindMain([uint32]$p.Id)
if ($main -eq [IntPtr]::Zero) {
    Write-Output ("FAIL 主窗未起 (pid=" + $p.Id + ")")
    try { $p.Kill() } catch {}
    exit 1
}
[PiePx]::ShowWindow($main, 9) | Out-Null
[PiePx]::SetForegroundWindow($main) | Out-Null
Start-Sleep -Milliseconds 600

$chart = [PiePx]::FindChart($main)
if ($chart -eq [IntPtr]::Zero) {
    Write-Output ("FAIL 主窗可见但无 VB6_TCHART 子窗 (pid=" + $p.Id + ")")
    try { $p.Kill() } catch {}
    exit 1
}
[PiePx]::RedrawWindow($chart, [IntPtr]::Zero, [IntPtr]::Zero, 0x0185) | Out-Null
Start-Sleep -Milliseconds 400

$chr = New-Object PiePx+RECT; [PiePx]::GetClientRect($chart, [ref]$chr) | Out-Null
$tw = $chr.R; $th = $chr.B
$dpi = [int][PiePx]::GetDpiForWindow($main)
if ($tw -le 0 -or $th -le 0) {
    Write-Output ("FAIL 图表客户区为空 ${tw}x${th}")
    try { $p.Kill() } catch {}
    exit 1
}

$bmp = New-Object System.Drawing.Bitmap($tw, $th)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$hdc = $g.GetHdc()
$srcDc = [PiePx]::GetDC($chart)
[PiePx]::BitBlt($hdc, 0, 0, $tw, $th, $srcDc, 0, 0, 0x00CC0020) | Out-Null
[PiePx]::ReleaseDC($chart, $srcDc) | Out-Null
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

if ($tw -ne 401 -or $th -ne 257) { $issues.Add("图表 ${tw}x${th} != 401x257（夹具几何变了，期望值要重录）") }
if ($dpi -ne 96) { $issues.Add("DPI=$dpi != 96 (判据按 96dpi 录制)") }

# 布局常量 → 饼心/半径（无 header、Panel 边距 12、图例在右 ⇒ plot=(28,27)-(302,220)）
$cx = 165; $cy = 123; $r = 94
$read.Add("chart=${tw}x${th} dpi=$dpi center=($cx,$cy) r=$r")

# tcPalette 前四色（内置实现的固定调色板）
$pal = @(
    @(0, 128, 255),   # 0 蓝  —— 被掰开的那块（右上象限，中线 -45°）
    @(255, 0, 0),     # 1 红
    @(0, 192, 0),     # 2 绿
    @(255, 192, 0)    # 3 黄
)
$maxD = @(-1.0, -1.0, -1.0, -1.0)
$cnt  = @(0, 0, 0, 0)
for ($y = 0; $y -lt $th; $y++) {
    $row = $y * $stride
    # 只扫绘图区左半：图例色块用的也是调色板 0 的蓝，扫进来会把第 0 块的
    # “最远点”拉到图例上（实测 197.6 vs 预期 112）。
    for ($x = 0; $x -lt 300; $x++) {
        $o = $row + $x * 4
        $b = [int]$buf[$o]; $g2 = [int]$buf[$o + 1]; $r2 = [int]$buf[$o + 2]
        for ($i = 0; $i -lt 4; $i++) {
            if ([Math]::Abs($r2 - $pal[$i][0]) -le 6 -and
                [Math]::Abs($g2 - $pal[$i][1]) -le 6 -and
                [Math]::Abs($b - $pal[$i][2]) -le 6) {
                $cnt[$i]++
                $d = [Math]::Sqrt(($x - $cx) * ($x - $cx) + ($y - $cy) * ($y - $cy))
                if ($d -gt $maxD[$i]) { $maxD[$i] = $d }
                break
            }
        }
    }
}
$read.Add(("radius0={0:N1}(n={1}) r1={2:N1} r2={3:N1} r3={4:N1}" -f $maxD[0], $cnt[0], $maxD[1], $maxD[2], $maxD[3]))

for ($i = 0; $i -lt 4; $i++) {
    if ($cnt[$i] -lt 400) { $issues.Add("第 $i 块颜色像素太少 n=$($cnt[$i]) (<400) —— 没画出来或调色板变了") }
}
$others = @($maxD[1], $maxD[2], $maxD[3]) | Measure-Object -Maximum | ForEach-Object { $_.Maximum }
if ($cnt[0] -ge 400) {
    if ([Math]::Abs($maxD[0] - ($r + 18)) -gt 8) { $issues.Add(("第 0 块可见半径 {0:N1} 与 r+18={1} 差超 8（掰开量没进绘制）" -f $maxD[0], ($r + 18))) }
    if (($maxD[0] - $others) -lt 10) { $issues.Add(("第 0 块半径 {0:N1} 不比其余三块(最大 {1:N1}) 多 10px 以上" -f $maxD[0], $others)) }
}
if ($others -lt ($r - 8) -or $others -gt ($r + 8)) { $issues.Add(("未掰开的扇区可见半径 {0:N1} 偏离 r={1} 超 8" -f $others, $r)) }

if (-not $p.HasExited) { try { $p.Kill() } catch {} }

if ($issues.Count -gt 0) {
    Write-Output ("FAIL tc_pie_px " + ($read -join " "))
    foreach ($it in $issues) { Write-Output ("  " + $it) }
    Write-Output ("  png=" + $pngPath)
    exit 1
}
Write-Output ("PASS tc_pie_px " + ($read -join " ") + " png=" + $pngPath)
exit 0
