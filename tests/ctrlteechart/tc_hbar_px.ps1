param(
    [Parameter(Mandatory=$true)][string]$Exe,
    [string]$OutDir = "",
    [int]$Sec = 4
)
# 内置 TChart 的横向柱 + As* 切面判据（tests\ctrlteechart\TcHbarApp.vbp）。
#
# 夹具：AddSeries 2（stHorizBar）、三值 20/30/25、AsHorizBar.BarWidthPercent = 40。
# 判据三件事，任何一条不成立都说明“横柱”没真画：
#   ① 柱从 plot.left(28) 长出去，最长那根到 plot.right(302) —— 横着长的证据
#   ② 在 x=40 竖着扫，柱色应成 3 段窄带（横条），不是一整片
#   ③ 每段厚度 ≈ band*40/200（band=(220-27)/4≈48 → ≈10±5）；BarWidthPercent 被忽略时
#      厚度会是 32 → 直接抓出来
#
# 同时核对 As* 切面的读回值（.AsCandle.CandleWidth / .AsSurface.WireFrame / …）：
# 这些不绘制，但必须“不是 438”，能存能读。

Add-Type -AssemblyName System.Drawing
Add-Type @"
using System; using System.Text; using System.Runtime.InteropServices;
public class HbPx {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder sb, int n);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool GetClientRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
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
if (-not $OutDir) { $OutDir = Join-Path $wd "tchbar" }
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$pngPath = Join-Path $OutDir "tchbar_chart.png"
$outFile = Join-Path $OutDir "tchbar.out"

$p = Start-Process -FilePath $exePath -WorkingDirectory $wd -PassThru `
     -RedirectStandardOutput $outFile -RedirectStandardError (Join-Path $OutDir "tchbar.err")
try { $p.WaitForInputIdle(15000) | Out-Null } catch {}
Start-Sleep -Seconds $Sec

$main = [HbPx]::FindMain([uint32]$p.Id)
if ($main -eq [IntPtr]::Zero) { Write-Output ("FAIL 主窗未起 (pid=" + $p.Id + ")"); try { $p.Kill() } catch {}; exit 1 }
[HbPx]::ShowWindow($main, 9) | Out-Null
Start-Sleep -Milliseconds 600
$chart = [HbPx]::FindChart($main)
if ($chart -eq [IntPtr]::Zero) { Write-Output "FAIL 无 VB6_TCHART 子窗"; try { $p.Kill() } catch {}; exit 1 }
[HbPx]::RedrawWindow($chart, [IntPtr]::Zero, [IntPtr]::Zero, 0x0185) | Out-Null
Start-Sleep -Milliseconds 400

$chr = New-Object HbPx+RECT; [HbPx]::GetClientRect($chart, [ref]$chr) | Out-Null
$tw = $chr.R; $th = $chr.B
$dpi = [int][HbPx]::GetDpiForWindow($main)

$bmp = New-Object System.Drawing.Bitmap($tw, $th)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$hdc = $g.GetHdc()
$srcDc = [HbPx]::GetDC($chart)
[HbPx]::BitBlt($hdc, 0, 0, $tw, $th, $srcDc, 0, 0, 0x00CC0020) | Out-Null
[HbPx]::ReleaseDC($chart, $srcDc) | Out-Null
$g.ReleaseHdc($hdc); $g.Dispose()
$bmp.Save($pngPath, [System.Drawing.Imaging.ImageFormat]::Png)

$rect = New-Object System.Drawing.Rectangle(0, 0, $tw, $th)
$bd = $bmp.LockBits($rect, [System.Drawing.Imaging.ImageLockMode]::ReadOnly, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
$stride = $bd.Stride
$buf = New-Object byte[] ($stride * $th)
[System.Runtime.InteropServices.Marshal]::Copy($bd.Scan0, $buf, 0, $buf.Length)
$bmp.UnlockBits($bd); $bmp.Dispose()

$issues = New-Object System.Collections.Generic.List[string]
$read   = New-Object System.Collections.Generic.List[string]
if ($tw -ne 401 -or $th -ne 257) { $issues.Add("图表 ${tw}x${th} != 401x257") }
if ($dpi -ne 96) { $issues.Add("DPI=$dpi != 96") }

function IsBar([int]$r, [int]$g2, [int]$b) {
    return ([Math]::Abs($r - 0) -le 6 -and [Math]::Abs($g2 - 128) -le 6 -and [Math]::Abs($b - 255) -le 6)
}

# ① 横向包络（只扫 x<300，避开图例）
$minX = 9999; $maxX = -1; $n = 0
for ($y = 0; $y -lt $th; $y++) {
    $row = $y * $stride
    for ($x = 0; $x -lt 306; $x++) {   # 306：最长柱画到 plot.right=302，但 300 会把右缘裁掉（v 第一版读数 299）
        $o = $row + $x * 4
        if (IsBar ([int]$buf[$o + 2]) ([int]$buf[$o + 1]) ([int]$buf[$o])) {
            $n++
            if ($x -lt $minX) { $minX = $x }
            if ($x -gt $maxX) { $maxX = $x }
        }
    }
}
$read.Add("hb n=$n x=$minX..$maxX")
if ($n -lt 500) { $issues.Add("柱色像素只 $n 个 —— 横向柱没画") }
if ($n -ge 500) {
    if ([Math]::Abs($minX - 28) -gt 3) { $issues.Add("柱左缘 $minX != plot.left(28±3)") }
    if ([Math]::Abs($maxX - 302) -gt 4) { $issues.Add("最长柱右缘 $maxX != plot.right(302±4)") }
}

# ②③ x=40 处竖扫：数段数 + 每段厚度
$bands = New-Object System.Collections.Generic.List[int]
$run = 0
for ($y = 0; $y -lt $th; $y++) {
    $o = $y * $stride + 40 * 4
    if (IsBar ([int]$buf[$o + 2]) ([int]$buf[$o + 1]) ([int]$buf[$o])) { $run++ }
    elseif ($run -gt 0) { $bands.Add($run); $run = 0 }
}
if ($run -gt 0) { $bands.Add($run) }
$read.Add("bands=" + ($bands -join ','))
if ($bands.Count -ne 3) {
    $issues.Add("x=40 竖扫到 $($bands.Count) 段柱色（期望 3 段横条）")
} else {
    foreach ($b in $bands) {
        if ([Math]::Abs($b - 10) -gt 5) { $issues.Add("横条厚度 $b px（BarWidthPercent=40、四值时 期望 10±5）"); break }
    }
}

# marker 与切面读回
$txt = ""
if (-not $p.HasExited) { try { $p.Kill() } catch {} }
Start-Sleep -Milliseconds 300
if (Test-Path $outFile) { $txt = [Text.Encoding]::GetEncoding(936).GetString([IO.File]::ReadAllBytes($outFile)) }
foreach ($kv in @(@('TC-HB-TYPE','2'), @('TC-HB-PCT','40'), @('TC-HB-ASBAR','40'), @('TC-HB-CANDLE','60'), @('TC-HB-SURF','1'), @('TC-HB-LINE','0'))) {
    if ($txt -notmatch ($kv[0] + '=' + $kv[1] + '\b')) { $issues.Add("切面读回不对：期望 " + $kv[0] + "=" + $kv[1]) }
}
if ($txt -notmatch 'TC-HB-READY') { $issues.Add("夹具没走到 TC-HB-READY") }

if ($issues.Count -gt 0) {
    Write-Output ("FAIL tc_hbar_px chart=${tw}x${th} dpi=$dpi " + ($read -join " | "))
    foreach ($it in $issues) { Write-Output ("  " + $it) }
    Write-Output ("  png=" + $pngPath)
    exit 1
}
Write-Output ("PASS tc_hbar_px chart=${tw}x${th} dpi=$dpi " + ($read -join " | ") + " png=" + $pngPath)
exit 0
