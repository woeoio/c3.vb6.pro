param(
    [Parameter(Mandatory=$true)][string]$Exe,
    [Parameter(Mandatory=$true)][ValidateSet('gantt','bubble','donut','badtype')][string]$Mode,
    [string]$OutDir = "",
    [int]$Sec = 4
)
# C30-10 第十八刀的四种判据（甘特 / 气泡 / 圆环 / 不支持类型报错）。
# 四种共用一套“起窗 → 抓图 → 读 marker”的骨架，判定各写各的：
#   gantt  ：三条横带，包络从 plot.left(28) 到 plot.right(302)
#   bubble ：气泡色像素够多，最大水平直径落在 2*(3..14) 这档（半径按 r 映射）
#   donut  ：圆心是 Panel 底色（洞），半径 0.75R 处是系列色（环）—— 缺一条就退化成饼图
#   badtype：Err.Number 必须是 5，报错后 Series.Count 仍是 0，再接 17 能建成
#
# 颜色口径：正面 RGB(0,128,255)；Panel 底 RGB(255,240,240)（台账 §5 记录）。

Add-Type -AssemblyName System.Drawing
Add-Type @"
using System; using System.Text; using System.Runtime.InteropServices;
public class TyPx {
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
if (-not $OutDir) { $OutDir = Join-Path $wd ("tc" + $Mode) }
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$pngPath = Join-Path $OutDir ("tc" + $Mode + "_chart.png")
$outFile = Join-Path $OutDir ("tc" + $Mode + ".out")

$p = Start-Process -FilePath $exePath -WorkingDirectory $wd -PassThru `
     -RedirectStandardOutput $outFile -RedirectStandardError (Join-Path $OutDir "err.txt")
try { $p.WaitForInputIdle(15000) | Out-Null } catch {}
Start-Sleep -Seconds $Sec

$issues = New-Object System.Collections.Generic.List[string]
$read   = New-Object System.Collections.Generic.List[string]

if ($Mode -eq 'badtype') {
    # 不看像素，只看报错语义
    if (-not $p.HasExited) { try { $p.Kill() } catch {} }
    Start-Sleep -Milliseconds 300
    $txt = ""
    if (Test-Path $outFile) { $txt = [Text.Encoding]::GetEncoding(936).GetString([IO.File]::ReadAllBytes($outFile)) }
    $read.Add("markers=" + (($txt -split "`r`n" | Where-Object { $_ -match 'TC-BAD' }) -join '+'))
    if ($txt -notmatch 'TC-BAD-ERR=5') { $issues.Add("AddSeries 8 没有抛 Err.Number=5（静默降级成折线了？）") }
    if ($txt -notmatch 'TC-BAD-COUNT=0') { $issues.Add("报错后 Series.Count 不是 0 —— 出错却建了系列") }
    if ($txt -notmatch 'TC-BAD-OK=17') { $issues.Add("报错后合法的 AddSeries 17 也不能用") }
    if ($txt -notmatch 'TC-BAD-READY') { $issues.Add("夹具没走到 TC-BAD-READY") }
    if ($issues.Count -gt 0) {
        Write-Output ("FAIL tc_type_px/$Mode " + ($read -join " | "))
        foreach ($it in $issues) { Write-Output ("  " + $it) }
        exit 1
    }
    Write-Output ("PASS tc_type_px/$Mode " + ($read -join " | "))
    exit 0
}

$main = [TyPx]::FindMain([uint32]$p.Id)
if ($main -eq [IntPtr]::Zero) { Write-Output ("FAIL 主窗未起 (pid=" + $p.Id + ")"); try { $p.Kill() } catch {}; exit 1 }
[TyPx]::ShowWindow($main, 9) | Out-Null
Start-Sleep -Milliseconds 600
$chart = [TyPx]::FindChart($main)
if ($chart -eq [IntPtr]::Zero) { Write-Output "FAIL 无 VB6_TCHART 子窗"; try { $p.Kill() } catch {}; exit 1 }
[TyPx]::RedrawWindow($chart, [IntPtr]::Zero, [IntPtr]::Zero, 0x0185) | Out-Null
Start-Sleep -Milliseconds 400

$chr = New-Object TyPx+RECT; [TyPx]::GetClientRect($chart, [ref]$chr) | Out-Null
$tw = $chr.R; $th = $chr.B
$dpi = [int][TyPx]::GetDpiForWindow($main)

$bmp = New-Object System.Drawing.Bitmap($tw, $th)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$hdc = $g.GetHdc()
$srcDc = [TyPx]::GetDC($chart)
[TyPx]::BitBlt($hdc, 0, 0, $tw, $th, $srcDc, 0, 0, 0x00CC0020) | Out-Null
[TyPx]::ReleaseDC($chart, $srcDc) | Out-Null
$g.ReleaseHdc($hdc); $g.Dispose()
$bmp.Save($pngPath, [System.Drawing.Imaging.ImageFormat]::Png)

$rect = New-Object System.Drawing.Rectangle(0, 0, $tw, $th)
$bd = $bmp.LockBits($rect, [System.Drawing.Imaging.ImageLockMode]::ReadOnly, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
$stride = $bd.Stride
$buf = New-Object byte[] ($stride * $th)
[System.Runtime.InteropServices.Marshal]::Copy($bd.Scan0, $buf, 0, $buf.Length)
$bmp.UnlockBits($bd); $bmp.Dispose()

function Px([int]$x, [int]$y) {
    $o = $y * $stride + $x * 4
    return @([int]$buf[$o + 2], [int]$buf[$o + 1], [int]$buf[$o])
}
function IsBar([int]$r, [int]$g2, [int]$b) {
    return ([Math]::Abs($r - 0) -le 6 -and [Math]::Abs($g2 - 128) -le 6 -and [Math]::Abs($b - 255) -le 6)
}

if ($tw -ne 401 -or $th -ne 257) { $issues.Add("图表 ${tw}x${th} != 401x257") }
if ($dpi -ne 96) { $issues.Add("DPI=$dpi != 96") }

switch ($Mode) {
'gantt' {
    $minX = 9999; $maxX = -1; $n = 0
    $bands = New-Object System.Collections.Generic.List[int]
    $run = 0
    for ($y = 0; $y -lt $th; $y++) {
        $rowHit = $false
        for ($x = 0; $x -lt 306; $x++) {
            $c = Px $x $y
            if (IsBar $c[0] $c[1] $c[2]) {
                $n++; $rowHit = $true
                if ($x -lt $minX) { $minX = $x }
                if ($x -gt $maxX) { $maxX = $x }
            }
        }
        if ($rowHit) { $run++ } elseif ($run -gt 0) { $bands.Add($run); $run = 0 }
    }
    if ($run -gt 0) { $bands.Add($run) }
    $read.Add("gantt n=$n x=$minX..$maxX bands=$($bands.Count)")
    if ($n -lt 500) { $issues.Add("甘特条像素只 $n 个 —— 没画") }
    if ($bands.Count -ne 3) { $issues.Add("横带数 $($bands.Count) != 3（三条任务）") }
    if ([Math]::Abs($minX - 28) -gt 3) { $issues.Add("最左 $minX != plot.left(28±3)") }
    if ([Math]::Abs($maxX - 302) -gt 4) { $issues.Add("最右 $maxX != plot.right(302±4)") }
}
'bubble' {
    $n = 0; $bestRun = 0; $bestRow = -1
    for ($y = 0; $y -lt $th; $y++) {
        $run = 0
        for ($x = 0; $x -lt 306; $x++) {
            $c = Px $x $y
            if (IsBar $c[0] $c[1] $c[2]) { $n++; $run++; if ($run -gt $bestRun) { $bestRun = $run; $bestRow = $y } }
            else { $run = 0 }
        }
    }
    $read.Add("bubble n=$n maxDia=$bestRun row=$bestRow")
    if ($n -lt 400) { $issues.Add("气泡像素只 $n 个 —— 没画") }
    # 最大气泡 r=45 → rr=12 → 直径约 24±6
    if ([Math]::Abs($bestRun - 24) -gt 6) { $issues.Add("最大气泡直径 $bestRun px（期望 24±6，半径按 r 映射）") }
}
'donut' {
    $cx = 165; $cy = 123
    $c = Px $cx $cy
    # Panel 底色从图内空白处取（夹具没设 Panel.Color，默认不是台账里那张图的 255,240,240）
    $bg = Px 6 250
    $read.Add("center=$($c[0]),$($c[1]),$($c[2]) bg=$($bg[0]),$($bg[1]),$($bg[2])")
    if (-not ([Math]::Abs($c[0] - $bg[0]) -le 6 -and [Math]::Abs($c[1] - $bg[1]) -le 6 -and [Math]::Abs($c[2] - $bg[2]) -le 6)) {
        $issues.Add("圆心（$($c[0]),$($c[1]),$($c[2])）不是 Panel 底色（$($bg[0]),$($bg[1]),$($bg[2])）—— 没有洞，还是张饼")
    }
    $n = 0; $ringHit = 0
    for ($y = 0; $y -lt $th; $y++) {
        for ($x = 0; $x -lt 306; $x++) {
            $c2 = Px $x $y
            if (IsBar $c2[0] $c2[1] $c2[2]) {
                $n++
                $dx = $x - $cx; $dy = $y - $cy
                if ([Math]::Abs([Math]::Sqrt($dx * $dx + $dy * $dy) - 70) -le 4) { $ringHit++ }
            }
        }
    }
    $read.Add("ring n=$n at70=$ringHit")
    if ($n -lt 1000) { $issues.Add("环色像素只 $n 个 —— 圆环没画") }
    if ($ringHit -lt 40) { $issues.Add("半径 70 处环色像素只 $ringHit 个（<40）—— 环不存在") }
}
}

# marker 校验
$txt2 = ""
if (-not $p.HasExited) { try { $p.Kill() } catch {} }
Start-Sleep -Milliseconds 300
if (Test-Path $outFile) { $txt2 = [Text.Encoding]::GetEncoding(936).GetString([IO.File]::ReadAllBytes($outFile)) }
$expect = @{ gantt = @('TC-GT-TYPE=9','TC-GT-READY'); bubble = @('TC-BB-TYPE=11','TC-BB-COUNT=3','TC-BB-READY'); donut = @('TC-DN-TYPE=17','TC-DN-READY') }[$Mode]
foreach ($e in $expect) { if ($txt2 -notmatch [regex]::Escape($e)) { $issues.Add("缺 marker $e") } }

if ($issues.Count -gt 0) {
    Write-Output ("FAIL tc_type_px/$Mode chart=${tw}x${th} dpi=$dpi " + ($read -join " | "))
    foreach ($it in $issues) { Write-Output ("  " + $it) }
    Write-Output ("  png=" + $pngPath)
    exit 1
}
Write-Output ("PASS tc_type_px/$Mode chart=${tw}x${th} dpi=$dpi " + ($read -join " | ") + " png=" + $pngPath)
exit 0
