param(
    [Parameter(Mandatory=$true)][string]$Exe,
    [Parameter(Mandatory=$true)][ValidateSet('fx','walls')][string]$Mode,
    [string]$OutDir = ""
)
# C30-10 第二十刀：勾选框那批属性的**绘制**判据。对应 ChartDisplay 示例里
#   Check1(渐变) / Check3(边框) / Check5(3D) / Check6(Frame) / Check11(Walls)
#   Check4(No Back colour) / Command4(LeftWall)
# 以前这些成员只存不画 —— 勾了没反应。夹具直接把属性设成非默认值，判据量像素。
#
# fx   （TcFxApp）：渐变红→蓝、外框与 Panel 边框都关、四壁关、2D
#   · (2,2) 必须是渐变顶端的红（边框关掉后这里不该是灰边 128,128,128）
#   · (2,h-3) 必须是渐变底端的蓝
#   · 绘图区右侧 (330,120) 不能是墙灰 (192,192,192) —— 四壁关 + 2D 之后那里应是渐变
# walls（TcFx2App）：3D、背板 Brush=vbGreen、左壁 Visible=False
#   · (310,100) 必须是绿（背板刷色真的生效）
#   · (33,100) 必须是面板底色而不是绿（左壁整面被关掉）
#
# 退出码：0 PASS / 1 FAIL

Add-Type -AssemblyName System.Drawing
Add-Type @"
using System; using System.Text; using System.Runtime.InteropServices;
public class TcFx {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder sb, int n);
  [DllImport("user32.dll")] public static extern bool GetClientRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern IntPtr GetDC(IntPtr h);
  [DllImport("user32.dll")] public static extern int ReleaseDC(IntPtr h, IntPtr dc);
  [DllImport("gdi32.dll")] public static extern bool BitBlt(IntPtr d, int x, int y, int w, int h, IntPtr s, int sx, int sy, int rop);
  public struct RECT { public int L, T, R, B; }
  public static string Cls(IntPtr h) { var sb = new StringBuilder(256); GetClassName(h, sb, 256); return sb.ToString(); }
  public static IntPtr FindChart(uint pid) {
    IntPtr found = IntPtr.Zero;
    EnumWindows((h, l) => {
      uint p; GetWindowThreadProcessId(h, out p);
      if (p != pid || !IsWindowVisible(h)) return true;
      EnumChildWindows(h, (ch, l2) => {
        var c = Cls(ch);
        if (c == "TChartChart" || c == "VB6_TCHART") { found = ch; return false; }
        return true;
      }, IntPtr.Zero);
      return found == IntPtr.Zero;
    }, IntPtr.Zero);
    return found;
  }
}
"@

$exePath = (Resolve-Path $Exe).Path
$wd = Split-Path $exePath
$proc = Start-Process -FilePath $exePath -WorkingDirectory $wd -PassThru
Start-Sleep -Milliseconds 1600
$hwnd = [TcFx]::FindChart($proc.Id)
if ($hwnd -eq [IntPtr]::Zero) { Write-Output "FAIL 未找到图表子窗（VB6_TCHART）"; try { $proc.Kill() } catch {}; exit 1 }

$r = New-Object TcFx+RECT
[void][TcFx]::GetClientRect($hwnd, [ref]$r)
$W = $r.R - $r.L; $H = $r.B - $r.T
$bmp = New-Object System.Drawing.Bitmap $W, $H
$g = [System.Drawing.Graphics]::FromImage($bmp)
$hdc = $g.GetHdc(); $src = [TcFx]::GetDC($hwnd)
[void][TcFx]::BitBlt($hdc, 0, 0, $W, $H, $src, 0, 0, 0x00CC0020)
[void][TcFx]::ReleaseDC($hwnd, $src); $g.ReleaseHdc($hdc); $g.Dispose()
if (-not $OutDir) { $OutDir = Join-Path $wd ("tc_fx_" + $Mode) }
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$png = Join-Path $OutDir ("tc_fx_" + $Mode + ".png")
$bmp.Save($png, [System.Drawing.Imaging.ImageFormat]::Png)

$bad = New-Object System.Collections.Generic.List[string]
$a = $bmp.GetPixel(2, 2)
$b = $bmp.GetPixel(2, $H - 3)
$c = $bmp.GetPixel(330, 120)
$d = $bmp.GetPixel(312, 150)
$e = $bmp.GetPixel(31, 45)
$read = "chart=${W}x${H} top(2,2)=$($a.R),$($a.G),$($a.B) bot(2,$($H-3))=$($b.R),$($b.G),$($b.B) right(330,120)=$($c.R),$($c.G),$($c.B) wall(312,150)=$($d.R),$($d.G),$($d.B) lwall(31,45)=$($e.R),$($e.G),$($e.B)"

if ($Mode -eq 'fx') {
    if (-not ($a.R -gt 150 -and $a.G -lt 90 -and $a.B -lt 90)) { $bad.Add("渐变顶端不是红 ($($a.R),$($a.G),$($a.B))；外框关了却是灰边？") }
    if (-not ($b.B -gt 150 -and $b.R -lt 90)) { $bad.Add("渐变底端不是蓝 ($($b.R),$($b.G),$($b.B))") }
    if ([Math]::Abs($c.R - 192) -lt 10 -and [Math]::Abs($c.G - 192) -lt 10 -and [Math]::Abs($c.B - 192) -lt 10) {
        $bad.Add("绘图区右侧还是墙灰 (192,192,192)：Walls/View3D 没生效")
    }
} else {
    if (-not ($d.G -gt 110 -and $d.R -lt 90 -and $d.B -lt 90)) { $bad.Add("背板不是 Brush 的绿 ($($d.R),$($d.G),$($d.B))") }
    if ($e.G -gt 110 -and $e.R -lt 90 -and $e.B -lt 90) { $bad.Add("左壁没关掉（(31,45) 还是绿）") }
}

try { $proc.Kill() } catch {}
$bmp.Dispose()
if ($bad.Count -eq 0) { Write-Output ("PASS tc_fx/" + $Mode + " " + $read + " png=" + $png); exit 0 }
Write-Output ("FAIL tc_fx/" + $Mode + " " + $read + " png=" + $png)
foreach ($x in $bad) { Write-Output ("  " + $x) }
exit 1
