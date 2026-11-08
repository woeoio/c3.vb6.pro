# ucslider_px.ps1 - ucSlider_2022 (Proyecto1.vbp) 像素判据
#
# 裁判 = 真 VB6 产物 D:\vb_yqt4qPac\ucSlider_2022\ucSliderDemo.exe（同相机重测的基准读数见
# §基线）。本脚本只做三件事，判据一律钉死、不放宽、不 SKIP：
#
#   A1  Picture1 矩形内 #393939 占比 >= 0.80
#        真 VB6 实测 0.881；旧 C3 产物实测 0.349（UC 宿主把 Picture1 底色刷成窗体 f0f0f0）。
#   A2  Picture1 矩形内「PNG 旋钮蓝灰体」像素数 >= 400
#        .frx 0x002C 的旋钮 PNG 体色是 (11,12,15)/(31,35,46)/(54,65,85) 一族蓝灰；
#        真 VB6 实测 1135 枚（bbox 833,240-892,265）；旧 C3 产物 0 枚（白色占位圆角块）。
#        只判「非白」会把白色占位块判绿，所以必须抓这族蓝色分量。
#   B   ucSlider1(4) 矩形内纯黑 (0,0,0) 像素数 == 0
#        旧 C3 产物多画一条 221x21 纯黑实心条（client 坐标 bbox 765..985 / 414..434 附近），
#        真 VB6 同区 0 枚。
#
# 采集口径：只抓窗体 **客户区**（GetClientRect + ClientToScreen），高度 539 < 576，
# 天然排除桌面壁纸那一条污染（不要用放宽阈值来绕）。
#
# 退出码：0 = 三条全绿；1 = 任一条红或前置条件不满足（本脚本**不提供 SKIP**）。
param(
    [Parameter(Mandatory=$true)][string]$Exe,
    [string]$OutDir = "",
    [int]$Sec = 10,
    [switch]$Baseline
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System; using System.Text; using System.Runtime.InteropServices;
public class UcsPx {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder sb, int n);
  [DllImport("user32.dll")] public static extern bool GetClientRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool ClientToScreen(IntPtr h, ref POINT p);
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h, int msg, IntPtr w, IntPtr l);
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
  [StructLayout(LayoutKind.Sequential)] public struct POINT { public int X, Y; }
}
"@

$p = Start-Process -FilePath $Exe -PassThru
[UcsPx]::SetProcessDPIAware() | Out-Null

$hwnd = [IntPtr]::Zero
$deadline = (Get-Date).AddSeconds($Sec)
while ((Get-Date) -lt $deadline) {
    $script:found = [IntPtr]::Zero
    [UcsPx]::EnumWindows({ param($h,$l)
        $pid2 = [uint32]0; [UcsPx]::GetWindowThreadProcessId($h, [ref]$pid2) | Out-Null
        if ($pid2 -eq $p.Id -and [UcsPx]::IsWindowVisible($h)) {
            $sb = New-Object Text.StringBuilder 256
            [UcsPx]::GetClassName($h, $sb, 256) | Out-Null
            $cn = $sb.ToString()
            if ($cn -eq '#32770') {
                # 模态框脚本化关闭：本工程正常启动不该有；真出现就关掉并记录，脚本不挂死。
                [UcsPx]::SendMessage($h, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null
                Write-Output ("WARN modal dialog closed: " + $h)
                return $true
            }
            if ($cn.StartsWith('VB6_Form_') -or $cn.StartsWith('ThunderRT6Form')) { $script:found = $h; return $false }
        }
        return $true
    }, [IntPtr]::Zero) | Out-Null
    if ($script:found -ne [IntPtr]::Zero) { $hwnd = $script:found; break }
    Start-Sleep -Milliseconds 100
}
if ($hwnd -eq [IntPtr]::Zero) {
    Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
    Write-Output "NO-WINDOW ($Exe)"
    exit 1
}

[UcsPx]::SetForegroundWindow($hwnd) | Out-Null
Start-Sleep -Milliseconds 700

$cr = New-Object UcsPx+RECT; [UcsPx]::GetClientRect($hwnd, [ref]$cr) | Out-Null
$pt = New-Object UcsPx+POINT; [UcsPx]::ClientToScreen($hwnd, [ref]$pt) | Out-Null
$cw = $cr.R - $cr.L; $ch = $cr.B - $cr.T
if ($cw -le 0 -or $ch -le 0) {
    Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
    Write-Output "BAD-CLIENT ${cw}x${ch}"
    exit 1
}
# 采集口径硬闸：
#   1) 只允许客户区（本脚本只有 GetClientRect 一条采集路径），高度必须 < 576 —— 桌面壁纸
#      那条污染进不来（这是采集侧的约束，不是靠放宽阈值绕）。
#   2) 还要核「这是 816x431 DIP 那张窗体的客户区」：按 cw 反推 scale，ch 必须等于
#      431*scale ±1。**不许写死 539**：本机是 125%(1020x539)，CI runner 是 96%(816x431)，
#      写死本机读数会让正确的判据在别的机器上"看起来红"。
$scaleProbe = [double]$cw / 816.0
$expectCh = [int][Math]::Round(431.0 * $scaleProbe)
if ($cw -lt 816 -or $ch -ge 576 -or [Math]::Abs($ch - $expectCh) -gt 1) {
    Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
    Write-Output ("CAPTURE-CONTRACT cw=${cw} ch=${ch} (期望 816x431 DIP 的客户区等比 => ch≈${expectCh}, 且 <576)")
    exit 1
}

$bmp = New-Object Drawing.Bitmap $cw, $ch
$g = [Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen($pt.X, $pt.Y, 0, 0, (New-Object Drawing.Size $cw, $ch))
$g.Dispose()
if ($OutDir -ne "" -and (Test-Path $OutDir)) {
    $bmp.Save((Join-Path $OutDir "ucslider_client.png"), [Drawing.Imaging.ImageFormat]::Png)
}
Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue

# 设计期几何（twips / 15 = DIP），再按客户区像素/816 缩放
$scale = [double]$cw / 816.0
function RectPx([double]$l, [double]$t, [double]$w, [double]$h) {
    return @([int][Math]::Round($l * $script:scale), [int][Math]::Round($t * $script:scale),
             [int][Math]::Round(($l + $w) * $script:scale), [int][Math]::Round(($t + $h) * $script:scale))
}
$pic1    = RectPx 600 56 177 161                             # Picture1 (9000,840,2655,2415 twips)
$slider4 = RectPx 608 304 177 73                             # ucSlider1(4) (9120,4560,2655,1095 twips)

function Stats([int[]]$rc) {
    $x0, $y0, $x1, $y1 = $rc
    $x1 = [Math]::Min($x1, $script:cw); $y1 = [Math]::Min($y1, $script:ch)
    $tot = 0; $bg = 0; $blue = 0; $black = 0
    for ($y = $y0; $y -lt $y1; $y++) {
        for ($x = $x0; $x -lt $x1; $x++) {
            $c = $script:bmp.GetPixel($x, $y); $tot++
            $r = $c.R; $gg = $c.G; $b = $c.B
            if ([Math]::Abs($r - 57) -le 5 -and [Math]::Abs($gg - 57) -le 5 -and [Math]::Abs($b - 57) -le 5) { $bg++ }
            if ($b -lt 140 -and $r -lt 90 -and $b -ge ($r + 6) -and $b -ge ($gg + 3)) { $blue++ }
            if ($r -eq 0 -and $gg -eq 0 -and $b -eq 0) { $black++ }
        }
    }
    return @{ Tot = $tot; Bg = $bg; Blue = $blue; Black = $black }
}

$script:bmp = $bmp; $script:cw = $cw; $script:ch = $ch; $script:scale = $scale
$s1 = Stats $pic1
$s4 = Stats $slider4
$frac = if ($s1.Tot -gt 0) { [Math]::Round($s1.Bg / $s1.Tot, 3) } else { 0 }

$fail = @()
if (-not ($frac -ge 0.80))  { $fail += ("A1 bg393939=$frac (<0.80, ref 0.881)") }
if (-not ($s1.Blue -ge 400)) { $fail += ("A2 knobPngBlue=" + $s1.Blue + " (<400, ref 1135)") }
if (-not ($s4.Black -eq 0))  { $fail += ("B slider4PureBlack=" + $s4.Black + " (must be 0)") }

$line = ("ucslider cw={0}x{1} scale={2} A1.bg={3} A2.knobPngBlue={4} B.slider4PureBlack={5}" -f $cw, $ch, $scale, $frac, $s1.Blue, $s4.Black)
if ($Baseline) { Write-Output ("BASELINE " + $line); exit 0 }
if ($fail.Count -gt 0) { Write-Output ($line + " FAIL: " + ($fail -join "; ")); exit 1 }
Write-Output ($line + " PASS")
exit 0
