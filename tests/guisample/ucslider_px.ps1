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
  [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
  [DllImport("user32.dll")] public static extern bool GetCursorPos(out POINT p);
  [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);
  [DllImport("user32.dll")] public static extern void mouse_event(uint flags, uint dx, uint dy, uint data, IntPtr extra);
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
  [StructLayout(LayoutKind.Sequential)] public struct POINT { public int X, Y; }
}
"@

# ---- 双击等价启动（用户 2026-11-09 报「双击后自己就关了」的回归闸）----
# 三条等价条件：
#   ① 工作目录 = exe 所在目录（Start-Process 默认继承父进程 cwd，那不是双击形态）；
#   ② 环境里没有本仓库/CI 注入的变量（C3_* 等），只留 OS 必需的最小集合；
#   ③ exe 旁边没有源工程的 img\/xml —— 先拷进一枚空目录再跑，任何「靠 cwd 找文件」
#      的依赖在裸目录里都会现形（缺文件必须红）。
$script:work = Join-Path ([IO.Path]::GetTempPath()) ("ucslider_dbl_" + [Guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Path $script:work -Force | Out-Null
$runExe = Join-Path $script:work (Split-Path -Leaf $Exe)
Copy-Item -LiteralPath $Exe -Destination $runExe -Force

$psi = New-Object Diagnostics.ProcessStartInfo
$psi.FileName = $runExe
$psi.WorkingDirectory = $script:work
$psi.UseShellExecute = $false
$psi.EnvironmentVariables.Clear()
$minEnv = @{
    'SystemRoot' = $env:SystemRoot; 'windir' = $env:windir; 'SystemDrive' = $env:SystemDrive
    'TEMP' = $env:TEMP; 'TMP' = $env:TMP; 'PATH' = $env:PATH; 'COMSPEC' = $env:COMSPEC
    'PATHEXT' = $env:PATHEXT; 'USERPROFILE' = $env:USERPROFILE; 'APPDATA' = $env:APPDATA
    'LOCALAPPDATA' = $env:LOCALAPPDATA; 'ProgramData' = $env:ProgramData
    'ALLUSERSPROFILE' = $env:ALLUSERSPROFILE
}
foreach ($k in $minEnv.Keys) {
    if ($minEnv[$k]) { $psi.EnvironmentVariables[$k] = [string]$minEnv[$k] }
}
function Stop-App {
    try { if ($script:appProc -and -not $script:appProc.HasExited) { Stop-Process -Id $script:appProc.Id -Force -ErrorAction SilentlyContinue } } catch { }
    try { Remove-Item -LiteralPath $script:work -Recurse -Force -ErrorAction SilentlyContinue } catch { }
}
$p = [Diagnostics.Process]::Start($psi)
$script:appProc = $p
$script:appStart = $p.StartTime
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
    Stop-App
    Write-Output "NO-WINDOW ($Exe)"
    exit 1
}

# ---- 双击后用户一定会**移动鼠标**：扫过客户区（60px 网格）触发每个 ucSlider 的 MouseEnter。
# 修前这条路径直接 0xC0000005（AddRef 写窗口内存），所以扫完必须还活着；再观察满启动后 20s
# （实测未修版的崩溃窗口 5~18s，MouseEnter 那一路是立刻崩）。期间每 250ms 探活，早退也是红。
[UcsPx]::SetForegroundWindow($hwnd) | Out-Null
Start-Sleep -Milliseconds 500
$cr0 = New-Object UcsPx+RECT; [UcsPx]::GetClientRect($hwnd, [ref]$cr0) | Out-Null
$pt0 = New-Object UcsPx+POINT; [UcsPx]::ClientToScreen($hwnd, [ref]$pt0) | Out-Null
$cw0 = $cr0.R - $cr0.L; $ch0 = $cr0.B - $cr0.T
if ($cw0 -gt 0 -and $ch0 -gt 0) {
    $saved = New-Object UcsPx+POINT; [UcsPx]::GetCursorPos([ref]$saved) | Out-Null
    for ($gy = 4; $gy -lt $ch0; $gy += 60) {
        for ($gx = 4; $gx -lt $cw0; $gx += 60) {
            [UcsPx]::SetCursorPos($pt0.X + $gx, $pt0.Y + $gy) | Out-Null
            Start-Sleep -Milliseconds 25
        }
    }
    [UcsPx]::SetCursorPos($saved.X, $saved.Y) | Out-Null
}
while (((Get-Date) - $script:appStart).TotalSeconds -lt 20) {
    if ($p.HasExited) { break }
    Start-Sleep -Milliseconds 250
}
if ($p.HasExited) {
    $codeTxt = "?"
    try {
        # PS 的 [uint32](-1073741819) 会抛转换异常 (按十进制字符串转), 手工补 2^32。
        $code = [int64]$p.ExitCode
        if ($code -lt 0) { $code += 0x100000000 }
        $codeTxt = ("0x{0:X8}" -f [int64]$code)
    } catch { }
    Stop-App
    Write-Output ("CRASH exit=${codeTxt} (双击等价启动后 20s 内进程自己退出)")
    exit 1
}

[UcsPx]::SetForegroundWindow($hwnd) | Out-Null
Start-Sleep -Milliseconds 700

$cr = New-Object UcsPx+RECT; [UcsPx]::GetClientRect($hwnd, [ref]$cr) | Out-Null
$pt = New-Object UcsPx+POINT; [UcsPx]::ClientToScreen($hwnd, [ref]$pt) | Out-Null
$cw = $cr.R - $cr.L; $ch = $cr.B - $cr.T
if ($cw -le 0 -or $ch -le 0) {
    Stop-App
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
    Stop-App
    Write-Output ("CAPTURE-CONTRACT cw=${cw} ch=${ch} (期望 816x431 DIP 的客户区等比 => ch≈${expectCh}, 且 <576)")
    exit 1
}

# 采集前置顶：桌面其它窗口（资源管理器等）压在上面会抓到别人的像素（实测过
# A1=0.005 / 主色 #191919 的假读数）。只改 Z 序，不改窗口几何。
[UcsPx]::SetWindowPos($hwnd, [IntPtr](-1), 0, 0, 0, 0, 0x43) | Out-Null
Start-Sleep -Milliseconds 200

function Grab([string]$name) {
    $b2 = New-Object Drawing.Bitmap $cw, $ch
    $g2 = [Drawing.Graphics]::FromImage($b2)
    $g2.CopyFromScreen($pt.X, $pt.Y, 0, 0, (New-Object Drawing.Size $cw, $ch))
    $g2.Dispose()
    if ($OutDir -ne "" -and (Test-Path $OutDir)) {
        $b2.Save((Join-Path $OutDir ($name + ".png")), [Drawing.Imaging.ImageFormat]::Png)
    }
    return $b2
}

$scale = [double]$cw / 816.0

# ---- D｜拖动门（ucSlider4 竖向: 456,32,113,329 DIP，Value=0 时旋钮贴底）----
# 判据取自真 VB6 exe 现场取证（同相机 125% 下 drag 150 DIP）：
#   拖动中/松手后 ucSlider4 矩形内的轨道绿 (57,140,0) 像素 = 1781，顶边 y=256；
#   Value=0 时 = 0。旧 C3 产物把 WM_MOUSEMOVE 的 Button 硬写成 0 ⇒
#   `UserControl_MouseMove` 里 `If Button = vbLeftButton` 永不成立 ⇒ 拖完仍是 0。
# 必须抓**按住期间**的 mid：只判松手后的话，MouseDown 里的轨道跳值也能装成拖动。
$s4vX0 = [int][Math]::Round(456 * $scale); $s4vY0 = [int][Math]::Round(32 * $scale)
$s4vX1 = [int][Math]::Round((456 + 113) * $scale); $s4vY1 = [int][Math]::Round((32 + 329) * $scale)
$s4v = @($s4vX0, $s4vY0, $s4vX1, $s4vY1)

$preBmp = Grab "ucslider_predrag"
$park = New-Object UcsPx+POINT; [UcsPx]::GetCursorPos([ref]$park) | Out-Null

$pressX = $pt.X + [int][Math]::Round((456 + 113 / 2.0) * $scale)
$pressY = $pt.Y + [int][Math]::Round((32 + 329) * $scale) - [int][Math]::Round(30 * $scale)
[UcsPx]::SetCursorPos($pressX, $pressY) | Out-Null
Start-Sleep -Milliseconds 200
[UcsPx]::mouse_event(0x0002, 0, 0, 0, [IntPtr]::Zero)      # LEFTDOWN
Start-Sleep -Milliseconds 150
for ($di = 1; $di -le 12; $di++) {
    $yy = $pressY - [int][Math]::Round(150 * $scale * $di / 12)
    [UcsPx]::SetCursorPos($pressX, $yy) | Out-Null
    Start-Sleep -Milliseconds 50
}
Start-Sleep -Milliseconds 200
$midBmp = Grab "ucslider_midrag"
[UcsPx]::mouse_event(0x0004, 0, 0, 0, [IntPtr]::Zero)      # LEFTUP
Start-Sleep -Milliseconds 300
# 光标挪到窗体空白处，让 LabelPlus 气泡收起来再做 A1/A2/B 采集
[UcsPx]::SetCursorPos($pt.X + 4, $pt.Y + ($ch - 4)) | Out-Null
Start-Sleep -Milliseconds 700
$bmp = Grab "ucslider_client"
Stop-App

# 设计期几何（twips / 15 = DIP），再按客户区像素/816 缩放
function RectPx([double]$l, [double]$t, [double]$w, [double]$h) {
    return @([int][Math]::Round($l * $script:scale), [int][Math]::Round($t * $script:scale),
             [int][Math]::Round(($l + $w) * $script:scale), [int][Math]::Round(($t + $h) * $script:scale))
}
$pic1    = RectPx 600 56 177 161                             # Picture1 (9000,840,2655,2415 twips)
$slider4 = RectPx 608 304 177 73                             # ucSlider1(4) (9120,4560,2655,1095 twips)

function Stats([int[]]$rc) {
    $x0, $y0, $x1, $y1 = $rc
    $x1 = [Math]::Min($x1, $script:cw); $y1 = [Math]::Min($y1, $script:ch)
    $tot = 0; $bg = 0; $blue = 0; $black = 0; $green = 0; $greenTop = 99999
    for ($y = $y0; $y -lt $y1; $y++) {
        for ($x = $x0; $x -lt $x1; $x++) {
            $c = $script:bmp.GetPixel($x, $y); $tot++
            $r = $c.R; $gg = $c.G; $b = $c.B
            if ([Math]::Abs($r - 57) -le 5 -and [Math]::Abs($gg - 57) -le 5 -and [Math]::Abs($b - 57) -le 5) { $bg++ }
            if ($b -lt 140 -and $r -lt 90 -and $b -ge ($r + 6) -and $b -ge ($gg + 3)) { $blue++ }
            if ($r -eq 0 -and $gg -eq 0 -and $b -eq 0) { $black++ }
            # 轨道正向前景色: ucSlider4 = 35897 = 0x8C39 -> RGB(57,140,0)
            if ($gg -gt 90 -and $r -lt 110 -and $b -lt 80 -and $gg -gt ($r + 40)) {
                $green++; if ($y -lt $greenTop) { $greenTop = $y }
            }
        }
    }
    return @{ Tot = $tot; Bg = $bg; Blue = $blue; Black = $black; Green = $green; GreenTop = $greenTop }
}

$script:bmp = $bmp; $script:cw = $cw; $script:ch = $ch; $script:scale = $scale
$s1 = Stats $pic1
$s4 = Stats $slider4
$sPre = $null; $sMid = $null; $sDrag = $null
$script:bmp = $preBmp; $sPre = Stats $s4v
$script:bmp = $midBmp; $sMid = Stats $s4v
$script:bmp = $bmp;    $sDrag = Stats $s4v
$frac = if ($s1.Tot -gt 0) { [Math]::Round($s1.Bg / $s1.Tot, 3) } else { 0 }

$fail = @()
if (-not ($frac -ge 0.80))  { $fail += ("A1 bg393939=$frac (<0.80, ref 0.881)") }
if (-not ($s1.Blue -ge 400)) { $fail += ("A2 knobPngBlue=" + $s1.Blue + " (<400, ref 1135)") }
if (-not ($s4.Black -eq 0))  { $fail += ("B slider4PureBlack=" + $s4.Black + " (must be 0)") }
if (-not ($sPre.Green -eq 0)) { $fail += ("D0 predragGreen=" + $sPre.Green + " (must be 0)") }
if (-not ($sMid.Green -ge 400)) { $fail += ("D1 midDragGreen=" + $sMid.Green + " (<400, ref 1781)") }
if (-not ($sDrag.Green -ge 400)) { $fail += ("D2 afterDragGreen=" + $sDrag.Green + " (<400, ref 1781)") }
if (-not ($sDrag.GreenTop -le 300)) { $fail += ("D3 afterDragGreenTop=" + $sDrag.GreenTop + " (>300, ref 256)") }

$line = ("ucslider cw={0}x{1} scale={2} A1.bg={3} A2.knobPngBlue={4} B.slider4PureBlack={5} D.dragGreen={6}->{7} top={8}" -f $cw, $ch, $scale, $frac, $s1.Blue, $s4.Black, $sMid.Green, $sDrag.Green, $sDrag.GreenTop)
if ($Baseline) { Write-Output ("BASELINE " + $line); exit 0 }
if ($fail.Count -gt 0) { Write-Output ($line + " FAIL: " + ($fail -join "; ")); exit 1 }
Write-Output ($line + " PASS")
exit 0
