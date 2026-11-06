param(
    [Parameter(Mandatory=$true)][string]$Exe,
    [string]$OutDir = "",
    [int]$Sec = 10
)
# uctreeview_px.ps1 — ucTreeView 树区**上屏**判据 (本机验收夹具; -Category guisample 分片调用).
#
# 抓的是**真桌面像素** (CopyFromScreen), 判 5 条:
#   1) 树区底色 = 纯白 255,255,255 —— 占比 >= 15%  (改前: 0)
#   2) 宿主灰底 240,240,240 占比 <= 40%            (改前: ~99%, 整块灰)
#   3) 节点文本行: 深色(lum<110) 像素行数 >= 10     (改前: 0)
#   4) 选中项高亮 (255,180,180) 采样点 >= 50        (改前: 0)
#   5) 图标彩色像素 (max-min>=60) >= 300            (改前: 0)
#
# 改前实测 (2026-11-06, hdc=0 那条): white=0 bg240=77703 dark=0 darkRows=0 hl=0 vivid=0
# 改后实测 (设计期 AutoRedraw + Collection ByRef 两刀): white=48052 bg240=5446
#            dark=15607 darkRows=445 hl=2521 vivid=3559
#
# 口径: 只有**缺产物 / 位宽不符 / 缺 img 资产 / 桌面捕获为空图**才 SKIP;
# "窗口起了但树区没画" 一律 FAIL (那就是本判据要红的东西).

Add-Type -AssemblyName System.Drawing
Add-Type @"
using System; using System.Text; using System.Runtime.InteropServices;
public class UP {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder sb, int n);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool GetClientRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
  [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr h);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, int flags);
  [DllImport("user32.dll")] public static extern bool RedrawWindow(IntPtr h, IntPtr r, IntPtr rgn, int flags);
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  [DllImport("user32.dll")] public static extern uint GetDpiForWindow(IntPtr h);
  public struct RECT { public int L, T, R, B; }
  public static string Cls(IntPtr h) { var sb = new StringBuilder(256); GetClassName(h, sb, 256); return sb.ToString(); }
  // 枚举放 C# 侧: PS 5.1 的 scriptblock 委托回调查不到调用方 pid (会误报"主窗未找到")
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
  // 树不是 SysTreeView32, 是 C3 的 UserControl 宿主窗 (自绘树)
  public static IntPtr FindUcHost(IntPtr main) {
    IntPtr best = IntPtr.Zero; long bestA = 0;
    EnumChildWindows(main, (h, l) => {
      if (IsWindowVisible(h) && Cls(h) == "VB6_UserControlHost") {
        RECT r; GetWindowRect(h, out r);
        long a = (long)(r.R - r.L) * (r.B - r.T);
        if (a > bestA) { bestA = a; best = h; }
      }
      return true;
    }, IntPtr.Zero);
    return best;
  }
  // 判据 #1/#2 的取样点来自真实子窗口坐标 (不用硬编码像素位置), 这样 96dpi/125dpi
  // 两种 CI 桌面上都能取到控件本体。
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder sb, int n);
  public static IntPtr FindChildByText(IntPtr parent, string text) {
    IntPtr found = IntPtr.Zero;
    EnumChildWindows(parent, (h, l) => {
      var sb = new StringBuilder(256); GetWindowText(h, sb, 256);
      if (sb.ToString() == text) { found = h; return false; }
      return true;
    }, IntPtr.Zero);
    return found;
  }
  public static IntPtr[] FindByClass(IntPtr parent, string cls) {
    var list = new System.Collections.Generic.List<IntPtr>();
    EnumChildWindows(parent, (h, l) => {
      if (!IsWindowVisible(h)) return true;
      if (Cls(h) != cls) return true;
      list.Add(h);
      return true;
    }, IntPtr.Zero);
    return list.ToArray();
  }
  public static IntPtr[] FindSwatches(IntPtr parent) {
    var list = new System.Collections.Generic.List<IntPtr>();
    EnumChildWindows(parent, (h, l) => {
      if (!IsWindowVisible(h)) return true;
      if (Cls(h) != "Static") return true;
      RECT r; GetWindowRect(h, out r);
      int w = r.R - r.L, hgt = r.B - r.T;
      if (hgt >= 15 && hgt <= 45 && w == hgt) list.Add(h);
      return true;
    }, IntPtr.Zero);
    var arr = list.ToArray();
    System.Array.Sort(arr, (a, b) => {
      RECT ra, rb; GetWindowRect(a, out ra); GetWindowRect(b, out rb);
      return ra.T.CompareTo(rb.T);
    });
    return arr;
  }
}
"@
[UP]::SetProcessDPIAware() | Out-Null

$exePath = (Resolve-Path $Exe -ErrorAction SilentlyContinue)
if (-not $exePath) { Write-Output "SKIP 缺产物 ($Exe 不存在)"; exit 2 }
$exePath = $exePath.Path
$wd = Split-Path $exePath

# 位宽: 32 位宿主装不下 64 位 exe (反之 WoW64 可以) —— 只有这一种按位宽 SKIP
$peBytes = [IO.File]::ReadAllBytes($exePath)
$peOff = [BitConverter]::ToInt32($peBytes, 0x3C)
$machine = [BitConverter]::ToUInt16($peBytes, $peOff + 4)
if ($machine -eq 0x8664 -and [IntPtr]::Size -eq 4) { Write-Output "SKIP 位宽不符 (64 位产物, 32 位宿主)"; exit 2 }

# 缺资产: App.Path\img 不在 exe 目录时, 图标必然取不到 —— 那是环境缺件
if (-not (Test-Path (Join-Path $wd "img"))) { Write-Output "SKIP 缺资产 ($wd 下没有 img 目录)"; exit 2 }

if (-not $OutDir) { $OutDir = Join-Path $wd "uctreeview_px" }
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$pngPath = Join-Path $OutDir "uctreeview_c3.png"

function Get-LaunchDiag {
    param($Proc, [string]$OutLog, [string]$ErrLog)
    $state = if ($Proc -and -not $Proc.HasExited) { 'alive' } else { 'exited' }
    $code = '?'
    if ($state -eq 'exited') {
        try { $Proc.WaitForExit(1000) | Out-Null } catch {}
        try { $code = ('0x{0:X8}' -f [uint32]$Proc.ExitCode) } catch { try { $code = [string]$Proc.ExitCode } catch {} }
    }
    $o = @(); $e = @()
    try { $o = @(Get-Content -LiteralPath $OutLog -ErrorAction SilentlyContinue | Where-Object { $_ }) } catch {}
    try { $e = @(Get-Content -LiteralPath $ErrLog -ErrorAction SilentlyContinue | Where-Object { $_ }) } catch {}
    $crash = @($e | Where-Object { $_ -match '\[C3_CRASH\]' } | Select-Object -First 3) -join ' | '
    $ol = if ($o.Count -gt 0) { @($o | Select-Object -Last 3) -join ' | ' } else { '(empty)' }
    $el = if ($e.Count -gt 0) { @($e | Select-Object -Last 3) -join ' | ' } else { '(empty)' }
    $s = "state=$state exitCode=$code outLines=$($o.Count) outTail='$ol' errLines=$($e.Count) errTail='$el'"
    if ($crash) { $s += " crash='$crash'" }
    return $s
}

# 必须走 CreateProcess: 带 -RedirectStandard* 时 Start-Process 用 CreateProcess, 不经
# ShellExecute, 不会撞仓库完整性标签引起的"打开文件-安全警告" (也不能同时给 -NoNewWindow)
$logOut = Join-Path $OutDir "sample.out"
$logErr = Join-Path $OutDir "sample.err"
$env:C3_CRASH_TRACE = '1'
$p = Start-Process -FilePath $exePath -WorkingDirectory $wd -PassThru `
     -RedirectStandardOutput $logOut -RedirectStandardError $logErr
try { $p.WaitForInputIdle(15000) | Out-Null } catch {}
Start-Sleep -Seconds $Sec

$main = [UP]::FindMain([uint32]$p.Id)
if ($main -eq [IntPtr]::Zero) {
    Write-Output ("FAIL 主窗未起 (pid=" + $p.Id + " " + (Get-LaunchDiag $p $logOut $logErr) + ")")
    try { $p.Kill() } catch {}
    exit 1
}

if ([UP]::IsIconic($main)) { [UP]::ShowWindow($main, 9) | Out-Null; Start-Sleep -Milliseconds 600 }
[UP]::ShowWindow($main, 5) | Out-Null
[UP]::SetForegroundWindow($main) | Out-Null
[UP]::SetWindowPos($main, [IntPtr](-1), 0, 0, 0, 0, 0x0003) | Out-Null   # TOPMOST|NOMOVE|NOSIZE
Start-Sleep -Milliseconds 900
[UP]::RedrawWindow($main, [IntPtr]::Zero, [IntPtr]::Zero, 0x0185) | Out-Null
Start-Sleep -Milliseconds 600

$uc = [UP]::FindUcHost($main)
if ($uc -eq [IntPtr]::Zero) {
    Write-Output ("FAIL 主窗可见但无可见 VB6_UserControlHost (pid=" + $p.Id + " " + (Get-LaunchDiag $p $logOut $logErr) + ")")
    try { $p.Kill() } catch {}
    exit 1
}

$wr = New-Object UP+RECT; [UP]::GetWindowRect($main, [ref]$wr) | Out-Null
$ur = New-Object UP+RECT; [UP]::GetWindowRect($uc, [ref]$ur) | Out-Null
$ww = $wr.R - $wr.L; $wh = $wr.B - $wr.T
if ($ww -le 0 -or $wh -le 0) {
    Write-Output "FAIL 主窗矩形非法 ${ww}x${wh}"
    try { $p.Kill() } catch {}
    exit 1
}

$bmp = New-Object System.Drawing.Bitmap($ww, $wh)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen($wr.L, $wr.T, 0, 0, (New-Object System.Drawing.Size($ww, $wh)))
$g.Dispose()
$bmp.Save($pngPath, [System.Drawing.Imaging.ImageFormat]::Png)

# 空图判据 (针对**整幅窗体**, 不是树区): 整幅只有一种颜色 ⇒ 桌面不可捕获
$first = $bmp.GetPixel(0, 0)
$blank = $true
for ($py = 0; $py -lt $wh -and $blank; $py += 4) {
    for ($px = 0; $px -lt $ww -and $blank; $px += 4) {
        $c = $bmp.GetPixel($px, $py)
        if ($c.R -ne $first.R -or $c.G -ne $first.G -or $c.B -ne $first.B) { $blank = $false }
    }
}
if ($blank) {
    Write-Output "SKIP 桌面捕获为空图 (单色 $($first.R),$($first.G),$($first.B) 于 ${ww}x${wh})"
    $bmp.Dispose(); try { $p.Kill() } catch {}
    exit 2
}

# --- 设计期背景色判据 (残差 #1/#2) ---
# 口径: 取样点来自运行时**真实子窗口坐标** (EnumChildWindows), 不写死像素位置,
#   96dpi / 125dpi 桌面都成立。色块 = 9 枚近正方形 STATIC (25x25 逻辑), 按屏幕 Y 排序;
#   期望色 = Form1.frm 设计期 BackColor 的 BBGGRR 换算 (&H80000005/0A 走 GetSysColor)。
$swatchExpect = @(
    @(255,255,255), @(180,180,180), @(99,99,255), @(0,0,0), @(99,99,255),
    @(204,204,204), @(0,0,0), @(204,204,204), @(204,232,255)
)
$swatchBad = @()
$swatchRead = @()
$swHwnds = [UP]::FindSwatches($main)
if ($swHwnds.Length -ne 9) { $swatchBad += "色块控件数 $($swHwnds.Length) != 9" }
for ($si = 0; $si -lt [Math]::Min(9, $swHwnds.Length); $si++) {
    $r = New-Object UP+RECT; [UP]::GetWindowRect($swHwnds[$si], [ref]$r) | Out-Null
    $px = [int](($r.L + $r.R) / 2 - $wr.L); $py = [int](($r.T + $r.B) / 2 - $wr.T)
    $c = $bmp.GetPixel($px, $py)
    $swatchRead += "$($c.R),$($c.G),$($c.B)"
    $e = $swatchExpect[$si]
    if ([Math]::Abs($c.R - $e[0]) -gt 3 -or [Math]::Abs($c.G - $e[1]) -gt 3 -or [Math]::Abs($c.B - $e[2]) -gt 3) {
        $swatchBad += ("swatch[$si]=$($c.R),$($c.G),$($c.B) != $($e[0]),$($e[1]),$($e[2])")
    }
}
# 判据一: 8 枚 CheckBox 的控件矩形。VB6 里 CheckBox 的默认 BackColor =
#   &H8000000F& (按钮面 240) 且不透明; 修复前 C3 走 DefWindowProcW 的类刷,
#   整块 255 白 (旧产物实测: 240 占比 0%, 纯白 84.7%)。
$chkCaptions = @(
    "AlternateNodeColors", "ShowTreeLines", "ShowPlusMinus", "CheckBoxes",
    "FullRowSelect", "PlusMinusStyle = 1 - TVP_PlusMinus", "PlusMinusAlignRight", "AllowEdit"
)
# 判据二: Label 控件矩形 (真 Label, 不是 CheckBox): 左列 6 条 + 右列 1 条。
$labelCaptions = @("BorderRadius", "BorderWidth", "Margin", "NodeSpacing", "ToggleIconSize", "MarginBottom", "SelectionColor")
function Measure-Hwnds([IntPtr[]]$hwnds, [System.Drawing.Bitmap]$bmp) {
    $a = 0; $c240 = 0; $c255 = 0
    foreach ($h in $hwnds) {
        if ($h -eq [IntPtr]::Zero) { continue }
        $r = New-Object UP+RECT; [UP]::GetWindowRect($h, [ref]$r) | Out-Null
        $x0 = [Math]::Max(0, $r.L - $wr.L + 1); $x1 = [Math]::Min($ww, $r.R - $wr.L - 1)
        $y0 = [Math]::Max(0, $r.T - $wr.T + 1); $y1 = [Math]::Min($wh, $r.B - $wr.T - 1)
        for ($yy = $y0; $yy -lt $y1; $yy++) {
            for ($xx = $x0; $xx -lt $x1; $xx++) {
                $c = $bmp.GetPixel($xx, $yy)
                $a++
                if ([Math]::Abs($c.R-240) -le 1 -and [Math]::Abs($c.G-240) -le 1 -and [Math]::Abs($c.B-240) -le 1) { $c240++ }
                if ($c.R -eq 255 -and $c.G -eq 255 -and $c.B -eq 255) { $c255++ }
            }
        }
    }
    if ($a -le 0) { return @(0, 1) }
    return @(($c240 / $a), ($c255 / $a))
}
$chkHwnds = @(); foreach ($t in $chkCaptions) { $chkHwnds += [UP]::FindChildByText($main, $t) }
$labHwnds = @(); foreach ($t in $labelCaptions) { $labHwnds += [UP]::FindChildByText($main, $t) }
$chkMissing = @($chkHwnds | Where-Object { $_ -eq [IntPtr]::Zero }).Count
$labMissing = @($labHwnds | Where-Object { $_ -eq [IntPtr]::Zero }).Count
$chk240, $chkWhite = Measure-Hwnds $chkHwnds $bmp
$labelRatio, $labelWhiteRatio = Measure-Hwnds $labHwnds $bmp

# --- 残差 #5/#6: TextBox 深框 + VScrollBar 自绘 (必须在 $bmp.Dispose() 之前量) ---
# 参考值来自真 VB6 运行产物 (同一 Form1.frm, 120dpi; 见 ai/C3_UCTREEVIEW_HANDOFF.md §22):
#   TextBox(Edit) 91x33: 周圈 (100,100,100) 深框像素 = 0 (旧产物 244/枚 = 周长),
#     顶行也没有深色像素 (旧产物 91/枚)。VB6 的 TextBox 自己画 1px 淡边 (236)。
#   VScrollBar 33x31: 240,240,240 = 960 + 255,255,255 = 63 (6 枚全等); 左边 1 列白 +
#     高/2 处 1 行白; **没有箭头/3D 边** —— VB6 这个尺寸下就不画箭头, 所以这里断言的是
#     "240 底 + 左列白 + 中线白" 三条实测特征, 不是箭头。(改前 C3: 1023/1023 纯白)
$editHwnds = [UP]::FindByClass($main, "Edit")
$sbHwnds = [UP]::FindByClass($main, "ScrollBar")
$edFrameDark = 0; $edTopDark = 0; $edEdge236 = 0; $edEdgeOk = 0
foreach ($h in $editHwnds) {
    $r = New-Object UP+RECT; [UP]::GetWindowRect($h, [ref]$r) | Out-Null
    $ex0 = [Math]::Max(0, $r.L - $wr.L); $ex1 = [Math]::Min($ww, $r.R - $wr.L)
    $ey0 = [Math]::Max(0, $r.T - $wr.T); $ey1 = [Math]::Min($wh, $r.B - $wr.T)
    $this236 = 0
    for ($yy = $ey0; $yy -lt $ey1; $yy++) {
        for ($xx = $ex0; $xx -lt $ex1; $xx++) {
            $c = $bmp.GetPixel($xx, $yy)
            if ([Math]::Abs($c.R-100) -le 2 -and [Math]::Abs($c.G-100) -le 2 -and [Math]::Abs($c.B-100) -le 2) { $edFrameDark++ }
            # VB6 ThunderRT6TextBox (ex=0x204=WS_EX_CLIENTEDGE) 的 3D 内凹: 上/左 1px 236 + 内 1px 254,
            # 下/右 1px 131; 客户区比控件每边小 2px. 这条 236 亮边是"深框已去掉"的正面证据。
            if ([Math]::Abs($c.R-236) -le 2 -and [Math]::Abs($c.G-236) -le 2 -and [Math]::Abs($c.B-236) -le 2) { $edEdge236++; $this236++ }
            if ($yy -eq $ey0) {
                $lum = 0.299*$c.R + 0.587*$c.G + 0.114*$c.B
                if ($lum -lt 200) { $edTopDark++ }
            }
        }
    }
    if ($this236 -ge 60) { $edEdgeOk++ }   # 单枚 91x33 实测 ~244px (周长), 阈值取 1/4
}
$sbArea = 0; $sb240 = 0; $sbWhite = 0; $sbCount = 0; $sbMidOk = 0; $sbLeftOk = 0
foreach ($h in $sbHwnds) {
    $r = New-Object UP+RECT; [UP]::GetWindowRect($h, [ref]$r) | Out-Null
    $sx0 = [Math]::Max(0, $r.L - $wr.L); $sx1 = [Math]::Min($ww, $r.R - $wr.L)
    $sy0 = [Math]::Max(0, $r.T - $wr.T); $sy1 = [Math]::Min($wh, $r.B - $wr.T)
    $sw = $sx1 - $sx0; $sh = $sy1 - $sy0
    if ($sw -le 0 -or $sh -le 0) { continue }
    $sbCount++
    # 不写死"第 15 行": GetWindowRect 与桌面像素偶尔差 1px, 所以判"存在一整行白 +
    # 存在一整列白" (VB6 实测就是 1 行 + 1 列, 共 63px)。
    $rowWhite = New-Object int[] $sh
    $colWhite = New-Object int[] $sw
    for ($yy = $sy0; $yy -lt $sy1; $yy++) {
        for ($xx = $sx0; $xx -lt $sx1; $xx++) {
            $c = $bmp.GetPixel($xx, $yy); $sbArea++
            if ([Math]::Abs($c.R-240) -le 1 -and [Math]::Abs($c.G-240) -le 1 -and [Math]::Abs($c.B-240) -le 1) { $sb240++ }
            if ($c.R -eq 255 -and $c.G -eq 255 -and $c.B -eq 255) {
                $sbWhite++
                $rowWhite[$yy - $sy0]++
                $colWhite[$xx - $sx0]++
            }
        }
    }
    $hasRow = $false; foreach ($n in $rowWhite) { if ($n -ge [Math]::Ceiling($sw * 0.9)) { $hasRow = $true } }
    $hasCol = $false; foreach ($n in $colWhite) { if ($n -ge [Math]::Ceiling($sh * 0.9)) { $hasCol = $true } }
    if ($hasRow) { $sbMidOk++ }
    if ($hasCol) { $sbLeftOk++ }
}
$sbR240 = 0; $sbRWhite = 0
if ($sbArea -gt 0) { $sbR240 = $sb240 / $sbArea; $sbRWhite = $sbWhite / $sbArea }

# --- 树区 (UC 宿主客户区, 坐标相对整幅窗体) ---
$x0 = [Math]::Max(0, $ur.L - $wr.L); $y0 = [Math]::Max(0, $ur.T - $wr.T)
$tw = [Math]::Min($ur.R - $ur.L, $ww - $x0); $th = [Math]::Min($ur.B - $ur.T, $wh - $y0)
if ($tw -le 0 -or $th -le 0) {
    Write-Output "FAIL 树区矩形非法 rel=($x0,$y0) ${tw}x${th}"
    $bmp.Dispose(); try { $p.Kill() } catch {}
    exit 1
}

$rect = New-Object System.Drawing.Rectangle $x0, $y0, $tw, $th
$bd = $bmp.LockBits($rect, [System.Drawing.Imaging.ImageLockMode]::ReadOnly, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
$stride = $bd.Stride; $ptr = $bd.Scan0
$len = [Math]::Abs($stride) * $th
$buf = New-Object byte[] $len
[System.Runtime.InteropServices.Marshal]::Copy($ptr, $buf, 0, $len)
$bmp.UnlockBits($bd)
$bmp.Dispose()

$area = $tw * $th
$white = 0; $bg240 = 0; $dark = 0; $hl = 0; $vivid = 0
$rowsArr = New-Object int[] $th
for ($y = 0; $y -lt $th; $y++) {
    $off = $y * [Math]::Abs($stride)
    for ($x = 0; $x -lt $tw; $x++) {
        $i = $off + $x * 4
        $b = $buf[$i]; $gg = $buf[$i+1]; $r = $buf[$i+2]
        if ($r -eq 255 -and $gg -eq 255 -and $b -eq 255) { $white++ }
        if ([Math]::Abs($r-240) -le 1 -and [Math]::Abs($gg-240) -le 1 -and [Math]::Abs($b-240) -le 1) { $bg240++ }
        $lum = 0.299*$r + 0.587*$gg + 0.114*$b
        if ($lum -lt 110) { $dark++; $rowsArr[$y]++ }
        if ([Math]::Abs($r-255) -le 8 -and [Math]::Abs($gg-180) -le 8 -and [Math]::Abs($b-180) -le 8) { $hl++ }
        $mx = [Math]::Max($r, [Math]::Max($gg, $b)); $mn = [Math]::Min($r, [Math]::Min($gg, $b))
        if (($mx - $mn) -ge 60) { $vivid++ }
    }
}
$darkRows = 0
for ($y = 0; $y -lt $th; $y++) { if ($rowsArr[$y] -ge 3) { $darkRows++ } }

# --- 残差 #3/#4: 选中行文字色/白底 + 高亮高度/行距 (对照 VB6 参考) ---
# 参考 (VB6 真机 125% DPI, 同一 Form1.frm):
#   高亮 205x41, 行距 47/44; 选中行文字=白 (纯黑 ~10 / dark ~21 只在图标边缘);
#   粉底上的纯白只在文字笔画上 (~258px), SubText 204 灰; 无 Opaque 文字底块。
# 改前 C3 (2026-11-06, 选中粉底 206x35): inner white=968..1410 (白底块),
#   pureblack=224 / dark=240 (黑字), 行距 38/38.
# 改后 (字体块进 bag + applyFontNow/SetForeColor 落 AutoRedraw DC):
#   inner white=355 (只剩白字笔画+复选框), pureblack=12 / dark=28, 高亮 206x39, 行距 42.
# 高亮 bbox 用**精确**粉 (255,180,180 = SelectionColor #FFB4B4): 带容差会把圆角
# 阴影 (255,217,217) 也算进来, bbox 下沿会压到下一行文字 ⇒ selDark 假高。
$hlMinX = 99999; $hlMinY = 99999; $hlMaxX = -1; $hlMaxY = -1
for ($y = 0; $y -lt $th; $y++) {
    $off = $y * [Math]::Abs($stride)
    for ($x = 0; $x -lt $tw; $x++) {
        $i = $off + $x * 4
        $b = $buf[$i]; $gg = $buf[$i+1]; $r = $buf[$i+2]
        if ($r -eq 255 -and $gg -eq 180 -and $b -eq 180) {
            if ($x -lt $hlMinX) { $hlMinX = $x }
            if ($x -gt $hlMaxX) { $hlMaxX = $x }
            if ($y -lt $hlMinY) { $hlMinY = $y }
            if ($y -gt $hlMaxY) { $hlMaxY = $y }
        }
    }
}
$selWhite = 0; $selDark = 0; $selBlack = 0
$hlH = 0
if ($hlMaxX -gt $hlMinX -and $hlMaxY -gt $hlMinY) {
    $hlH = $hlMaxY - $hlMinY + 1
    for ($y = $hlMinY + 2; $y -le $hlMaxY - 2; $y++) {
        $off = $y * [Math]::Abs($stride)
        for ($x = $hlMinX + 2; $x -le $hlMaxX - 2; $x++) {
            $i = $off + $x * 4
            $b = $buf[$i]; $gg = $buf[$i+1]; $r = $buf[$i+2]
            if ($r -eq 255 -and $gg -eq 255 -and $b -eq 255) { $selWhite++ }
            if ((0.299*$r + 0.587*$gg + 0.114*$b) -lt 110) { $selDark++ }
            if (($r + $gg + $b) -lt 20) { $selBlack++ }
        }
    }
}
# 行距 = 高亮以下两条 "墨迹带" 起点的差。带 = 该行 (r+g+b)<450 的像素 >0
# (文字/图标都算); 这个阈值跟我方 C# 分析脚本同一口径, 图标行不会把间距切碎。
$rowsInk = New-Object int[] $th
for ($y = 0; $y -lt $th; $y++) {
    $off = $y * [Math]::Abs($stride)
    for ($x = 0; $x -lt $tw; $x++) {
        $i = $off + $x * 4
        $b = $buf[$i]; $gg = $buf[$i+1]; $r = $buf[$i+2]
        if (($r + $gg + $b) -lt 450) { $rowsInk[$y]++ }
    }
}
$bandStarts = @()
for ($y = 0; $y -lt $th; $y++) {
    if ($rowsInk[$y] -gt 0 -and ($y -eq 0 -or $rowsInk[$y-1] -eq 0)) { $bandStarts += $y }
}
$pitch = 0
$outside = @($bandStarts | Where-Object { $_ -gt $hlMaxY + 2 })
if ($outside.Count -ge 2) { $pitch = $outside[1] - $outside[0] }
$expDpi = [UP]::GetDpiForWindow($main)
if ($expDpi -lt 96) { $expDpi = 96 }
$expHl = [int][Math]::Round(41 * $expDpi / 120)
$expPitch = [int][Math]::Round(44 * $expDpi / 120)

$bad = @()
if (($white / $area) -lt 0.15) { $bad += "白底占比 $([Math]::Round(100*$white/$area,1))% < 15%" }
if (($bg240 / $area) -gt 0.40) { $bad += "宿主灰底占比 $([Math]::Round(100*$bg240/$area,1))% > 40%" }
if ($darkRows -lt 10) { $bad += "节点文本行数 $darkRows < 10" }
if ($hl -lt 50) { $bad += "选中高亮采样 $hl < 50" }
if ($vivid -lt 300) { $bad += "图标彩色像素 $vivid < 300" }
if ($swatchBad.Count -gt 0) { $bad += ("色块中心色不符 (" + ($swatchBad -join ", ") + ")") }
if ($chkMissing -gt 0) { $bad += "CheckBox 定位失败 $chkMissing/8" }
if ($labMissing -gt 0) { $bad += "Label 定位失败 $labMissing/7" }
if ($chk240 -lt 0.60) { $bad += "CheckBox 区 240 占比 $([Math]::Round(100*$chk240,1))% < 60%" }
if ($chkWhite -gt 0.05) { $bad += "CheckBox 区纯白占比 $([Math]::Round(100*$chkWhite,1))% > 5%" }
if ($labelRatio -lt 0.80) { $bad += "Label 区 240 占比 $([Math]::Round(100*$labelRatio,1))% < 80%" }
if ($labelWhiteRatio -gt 0.02) { $bad += "Label 区纯白占比 $([Math]::Round(100*$labelWhiteRatio,1))% > 2%" }
# 残差 #5: TextBox 周圈不能有 (100,100,100) 深框 (VB6 自己画 1px 淡边 236)
if ($editHwnds.Length -lt 6) { $bad += "TextBox 定位失败 $($editHwnds.Length)/6" }
if ($edFrameDark -gt 0) { $bad += "TextBox 深灰边框(100,100,100) ${edFrameDark}px > 0" }
if ($edTopDark -gt 0) { $bad += "TextBox 顶行深色 ${edTopDark}px > 0" }
# 另一面: 深框去掉之后必须**还留着** WS_EX_CLIENTEDGE 的 3D 内凹 (上/左 236 亮边),
# 否则"没深框"是靠"连边一起没有"换来的 —— 那是把 VB6 的 2px 内凹一起丢了。
if ($edEdgeOk -lt $editHwnds.Length) { $bad += "TextBox 236 亮边缺失 ($edEdgeOk/$($editHwnds.Length) 枚, 共 ${edEdge236}px)" }
# 残差 #6: VScrollBar 必须是 VB6 的 240 底 + 左列白 + 中线白 (不是 1023px 纯白)
if ($sbHwnds.Length -lt 6) { $bad += "VScrollBar 定位失败 $($sbHwnds.Length)/6" }
if ($sbR240 -lt 0.85) { $bad += "VScrollBar 240 占比 $([Math]::Round(100*$sbR240,1))% < 85%" }
if ($sbRWhite -lt 0.03 -or $sbRWhite -gt 0.10) { $bad += "VScrollBar 纯白占比 $([Math]::Round(100*$sbRWhite,1))% 不在 3%~10%" }
if ($sbMidOk -lt $sbCount) { $bad += "VScrollBar 中线白线缺失 ($sbMidOk/$sbCount)" }
if ($sbLeftOk -lt $sbCount) { $bad += "VScrollBar 左边白线缺失 ($sbLeftOk/$sbCount)" }
# 残差 #3: 选中行文字色/白底 —— VB6 是白字透明底, C3 改前是黑字 + Opaque 白底块
if ($hlH -le 0) { $bad += "选中高亮未找到" }
else {
    if ($selWhite -gt 600) { $bad += "粉底上纯白像素 $selWhite > 600 (Opaque 文字底块)" }
    # 文字色判据用**纯黑** (r+g+b<20): 选中行里的图标本身也是深色 (Facebook 蓝 lum~101),
    # 用 lum 阈值会把图标算成"黑字"。
    if ($selBlack -gt 30) { $bad += "选中行纯黑像素 $selBlack > 30 (应为白字)" }
    if ($selWhite -lt 100) { $bad += "选中行白字笔画 $selWhite < 100 (文字色不是白)" }
    # 残差 #4: 高亮 bbox 高度 / 行距 与 VB6 参考 ±2px (按 DPI 折算)
    if ([Math]::Abs($hlH - $expHl) -gt 2) { $bad += "高亮高度 ${hlH}px 偏离参考 ${expHl}px 超过 ±2" }
    if ($pitch -gt 0 -and [Math]::Abs($pitch - $expPitch) -gt 2) {
        $bad += "行距 ${pitch}px 偏离参考 ${expPitch}px 超过 ±2"
    }
}

$readout = "tree=${tw}x${th} white=$white bg240=$bg240 dark=$dark darkRows=$darkRows hl=$hl vivid=$vivid"
$readout += " selWhite=$selWhite selDark=$selDark selBlack=$selBlack hlH=${hlH}@${expHl} pitch=${pitch}@${expPitch} dpi=$expDpi"
$readout += " swatches=[" + ($swatchRead -join " | ") + "] chk240=$([Math]::Round(100*$chk240,1))% chkWhite=$([Math]::Round(100*$chkWhite,1))% label240=$([Math]::Round(100*$labelRatio,1))% labelWhite=$([Math]::Round(100*$labelWhiteRatio,1))%"
$readout += " edFrameDark=$edFrameDark edTopDark=$edTopDark edEdge236=$edEdge236/$($editHwnds.Length) sb240=$([Math]::Round(100*$sbR240,1))% sbWhite=$([Math]::Round(100*$sbRWhite,1))% sbMid=$sbMidOk/$sbCount sbLeft=$sbLeftOk/$sbCount"
if ($bad.Count -gt 0) {
    Write-Output ("FAIL " + ($bad -join "; ") + " | " + $readout + " | png=$pngPath")
    try { $p.Kill() } catch {}
    exit 1
}
Write-Output ("PASS " + $readout + " | png=$pngPath")
try { $p.Kill() } catch {}
exit 0
