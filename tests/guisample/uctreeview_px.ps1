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

$bad = @()
if (($white / $area) -lt 0.15) { $bad += "白底占比 $([Math]::Round(100*$white/$area,1))% < 15%" }
if (($bg240 / $area) -gt 0.40) { $bad += "宿主灰底占比 $([Math]::Round(100*$bg240/$area,1))% > 40%" }
if ($darkRows -lt 10) { $bad += "节点文本行数 $darkRows < 10" }
if ($hl -lt 50) { $bad += "选中高亮采样 $hl < 50" }
if ($vivid -lt 300) { $bad += "图标彩色像素 $vivid < 300" }

$readout = "tree=${tw}x${th} white=$white bg240=$bg240 dark=$dark darkRows=$darkRows hl=$hl vivid=$vivid"
if ($bad.Count -gt 0) {
    Write-Output ("FAIL " + ($bad -join "; ") + " | " + $readout + " | png=$pngPath")
    try { $p.Kill() } catch {}
    exit 1
}
Write-Output ("PASS " + $readout + " | png=$pngPath")
try { $p.Kill() } catch {}
exit 0
