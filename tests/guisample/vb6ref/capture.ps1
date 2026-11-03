param(
    [Parameter(Mandatory=$true)][string]$ExeDir,
    [Parameter(Mandatory=$true)][string]$ExeName,
    [Parameter(Mandatory=$true)][string]$Out,
    [string]$Label = "",
    [int]$Sec = 12
)
# 统一采集脚本: VB6 侧与 C3 侧用同一份, 输出同格式文本 + 状态栏窗口 DC 位图.
# 采集方式: BitBlt(窗口 DC) —— 不用 PrintWindow (抓不到 Codejock 子窗), 不用桌面截屏.
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System; using System.Runtime.InteropServices; using System.Text;
public class CP {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern int GetWindowThreadProcessId(IntPtr h, out int pid);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder sb, int n);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder sb, int n);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool GetClientRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern IntPtr GetParent(IntPtr h);
  [DllImport("user32.dll")] public static extern bool RedrawWindow(IntPtr h, IntPtr r, IntPtr rgn, int flags);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, int flags);
  [DllImport("user32.dll")] public static extern IntPtr GetDC(IntPtr h);
  [DllImport("user32.dll")] public static extern int ReleaseDC(IntPtr h, IntPtr dc);
  [DllImport("gdi32.dll")] public static extern bool BitBlt(IntPtr d, int x, int y, int w, int h, IntPtr s, int sx, int sy, int rop);
  [DllImport("user32.dll")] public static extern IntPtr WindowFromPoint(POINT p);
  public struct POINT { public int X, Y; }
  [DllImport("kernel32.dll")] public static extern IntPtr OpenProcess(int access, bool inherit, int pid);
  [DllImport("kernel32.dll")] public static extern IntPtr VirtualAllocEx(IntPtr proc, IntPtr addr, int size, int type, int protect);
  [DllImport("kernel32.dll")] public static extern bool ReadProcessMemory(IntPtr proc, IntPtr addr, byte[] buf, int size, out int read);
  [DllImport("kernel32.dll")] public static extern bool VirtualFreeEx(IntPtr proc, IntPtr addr, int size, int type);
  [DllImport("kernel32.dll")] public static extern bool CloseHandle(IntPtr h);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern IntPtr SendMessageW(IntPtr h, int m, IntPtr w, IntPtr l);
  // 跨进程 WM_GETTEXT: 远端缓冲区 + ReadProcessMemory. GetWindowText 对他进程控件只回缓存标题(设计期 "FlatEdit1"),
  // 这才是 "WM_GETTEXT 全文" 的正确读法.
  public static string EditTextRemote(IntPtr hwnd) {
    int pid; GetWindowThreadProcessId(hwnd, out pid);
    IntPtr proc = OpenProcess(0x0038, false, pid);
    if (proc == IntPtr.Zero) return "<OpenProcess failed>";
    IntPtr lr = SendMessageW(hwnd, 0x000E, IntPtr.Zero, IntPtr.Zero); // WM_GETTEXTLENGTH
    int len = (int)lr;
    if (len <= 0) { CloseHandle(proc); return ""; }
    int size = (len + 2) * 2;
    IntPtr buf = VirtualAllocEx(proc, IntPtr.Zero, size, 0x1000, 0x04);
    if (buf == IntPtr.Zero) { CloseHandle(proc); return "<alloc failed>"; }
    SendMessageW(hwnd, 0x000D, (IntPtr)size, buf); // WM_GETTEXT
    byte[] data = new byte[size];
    int got = 0;
    ReadProcessMemory(proc, buf, data, size, out got);
    VirtualFreeEx(proc, buf, 0, 0x8000);
    CloseHandle(proc);
    var sb = new StringBuilder();
    for (int i = 0; i + 1 < got; i += 2) { char c = (char)(data[i] | (data[i+1] << 8)); if (c == '\0') break; sb.Append(c); }
    return sb.ToString();
  }
  [DllImport("user32.dll")] public static extern uint GetDpiForWindow(IntPtr h);
  public struct RECT { public int L, T, R, B; }
  public static string Cls(IntPtr h) { var sb = new StringBuilder(256); GetClassName(h, sb, 256); return sb.ToString(); }
  public static string Txt(IntPtr h) { var sb = new StringBuilder(4096); GetWindowText(h, sb, 4096); return sb.ToString(); }
  public static string Rc(IntPtr h) { RECT r; GetWindowRect(h, out r); return (r.R-r.L) + "x" + (r.B-r.T) + "@" + r.L + "," + r.T; }
  public static int W(IntPtr h) { RECT r; GetWindowRect(h, out r); return r.R-r.L; }
  public static int H(IntPtr h) { RECT r; GetWindowRect(h, out r); return r.B-r.T; }
  public static int A(IntPtr h) { RECT r; GetWindowRect(h, out r); return (r.R-r.L)*(r.B-r.T); }
  public static int Dpi(IntPtr h) { try { return (int)GetDpiForWindow(h); } catch { return 0; } }
}
"@
$lines = New-Object System.Collections.Generic.List[string]
function Cap($exeDir, $exeName, $sec, $label) {
    $exe = (Resolve-Path (Join-Path $exeDir $exeName)).Path
    $wd  = (Resolve-Path $exeDir).Path
    $errF = Join-Path $wd ("cap_" + $label + ".err")
    $outF = Join-Path $wd ("cap_" + $label + ".out")
    $p = Start-Process -FilePath $exe -WorkingDirectory $wd -RedirectStandardError $errF -RedirectStandardOutput $outF -PassThru
    Start-Sleep -Seconds $sec
    # 主窗: 该进程可见顶层窗中类名匹配窗体类 (C3 = VB6_Form_* / VB6 = Thunder*),
    # 不能只取面积最大 —— Codejock 的 XTPFrameShadow 是更大的 Afx 窗口, 会把主窗挤掉.
    $script:main = [IntPtr]::Zero; $script:bestA = 0
    $script:anyMain = [IntPtr]::Zero; $script:anyA = 0
    $cb = [CP+EnumProc]{ param($h,$l)
        $pp = 0; [CP]::GetWindowThreadProcessId($h, [ref]$pp) | Out-Null
        if ($pp -eq $p.Id -and [CP]::IsWindowVisible($h)) {
            $a = [CP]::A($h); $c = [CP]::Cls($h)
            if ($c -like 'VB6_Form_*' -or $c -like 'Thunder*') {
                if ($a -gt $script:bestA) { $script:bestA = $a; $script:main = $h }
            }
            if ($a -gt $script:anyA) { $script:anyA = $a; $script:anyMain = $h }
        }
        return $true }
    [CP]::EnumWindows($cb, [IntPtr]::Zero) | Out-Null
    if ($script:main -eq [IntPtr]::Zero) { $script:main = $script:anyMain }
    # BitBlt 读的是屏幕像素, 目标窗一旦被别的窗口盖住就会读到别人 —— 先抬到最前再采.
    if ($script:main -ne [IntPtr]::Zero) {
        [CP]::SetForegroundWindow($script:main) | Out-Null
        [CP]::SetWindowPos($script:main, [IntPtr](-1), 0, 0, 0, 0, 0x0001 -bor 0x0002 -bor 0x0040) | Out-Null
        Start-Sleep -Milliseconds 600
    }
    $script:kids = @()
    if ($script:main -ne [IntPtr]::Zero) {
        [CP]::RedrawWindow($script:main, [IntPtr]::Zero, [IntPtr]::Zero, 0x0185) | Out-Null
        Start-Sleep -Milliseconds 700
        $cb2 = [CP+EnumProc]{ param($h,$l)
            $script:kids += ,@{ c=[CP]::Cls($h); h=$h; a=[CP]::A($h) }
            return $true }
        [CP]::EnumChildWindows($script:main, $cb2, [IntPtr]::Zero) | Out-Null
    }
    $dpi = if ($script:main -ne [IntPtr]::Zero) { [CP]::Dpi($script:main) } else { 0 }
    $lines.Add("# CAPTURE label=$label side_dir=$exeDir")
    $lines.Add("DPI=$dpi MAIN=$([CP]::Rc($script:main)) mainclass=$([CP]::Cls($script:main)) pid=$($p.Id)")
    $lines.Add("KIDS=$($script:kids.Count) KIDCLASSES=" + (($script:kids | ForEach-Object { "$($_.c)($($_.a))" }) -join ','))
    # 1) 状态栏: 面积最大的 XTPStatusBar (显式取最大, 不依赖 Sort-Object 对 hashtable 的按键排序)
    function PickMax($pred) {
        $best = $null; $ba = -1
        foreach ($k in $script:kids) { if (& $pred $k) { if ($k.a -gt $ba) { $ba = $k.a; $best = $k } } }
        return $best
    }
    $sb = PickMax { param($k) $k.c -like 'XTPStatusBar*' }
    $sbH = [IntPtr]::Zero; $sbRect = "none"; $sbCli = "none"
    if ($sb) {
        $sbH = $sb.h; $sbRect = [CP]::Rc($sbH)
        $cr = New-Object CP+RECT; [CP]::GetClientRect($sbH, [ref]$cr) | Out-Null
        $sbCli = "$($cr.R)x$($cr.B)"
    }
    $lines.Add("STATUSBAR rect=$sbRect client=$sbCli")
    $lines.Add("STATUSBAR_TITLE=[$(if($sbH -ne [IntPtr]::Zero){[CP]::Txt($sbH)})]")
    if ($sbH -ne [IntPtr]::Zero) {
        $r2 = New-Object CP+RECT; [CP]::GetWindowRect($sbH, [ref]$r2) | Out-Null
        $pt = New-Object CP+POINT; $pt.X = [int](($r2.L + $r2.R) / 2); $pt.Y = [int](($r2.T + $r2.B) / 2)
        $hit = [CP]::WindowFromPoint($pt)
        $lines.Add("SB_HITTEST cls=$([CP]::Cls($hit)) rect=$([CP]::Rc($hit)) same=$($hit -eq $sbH)")
    }
    # 2) 窗口 DC 全量像素 (窗口 DC + BitBlt, 不用 PrintWindow, 不依赖桌面截屏)
    function MeasurePix($hwnd, $tag, $pngName) {
        $w = [CP]::W($hwnd); $h = [CP]::H($hwnd)
        if ($w -le 0 -or $h -le 0) {
            $lines.Add("$($tag)_COLORS=0 $($tag)_DOMINANT=none $($tag)_PIXELS=0 $($tag)_CONTRAST=0")
            $lines.Add("$($tag)_ROWS="); $lines.Add("$($tag)_TOP5="); return
        }
        $bmp = New-Object System.Drawing.Bitmap($w, $h)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $dc = $g.GetHdc(); $src = [CP]::GetDC($hwnd)
        [CP]::BitBlt($dc, 0, 0, $w, $h, $src, 0, 0, 0x00CC0020) | Out-Null
        [CP]::ReleaseDC($hwnd, $src) | Out-Null
        $g.ReleaseHdc($dc); $g.Dispose()
        $rect = New-Object System.Drawing.Rectangle(0, 0, $w, $h)
        $bd = $bmp.LockBits($rect, [System.Drawing.Imaging.ImageLockMode]::ReadOnly, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
        $buf = New-Object byte[] ($bd.Stride * $h)
        [System.Runtime.InteropServices.Marshal]::Copy($bd.Scan0, $buf, 0, $buf.Length)
        $bmp.UnlockBits($bd)
        $hist = @{}
        for ($y = 0; $y -lt $h; $y++) { $row = $y * $bd.Stride
            for ($x = 0; $x -lt $w; $x++) { $o = $row + $x*4
                $k = "$($buf[$o+3]),$($buf[$o+2]),$($buf[$o+1]),$($buf[$o])"
                if ($hist.ContainsKey($k)) { $hist[$k]++ } else { $hist[$k] = 1 } } }
        $dom = ($hist.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1)
        $dv = $dom.Key -split ','
        $dr = [int]$dv[1]; $dg = [int]$dv[2]; $db = [int]$dv[3]
        $contrast = 0; $rows = @()
        for ($y = 0; $y -lt $h; $y++) { $row = $y * $bd.Stride; $n = 0; $xmin = 99999; $xmax = -1
            for ($x = 0; $x -lt $w; $x++) { $o = $row + $x*4
                $r = $buf[$o+2]; $gg = $buf[$o+1]; $b = $buf[$o]
                $d1 = [Math]::Abs($r-$dr); $d2 = [Math]::Abs($gg-$dg); $d3 = [Math]::Abs($b-$db)
                $mx = [Math]::Max($d1, [Math]::Max($d2, $d3))
                if ($mx -gt 32) { $n++; $contrast++; if ($x -lt $xmin){$xmin=$x}; if ($x -gt $xmax){$xmax=$x} } }
            if ($n -gt 0) { $rows += "y$y=${n}:$xmin-$xmax" } }
        $lines.Add("$($tag)_COLORS=$($hist.Count) $($tag)_DOMINANT=$($dom.Key)x$($dom.Value) $($tag)_PIXELS=$($w*$h) $($tag)_CONTRAST=$contrast")
        $lines.Add("$($tag)_ROWS=" + ($rows -join '|'))
        $top = $hist.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 5
        $lines.Add("$($tag)_TOP5=" + (($top | ForEach-Object { "$($_.Key)x$($_.Value)" }) -join ' '))
        $bmp.Save((Join-Path $wd ("cap_" + $label + "_" + $pngName + ".png")), [System.Drawing.Imaging.ImageFormat]::Png)
        $bmp.Dispose()
    }
    if ($sbH -ne [IntPtr]::Zero) { MeasurePix $sbH 'SB' 'sbdump' }
    else { $lines.Add("SB_COLORS=0 SB_DOMINANT=none SB_PIXELS=0 SB_CONTRAST=0"); $lines.Add("SB_ROWS="); $lines.Add("SB_TOP5=") }
    # 3) Output 面板 Edit 子窗文本
    $ed = PickMax { param($k) $k.c -eq 'Edit' }
    if ($ed) {
        $t = [CP]::Txt($ed.h)
        $tm = [CP]::EditTextRemote($ed.h)
        $head = $tm.Replace("`r",'\r').Replace("`n",'\n')
        if ($head.Length -gt 60) { $head = $head.Substring(0, 60) }
        $lines.Add("OUTEDIT rect=$([CP]::Rc($ed.h)) len=$($tm.Length) head=[$head]")
        $lines.Add("OUTEDIT_WTXT len=$($t.Length) head=[$($t.Replace("`r",'\r').Replace("`n",'\n'))]")
        $re = New-Object CP+RECT; [CP]::GetWindowRect($ed.h, [ref]$re) | Out-Null
        $pe = New-Object CP+POINT; $pe.X = [int](($re.L + $re.R) / 2); $pe.Y = [int](($re.T + $re.B) / 2)
        $hit = [CP]::WindowFromPoint($pe)
        $lines.Add("OUTEDIT_HITTEST cls=$([CP]::Cls($hit)) rect=$([CP]::Rc($hit)) same=$($hit -eq $ed.h)")
        MeasurePix $ed.h 'OUTEDIT' 'outdump'
    } else { $lines.Add("OUTEDIT rect=none len=0 head=[]") }
    # 4) 布局 rect
    $mc = PickMax { param($k) $k.c -eq 'MDIClient' }
    $doc = PickMax { param($k) $k.c -like '*frmDocument*' }
    if (-not $doc -and $mc) { $doc = PickMax { param($k) ($k.c -eq 'ThunderRT6FormDC' -and [CP]::GetParent($k.h) -eq $mc.h) } }
    $pans = @($script:kids | Where-Object { $_.c -like 'XTPDockingPaneTabbedContainer*' } | Sort-Object { $_.a } -Descending)
    $lines.Add("MDICLIENT rect=$(if($mc){[CP]::Rc($mc.h)}else{'none'})")
    $lines.Add("FRMDOCUMENT rect=$(if($doc){[CP]::Rc($doc.h)}else{'none'})")
    $pi = 0
    foreach ($pn in $pans) { $pi++; $lines.Add("DOCKPANE$pi rect=$([CP]::Rc($pn.h))") }
    if (-not $p.HasExited) { try { $p.Kill() } catch {} }
    return $p.Id
}
Cap $ExeDir $ExeName $Sec $Label | Out-Null
$lines | Set-Content -Path $Out -Encoding UTF8
$lines | ForEach-Object { $_ }
