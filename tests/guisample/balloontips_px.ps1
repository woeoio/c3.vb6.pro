param(
    [Parameter(Mandatory=$true)][string]$Exe,
    [string]$OutDir = "",
    [int]$Sec = 10
)
# BalloonTooltips 夹具的 tooltip 视觉判据 (tests\BalloonTooltips)。
#
# 参考物: 同目录的 VB6 产物 BalloonTooltips.exe (D:\vb_yqt4qPac\BalloonTooltips),
# 本机 1920x1200 / 120dpi(125%) 下逐控件实测 (TTM_GETTIPBKCOLOR/GETTIPTEXTCOLOR +
# WM_GETFONT->LOGFONT + 自身 DC 的 BitBlt 主色):
#
#   悬停控件                                 VB6 底色    C3 改前      C3 改后
#   Label "Label Tooltip"                    #FFFFE1     #FFFFFF      #FFFFE1
#   TextBox (Edit, 初始)                     #FFFF00     #FFFFFF      #FFFF00
#   Button "Change TextBox Tooltip"           #00FFFF     #FFFFFF      #00FFFF
#   Button "Enable/Disable Tooltip"           #00FF00     #FFFFFF      #00FF00
#   ComboBox                                 #FFFFE1     #FFFFFF      #FFFFE1
#   Label "Label In Frame Tooltip"           #FFFFE1     #FFFFFF      #FFFFE1
#   点 "Change TextBox Tooltip" 之后:
#     TextBox / Label                        #FFFFE1     #FFFFFF      #FFFFE1
#
# 改前的 getter TTM_GETTIPBKCOLOR/GETTIPTEXTCOLOR、WM_GETFONT(LOGFONT 面名/lfHeight/
# 字重/字符集)、style(exstyle) 与 VB6 **逐项相同**; 唯一差异是绘制层: 白底、客户区高
# +8..9px、lfQuality=5(ClearType)。根因 = cTT.cls 的
#      SetWindowTheme hWndTT, StrPtr(sEmpty), StrPtr(sEmpty)   (sEmpty = Const "")
# 在 VB6 里 StrPtr("") 答非 0 ⇒ uxtheme 关掉视觉样式 ⇒ 经典气泡按 TTM_SETTIP*COLOR 上色;
# C3 旧口径按长度判空 ⇒ 实际传 NULL ⇒ comctl32 保留视觉样式 ⇒ 主题白底。
# 修在 vb6_StrPtr (只看 NULL) + 发码侧 defaultValue(String)/vbNullString 发 NULL。
#
# 判据:
#   · 每组 tooltip 弹出后, 取 tooltip **自身 DC** 的 BitBlt 主色 (= 实际可见底色),
#     与上表期望值逐通道 ±3; 同时 TTM_GETTIPBKCOLOR 必须精确等于该 COLORREF
#     (0x00BBGGRR: #FFFFE1=0xE1FFFF / #FFFF00=0xFFFF / #00FFFF=0xFFFF00 / #00FF00=0xFF00);
#   · 字体面名 == 运行期 NONCLIENTMETRICS.lfStatusFont 面名, lfHeight 与该状态字体 ±1
#     (实测 VB6 与本机 C3 都是 Microsoft YaHei UI / -15; 面名/字号来自系统状态字体,
#      所以 CI 换语言/换 DPI 也成立, 不是把本机读数写死);
#   · style 必须含 TTS_BALLOON|TTS_ALWAYSTIP|TTS_NOPREFIX (0x43);
#   · 至少 6 组读数 + "改色后" 2 组。
#
# SKIP 口径 (rc=2): 只在取不到交互桌面时 (session 0: 屏幕 BitBlt 全黑) —— 环境不足,
# 不是用例失败。产物起不来/气泡弹不出/颜色不符一律 FAIL(rc=1), 不放水。

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing
Add-Type -ReferencedAssemblies 'System.Drawing' @"
using System; using System.Text; using System.Runtime.InteropServices; using System.Drawing; using System.Collections.Generic;
public class BTP {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder sb, int n);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowTextW(IntPtr h, StringBuilder sb, int n);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool GetClientRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
  [DllImport("user32.dll")] public static extern IntPtr WindowFromPoint(POINT p);
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr a, int x, int y, int cx, int cy, int f);
  [DllImport("user32.dll")] public static extern IntPtr SendMessageW(IntPtr h, int m, IntPtr w, IntPtr l);
  [DllImport("user32.dll")] public static extern bool SystemParametersInfoW(int a, int b, IntPtr c, int d);
  [DllImport("user32.dll")] public static extern int GetWindowLongW(IntPtr h, int i);
  [DllImport("gdi32.dll")] public static extern int GetObjectW(IntPtr h, int cb, IntPtr buf);
  [DllImport("user32.dll")] public static extern IntPtr GetDC(IntPtr h);
  [DllImport("user32.dll")] public static extern int ReleaseDC(IntPtr h, IntPtr dc);
  [DllImport("gdi32.dll")] public static extern bool BitBlt(IntPtr d, int x, int y, int w, int h, IntPtr s, int sx, int sy, int rop);
  [DllImport("gdi32.dll")] public static extern IntPtr CreateCompatibleDC(IntPtr dc);
  [DllImport("gdi32.dll")] public static extern IntPtr CreateCompatibleBitmap(IntPtr dc, int w, int h);
  [DllImport("gdi32.dll")] public static extern IntPtr SelectObject(IntPtr dc, IntPtr o);
  [DllImport("gdi32.dll")] public static extern bool DeleteObject(IntPtr o);
  [DllImport("gdi32.dll")] public static extern bool DeleteDC(IntPtr dc);
  [DllImport("gdi32.dll")] public static extern int GetTextExtentPoint32W(IntPtr dc, string s, int c, out SIZE sz);
  public struct POINT { public int X, Y; }
  public struct RECT { public int L, T, R, B; }
  public struct SIZE { public int CX, CY; }
  // VB6 TextHeight("Ag") 口径: 用 tooltip 那枚字体在屏幕 DC 上量字符串高
  public static int TextH(IntPtr hfont) {
    if (hfont == IntPtr.Zero) return -1;
    IntPtr sdc = GetDC(IntPtr.Zero);
    IntPtr mdc = CreateCompatibleDC(sdc);
    IntPtr old = SelectObject(mdc, hfont);
    SIZE sz; GetTextExtentPoint32W(mdc, "Ag", 2, out sz);
    SelectObject(mdc, old);
    DeleteDC(mdc); ReleaseDC(IntPtr.Zero, sdc);
    return sz.CY;
  }
  public static string Cls(IntPtr h) { var sb = new StringBuilder(256); GetClassName(h, sb, 256); return sb.ToString(); }
  public static string Txt(IntPtr h) { var sb = new StringBuilder(1024); GetWindowTextW(h, sb, 1024); return sb.ToString(); }
  public static int A(IntPtr h) { RECT r; GetWindowRect(h, out r); return (r.R-r.L)*(r.B-r.T); }
  public static int W(IntPtr h) { RECT r; GetWindowRect(h, out r); return r.R-r.L; }
  public static int H(IntPtr h) { RECT r; GetWindowRect(h, out r); return r.B-r.T; }
  public static IntPtr FindMain(uint pid) {
    IntPtr best = IntPtr.Zero; int bestA = 0;
    EnumWindows((h, l) => { uint p; GetWindowThreadProcessId(h, out p);
      if (p == pid && IsWindowVisible(h)) { string c = Cls(h);
        if (c.StartsWith("VB6_Form_") || c.StartsWith("Thunder") || c.StartsWith("VB6_MDI")) { int a = A(h); if (a > bestA) { bestA = a; best = h; } } }
      return true; }, IntPtr.Zero);
    return best;
  }
  public static IntPtr FindTip(uint pid) {
    IntPtr found = IntPtr.Zero; int bestA = 0;
    EnumWindows((h, l) => { uint p; GetWindowThreadProcessId(h, out p);
      if (p == pid && IsWindowVisible(h) && Cls(h) == "tooltips_class32") { int a = A(h); if (a > bestA) { bestA = a; found = h; } }
      return true; }, IntPtr.Zero);
    return found;
  }
  public static List<IntPtr> Kids(IntPtr main) { var list = new List<IntPtr>(); EnumChildWindows(main, (h,l) => { if (Cls(h) != "tooltips_class32") list.Add(h); return true; }, IntPtr.Zero); return list; }
  public static bool BringTop(IntPtr h) { return SetWindowPos(h, (IntPtr)(-1), 0, 0, 0, 0, 0x0001 | 0x0002 | 0x0010 | 0x0040); }
  public static uint PidOf(IntPtr h) { uint p; GetWindowThreadProcessId(h, out p); return p; }
  public static IntPtr WFP(int x, int y) { POINT p; p.X = x; p.Y = y; return WindowFromPoint(p); }
  public static int MakeLParam(int lo, int hi) { return (lo & 0xFFFF) | ((hi & 0xFFFF) << 16); }
  public static string LogFont(IntPtr hfont) {
    if (hfont == IntPtr.Zero) return "hfont=NULL";
    IntPtr buf = Marshal.AllocHGlobal(96);
    for (int i = 0; i < 96; i++) Marshal.WriteByte(buf, i, 0);
    int rc = GetObjectW(hfont, 92, buf);
    string face = "";
    for (int i = 0; i < 32; i++) { char c = (char)Marshal.ReadInt16(buf, 28 + i*2); if (c == 0) break; face += c; }
    int h = Marshal.ReadInt32(buf, 0); int weight = Marshal.ReadInt32(buf, 16);
    byte charset = Marshal.ReadByte(buf, 23); byte quality = Marshal.ReadByte(buf, 26);
    Marshal.FreeHGlobal(buf);
    return string.Format("rc={0} face='{1}' lfHeight={2} weight={3} charset={4} quality={5}", rc, face, h, weight, charset, quality);
  }
  // 运行期状态字体 (tooltip 用的就是它): NONCLIENTMETRICS.lfStatusFont, 偏移 316
  public static string StatusFont() {
    IntPtr buf = Marshal.AllocHGlobal(512);
    for (int i = 0; i < 512; i++) Marshal.WriteByte(buf, i, 0);
    Marshal.WriteInt32(buf, 0, 500);
    SystemParametersInfoW(0x0029, 500, buf, 0);
    int off = 316;
    string face = "";
    int h = Marshal.ReadInt32(buf, off);
    for (int i = 0; i < 32; i++) { char c = (char)Marshal.ReadInt16(buf, off + 28 + i*2); if (c == 0) break; face += c; }
    string s = string.Format("face='{0}' lfHeight={1} weight={2} charset={3}", face, h, Marshal.ReadInt32(buf, off+16), Marshal.ReadByte(buf, off+23));
    Marshal.FreeHGlobal(buf);
    return s;
  }
  // tooltip 自身 DC 的主色 (不算边框/标题区); 同时存 png 备查
  public static string Shot(IntPtr h, string png) {
    RECT r; GetClientRect(h, out r);
    int w = r.R, hh = r.B;
    if (w <= 0 || hh <= 0) return "zero-size";
    IntPtr sdc = GetDC(h);
    IntPtr mdc = CreateCompatibleDC(sdc);
    IntPtr bmp = CreateCompatibleBitmap(sdc, w, hh);
    IntPtr old = SelectObject(mdc, bmp);
    bool ok = BitBlt(mdc, 0, 0, w, hh, sdc, 0, 0, 0x00CC0020);
    SelectObject(mdc, old);
    ReleaseDC(h, sdc);
    var dict = new Dictionary<int, int>();
    using (var image = Image.FromHbitmap(bmp)) {
      if (png != "") image.Save(png, System.Drawing.Imaging.ImageFormat.Png);
      int x0 = w / 3, y0 = hh / 3, x1 = w - 2, y1 = hh - 2;
      for (int y = y0; y < y1; y += 2)
        for (int x = x0; x < x1; x += 2) {
          Color c = image.GetPixel(x, y);
          int key = (c.R << 16) | (c.G << 8) | c.B;
          if (dict.ContainsKey(key)) dict[key]++; else dict[key] = 1;
        }
    }
    DeleteObject(bmp); DeleteDC(mdc);
    int bestKey = 0, bestN = 0, total = 0;
    foreach (var kv in dict) { total += kv.Value; if (kv.Value > bestN) { bestN = kv.Value; bestKey = kv.Key; } }
    int R = (bestKey >> 16) & 0xFF, G = (bestKey >> 8) & 0xFF, B = bestKey & 0xFF;
    return string.Format("ok={0} size={1}x{2} mode=#{3:X2}{4:X2}{5:X2} frac={6:P0} colors={7}",
      ok, w, hh, R, G, B, (double)bestN / (total == 0 ? 1 : total), dict.Count);
  }
  // 交互桌面探针: 屏幕左上角 BitBlt 全黑 = session 0 / 无桌面
  public static bool ScreenLive() {
    IntPtr sdc = GetDC(IntPtr.Zero);
    if (sdc == IntPtr.Zero) return false;
    IntPtr mdc = CreateCompatibleDC(sdc);
    IntPtr bmp = CreateCompatibleBitmap(sdc, 64, 64);
    IntPtr old = SelectObject(mdc, bmp);
    bool ok = BitBlt(mdc, 0, 0, 64, 64, sdc, 0, 0, 0x00CC0020);
    SelectObject(mdc, old);
    ReleaseDC(IntPtr.Zero, sdc);
    int nonblack = 0;
    using (var image = Image.FromHbitmap(bmp)) {
      for (int y = 0; y < 64; y += 4)
        for (int x = 0; x < 64; x += 4) {
          Color c = image.GetPixel(x, y);
          if (c.R != 0 || c.G != 0 || c.B != 0) nonblack++;
        }
    }
    DeleteObject(bmp); DeleteDC(mdc);
    return ok && nonblack > 0;
  }
}
"@
[BTP]::SetProcessDPIAware() | Out-Null

$fail = New-Object System.Collections.ArrayList
$read = New-Object System.Collections.ArrayList
function Add-Fail([string]$m) { [void]$fail.Add($m) }
function Add-Read([string]$m) { [void]$read.Add($m) }

$exePath = (Resolve-Path $Exe -ErrorAction SilentlyContinue)
if (-not $exePath) { Write-Output "FAIL 缺产物 ($Exe 不存在)"; exit 1 }
$exePath = $exePath.Path
# 位宽: 本夹具按 --arch x86 编译, 与 VB6 参考产物同宽 (32 位 PE)。
$peBytes = [IO.File]::ReadAllBytes($exePath)
$peMachine = [BitConverter]::ToUInt16($peBytes, [BitConverter]::ToInt32($peBytes, 0x3C) + 4)
if ($peMachine -ne 0x014C) {
    Write-Output ("FAIL 位宽不符 (machine=0x{0:X4}, 期望 0x014C x86; 本夹具只按 --arch x86 编)" -f $peMachine)
    exit 1
}
$wd = Split-Path $exePath
if (-not $OutDir) { $OutDir = Join-Path $wd "balloontips_px" }
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

if (-not [BTP]::ScreenLive()) {
    Write-Output "SKIP 无交互桌面 (屏幕 BitBlt 全黑, session 0); 本判据需要真桌面才能弹 tooltip 采样"
    exit 2
}

$logOut = Join-Path $OutDir "balloontips.out"
$logErr = Join-Path $OutDir "balloontips.err"
$proc = Start-Process -FilePath $exePath -WorkingDirectory $wd `
    -RedirectStandardOutput $logOut -RedirectStandardError $logErr -PassThru
if (-not $proc) { Write-Output "FAIL 无法启动 $exePath"; exit 1 }
try { $proc.WaitForInputIdle(8000) | Out-Null } catch { }

$main = [IntPtr]::Zero
$deadline = (Get-Date).AddSeconds($Sec)
while ((Get-Date) -lt $deadline -and $main -eq [IntPtr]::Zero) {
    Start-Sleep -Milliseconds 150
    if ($proc.HasExited) { break }
    $main = [BTP]::FindMain([uint32]$proc.Id)
}
if ($main -eq [IntPtr]::Zero) {
    $errTail = ""
    if (Test-Path $logErr) { $errTail = (@((Get-Content $logErr -ErrorAction SilentlyContinue) | Where-Object { $_ }) | Select-Object -Last 3) -join " | " }
    try { if (-not $proc.HasExited) { $proc.Kill(); $proc.WaitForExit(3000) | Out-Null } } catch { }
    Write-Output ("FAIL 主窗在 ${Sec}s 内未出现 (exit=$($proc.ExitCode)) " + $errTail)
    exit 1
}
Add-Read ("main class=" + [BTP]::Cls($main) + " title=[" + [BTP]::Txt($main) + "] size=" + [BTP]::W($main) + "x" + [BTP]::H($main))
Start-Sleep -Milliseconds 600   # 首轮 tooltip 计时器要先吃一次真实鼠标移动, 让窗口先稳住

$statusFont = [BTP]::StatusFont()
Add-Read ("status font: " + $statusFont)
$sfFace = ($statusFont -replace ".*face='([^']*)'.*", '$1')
$sfHeight = [int]($statusFont -replace ".*lfHeight=(-?\d+).*", '$1')

$pid32 = [uint32]$proc.Id
function Reset-Hover {
    [BTP]::SetCursorPos(5, 5) | Out-Null
    [BTP]::SetForegroundWindow($main) | Out-Null
    [BTP]::BringTop($main) | Out-Null
    Start-Sleep -Milliseconds 400
}
# 悬停: 先把 app 抬到光标下 (本机 IDE/浏览器常在最前), 再直接 WM_MOUSEMOVE ——
# TTF_SUBCLASS 宿主自己会复核 GetCursorPos, 所以两条都要给。
function Invoke-Hover([IntPtr]$hwnd) {
    $r = New-Object BTP+RECT; [BTP]::GetWindowRect($hwnd, [ref]$r) | Out-Null
    $cx = [int](($r.L + $r.R) / 2); $cy = [int](($r.T + $r.B) / 2)
    $cr = New-Object BTP+RECT; [BTP]::GetClientRect($hwnd, [ref]$cr) | Out-Null
    $ccx = [int]($cr.R / 2); $ccy = [int]($cr.B / 2)
    $own = $false
    $tip = [IntPtr]::Zero
    # 第一枚气泡实测偶发不弹 (tooltip 的 initial 计时器要先看到一次**真实**鼠标移动)。
    # 解法: 最多 3 轮 "移开→回位 (抖动 1px)→WM_MOUSEMOVE", 只在真弹出来时退出;
    # 这不是放宽判据 —— 3 轮都不弹仍判 FAIL。
    for ($try = 1; $try -le 3; $try++) {
        Reset-Hover
        for ($i = 0; $i -lt 6; $i++) {
            [BTP]::SetCursorPos($cx, $cy) | Out-Null
            [BTP]::SetForegroundWindow($main) | Out-Null
            [BTP]::BringTop($main) | Out-Null
            Start-Sleep -Milliseconds 200
            if ([BTP]::PidOf([BTP]::WFP($cx, $cy)) -eq $pid32) { $own = $true; break }
        }
        # 抖动一下 (= 真实 WM_MOUSEMOVE), 再补一记显式消息
        [BTP]::SetCursorPos($cx + 1, $cy) | Out-Null
        Start-Sleep -Milliseconds 100
        [BTP]::SetCursorPos($cx, $cy) | Out-Null
        [BTP]::SendMessageW($hwnd, 0x200, [IntPtr]::Zero, [IntPtr][BTP]::MakeLParam($ccx, $ccy)) | Out-Null
        Start-Sleep -Milliseconds 1400
        $tip = [BTP]::FindTip($pid32)
        if ($tip -ne [IntPtr]::Zero) { break }
    }
    return @{ Tip = $tip; Own = $own; X = $cx; Y = $cy }
}

# C3 的 TextBox 窗口自身 GetWindowText 答空 (文本在内层), 所以 $txt 为空 = 按类名匹配
# (本窗体每类只有一个这样的控件: Edit/ComboBox; Static 有多个, 必须带文本)。
function Find-Kid([string]$cls, [string]$txt) {
    foreach ($k in ([BTP]::Kids($main))) {
        if ([BTP]::Cls($k) -ne $cls) { continue }
        if ($txt -eq "" -or [BTP]::Txt($k) -eq $txt) { return $k }
    }
    return [IntPtr]::Zero
}

# 期望表: 颜色来自 frmBalloonTooltips.frm 的 CreateToolTip 实参 (代码里的 tooltip 属性块)
#   cmdButtonInPic  vbGreen   -> #00FF00
#   cmdBalloonTooltip vbCyan  -> #00FFFF
#   txtBalloonTooltip vbYellow-> #FFFF00
#   其余 (cmb/lbl/lblInFrame) 不传颜色 -> 默认 #FFFFE1
$groups = @(
    @{ Cls = "Static";   Txt = "Label Tooltip";           RGB = 0xFFFFE1; Why = "无颜色 (默认)" },
    @{ Cls = "Edit";     Txt = "";                        RGB = 0xFFFF00; Why = "vbYellow 背色" },
    @{ Cls = "Button";   Txt = "Change TextBox Tooltip";  RGB = 0x00FFFF; Why = "vbCyan 背色" },
    @{ Cls = "Button";   Txt = "Enable/Disable Tooltip";  RGB = 0x00FF00; Why = "vbGreen 背色" },
    @{ Cls = "ComboBox"; Txt = "";                        RGB = 0xFFFFE1; Why = "无颜色 (默认)" },
    @{ Cls = "Static";   Txt = "Label In Frame Tooltip";  RGB = 0xFFFFE1; Why = "无颜色 (默认)" }
)

$measured = 0
$idx = 0
foreach ($g in $groups) {
    $idx++
    $k = Find-Kid $g.Cls $g.Txt
    if ($k -eq [IntPtr]::Zero) { Add-Fail ("G$idx 找不到控件 " + $g.Cls + " '" + $g.Txt + "'"); continue }
    $h = Invoke-Hover $k
    if ($h.Tip -eq [IntPtr]::Zero) {
        Add-Fail ("G$idx '" + $g.Txt + "' 悬停后没有 tooltip (own=" + $h.Own + " at " + $h.X + "," + $h.Y + ")")
        continue
    }
    $tip = $h.Tip
    $style = [BTP]::GetWindowLongW($tip, -16)
    # TTM_GETTIPBKCOLOR=0x416 / TTM_GETTIPTEXTCOLOR=0x417 / WM_GETFONT=0x31
    $bk = [int64][BTP]::SendMessageW($tip, 0x416, [IntPtr]::Zero, [IntPtr]::Zero)
    $fc = [int64][BTP]::SendMessageW($tip, 0x417, [IntPtr]::Zero, [IntPtr]::Zero)
    $hf = [BTP]::SendMessageW($tip, 0x31, [IntPtr]::Zero, [IntPtr]::Zero)
    $png = Join-Path $OutDir ("g" + $idx + "_" + ($g.Cls) + ".png")
    $shot = [BTP]::Shot($tip, $png)
    $lf = [BTP]::LogFont($hf)
    # COLORREF = 0x00BBGGRR: R | G<<8 | B<<16
    $expRef = (($g.RGB -shr 16) -band 0xFF) -bor (((($g.RGB -shr 8) -band 0xFF)) -shl 8) -bor (((($g.RGB -band 0xFF)) -shl 16))
    $th = [BTP]::TextH($hf)
    Add-Read ("G$idx '$($g.Txt)' [$($g.Why)] expect=#$('{0:X6}' -f $g.RGB) bk=$bk refBk=$expRef fc=$fc style=0x$('{0:X8}' -f $style) TextHeight=$th $shot font: $lf")
    $measured++

    if (($style -band 0x43) -ne 0x43) { Add-Fail ("G$idx style=0x{0:X8} 缺 TTS_BALLOON|TTS_ALWAYSTIP|TTS_NOPREFIX (mask 0x43)" -f $style) }
    if ($bk -ne $expRef) { Add-Fail ("G$idx TTM_GETTIPBKCOLOR=$bk 期望 $expRef (#$('{0:X6}' -f $g.RGB))") }

    # 实际可见像素: 主色与期望逐通道 ±3
    if ($shot -match 'mode=#([0-9A-F]{6})') {
        $mr = [Convert]::ToInt32($Matches[1].Substring(0,2),16)
        $mg = [Convert]::ToInt32($Matches[1].Substring(2,2),16)
        $mb = [Convert]::ToInt32($Matches[1].Substring(4,2),16)
        $er = ($g.RGB -shr 16) -band 0xFF; $eg = ($g.RGB -shr 8) -band 0xFF; $eb = $g.RGB -band 0xFF
        if ([Math]::Abs($mr-$er) -gt 3 -or [Math]::Abs($mg-$eg) -gt 3 -or [Math]::Abs($mb-$eb) -gt 3) {
            Add-Fail ("G$idx 可见底色 #$('{0:X2}{1:X2}{2:X2}' -f $mr,$mg,$mb) != 期望 #$('{0:X6}' -f $g.RGB) (逐通道 ±3)")
        }
        if (-not $h.Own) { Add-Read ("G$idx 提示: 悬停时指针下的窗口不属于被测进程 (own=false), 像素取自 tooltip 自身 DC 故不受遮挡影响") }
    } else {
        Add-Fail ("G$idx 取不到 tooltip 像素 ($shot)")
    }

    # 字体: 状态字体口径 (VB6 与本机实测同为 Microsoft YaHei UI / -15)
    if ($lf -match "face='([^']*)'") {
        $ttFace = $Matches[1]
        if ($ttFace -ne $sfFace) { Add-Fail ("G$idx 字体面名 '$ttFace' != 系统状态字体 '$sfFace'") }
    } else { Add-Fail ("G$idx 读不到字体 LOGFONT ($lf)") }
    if ($lf -match 'lfHeight=(-?\d+)') {
        $ttH = [int]$Matches[1]
        if ([Math]::Abs($ttH - $sfHeight) -gt 1) { Add-Fail ("G$idx lfHeight=$ttH 与状态字体 $sfHeight 差超过 1px") }
    }
}

# 点 "Change TextBox Tooltip" 之后: txt 与 lbl 两个 tooltip 被重建, 底色回默认 #FFFFE1
$btn = Find-Kid "Button" "Change TextBox Tooltip"
if ($btn -ne [IntPtr]::Zero) {
    Reset-Hover
    # BM_CLICK = 0x00F5
    [BTP]::SendMessageW($btn, 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null
    Start-Sleep -Milliseconds 800
    $after = @(
        @{ Cls = "Edit";   Txt = "";                RGB = 0xFFFFE1 },
        @{ Cls = "Static"; Txt = "Label Tooltip";   RGB = 0xFFFFE1 }
    )
    $j = 0
    foreach ($g in $after) {
        $j++
        $k = Find-Kid $g.Cls $g.Txt
        if ($k -eq [IntPtr]::Zero) { Add-Fail ("G7.$j 找不到控件 " + $g.Cls); continue }
        $h = Invoke-Hover $k
        if ($h.Tip -eq [IntPtr]::Zero) { Add-Fail ("G7.$j '" + $g.Txt + "' 改色后没有 tooltip"); continue }
        $shot = [BTP]::Shot($h.Tip, (Join-Path $OutDir ("g7_" + $j + ".png")))
        Add-Read ("G7.$j '$($g.Txt)' 改色后 expect=#FFFFE1 $shot")
        $measured++
        if ($shot -match 'mode=#([0-9A-F]{6})') {
            $mr = [Convert]::ToInt32($Matches[1].Substring(0,2),16)
            $mg = [Convert]::ToInt32($Matches[1].Substring(2,2),16)
            $mb = [Convert]::ToInt32($Matches[1].Substring(4,2),16)
            if ([Math]::Abs($mr-0xFF) -gt 3 -or [Math]::Abs($mg-0xFF) -gt 3 -or [Math]::Abs($mb-0xE1) -gt 3) {
                Add-Fail ("G7.$j 改色后可见底色 #$('{0:X2}{1:X2}{2:X2}' -f $mr,$mg,$mb) != #FFFFE1")
            }
        } else { Add-Fail ("G7.$j 取不到改色后的 tooltip 像素 ($shot)") }
    }
} else {
    Add-Fail "找不到按钮 'Change TextBox Tooltip'"
}

[BTP]::SetCursorPos(5, 5) | Out-Null
try { if (-not $proc.HasExited) { $proc.Kill(); $proc.WaitForExit(3000) | Out-Null } } catch { }

if ($measured -lt 6) { Add-Fail ("有效读数只有 $measured 组 (<6, 判据没被真正覆盖)") }
$txt = ($read -join "`r`n")
[IO.File]::WriteAllText((Join-Path $OutDir "balloontips.readings.txt"), $txt, (New-Object Text.UTF8Encoding($true)))

if ($fail.Count -eq 0) {
    Write-Output ("PASS tooltip 视觉: $measured 组底色/字体与参考 VB6 一致 (status font " + $sfFace + " " + $sfHeight + ")")
    exit 0
}
Write-Output ("FAIL " + $fail[0])
foreach ($f in ($fail | Select-Object -Skip 1 -First 6)) { Write-Output ("  " + $f) }
Write-Output ("  (读数 " + $measured + " 组, 详见 " + (Join-Path $OutDir "balloontips.readings.txt") + ")")
exit 1
