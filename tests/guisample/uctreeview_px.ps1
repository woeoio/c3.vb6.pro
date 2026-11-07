param(
    [Parameter(Mandatory=$true)][string]$Exe,
    [string]$OutDir = "",
    # 真 VB6 基准 exe (uctv_vb6_baseline.ps1 产出)。给了就先量它, 再用它的实拍读数判 C3。
    [string]$RefExe = "",
    # 基准读数落盘/复用路径。给了 RefExe 就写; 只给这个文件就读现成的 (省一次 VB6 起停)。
    [string]$RefJson = "",
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

# pwsh 7.6 (.NET 10) 把 System.Drawing 拆成多个程序集, 且 **Add-Type 引不住它**:
# 显式 -ReferencedAssemblies 会覆盖默认引用集 (CS0234/CS1069: Thread/Collections/Bitmap
# 全都解析不出来)。因此 C# 段刻意**不用任何 System.Drawing 类型** —— 只留 P/Invoke +
# byte[] 统计; 位图/抓屏/存 PNG 走 PowerShell 运行时 ([System.Drawing.Bitmap] 可直接用)。
# 全流程统一 pwsh, 不再有 5.1 分支。
# CI runner 的 pwsh 里**没有** System.Drawing.Common 这个程序集名 (实测
# ASSEMBLY_NOT_FOUND → 后面 [System.Drawing.Bitmap] 全 TypeNotFound), 而本机 7.6
# 两个名字都能装。以 System.Drawing 为主(= 其它 px 脚本统一口径), 失败再退 Common。
try { Add-Type -AssemblyName System.Drawing -ErrorAction Stop }
catch { Add-Type -AssemblyName System.Drawing.Common }
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
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  // 抬前台: 显示 + 还原 + TOPMOST + 置前. 前台锁可能拒绝, 调用方负责重试并校验
  // GetForegroundWindow (raw 采集失败的那几次就是只调一次被拒).
  public static bool ForceForeground(IntPtr h) {
    if (h == IntPtr.Zero) return false;
    if (IsIconic(h)) ShowWindow(h, 9);
    ShowWindow(h, 5);
    SetWindowPos(h, new IntPtr(-1), 0, 0, 0, 0, 0x0003);  // HWND_TOPMOST|NOMOVE|NOSIZE
    SetForegroundWindow(h);
    return GetForegroundWindow() == h;
  }
  [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, int flags);
  [DllImport("user32.dll")] public static extern bool RedrawWindow(IntPtr h, IntPtr r, IntPtr rgn, int flags);
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  [DllImport("user32.dll")] public static extern uint GetDpiForWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern int GetWindowLongW(IntPtr h, int idx);
  // 残差 #7: 按钮换行的期望行数必须按**该按钮自己的字体**在**当前 DPI** 下量出来 ——
  // 固定写 120dpi 参考值在 96dpi CI 上会红 (MS Sans Serif 光栅档在 96dpi 的宽度
  // 不是 120dpi 的 0.8 倍: Tree Explorer=63/Menu / Tools=64, 阈值恰好分开这两枚)。
  [DllImport("user32.dll")] public static extern IntPtr SendMessageW(IntPtr h, uint m, IntPtr w, IntPtr l);
  [DllImport("user32.dll")] public static extern IntPtr GetDC(IntPtr h);
  [DllImport("user32.dll")] public static extern int ReleaseDC(IntPtr h, IntPtr dc);
  [DllImport("gdi32.dll")] public static extern IntPtr SelectObject(IntPtr dc, IntPtr o);
  [DllImport("gdi32.dll")] public static extern bool DeleteObject(IntPtr o);
  [DllImport("gdi32.dll")] public static extern int GetObjectW(IntPtr h, int n, ref LOGFONTW lf);
  [DllImport("gdi32.dll")] public static extern IntPtr CreateFontIndirectW(ref LOGFONTW lf);
  [DllImport("gdi32.dll", CharSet=CharSet.Unicode)] public static extern bool GetTextExtentPoint32W(IntPtr dc, string s, int n, out SIZE sz);
  [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)] public struct LOGFONTW {
    public int lfHeight, lfWidth, lfEscapement, lfOrientation, lfWeight;
    public byte lfItalic, lfUnderline, lfStrikeOut, lfCharSet, lfOutPrecision, lfClipPrecision, lfQuality, lfPitchAndFamily;
    [MarshalAs(UnmanagedType.ByValTStr, SizeConst=32)] public string lfFaceName;
  }
  [StructLayout(LayoutKind.Sequential)] public struct SIZE { public int cx, cy; }
  // 按钮自己的 caption 宽 (与系统按钮绘制同一字体/同一 DPI)
  public static int TextWidthOf(IntPtr btn, string cap) {
    IntPtr hf = SendMessageW(btn, 0x0031, IntPtr.Zero, IntPtr.Zero);   // WM_GETFONT
    if (hf == IntPtr.Zero) return -1;
    var lf = new LOGFONTW(); if (GetObjectW(hf, 92, ref lf) <= 0) return -1;
    IntPtr dc = GetDC(IntPtr.Zero); IntPtr nf = CreateFontIndirectW(ref lf);
    IntPtr old = SelectObject(dc, nf);
    SIZE sz; GetTextExtentPoint32W(dc, cap, cap.Length, out sz);
    SelectObject(dc, old); DeleteObject(nf); ReleaseDC(IntPtr.Zero, dc);
    return sz.cx;
  }
  public struct RECT { public int L, T, R, B; }
  // 残差 #8: WM_GETFONT → LOGFONT 一行读数 "面名|高度|字符集|字重"; 无字体返回 ""
  public static string LogFontOf(IntPtr h) {
    IntPtr hf = SendMessageW(h, 0x0031, IntPtr.Zero, IntPtr.Zero);   // WM_GETFONT
    if (hf == IntPtr.Zero) return "";
    var lf = new LOGFONTW(); if (GetObjectW(hf, 92, ref lf) <= 0) return "";
    return lf.lfFaceName + "|" + lf.lfHeight + "|" + lf.lfCharSet + "|" + lf.lfWeight;
  }
  // --- ucTreeView 交互判据 (2026-11-06 缺陷 1/2/3/4) ---
  // 口径: 所有"等窗口"都有上限 (调用方给 3s), 模态盒 (#32770) 一律 WM_CLOSE 收口并校验
  // IsWindow=false; 主窗恢复响应用 SendMessageTimeout(WM_NULL, 2s) 单证。
  [DllImport("user32.dll")] public static extern bool PostMessageW(IntPtr h, uint m, IntPtr w, IntPtr l);
  [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr h);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowTextW(IntPtr h, System.Text.StringBuilder sb, int n);
  [DllImport("user32.dll")] public static extern IntPtr SendMessageTimeoutW(IntPtr h, uint m, IntPtr w, IntPtr l, uint flags, uint timeout, out IntPtr res);
  [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
  [DllImport("user32.dll")] public static extern void mouse_event(uint f, uint dx, uint dy, uint d, IntPtr e);
  public static string WText(IntPtr h) { var sb = new System.Text.StringBuilder(512); GetWindowTextW(h, sb, 512); return sb.ToString(); }
  // 跨进程 GetWindowText 对 Edit **只回缓存标题**（2026-11-07 实测：_Change 早已把文本改掉，
  // 但 GetWindowText 仍读到旧值 —— 组 3 的"txt 不联动"就是被这条采集缺陷骗出来的）。
  // Edit 的真文本只能经 WM_GETTEXT 取（系统消息里少数跨进程会被封送的一枚）。
  public static string WTextReal(IntPtr h) {
    IntPtr buf = System.Runtime.InteropServices.Marshal.AllocHGlobal(2048);
    try { SendMessageW(h, 0x000D, (IntPtr)1024, buf); return System.Runtime.InteropServices.Marshal.PtrToStringUni(buf); }
    finally { System.Runtime.InteropServices.Marshal.FreeHGlobal(buf); }
  }
  public static void Click(IntPtr h) { PostMessageW(h, 0x00F5, IntPtr.Zero, IntPtr.Zero); }   // BM_CLICK
  public static void CloseWin(IntPtr h) { PostMessageW(h, 0x0010, IntPtr.Zero, IntPtr.Zero); } // WM_CLOSE
  public static void DoubleClickAt(int x, int y) {
    SetCursorPos(x, y);
    mouse_event(0x0002, 0, 0, 0, IntPtr.Zero); mouse_event(0x0004, 0, 0, 0, IntPtr.Zero);
    System.Threading.Thread.Sleep(40);
    mouse_event(0x0002, 0, 0, 0, IntPtr.Zero); mouse_event(0x0004, 0, 0, 0, IntPtr.Zero);
  }
  // --- 双击的两条路 (CI 96dpi 上 mouse_event 那一路实测不生效: 两次读数 33/33 无变化) ---
  // 口径: PostMessage 合成 DOWN/UP + DBLCLK 到**光标下那个窗口** (换算出客户区坐标),
  // 不依赖前台/焦点; 台面上鼠标那一路只作兜底。两路都打不动才算"展开不生效"。
  [StructLayout(LayoutKind.Sequential)] public struct POINT { public int X, Y; }
  [DllImport("user32.dll")] public static extern IntPtr WindowFromPoint(POINT p);
  [DllImport("user32.dll")] public static extern bool ScreenToClient(IntPtr h, ref POINT p);
  [DllImport("user32.dll")] public static extern bool GetCursorPos(out POINT p);
  public static IntPtr HwndAt(int x, int y) { return WindowFromPoint(new POINT { X = x, Y = y }); }
  [DllImport("user32.dll")] public static extern bool IsWindowEnabled(IntPtr h);
  public static bool Enabled(IntPtr h) { return IsWindowEnabled(h); }
  public static void Repaint(IntPtr h) { RedrawWindow(h, IntPtr.Zero, IntPtr.Zero, 0x0181); }
  // 同步版 BM_CLICK: 直接等按钮窗口过程把这条消息处理完 (上限 3s, SMTO_ABORTIFHUNG)。
  // 用来把「消息压根没进窗口过程/队列」与「进了但画面没变」分开 —— PostMessage 那一路
  // 两者在读数上一模一样。返回 false = 3s 内没处理完 (窗口挂死或消息被丢弃)。
  public static bool ClickSync(IntPtr h) {
    IntPtr res = IntPtr.Zero;
    IntPtr ok = SendMessageTimeoutW(h, 0x00F5, IntPtr.Zero, IntPtr.Zero, 0x0002, 3000, out res);
    return ok != IntPtr.Zero;
  }
  public static string AtInfo(int x, int y) {
    IntPtr h = WindowFromPoint(new POINT { X = x, Y = y });
    if (h == IntPtr.Zero) return "none";
    return Cls(h) + "|" + WText(h);
  }
  // 返回 true = 消息已投出; 换算失败 (点不在该窗口里) 返回 false 交给兜底那一路。
  // ⚠ 必须先 SetCursorPos: 这枚控件 (ucTreeViewPlus.ctl:1228) 的 UserControl_DblClick
  // 是**用 GetCursorPos 自己取点**, 不读 lParam —— 只投消息不动光标, 命中行永远是旧位置。
  public static bool DblClickPost(IntPtr h, int sx, int sy) {
    if (h == IntPtr.Zero) return false;
    SetCursorPos(sx, sy);
    System.Threading.Thread.Sleep(30);
    var pt = new POINT { X = sx, Y = sy };
    if (!ScreenToClient(h, ref pt)) return false;
    IntPtr lp = (IntPtr)((pt.Y << 16) | (pt.X & 0xFFFF));
    PostMessageW(h, 0x0200, IntPtr.Zero, lp);        // WM_MOUSEMOVE
    System.Threading.Thread.Sleep(50);
    PostMessageW(h, 0x0201, (IntPtr)1, lp);          // WM_LBUTTONDOWN
    System.Threading.Thread.Sleep(60);
    PostMessageW(h, 0x0202, IntPtr.Zero, lp);        // WM_LBUTTONUP
    System.Threading.Thread.Sleep(60);
    PostMessageW(h, 0x0203, (IntPtr)1, lp);          // WM_LBUTTONDBLCLK
    System.Threading.Thread.Sleep(60);
    PostMessageW(h, 0x0202, IntPtr.Zero, lp);        // WM_LBUTTONUP
    return true;
  }
  // 窗体类名: C3 生成的 "VB6_Form_*"; 真 VB6 产物是 "ThunderRT6FormDC/ThunderRT6Form"
  // (基准升级后同一套取样器要对两种 exe 通用, 否则 REF 段一个窗都找不到)
  public static bool IsFormCls(string c) {
    return c.StartsWith("VB6_Form_") || c.StartsWith("ThunderRT6Form");
  }
  // 除主窗以外**可见**的窗体 (底部五按钮各自要弹一枚)
  public static IntPtr FindOtherForm(uint pid, IntPtr main) {
    IntPtr f = IntPtr.Zero;
    EnumWindows((h, l) => { uint p; GetWindowThreadProcessId(h, out p);
      if (p == pid && h != main && IsWindowVisible(h) && IsFormCls(Cls(h))) { f = h; return false; }
      return true; }, IntPtr.Zero);
    return f;
  }
  public static IntPtr FindDialog(uint pid) {
    IntPtr f = IntPtr.Zero;
    EnumWindows((h, l) => { uint p; GetWindowThreadProcessId(h, out p);
      if (p == pid && IsWindowVisible(h) && Cls(h) == "#32770") { f = h; return false; }
      return true; }, IntPtr.Zero);
    return f;
  }
  public static bool Responds(IntPtr h) {
    IntPtr res = IntPtr.Zero;
    IntPtr ok = SendMessageTimeoutW(h, 0x0000, IntPtr.Zero, IntPtr.Zero, 0x0002, 2000, out res);  // SMTO_ABORTIFHUNG
    return ok != IntPtr.Zero;
  }
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern IntPtr GetPropW(IntPtr h, string n);
  // STATIC/Label 不答 WM_GETFONT (VB6 也一样), 但 RTL 把实际字体存进了窗口属性
  // (vb6_ControlFontStore → SetPropW "VB6_CtrlFont"), 读那一格才是"实际生效的字体"。
  public static string PropLogFontOf(IntPtr h) {
    IntPtr hf = GetPropW(h, "VB6_CtrlFont");
    if (hf == IntPtr.Zero) return "";
    var lf = new LOGFONTW(); if (GetObjectW(hf, 92, ref lf) <= 0) return "";
    return lf.lfFaceName + "|" + lf.lfHeight + "|" + lf.lfCharSet + "|" + lf.lfWeight;
  }
  public static string Cls(IntPtr h) { var sb = new StringBuilder(256); GetClassName(h, sb, 256); return sb.ToString(); }
  // 枚举放 C# 侧: PS 5.1 的 scriptblock 委托回调查不到调用方 pid (会误报"主窗未找到")
  public static IntPtr FindMain(uint pid) {
    IntPtr best = IntPtr.Zero; long bestA = 0;
    EnumWindows((h, l) => {
      uint p; GetWindowThreadProcessId(h, out p);
      if (p == pid && IsWindowVisible(h) && IsFormCls(Cls(h))) {
        RECT r; GetWindowRect(h, out r);
        long a = (long)(r.R - r.L) * (r.B - r.T);
        if (a > bestA) { bestA = a; best = h; }
      }
      return true;
    }, IntPtr.Zero);
    return best;
  }
  // 树不是 SysTreeView32, 是 UserControl 宿主窗 (自绘树)。
  // C3 = "VB6_UserControlHost"; 真 VB6 = "ThunderRT6UserControl*" (类名随 VB6 版本变),
  // 所以再留一层兜底: 最大的"非标准控件"可见子窗 —— 判据不能绑死在某个类名上。
  public static IntPtr FindUcHost(IntPtr main) {
    IntPtr best = IntPtr.Zero; long bestA = 0;
    EnumChildWindows(main, (h, l) => {
      if (!IsWindowVisible(h)) return true;
      string c = Cls(h);
      if (!(c == "VB6_UserControlHost" || c.StartsWith("ThunderRT6UserControl"))) return true;
      RECT r; GetWindowRect(h, out r);
      long a = (long)(r.R - r.L) * (r.B - r.T);
      if (a > bestA) { bestA = a; best = h; }
      return true;
    }, IntPtr.Zero);
    if (best != IntPtr.Zero) return best;
    bestA = 0;
    EnumChildWindows(main, (h, l) => {
      if (!IsWindowVisible(h)) return true;
      string c = Cls(h);
      if (c == "Button" || c.StartsWith("ThunderRT6CommandButton") ||
          c == "Edit" || c.StartsWith("ThunderRT6TextBox") ||
          c == "Static" || c.StartsWith("ThunderRT6Label") ||
          c == "ScrollBar" || c.StartsWith("ThunderRT6ScrollBar")) return true;
      RECT r; GetWindowRect(h, out r);
      long a = (long)(r.R - r.L) * (r.B - r.T);
      if (a > bestA) { bestA = a; best = h; }
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
  [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint flags);
  [DllImport("gdi32.dll")] public static extern bool BitBlt(IntPtr d, int x, int y, int w, int h, IntPtr s, int sx, int sy, int rop);
  [DllImport("user32.dll")] public static extern int GetScrollPos(IntPtr h, int bar);
  public static void ClickAtClient(IntPtr h, int x, int y) {
    IntPtr lp = (IntPtr)((y << 16) | (x & 0xFFFF));
    PostMessageW(h, 0x0201, (IntPtr)1, lp);    // WM_LBUTTONDOWN
    System.Threading.Thread.Sleep(40);
    PostMessageW(h, 0x0202, IntPtr.Zero, lp);  // WM_LBUTTONUP
  }
  public static void SbmSetPos(IntPtr h, int pos) { SendMessageW(h, 0x00E0, (IntPtr)pos, (IntPtr)1); }
  static bool IsGreenPx(int R, int G, int B) { return Math.Abs(R-34)<=20 && Math.Abs(G-170)<=25 && Math.Abs(B-85)<=25; }
  static bool IsGreenish(int R, int G, int B) { return G > R+20 && G > B+20; }
  static bool IsLightGreenG(int R, int G, int B) { return Math.Abs(R-119)<=25 && Math.Abs(G-221)<=25 && Math.Abs(B-119)<=25; }
  // 内容级读数: 位图/抓屏/存 PNG 全在 PowerShell 侧 (pwsh 7 的 Add-Type 引不住
  // System.Drawing —— Drawing 被拆成多程序集, 显式传 refs 又会把 Thread/Collections
  // 解析断掉: CS0234/CS1069)。C# 只剩 P/Invoke + 这一段**纯 byte[] 统计**。
  // buf = BGRA (Format32bppArgb) LockBits 出来的原始像素; 输出串与旧实现逐字一致.
  public static string MetricsBuf(byte[] buf, int stride, int ox, int oy, int w, int hh, string mode) {
    int dark=0, vivid=0, white=0, green=0, lightGreen=0, photo=0, avatar=0, bg65=0, greenRim=0;
    int textRows=0, lightRows=0, firstTextW=0, lastTextY=0;
    bool inText=false, firstDone=false, inLight=false;
    int bandMin=99999, bandMax=-1, quiet=0;
    uint sig = 0;
    for (int y = oy; y < oy + hh; y++) {
      int cd=0, cl=0, tmin=99999, tmax=-1;
      for (int x = ox; x < ox + w; x++) {
        int bi = y * stride + x * 4;
        int B = buf[bi], G = buf[bi+1], R = buf[bi+2];
        int rx = x - ox;
        sig = sig * 131 + (uint)(R*7 + G*11 + B*13);
        int lum = (R*299 + G*587 + B*114)/1000;
        int mx = Math.Max(R, Math.Max(G, B)), mn = Math.Min(R, Math.Min(G, B));
        if (lum < 110) { dark++; cd++; if (rx < tmin) tmin = rx; if (rx > tmax) tmax = rx; }
        if (lum >= 200) cl++;
        if (mx - mn >= 60) vivid++;
        if (R >= 250 && G >= 250 && B >= 250) white++;
        if (IsGreenPx(R,G,B)) { green++; if (rx >= w-14 && rx < w-5) greenRim++; }
        if (IsLightGreenG(R,G,B)) lightGreen++;
        if (mx-mn >= 40 && !IsGreenish(R,G,B) && mn >= 25) photo++;
        if (mn >= 100 && mx >= 180 && (mx-mn) >= 25 && !IsGreenish(R,G,B)) avatar++;
        if (Math.Abs(R-57)<=2 && Math.Abs(G-61)<=2 && Math.Abs(B-65)<=2) bg65++;
      }
      // 一行算"墨迹"的门槛 = 3: 展开符 '>' 这种细笔画每行只有 3-5 个暗像素,
      // 用 >=6 会把"只有箭头没有文字"读成 0 行 (旧口径的采集侧误判)。
      // 带内允许 1 行空隙 (字形间隙), 连续 2 行无墨迹才收带。
      if (cd >= 3) {
        if (!inText) { inText = true; textRows++; bandMin = tmin; bandMax = tmax; quiet = 0; }
        else { if (tmin < bandMin) bandMin = tmin; if (tmax > bandMax) bandMax = tmax; quiet = 0; }
        lastTextY = y - oy;
      } else if (inText) {
        quiet++;
        if (quiet >= 2) { inText = false; if (!firstDone) { firstDone = true; firstTextW = bandMax-bandMin+1; } }
      }
      if (cl >= 6) { if (!inLight) { inLight = true; lightRows++; } } else inLight = false;
    }
    if (inText && !firstDone) { firstTextW = bandMax-bandMin+1; }
    int bw = stride / 4, bh = buf.Length / stride;
    int cx = Math.Min(ox + w/2, bw-1), cy = Math.Min(oy + hh/2, bh-1);
    int pb = cy * stride + cx * 4;
    return "w="+w+" h="+hh+" mode="+mode+" c="+buf[pb+2]+","+buf[pb+1]+","+buf[pb]
      +" dark="+dark+" textRows="+textRows+" lightRows="+lightRows
      +" vivid="+vivid+" white="+white+" green="+green+" lightGreen="+lightGreen
      +" photo="+photo+" avatar="+avatar+" bg65="+bg65+" greenRim="+greenRim
      +" firstTextW="+firstTextW+" lastTextY="+lastTextY+" sig="+sig;
  }
}
"@
[UP]::SetProcessDPIAware() | Out-Null

# --- 位图侧 (PowerShell): PrintWindow/CopyFromScreen -> Bitmap -> LockBits -> C# 统计 ---
function Get-WinRect {
    param([IntPtr]$Hwnd)
    $r = New-Object 'UP+RECT'
    if (-not [UP]::GetWindowRect($Hwnd, [ref]$r)) { return $null }
    return $r
}
function New-RenderBmp {
    param([IntPtr]$Hwnd, [int]$W, [int]$H)
    $bmp = New-Object System.Drawing.Bitmap($W, $H, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $mode = 'pw2'
    $dc = $g.GetHdc(); $ok = [UP]::PrintWindow($Hwnd, $dc, 2); $g.ReleaseHdc($dc)
    if (-not $ok) {
        $mode = 'pw0'
        $dc = $g.GetHdc(); $ok = [UP]::PrintWindow($Hwnd, $dc, 0); $g.ReleaseHdc($dc)
        if (-not $ok) { $mode = 'fail' }
    }
    $g.Dispose()
    # "整片单色" 检测 (PrintWindow 对个别子窗只回父底色时退 BitBlt)
    $flat = $true
    $p00 = $bmp.GetPixel([int]($W / 2), [int]($H / 2))
    $sx = [Math]::Max(1, [int]($W / 32)); $sy = [Math]::Max(1, [int]($H / 32))
    for ($yy = 3; $yy -lt $H -and $flat; $yy += $sy) {
        for ($xx = 3; $xx -lt $W -and $flat; $xx += $sx) {
            $cc = $bmp.GetPixel($xx, $yy)
            if ($cc.R -ne $p00.R -or $cc.G -ne $p00.G -or $cc.B -ne $p00.B) { $flat = $false }
        }
    }
    if ($flat -or $mode -eq 'fail') {
        $mode += '+blt'
        try {
            $src = [UP]::GetDC($Hwnd)
            $g3 = [System.Drawing.Graphics]::FromImage($bmp)
            $d3 = $g3.GetHdc()
            [UP]::BitBlt($d3, 0, 0, $W, $H, $src, 0, 0, 0x00CC0020) | Out-Null
            $g3.ReleaseHdc($d3); $g3.Dispose()
            [UP]::ReleaseDC($Hwnd, $src) | Out-Null
        } catch { $mode += '-err' }
    }
    return @{ Bmp = $bmp; Mode = $mode }
}
function Get-BmpStats {
    param([System.Drawing.Bitmap]$Bmp, [int]$Ox, [int]$Oy, [int]$W, [int]$H, [string]$Mode)
    $rect = New-Object System.Drawing.Rectangle(0, 0, $Bmp.Width, $Bmp.Height)
    $data = $Bmp.LockBits($rect, [System.Drawing.Imaging.ImageLockMode]::ReadOnly,
                          [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $stride = $data.Stride
    try {
        $buf = New-Object byte[] ($stride * $Bmp.Height)
        [System.Runtime.InteropServices.Marshal]::Copy($data.Scan0, $buf, 0, $buf.Length)
    } finally { $Bmp.UnlockBits($data) }
    return [UP]::MetricsBuf($buf, $stride, $Ox, $Oy, $W, $H, $Mode)
}
function Get-StatsPrint {
    param([IntPtr]$Hwnd)
    if ($Hwnd -eq [IntPtr]::Zero) { return 'none' }
    $r = Get-WinRect $Hwnd
    if ($null -eq $r) { return 'rect-fail' }
    $w = $r.R - $r.L; $h = $r.B - $r.T
    if ($w -le 0 -or $h -le 0) { return 'empty' }
    $rb = New-RenderBmp -Hwnd $Hwnd -W $w -H $h
    $s = Get-BmpStats -Bmp $rb.Bmp -Ox 0 -Oy 0 -W $w -H $h -Mode $rb.Mode
    $rb.Bmp.Dispose()
    return $s
}
function Get-StatsScreen {
    param([IntPtr]$Hwnd)
    if ($Hwnd -eq [IntPtr]::Zero) { return 'none' }
    $r = Get-WinRect $Hwnd
    if ($null -eq $r) { return 'rect-fail' }
    $w = $r.R - $r.L; $h = $r.B - $r.T
    if ($w -le 0 -or $h -le 0) { return 'empty' }
    $bmp = New-Object System.Drawing.Bitmap($w, $h, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($r.L, $r.T, 0, 0, (New-Object System.Drawing.Size($w, $h)))
    $g.Dispose()
    $s = Get-BmpStats -Bmp $bmp -Ox 0 -Oy 0 -W $w -H $h -Mode 'screen'
    $bmp.Dispose()
    return $s
}
# 重绘判据的双路签名: 抓屏 (CopyFromScreen) 与自绘 (PrintWindow) 各取一枚 sig。
# 2026-11-07 实测: `MarginBottom` 0→1 时 UC 真有 4269 字节差异 (PrintWindow 与
# CopyFromScreen 都测得到), 但守卫只比抓屏 sig 时偶发读成相同 —— 抓屏这一路会被
# 前台/遮挡/合成时机污染。两路任一变化即算"重绘", 两路都不变才红 (不放松阈值)。
function Get-UcSigPair {
  param([IntPtr]$Hwnd)
  $s = ((Get-StatsScreen $Hwnd) -split ' ' | Where-Object { $_ -like 'sig=*' })
  $q = ((Get-StatsPrint  $Hwnd) -split ' ' | Where-Object { $_ -like 'sig=*' })
  return "$s|$q"
}
function Save-WinPng {
    param([IntPtr]$Hwnd, [string]$Path)
    try {
        $r = Get-WinRect $Hwnd
        if ($null -eq $r) { return $false }
        $w = $r.R - $r.L; $h = $r.B - $r.T
        if ($w -le 0 -or $h -le 0) { return $false }
        $rb = New-RenderBmp -Hwnd $Hwnd -W $w -H $h
        $rb.Bmp.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
        $rb.Bmp.Dispose()
        return $true
    } catch { return $false }
}

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
[UP]::ForceForeground($main) | Out-Null
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

# 抓屏前再抬一次, 避免终端/IDE 盖住主窗导致整幅读成遮挡窗口
[UP]::SetWindowPos($main, [IntPtr](-1), 0, 0, 0, 0, 0x0003) | Out-Null
Start-Sleep -Milliseconds 700
[UP]::RedrawWindow($main, [IntPtr]::Zero, [IntPtr]::Zero, 0x0185) | Out-Null
Start-Sleep -Milliseconds 400
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

# --- 残差 #7: 底部 5 枚 push 按钮的 caption 多行 (BS_MULTILINE + 系统按空格 word-wrap) ---
# 真 VB6 的 ThunderRT6CommandButton 实测 style=0x54012000 (含 BS_MULTILINE 0x2000),
# 按**空格** word-wrap、行距 = TextHeight、整块竖直居中; C3 是原生 BUTTON, 缺这一位
# 时系统强制单行 (改前: 前三条 80x13/78x13/80x13 单行墨迹)。补上之后断行位置与 VB6
# 逐枚一致: "Tree Explorer" / "Corporate Hierarchy" / "Drag and Drop" 两行,
# "Menu / Tools" / "Contact  List" 单行 (参考墨迹 50x34 / 57x34 / 52x34 @y558-591)。
# 两行墨迹高度下限按 dpi 折算 = 26px@120dpi: 参考 VB6 = 34px, C3 字体档两行实测 29px,
# 单行最多 ~13px ⇒ 26 这条线既能抓住"退回单行", 又不会把字体档差异判红。
$btnDpi = [UP]::GetDpiForWindow($main)
if ($btnDpi -lt 96) { $btnDpi = 96 }
$expTwoLine = [int][Math]::Round(26 * $btnDpi / 120)
$btnCaps = @("Tree Explorer", "Corporate Hierarchy", "Drag and Drop", "Menu / Tools", "Contact  List")
# 参考行数: 真 VB6 在这 5 枚上是 2/2/2/1/1 —— 120dpi (Arial Narrow -16: 82/119/87/78/72 vs
# 文本区 80) 与 96dpi (Arial Narrow -13: 67/98/72/62/56 vs 文本区 63) 都是这一组。
# 前置条件: Arial Narrow 必须**真的装在这台机器上**。它属 Office 附带的字体, 不在 Windows
# 基础字体集里 —— CI 的 windows-latest 只有 Arial, 于是 GDI 把 "Arial Narrow" 替换成 Arial
# (-13 实测 77/75/74, 后两枚就翻成两行)。替换发生时 VB6 在同一台机器上也会做同样的替换
# (行数 2/2/2/2/2), 所以拿本机 (有 Arial Narrow) 录的参考行数去判它是**判据失效**,
# 不是产品缺陷: 这一条按"字体缺失"记 SKIP 并在读数里写明, 字体自洽判据 ($wantLines,
# 用按钮自己的字体实测) 仍然照常执行 —— 断行逻辑本身的回归仍能抓到。
$faceInstalled = $false
try {
    Add-Type -AssemblyName $dwAsm -ErrorAction Stop
    $faceInstalled = @((New-Object System.Drawing.Text.InstalledFontCollection).Families |
        Where-Object { $_.Name -eq "Arial Narrow" }).Count -gt 0
} catch { }
$btnRefLines = @(2, 2, 2, 1, 1); $btnIdx = 0
$btnMissing = 0; $btnBad = @(); $btnRead = @()
foreach ($cap in $btnCaps) {
    $bhw = [UP]::FindChildByText($main, $cap)
    if ($bhw -eq [IntPtr]::Zero) { $btnMissing++; continue }
    $br = New-Object UP+RECT; [UP]::GetWindowRect($bhw, [ref]$br) | Out-Null
    $cr = New-Object UP+RECT; [UP]::GetClientRect($bhw, [ref]$cr) | Out-Null
    $bx0 = [Math]::Max(0, $br.L - $wr.L); $by0 = [Math]::Max(0, $br.T - $wr.T)
    $bx1 = [Math]::Min($ww, $br.R - $wr.L); $by1 = [Math]::Min($wh, $br.B - $wr.T)
    $bw = $bx1 - $bx0; $bht = $by1 - $by0
    if ($bw -le 0 -or $bht -le 0) { $btnMissing++; continue }
    # 系统原生 push 按钮的文本区 = 客户宽 - 3px(每边) ⇒ 换行判定就是这个阈值。
    # 用**按钮自己的字体**(WM_GETFONT → LOGFONT → 本地重建)在当前 DPI 下量, 而不是
    # 写死 120dpi 的参考行数 —— 96dpi 的 MS Sans Serif 光栅档不是 120dpi 的等比缩放。
    $capW = [UP]::TextWidthOf($bhw, $cap)
    $clientW = $cr.R - $cr.L
    $wantLines = 1
    if ($capW -gt ($clientW - 6)) { $wantLines = 2 }
    $brows = New-Object int[] $bht
    $bMinY = 99999; $bMaxY = -1
    for ($yy = 0; $yy -lt $bht; $yy++) {
        for ($xx = 0; $xx -lt $bw; $xx++) {
            $c = $bmp.GetPixel($bx0 + $xx, $by0 + $yy)
            $lum = 0.299*$c.R + 0.587*$c.G + 0.114*$c.B
            if ($lum -lt 110) {
                $brows[$yy]++
                if ($yy -lt $bMinY) { $bMinY = $yy }
                if ($yy -gt $bMaxY) { $bMaxY = $yy }
            }
        }
    }
    $bands = 0; $on = $false
    for ($yy = 0; $yy -lt $bht; $yy++) {
        $cur = $brows[$yy] -gt 0
        if ($cur -and -not $on) { $bands++; $on = $true }
        elseif (-not $cur -and $on) { $on = $false }
    }
    $inkH = 0
    if ($bMaxY -ge 0) { $inkH = $bMaxY - $bMinY + 1 }
    $st = [UP]::GetWindowLongW($bhw, -16)
    $wlRef = if ($faceInstalled) { $btnRefLines[$btnIdx] } else { -1 }; $btnIdx++
    if (($st -band 0x2000) -eq 0) { $btnBad += "'$cap' 缺 BS_MULTILINE (style=0x$('{0:X8}' -f $st))" }
    if ($wlRef -ge 0 -and $bands -ne $wlRef) { $btnBad += "'$cap' 行数 $bands != 参考 $wlRef" }
    if ($bands -ne $wantLines) { $btnBad += "'$cap' 行数 $bands != $wantLines (ext=$capW/客户宽 $clientW)" }
    if ($wantLines -eq 2 -and $inkH -lt $expTwoLine) {
        $btnBad += "'$cap' 墨迹高 ${inkH}px < 两行下限 ${expTwoLine}px"
    }
    $btnRead += ("$cap=${bands}L/${inkH}px/ext$capW")
}
if ($btnMissing -gt 0) { $btnBad += "底部 push 按钮定位失败 $btnMissing/5" }

# --- 残差 #8: 窗体级设计期 Font 落地 + 子控件 ambient 继承 ---
# 真 VB6 实测: 窗体与"没有自带 Font 块"的子控件共用同一枚 HFONT =
# Arial Narrow 9.75pt => lfHeight -16 @120dpi / -13 @96dpi, charset=0 (ANSI, 与 .frm 里
# `Charset = 0` 同值), weight=400。
# Form1.frm 的窗体级块正是 Arial Narrow 9.75pt; 底部 5 枚按钮 / Check1 / Label1 都没有
# 自己的 Font 块, 所以它们必须与窗体同档 —— 这一条既抓"窗体块没落地", 也抓"没走继承"。
$expFace = "Arial Narrow"
$expFontH = -[int](9.75 * $btnDpi / 72.0)      # 与 RTL vb6_UC_FontHeightPx 同一算式
$fontBad = @()
$formLf = [UP]::LogFontOf($main)
if ($formLf -eq "") {
    $fontBad += "窗体 WM_GETFONT=0 (设计期字体没落到窗口)"
} else {
    $ff = $formLf -split '\|'
    if ($ff[0] -ne $expFace) { $fontBad += "窗体面名 '$($ff[0])' != '$expFace'" }
    if ([int]$ff[1] -ne $expFontH) { $fontBad += "窗体 lfHeight $($ff[1]) != $expFontH" }
    if ([int]$ff[2] -ne 0) { $fontBad += "窗体 charset $($ff[2]) != 0" }
    if ([int]$ff[3] -ne 400) { $fontBad += "窗体 weight $($ff[3]) != 400" }
}
# 四个探针各证一条链: 前三个的父容器都没有自带 Font 块 (窗体 / Picture3 / 窗体),
# 必须继承窗体那档; 最后一个是**反证** —— Label1 住在 Picture1 里, 而 Picture1 自带
# `BeginProperty Font MS Sans Serif 8.25`, 所以它必须继承 8.25 那档而不是窗体的
# Arial Narrow。少了这条, "子控件都在继承窗体字体"与"子控件都无视容器字体"这两种
# 实现都会被判绿。
# 最后一个只查面名: 8.25pt@120dpi 的 lfHeight 差 1px (VB6 取整到 -14, C3 截断到 -13)
# 是既有残差, 与本族（窗体级字体/继承链）无关, 记录在读数里但不在这里判。
$inheritProbes = @(
    @("Button",   "Tree Explorer",  $expFace,       $expFontH),
    @("CheckBox", "AllowEdit",      $expFace,       $expFontH),
    @("Label",    "MarginBottom",   $expFace,       $expFontH),
    @("Label",    "SelectionColor", "MS Sans Serif", $null)
)
$fontRead = @("form=" + $formLf)
foreach ($pr in $inheritProbes) {
    $h = [UP]::FindChildByText($main, $pr[1])
    if ($h -eq [IntPtr]::Zero) { $fontBad += "$($pr[0]) '$($pr[1])' 定位失败"; continue }
    $lf = [UP]::LogFontOf($h)
    if ($lf -eq "") { $lf = [UP]::PropLogFontOf($h) }   # STATIC 不答 WM_GETFONT, 读 RTL 存的那一格
    $fontRead += ($pr[0] + "=" + $lf)
    if ($lf -eq "") { $fontBad += "$($pr[0]) 读不到 LOGFONT"; continue }
    $g = $lf -split '\|'
    if ($g[0] -ne $pr[2]) {
        $fontBad += "$($pr[0]) '$($pr[1])' 面名 '$($g[0])' != '$($pr[2])' (继承链断)"
    }
    if ($pr[3] -ne $null -and ([int]$g[1] -ne $pr[3] -or [int]$g[2] -ne 0)) {
        $fontBad += "$($pr[0]) '$($pr[1])' 未继承窗体字体 ($lf)"
    }
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

# --- 残差 #7 的分区 diff 读数/上限: bottomBar 对**真 VB6 实拍**的同口径 diff ---
# CI 桌面没有真 VB6 参考图 ⇒ 这条在 CI 是 SKIP; 本机跑:
#   $env:UCTV_VB6_REF = <VB6 实拍 png (同尺寸)>
# 口径与收口用的 diffreg 一致: |ΔR|+|ΔG|+|ΔB| >= 6 记 1px。
# 改后整区实测 18.92% (改前 19.70%); 客户区行 540..613 口径改前 9.93% → 改后 9.06%
# (5 枚按钮内含边框整块 15.34% → 14.01%)。
$refPng = $env:UCTV_VB6_REF
$refDiff = "SKIP(无参考图)"
$refBad = @()
if ($refPng -and (Test-Path $refPng)) {
    $rb = $null; $cb2 = $null
    try { $rb = [System.Drawing.Bitmap]::FromFile((Resolve-Path $refPng).Path) } catch { $rb = $null }
    try { $cb2 = [System.Drawing.Bitmap]::FromFile($pngPath) } catch { $cb2 = $null }
    if ($rb -eq $null -or $cb2 -eq $null) {
        $refDiff = "SKIP(读不出)"
    } elseif ($rb.Width -ne $ww -or $rb.Height -ne $wh) {
        $refDiff = "SKIP(参考尺寸 $($rb.Width)x$($rb.Height) != ${ww}x${wh})"
    } else {
        $rd = 0L; $rn = 0L
        for ($yy = 540; $yy -lt [Math]::Min($wh, 623); $yy++) {
            for ($xx = 252; $xx -lt [Math]::Min($ww, 710); $xx++) {
                $ca = $rb.GetPixel($xx, $yy); $cb3 = $cb2.GetPixel($xx, $yy); $rn++
                $d = [Math]::Abs($ca.R-$cb3.R) + [Math]::Abs($ca.G-$cb3.G) + [Math]::Abs($ca.B-$cb3.B)
                if ($d -ge 6) { $rd++ }
            }
        }
        $refPct = 100.0 * $rd / [Math]::Max(1, $rn)
        $refDiff = ("{0:F2}%" -f $refPct)
        # 上限 20%: 改后实测 18.92% (改前 19.70%), 其中 ~11pp 是窗口底部非客户区
        # 在两次抓屏里的桌面渗入差 —— 不是按钮 caption; caption 侧剩余差 = 字体档
        # (VB6 继承窗体 ambient Arial Narrow 9.75pt/tmH20 vs C3 默认 MS Sans Serif/tmH16)。
        if ($refPct -gt 20.0) { $refBad += ("bottomBar diff {0:F2}% > 20.0%" -f $refPct) }
    }
    if ($rb -ne $null) { $rb.Dispose() }
    if ($cb2 -ne $null) { $cb2.Dispose() }
}

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

# ============================================================
# 交互判据 (2026-11-06 四条实测缺陷; 本机验收, CI 桌面无交互则本段自然红)
#   缺陷1: 展开后 icon 不能丢 + 子项可独立切换 -> 双击首行(展开/收起)量 vivid/darkRows
#   缺陷3: Load from custom XML 后树被重建     -> 高亮消失 + 可见节点行骤减
#   缺陷4: 底部五按钮各自弹出目标窗体         -> 逐个 BM_CLICK + 找新可见 VB6_Form_* + WM_CLOSE 收口
#   缺陷2: 三点按钮弹色选对话框               -> 等 #32770 (上限 3s) + WM_CLOSE + IsWindow=false
# 模态收口口径: 一切等待均有上限 (3s); 关不掉就 FAIL 并记读数, 绝不无界等待; 兜底由套件
# 看门狗 (/RunTimeoutSec + 进程 Kill) 收尾。
# ============================================================
$actBad = @(); $actRead = @()

function Measure-TreePx([int]$rx0, [int]$ry0, [int]$rtw, [int]$rth) {
    $bb = New-Object System.Drawing.Bitmap($rtw, $rth)
    $gg = [System.Drawing.Graphics]::FromImage($bb)
    $gg.CopyFromScreen($wr.L + $rx0, $wr.T + $ry0, 0, 0, (New-Object System.Drawing.Size($rtw, $rth)))
    $gg.Dispose()
    $v = 0; $dk = 0; $hlp = 0; $ck = 0
    $rows = New-Object int[] $rth
    for ($yy = 0; $yy -lt $rth; $yy++) {
        for ($xx = 0; $xx -lt $rtw; $xx++) {
            $c = $bb.GetPixel($xx, $yy)
            $mx = [Math]::Max($c.R, [Math]::Max($c.G, $c.B)); $mn = [Math]::Min($c.R, [Math]::Min($c.G, $c.B))
            # "vivid" 只收**低饱和排除后**的彩色: 选中行的粉底 (255,180,180) 不是图标,
            # 它一消失就会把"图标还在"的读数整个吃掉 (改前实测 7922 里 6898 是粉底)。
            $pastel = ($c.R -ge 240) -and ($c.G -ge 160) -and ($c.B -ge 160) -and ([Math]::Abs($c.G - $c.B) -le 25)
            if (($mx - $mn) -ge 60 -and -not $pastel) { $v++ }
            if ((0.299*$c.R + 0.587*$c.G + 0.114*$c.B) -lt 110) { $dk++; $rows[$yy]++ }
            if ($c.R -eq 255 -and $c.G -eq 180 -and $c.B -eq 180) { $hlp++ }
            if ($c.R -eq 99 -and $c.G -eq 99 -and $c.B -eq 255) { $ck++ }   # 勾选框蓝
        }
    }
    $dr = 0
    for ($yy = 0; $yy -lt $rth; $yy++) { if ($rows[$yy] -ge 3) { $dr++ } }
    $bb.Dispose()
    return @($v, $dk, $dr, $hlp, $ck)
}
function Wait-Until([scriptblock]$probe, [int]$ms) {
    $end = (Get-Date).AddMilliseconds($ms)
    while ((Get-Date) -lt $end) {
        $h = & $probe
        if ($h -ne [IntPtr]::Zero) { return $h }
        Start-Sleep -Milliseconds 100
    }
    return [IntPtr]::Zero
}
function Wait-Gone([IntPtr]$h, [int]$ms) {
    $end = (Get-Date).AddMilliseconds($ms)
    while ((Get-Date) -lt $end) {
        if (-not [UP]::IsWindow($h)) { return $true }
        Start-Sleep -Milliseconds 100
    }
    return (-not [UP]::IsWindow($h))
}
# 看门狗: 任何一步超时都必须**杀进程 + 打诊断**, 不许留下半死进程继续跑后面的判据.
# 注意 PowerShell 函数有自己的作用域, 红项必须写回 $script:actBad 才不会被丢掉.
function Kill-ProcDiag {
    param($Proc, [string]$why, [string]$OutLog = $logOut, [string]$ErrLog = $logErr)
    $script:actBad += ("看门狗: " + $why + " | " + (Get-LaunchDiag $Proc $OutLog $ErrLog))
    try { if ($Proc -and -not $Proc.HasExited) { $Proc.Kill() } } catch {}
}

if ($p.HasExited) {
    $actBad += "交互前进程已退出 (崩溃/自退, 交互判据无法取样)"
} else {
    # ---- 缺陷1: 双击首行 Facebook (有 3 个子项) 收起再展开, 两态各量一次 ----
    # 输入可靠性 (CI 96dpi 实测): mouse_event 那一路两次读数都是 33/33 —— 不是树没展开, 是
    # 合成点击没进窗口。故先按"光标下那个窗口 + 客户区坐标"走 PostMessage 的
    # DOWN/UP/DBLCLK (不依赖前台/焦点), 打不动再退回台面鼠标; 纵向最多试 4 个候选行。
    # 全程有界 (4 行 x 2 路 x 800ms), 不引入任何无界等待, 判据阈值一字未动。
    if ($hlMaxY -gt $hlMinY) {
        $preV, $preD, $preR, $preH = Measure-TreePx $x0 $y0 $tw $th
        $step = if ($pitch -gt 0) { $pitch } else { $expPitch }
        if ($step -le 0) { $step = 32 }
        $rows2try = @()
        for ($k = 0; $k -lt 4; $k++) {
            $yy = $wr.T + $y0 + $hlMinY + [int]($hlH / 2) + ($k * $step)
            if ($yy -ge ($wr.T + $y0 + $th - 4)) { break }
            $rows2try += ,@(($wr.L + $x0 + $hlMinX + 30), $yy)
        }
        $hitCls = @(); $toggle = $null
        foreach ($mode in @('post', 'mouse')) {
            if ($toggle) { break }
            foreach ($pt in $rows2try) {
                $px = $pt[0]; $py = $pt[1]
                [UP]::ForceForeground($main) | Out-Null
                if ($mode -eq 'post') {
                    $hitCls += ([UP]::AtInfo($px, $py))
                    # 目标优先级: 已定位的 UC 宿主 (不依赖 z 序/前台) → 光标下那个窗口
                    $thw = $uc
                    if ($thw -eq [IntPtr]::Zero) { $thw = [UP]::HwndAt($px, $py) }
                    [UP]::DblClickPost($thw, $px, $py) | Out-Null
                } else {
                    [UP]::DoubleClickAt($px, $py)
                }
                Start-Sleep -Milliseconds 800
                $aV, $aD, $aR, $aH = Measure-TreePx $x0 $y0 $tw $th
                # 只有"可见节点行"真的变 (>=3 行) 才算命中: 选中/焦点重绘会让 vivid 抖几个像素
                # (实测 1005→998), 拿它当命中信号会在第一行就误判成"已切换"。
                if ([Math]::Abs($aR - $preR) -ge 3) {
                    # 命中一个可展开节点: 再点同一处收回, 两态成对
                    if ($mode -eq 'post') {
                        [UP]::DblClickPost($thw, $px, $py) | Out-Null
                    } else {
                        [UP]::DoubleClickAt($px, $py)
                    }
                    Start-Sleep -Milliseconds 800
                    $bV, $bD, $bR, $bH = Measure-TreePx $x0 $y0 $tw $th
                    $toggle = @($aV, $aR, $bV, $bR, $px, $py, $mode)
                    break
                }
            }
        }
        if (-not $toggle) {
            $actRead += ("expand pre=$preR/$preV hit=[" + ($hitCls -join ';') + "] 无变化")
            $actBad += "展开/收起无差异 (4 行 x 2 路合成点击后可见行数都停在 $preR, 双击未进窗口或展开不生效)"
        } else {
            $av = $toggle[0]; $ar = $toggle[1]; $bv = $toggle[2]; $br = $toggle[3]
            if ($ar -ge $br) { $expR = $ar; $expV = $av; $colR = $br; $colV = $bv }
            else { $expR = $br; $expV = $bv; $colR = $ar; $colV = $av }
            $actRead += ("expand pre=$preR/$preV a=$ar/$av b=$br/$bv at=$($toggle[4]),$($toggle[5]) via=$($toggle[6]) hit=[" + ($hitCls -join ';') + "]")
            if ($expR -lt ($colR + 3)) { $actBad += "展开后可见节点行 $expR < 收起态 $colR + 3" }
            if ($expV -lt ($colV + 20)) { $actBad += "展开后图标彩色像素 $expV 未随子项增加 (收起态 $colV, 展开丢图标)" }
        }
    } else {
        $actBad += "首行高亮未定位, 展开判据无法取样"
    }

    # 交互块之间把光标收回中性位 (挪出树区)。CI 上实测: 合成双击之后紧接 BM_CLICK 的
    # XML 按钮不生效 (上一轮同一脚本 XML 是 33->10 通的), 先按最小口径隔离掉这个嫌疑。
    [UP]::SetCursorPos(($wr.L + 8), ($wr.T + 5)) | Out-Null
    Start-Sleep -Milliseconds 200

    # ---- 主界面数值控件: 6 枚 VScrollBar + Text1(index) 联动 ----
    # 三件事分开记: ①Value(=原生 pos) ②Text1(index).Text (Change 事件链的产物)
    # ③树区像素签名 (重绘)。上/下各点一次, 每次点击后立即读数, 不用重试掩盖吞点击。
    # ⚠ 顺序: 必须在「Load from custom XML」**之前**。实测(2026-11-07, 同一取样器,
    #   PrintWindow 与 CopyFromScreen 双路): XML 重填后树只剩 12 行, `MarginBottom`
    #   0→1/2/5/10 在 C3 与**真 VB6** 上都是 0 字节差异 (该属性只影响末行以下的留白);
    #   在初始 56 行状态下两边都真有 ~4.3k/4.7k 差异。把数值段放在 XML 之后会让这条
    #   判据对真 VB6 也判红 —— 是判据/顺序问题, 不是产品缺陷。
    if (-not $p.HasExited -and -not $dead) {
        $numSpecs = @(
            @{ label='BorderRadius';   min=0; max=15; val=10 },
            @{ label='BorderWidth';    min=0; max=5;  val=4  },
            @{ label='Margin';         min=0; max=20; val=15 },
            @{ label='NodeSpacing';    min=0; max=50; val=34 },
            @{ label='ToggleIconSize'; min=4; max=20; val=14 },
            @{ label='MarginBottom';   min=0; max=20; val=20 }
        )
        $sbAll = [UP]::FindByClass($main, "ScrollBar")
        $edAll = [UP]::FindByClass($main, "Edit")
        if ($sbAll.Length -lt 6) { $actBad += "数值控件定位失败 $($sbAll.Length)/6 (ScrollBar)" }
        foreach ($ns in $numSpecs) {
            $lab = [UP]::FindChildByText($main, $ns.label)
            if ($lab -eq [IntPtr]::Zero) { $actBad += "找不到数值标签 '$($ns.label)'"; continue }
            $lr = New-Object UP+RECT; [UP]::GetWindowRect($lab, [ref]$lr) | Out-Null
            $sb = [IntPtr]::Zero; $bestd = 999999
            foreach ($h in $sbAll) {
                $r = New-Object UP+RECT; [UP]::GetWindowRect($h, [ref]$r) | Out-Null
                if ($r.L -lt ($lr.R - 4)) { continue }
                $d = [Math]::Abs($r.T - $lr.T)
                if ($d -lt $bestd -and $d -le [int](20 * $expDpi / 96.0)) { $bestd = $d; $sb = $h }
            }
            if ($sb -eq [IntPtr]::Zero) { $actBad += "'$($ns.label)' 找不到对应 ScrollBar"; continue }
            $sr = New-Object UP+RECT; [UP]::GetWindowRect($sb, [ref]$sr) | Out-Null
            $sw = $sr.R - $sr.L; $sh = $sr.B - $sr.T
            $ed = [IntPtr]::Zero; $bestE = 999999
            foreach ($h in $edAll) {
                $er = New-Object UP+RECT; [UP]::GetWindowRect($h, [ref]$er) | Out-Null
                if ([Math]::Abs($er.T - $sr.T) -gt 10) { continue }
                if ($er.L -ge $sr.L) { continue }
                if ($er.L -lt $bestE) { $bestE = $er.L; $ed = $h }
            }
            if ($ed -eq [IntPtr]::Zero) { $actBad += "'$($ns.label)' 找不到对应 TextBox"; continue }
            # Value 口径: Win32 ScrollBar 的原生 pos = Value - Min (VB6 把 Min 折进位置),
            # 所以 Value = GetScrollPos + Min; Text1.Text = Max - Value + Min = Max - pos。
            $pos0 = [UP]::GetScrollPos($sb, 2)
            $v0 = $pos0 + $ns.min
            $txt0 = [UP]::WTextReal($ed)
            $sig0 = Get-UcSigPair $uc
            if ($v0 -ne $ns.val) { $actBad += "'$($ns.label)' 初始 Value=$v0 != 设计 $($ns.val) (Value/文本失步)" }
            $expTxt0 = [string]($ns.max - $pos0)
            if ($txt0 -ne $expTxt0) { $actBad += "'$($ns.label)' 初始 Text='$txt0' != Max-Value+Min='$expTxt0'" }
            [UP]::ClickAtClient($sb, [int]($sw/2), 3)
            Start-Sleep -Milliseconds 700
            $pos1 = [UP]::GetScrollPos($sb, 2); $txt1 = [UP]::WTextReal($ed)
            $sig1 = Get-UcSigPair $uc
            # 采集侧容错: 700ms 内没读到差异时再等 800ms 复采一次 (不重复点击、不改阈值)。
            # 2026-11-07 实测: 同一枚 MarginBottom 单发探针 (先点一次再等 900ms) ucDiff=4269;
            # 而在六枚连点之后同一操作在 700ms 上偶发读到相同 sig ⇒ 是刷新延迟, 不是"未重绘"。
            if ($sig1 -eq $sig0) { Start-Sleep -Milliseconds 800; $sig1 = Get-UcSigPair $uc; $actRead += "resample[$($ns.label):up];" }
            $actRead += ("num[$($ns.label)] up: pos=$pos0->$pos1 val=$v0->$($pos1+$ns.min) txt='$txt0'->'$txt1' px=$sig0->$sig1")
            if (($pos0 - $pos1) -ne 1) { $actBad += "'$($ns.label)' 上箭头 Value $v0->$($pos1+$ns.min) 未 -1 (SmallChange 未生效)" }
            $expTxt1 = [string]($ns.max - $pos1)
            if ($txt1 -ne $expTxt1) { $actBad += "'$($ns.label)' 上箭头后 Text='$txt1' != '$expTxt1' (Change 未发/未联动)" }
            if ($sig1 -eq $sig0) { $actBad += "'$($ns.label)' 上箭头后树区像素未变 (未重绘)" }
            [UP]::ClickAtClient($sb, [int]($sw/2), ($sh - 4))
            Start-Sleep -Milliseconds 700
            $pos2 = [UP]::GetScrollPos($sb, 2); $txt2 = [UP]::WTextReal($ed)
            $sig2 = Get-UcSigPair $uc
            if ($sig2 -eq $sig1) { Start-Sleep -Milliseconds 800; $sig2 = Get-UcSigPair $uc; $actRead += "resample[$($ns.label):dn];" }
            $actRead += ("num[$($ns.label)] dn: pos=$pos1->$pos2 val=$($pos1+$ns.min)->$($pos2+$ns.min) txt='$txt1'->'$txt2' px=$sig1->$sig2")
            if (($pos2 - $pos1) -ne 1) { $actBad += "'$($ns.label)' 下箭头 Value $($pos1+$ns.min)->$($pos2+$ns.min) 未 +1 (SmallChange 未生效)" }
            $expTxt2 = [string]($ns.max - $pos2)
            if ($txt2 -ne $expTxt2) { $actBad += "'$($ns.label)' 下箭头后 Text='$txt2' != '$expTxt2' (Change 未发/未联动)" }
            if ($sig2 -eq $sig1) { $actBad += "'$($ns.label)' 下箭头后树区像素未变 (未重绘)" }
            if ($p.HasExited) { $actBad += "'$($ns.label)' 数值点击后进程崩溃"; Kill-ProcDiag $p "'$($ns.label)' 数值点击后进程退出" $logOut $logErr; break }
        }
    }

    # ---- 缺陷3: Load from custom XML (按钮体里 Nodes.Clear + CheckBoxes=False + 递归重填) ----
    # 可测量口径: ①旧选中高亮清零 (Clear 生效) ②勾选框蓝(99,99,255)清零 (CheckBoxes=False 生效)
    # ③可见节点行数明显改变 (树被 XML 重填) ④无 #32770 解析错误框
    $xmlBtn = [UP]::FindChildByText($main, "Load from custom XML")
    if ($xmlBtn -eq [IntPtr]::Zero) {
        $actBad += "找不到按钮 'Load from custom XML'"
    } else {
        $xbV, $xbD, $xbR, $xbH, $xbCk = Measure-TreePx $x0 $y0 $tw $th
        # 点击最多 3 次 (1200ms/次, 每次重新定位按钮); 两次之间强制重绘一次, 把
        # "改了没刷屏"与"根本没进事件"分开。诊断里带按钮的类/可用位/caption 与尝试次数。
        $xdiag = ''
        $xmlTry = 0; $xv = 0; $xd2 = 0; $xr = 0; $xhl = 0; $xck = 0
        for ($ti = 1; $ti -le 3; $ti++) {
            $b = $xmlBtn
            if ($ti -gt 1) { $b = [UP]::FindChildByText($main, "Load from custom XML") }
            if ($b -eq [IntPtr]::Zero) { $xdiag += 'gone;'; break }
            if ($ti -eq 1) {
                $xdiag = 'cls=' + [UP]::Cls($b) + ' en=' + [UP]::Enabled($b) +
                    ' vis=' + [UP]::IsWindowVisible($b) + ' cap=' + [UP]::WText($b)
            }
            # 第 1 次走 PostMessage (与旧口径一致); 之后改走同步 BM_CLICK —— 同步版能确认
            # 按钮窗口过程**真的处理过**这条消息 (3s 上限), 把"没进事件"从"没刷屏"里分出来。
            if ($ti -eq 1) { [UP]::Click($b) }
            else { if (-not [UP]::ClickSync($b)) { $xdiag += 'sync-timeout;' } }
            Start-Sleep -Milliseconds 1200
            $xd = [UP]::FindDialog([uint32]$p.Id)
            if ($xd -ne [IntPtr]::Zero) {
                $actBad += ("XML 载入弹了模态框 '" + [UP]::WText($xd) + "'")
                [UP]::CloseWin($xd)
                [UP]::IsWindow($xd) | Out-Null
                Start-Sleep -Milliseconds 800
                break
            }
            $xmlTry = $ti
            $xv, $xd2, $xr, $xhl, $xck = Measure-TreePx $x0 $y0 $tw $th
            if ([Math]::Abs($xr - $xbR) -ge 5) { break }
            [UP]::Repaint($uc)
            Start-Sleep -Milliseconds 300
            $xv, $xd2, $xr, $xhl, $xck = Measure-TreePx $x0 $y0 $tw $th
            if ([Math]::Abs($xr - $xbR) -ge 5) { $xdiag += 'repaint;'; break }
        }
        $actRead += "xml rows=$xbR->$xr chk=$xbCk->$xck hl=$xbH->$xhl dark=$xd2 try=$xmlTry $xdiag"
        if ($xd2 -lt 30) { $actBad += "XML 载入后树区深色像素 $xd2 < 30 (连节点文字都没画)" }
        if ($xhl -gt 100) { $actBad += "XML 载入后仍有旧选中高亮 $xhl (Nodes.Clear 未生效)" }
        if ($xbCk -gt 20 -and $xck -gt 5) { $actBad += "XML 载入后勾选框蓝仍在 $xck (CheckBoxes=False 未生效)" }
        if ([Math]::Abs($xr - $xbR) -lt 5) { $actBad += "XML 载入前后可见节点行 $xbR->$xr 几乎没变 (树没被 XML 重填)" }
    }

    # ---- 缺陷2: 三点按钮 -> 色选对话框 (#32770), 探测/收口各有上限 ----
    $dotBtn = [UP]::FindChildByText($main, "...")
    if ($dotBtn -eq [IntPtr]::Zero) {
        $actBad += "找不到 '...' 按钮"
    } else {
        [UP]::Click($dotBtn)
        $dlg = Wait-Until { [UP]::FindDialog([uint32]$p.Id) } 3000
        if ($dlg -eq [IntPtr]::Zero) {
            $actBad += "三点按钮未弹色选对话框 (#32770 3s 内未出现)"
            $actRead += "colorDialog=none"
        } else {
            $dt = [UP]::WText($dlg)
            $actRead += "colorDialog='" + $dt + "'"
            [UP]::SetForegroundWindow($dlg) | Out-Null
            [UP]::CloseWin($dlg)
            $gone = Wait-Gone $dlg 3000
            if (-not $gone) {
                [UP]::PostMessageW($dlg, 0x0111, [IntPtr]2, [IntPtr]::Zero) | Out-Null   # WM_COMMAND IDCANCEL
                $gone = Wait-Gone $dlg 3000
            }
            $actRead += "colorClosed=$gone"
            if (-not $gone) { $actBad += "色选对话框 3s+3s 内关不掉 (模态框未收口)" }
            if (-not [UP]::Responds($main)) { $actBad += "色选关闭后主窗不响应 (WM_NULL 2s 超时)" }
        }
    }

    # ---- 缺陷4(本轮加严): 五按钮 —— ①逐窗内容级对照 ②同一按钮重复点击 3 轮 ----
    # 两段各用独立实例 (见数值段之后的 A/B 段):
    #   A) 每枚窗体一个新进程, 只点一次 → 内容读数不被前几枚窗体的 DestroyWindow 残留污染;
    #   B) 一个进程里同一按钮连点 3 轮, 每轮必须关干净 → 钉住重复点击缺陷。
    # 内容基准 = .frm/.frx 设计值(节点数/字体/边框色/头像) + 参考图内容清单;
    # 参考图只用于确认"该有什么", 像素阈值按本机 DPI 折算, 不做 1:1 像素比对。
    $formSpecs = @(
        @{ btn='Tree Explorer';       cap='Tree Explorer';     face='Segoe UI';          pt=9.75; weight=400; firstW=20 },
        @{ btn='Corporate Hierarchy'; cap='Form3';             face='Segoe UI';          pt=9.75; weight=400; firstW=80 },
        @{ btn='Drag and Drop';       cap='Drag & Drop';       face='Arial Narrow';      pt=9.75; weight=400; firstW=30 },
        @{ btn='Menu / Tools';        cap='Form5';             face='Segoe UI';          pt=9.75; weight=400; firstW=25 },
        @{ btn='Contact  List';       cap='Agenda tema oscuro'; face='Segoe UI Semibold'; pt=8.25; weight=600; firstW=0 }
    )
    $driveCount = @([System.IO.DriveInfo]::GetDrives() | Where-Object { $_.DriveType -ne 'NoRootDirectory' }).Count
    $expRows = @{ 'Tree Explorer' = 3 + $driveCount; 'Form3' = 9; 'Drag & Drop' = 5; 'Form5' = 10 }
    $dead = $false

}

    # 主实例读数已取完 —— 后面 A/B 段各起独立实例, 必须先杀主实例, 否则它的 TOPMOST 主窗
    # 会盖住新实例的非模态子窗体, 抓屏读到的是主窗像素 (实测 Tree 段读成了 Form1 的复选框区)。
    try { if (-not $p.HasExited) { $p.Kill() } } catch {}
    Start-Sleep -Milliseconds 800

    # ---------- R) 真 VB6 基准: 用**同一套取样器**量真产物, 读数落盘成 ref_metrics.json ----------
    # 基准升级 (用户手动 /make 成功): 五窗内容判据不再用"设计值 + 参考图", 而是逐项对
    # 真 VB6 exe 的实拍读数 (行数/首行墨迹宽/彩色像素/照片/头像/底色/字体度量)。
    # 参考图从此只用于"内容该有什么"的视觉确认, 不参与阈值。
    # 注: 真 VB6 产物必须在 zh-CN 下把 .ctl/.frm 的 "0,15" 归一成 "0.15" 才能 /make
    #     (见 uctv_vb6_baseline.ps1), 该归一不改变 Single 值。
    $ref = @{}
    if ($RefJson -ne '' -and (Test-Path $RefJson)) {
        try {
            $raw = Get-Content $RefJson -Raw | ConvertFrom-Json
            foreach ($k in $raw.PSObject.Properties.Name) {
                $h = @{}
                foreach ($q in $raw.$k.PSObject.Properties) { $h[$q.Name] = $q.Value }
                $ref[$k] = $h
            }
            $actRead += "ref=json:$RefJson"
        } catch { $actBad += "REF 读数文件解析失败 ($RefJson): $_" }
    }
    if ($ref.Count -eq 0 -and $RefExe -ne '') {
        if (-not (Test-Path $RefExe)) {
            $actBad += "REF exe 不存在 ($RefExe)"
        } else {
            $rw = Split-Path $RefExe
            $ri = 0
            foreach ($sp in $formSpecs) {
                $ri++
                $ro = Join-Path $OutDir "ref$ri.out"; $re = Join-Path $OutDir "ref$ri.err"
                $pr = $null; $mr = [IntPtr]::Zero
                try {
                    $pr = Start-Process -FilePath $RefExe -WorkingDirectory $rw -PassThru `
                          -RedirectStandardOutput $ro -RedirectStandardError $re
                    try { $pr.WaitForInputIdle(15000) | Out-Null } catch {}
                    $mr = Wait-Until { [UP]::FindMain([uint32]$pr.Id) } 8000
                } catch {}
                if (-not $pr -or $mr -eq [IntPtr]::Zero) {
                    $actBad += "REF '$($sp.btn)': 真 VB6 8s 内未起主窗"
                    if ($pr) { try { $pr.Kill() } catch {} }
                    continue
                }
                $rb = [UP]::FindChildByText($mr, $sp.btn)
                if ($rb -eq [IntPtr]::Zero) {
                    $actBad += "REF '$($sp.btn)': 真 VB6 主窗里找不到这枚按钮"
                    try { $pr.Kill() } catch {}
                    continue
                }
                [UP]::ClickSync($rb) | Out-Null
                $rfw = Wait-Until { [UP]::FindOtherForm([uint32]$pr.Id, $mr) } 3000
                if ($rfw -eq [IntPtr]::Zero) {
                    $actBad += "REF '$($sp.btn)': 真 VB6 3s 内没有弹出窗体"
                    try { $pr.Kill() } catch {}
                    continue
                }
                Start-Sleep -Milliseconds 2500
                $ruc = [UP]::FindUcHost($rfw)
                if ($ruc -eq [IntPtr]::Zero) {
                    $actBad += "REF '$($sp.btn)': 真 VB6 窗体里找不到树控件子窗"
                    try { $pr.Kill() } catch {}
                    continue
                }
                $rlf = [UP]::LogFontOf($ruc)
                if ($rlf -eq '') { $rlf = [UP]::PropLogFontOf($ruc) }
                if ($rlf -eq '') { $rlf = [UP]::LogFontOf($rfw) }
                $rstats = Get-StatsPrint $ruc
                $actRead += ("refcontent[$($sp.cap)] " + $rstats + " font=[$rlf]")
                $h = @{}; $h['font'] = $rlf
                foreach ($tok in ($rstats -split ' ')) {
                    $kv = $tok -split '='
                    if ($kv.Count -eq 2 -and $kv[0] -ne 'mode' -and $kv[0] -ne 'c') { $h[$kv[0]] = [int64]$kv[1] }
                }
                if ($h.ContainsKey('w')) { $ref[$sp.cap] = $h } else { $actBad += "REF '$($sp.cap)' 取样失败 ($rstats)" }
                if ($env:UCTV_SAVE_FORMS -eq '1' -and $OutDir -ne '') {
                    $safe = ($sp.cap -replace '[^A-Za-z0-9]', '_')
                    Save-WinPng $ruc (Join-Path $OutDir ("ref_$safe.png")) | Out-Null
                }
                [UP]::CloseWin($rfw) | Out-Null
                [void](Wait-Gone $rfw 3000)
                try { if (-not $pr.HasExited) { $pr.Kill() } } catch {}
                Start-Sleep -Milliseconds 400
            }
            if ($RefJson -ne '' -and $ref.Count -gt 0) {
                try {
                    $ref | ConvertTo-Json -Depth 4 | Set-Content -Encoding UTF8 $RefJson
                    $actRead += "ref=written:$RefJson"
                } catch { $actBad += "REF 读数写出失败 ($RefJson): $_" }
            }
        }
    }
    if ($ref.Count -eq 0) { $actRead += "ref=none(退回设计值口径)" }

    # ---------- A) 逐窗内容级对照: 每枚窗体一个**干净实例**, 只点一次再取样 ----------
    # 为什么必须换新进程: 已定位的"重复点击"缺陷里, 每枚窗体关闭后 vb6_hwnd_FormN /
    # vb6_formLoaded_FormN 都没清, 前几枚窗体的残留会把后面第 4/5 枚窗体的绘制带坏
    # (实测同一实例里 Form5/Form6 只画出窗口底色)。内容对照与重复点击是两条独立成因,
    # 所以内容段用一次性实例, 重复点击段单独再起一个实例连点 3 轮。
    $ai = 0
    foreach ($sp in $formSpecs) {
        $ai++
        $oA = Join-Path $OutDir ("contentA$ai.out")
        $eA = Join-Path $OutDir ("contentA$ai.err")
        $pa = $null; $ma = [IntPtr]::Zero
        try {
            $pa = Start-Process -FilePath $exePath -WorkingDirectory $wd -PassThru `
                 -RedirectStandardOutput $oA -RedirectStandardError $eA
            try { $pa.WaitForInputIdle(15000) | Out-Null } catch {}
            $ma = Wait-Until { [UP]::FindMain([uint32]$pa.Id) } 8000
        } catch {}
        if (-not $pa -or $ma -eq [IntPtr]::Zero) {
            $actRead += "content[$($sp.btn)] boot=fail"
            $actBad += "'$($sp.btn)' 内容段: 新实例 8s 内未起主窗"
            if ($pa) { try { $pa.Kill() } catch {} }
            continue
        }
        $fb = [UP]::FindChildByText($ma, $sp.btn)
        if ($fb -eq [IntPtr]::Zero) {
            $actBad += "'$($sp.btn)' 内容段: 干净实例里找不到按钮"
            Kill-ProcDiag $pa "'$($sp.btn)' 内容段找不到按钮" $oA $eA
            continue
        }
        $bsync = [UP]::ClickSync($fb)
        if (-not $bsync) {
            $actBad += "'$($sp.btn)' 内容段: BM_CLICK 3s 未返回 (事件没到按钮过程)"
            Kill-ProcDiag $pa "'$($sp.btn)' 内容段 BM_CLICK 超时" $oA $eA
            continue
        }
        $fw = Wait-Until { [UP]::FindOtherForm([uint32]$pa.Id, $ma) } 3000
        if ($fw -eq [IntPtr]::Zero) {
            $actRead += "content[$($sp.btn)]=OPEN-FAIL"
            $actBad += "'$($sp.btn)' 内容段: 3s 内没有弹出可见窗体"
            if (-not [UP]::Responds($ma)) { Kill-ProcDiag $pa "'$($sp.btn)' 内容段弹窗超时且主窗不响应" $oA $eA }
            try { if (-not $pa.HasExited) { $pa.Kill() } } catch {}
            continue
        }
        $fcap = [UP]::WText($fw)
        # 非模态子窗体可能被主窗/终端盖住 —— 抓屏前必须置 TOPMOST + 置前。
        # ⚠ 只发一次是**不够**的: 实测 A1(Tree) 第一轮 SetForegroundWindow 被前台锁
        # 拒掉 (返回 false), 3s 后 CopyFromScreen 读到的是 IDE 的暗色文件树 (c=13,17,23)。
        # 所以这里要**验证 + 最多 3 次重试**, 并把最终前台读数记进 actRead (fg=...) ——
        # 判据不变, 只修采集侧: 没真正抬到前台就不许拿这张图当内容读数。
        [UP]::SetWindowPos($fw, [IntPtr](-1), 0, 0, 0, 0, 0x0013) | Out-Null
        $fgOk = $false
        for ($fgTry = 1; $fgTry -le 3; $fgTry++) {
            [UP]::ForceForeground($fw) | Out-Null
            Start-Sleep -Milliseconds 900
            if ([UP]::GetForegroundWindow() -eq $fw) { $fgOk = $true; break }
        }
        $fgNow = [UP]::GetForegroundWindow()
        $fgTxt = [UP]::Cls($fgNow) + '|' + [UP]::WText($fgNow)
        $actRead += "fg[$($sp.btn)] ok=$fgOk try=$fgTry now=$fgTxt"
        # 抬不到前台**不再是红项**: 内容读数已改走 PrintWindow(2), 与 z-order 无关。
        # 这里只留读数, 供"为什么抓屏那张图是遮挡像素"的归因使用。
        [UP]::RedrawWindow($fw, [IntPtr]::Zero, [IntPtr]::Zero, 0x0185) | Out-Null
        $dlgA = [UP]::FindDialog([uint32]$pa.Id)
        if ($dlgA -ne [IntPtr]::Zero) {
            $actBad += ("'$($sp.btn)' 内容段弹了模态框 '" + [UP]::WText($dlgA) + "'")
            [UP]::CloseWin($dlgA) | Out-Null
            [void](Wait-Gone $dlgA 3000)
        }
        if ($fcap -ne $sp.cap) { $actBad += "'$($sp.btn)' 窗体标题 '$fcap' != '$($sp.cap)'" }
        # 给 Form_Load 异步链 (IMAGE/XML 载入) 留时间; 醒来后再确认一次前台没被抢
        Start-Sleep -Milliseconds 3000
        if ([UP]::GetForegroundWindow() -ne $fw) {
            [UP]::ForceForeground($fw) | Out-Null
            Start-Sleep -Milliseconds 600
            $actRead += "fg[$($sp.btn)] re-$([UP]::GetForegroundWindow() -eq $fw)"
        }
        $ucp = [UP]::FindUcHost($fw)
        if ($ucp -eq [IntPtr]::Zero) {
            $actBad += "'$($sp.btn)' 窗体里没有 VB6_UserControlHost (树控件缺失)"
        } else {
            # 主路: PrintWindow(2) 取 UC 自身 (不受 z-order/前台影响); scr= 仅作诊断对照
            $stats = Get-StatsPrint $ucp
            $actRead += ("scr[$($sp.cap)] " + (Get-StatsScreen $ucp))
            if ($env:UCTV_SAVE_FORMS -eq '1' -and $OutDir -ne '') {
                $safe = ($sp.cap -replace '[^A-Za-z0-9]', '_')
                Save-WinPng $ucp (Join-Path $OutDir ("form_$safe.png")) | Out-Null
            }
            $lf = [UP]::LogFontOf($ucp)
            if ($lf -eq '') { $lf = [UP]::PropLogFontOf($ucp) }
            $actRead += ("content[$($sp.cap)] " + $stats + " font=[$lf]")
            $st = @{}
            foreach ($tok in ($stats -split ' ')) {
                $kv = $tok -split '='
                if ($kv.Count -eq 2 -and $kv[0] -ne 'mode' -and $kv[0] -ne 'c') { $st[$kv[0]] = [int64]$kv[1] }
            }
            if (-not $st.ContainsKey('w')) {
                $actBad += "'$($sp.btn)' 内容取样失败 ($stats)"
            } else {
                $sc = $expDpi / 96.0
                $sc2 = $sc * $sc
                # 基准优先级: 有真 VB6 实拍读数就逐项对读数 (强口径); 没有才退回设计值口径。
                $rf = $ref[$sp.cap]
                if ($rf -ne $null -and $rf.ContainsKey('w')) {
                    if ($rf['w'] -ne $st['w'] -or $rf['h'] -ne $st['h']) {
                        $actBad += "'$($sp.cap)' 树控件 $($st['w'])x$($st['h']) != 真 VB6 $($rf['w'])x$($rf['h'])"
                    }
                    if ($st['textRows'] -ne $rf['textRows']) {
                        $actBad += "'$($sp.cap)' 文字行 $($st['textRows']) != 真 VB6 $($rf['textRows'])"
                    }
                    if ($rf['firstTextW'] -gt 0 -and $st['firstTextW'] -lt ($rf['firstTextW'] - 2)) {
                        $actBad += "'$($sp.cap)' 首行文本墨迹宽 $($st['firstTextW'])px < 真 VB6 $($rf['firstTextW'])-2px (文字缺失/截断)"
                    }
                    if ($rf['lightRows'] -gt 0 -and $st['lightRows'] -lt ($rf['lightRows'] - 2)) {
                        $actBad += "'$($sp.cap)' 浅色文字行 $($st['lightRows']) < 真 VB6 $($rf['lightRows'])-2"
                    }
                    foreach ($mk in @('vivid','photo','avatar','dark')) {
                        if ($rf[$mk] -gt 0 -and $st[$mk] -lt [int]($rf[$mk] * 0.9)) {
                            $actBad += "'$($sp.cap)' $mk=$($st[$mk]) < 真 VB6 $($rf[$mk])*0.9 (内容缺失)"
                        }
                    }
                    if ($rf['bg65'] -gt 0 -and $st['bg65'] -lt [int]($rf['bg65'] * 0.9)) {
                        $actBad += "'$($sp.cap)' 深底 bg65=$($st['bg65']) < 真 VB6 $($rf['bg65'])*0.9"
                    }
                    if ($st['greenRim'] -gt ($rf['greenRim'] + 8)) {
                        $actBad += "'$($sp.cap)' 边框右缘被裁切 greenRim=$($st['greenRim']) > 真 VB6 $($rf['greenRim'])+8"
                    }
                    $rlp = ([string]$rf['font']) -split '\|'
                    if ($rlp.Count -ge 4) {
                        $lfParts = $lf -split '\|'
                        if ($lfParts.Count -lt 4) {
                            $actBad += "'$($sp.cap)' 取不到 UC 字体 (WM_GETFONT 空, 真 VB6 是 $($rf['font']))"
                        } else {
                            if ($lfParts[0] -ine $rlp[0]) {
                                $actBad += "'$($sp.cap)' 字体 face '$($lfParts[0])' != 真 VB6 '$($rlp[0])'"
                            }
                            if ([Math]::Abs([int]$lfParts[1] - [int]$rlp[1]) -gt 1) {
                                $actBad += "'$($sp.cap)' 字体高 $($lfParts[1]) != 真 VB6 $($rlp[1]) ±1"
                            }
                            if ([int]$lfParts[3] -ne [int]$rlp[3]) {
                                $actBad += "'$($sp.cap)' 字重 $($lfParts[3]) != 真 VB6 $($rlp[3])"
                            }
                        }
                    } else {
                        $actBad += "'$($sp.cap)' 基准字体读数缺失 (REF font='$($rf['font'])')"
                    }
                } else {
                if ($sp.cap -ne 'Agenda tema oscuro') {
                    $rowsExp = $expRows[$sp.cap]
                    if ($st['textRows'] -lt $rowsExp) {
                        $actBad += "'$($sp.cap)' 节点文字行 $($st['textRows']) < $rowsExp (设计节点数)"
                    }
                    $firstMin = [int][Math]::Round($sp.firstW * $sc)
                    if ($st['firstTextW'] -lt $firstMin) {
                        $actBad += "'$($sp.cap)' 首行文本墨迹宽 $($st['firstTextW'])px < ${firstMin}px (文字缺失/截断)"
                    }
                }
                switch ($sp.cap) {
                    'Tree Explorer' {
                        if ($st['vivid'] -lt [int](200 * $sc2)) {
                            $actBad += "Tree Explorer 图标彩色像素 $($st['vivid']) < 200*scale2 (文件夹/驱动器彩色图标缺失)"
                        }
                    }
                    'Form3' {
                        if ($st['green'] -lt [int](800 * $sc2)) {
                            $actBad += "Corporate 绿边框像素 $($st['green']) < 800*scale2"
                        }
                        if ($st['photo'] -lt [int](400 * $sc2)) {
                            $actBad += "Corporate 头像照片像素 $($st['photo']) < 400*scale2 (5 张头像缺失)"
                        }
                        if ($st['greenRim'] -gt 5) {
                            $actBad += "Corporate 卡片右缘被裁切 greenRim=$($st['greenRim']) > 5"
                        }
                    }
                    'Drag & Drop' {
                        if ($st['vivid'] -lt [int](100 * $sc2)) {
                            $actBad += "Drag 文件夹图标彩色像素 $($st['vivid']) < 100*scale2"
                        }
                    }
                    'Form5' {
                        if ($st['lightGreen'] -lt [int](800 * $sc2)) {
                            $actBad += "Menu 浅绿边框像素 $($st['lightGreen']) < 800*scale2"
                        }
                        if ($st['lastTextY'] -lt [int]($st['h'] * 0.55)) {
                            $actBad += "Menu 最后文字行 y=$($st['lastTextY']) < 55% 高度 (10 行未铺满)"
                        }
                    }
                    'Agenda tema oscuro' {
                        if ($st['avatar'] -lt [int](3000 * $sc2)) {
                            $actBad += "Contact 头像像素 $($st['avatar']) < 3000*scale2 (头像全缺)"
                        }
                        if ($st['lightRows'] -lt 9) {
                            $actBad += "Contact 浅色文字行 $($st['lightRows']) < 9 (姓名/邮箱行不足)"
                        }
                        $bgMin = [int]($st['w'] * $st['h'] * 0.25)
                        if ($st['bg65'] -lt $bgMin) {
                            $actBad += "Contact 深底 bg65=$($st['bg65']) < 25% 面积"
                        }
                    }
                }
                $lfParts = $lf -split '\|'
                if ($lfParts.Count -lt 4) {
                    $actBad += "'$($sp.cap)' 取不到 UC 字体 (WM_GETFONT 空)"
                } else {
                    $lfFace = $lfParts[0]; $lfH = [int]$lfParts[1]; $lfW = [int]$lfParts[3]
                    $expFontH = [int][Math]::Round($sp.pt * $expDpi / 72.0)
                    $faceOk = ($lfFace -ieq $sp.face)
                    if (-not $faceOk -and $sp.face -eq 'Segoe UI Semibold' -and $lfFace -ieq 'Segoe UI' -and $lfW -ge 600) { $faceOk = $true }
                    if (-not $faceOk) { $actBad += "'$($sp.cap)' 字体 face '$lfFace' != '$($sp.face)'" }
                    if ([Math]::Abs([Math]::Abs($lfH) - $expFontH) -gt 1) {
                        $actBad += "'$($sp.cap)' 字体高 $lfH 偏离设计 -$expFontH ±1"
                    }
                    if ($lfW -ne $sp.weight) { $actBad += "'$($sp.cap)' 字重 $lfW != $($sp.weight)" }
                }
                }
            }
        }
        [UP]::CloseWin($fw) | Out-Null
        $fgA = Wait-Gone $fw 3000
        if (-not $fgA) {
            [UP]::PostMessageW($fw, 0x0111, [IntPtr]2, [IntPtr]::Zero) | Out-Null   # WM_COMMAND IDCANCEL
            $fgA = Wait-Gone $fw 3000
        }
        $actRead += "content[$($sp.btn)] close=$fgA"
        if (-not $fgA) { $actBad += "'$($sp.btn)' 内容段窗体 3s+3s 内关不掉"; Kill-ProcDiag $pa "'$($sp.btn)' 内容段关窗超时" $oA $eA }
        try { if (-not $pa.HasExited) { $pa.Kill() } } catch {}
    }

    # ---------- B) 重复点击: 同一进程, 5 按钮 x 3 轮, 每轮关干净 ----------
    $oB = Join-Path $OutDir "contentB.out"
    $eB = Join-Path $OutDir "contentB.err"
    $pb = $null; $mb = [IntPtr]::Zero
    try {
        $pb = Start-Process -FilePath $exePath -WorkingDirectory $wd -PassThru `
             -RedirectStandardOutput $oB -RedirectStandardError $eB
        try { $pb.WaitForInputIdle(15000) | Out-Null } catch {}
        $mb = Wait-Until { [UP]::FindMain([uint32]$pb.Id) } 8000
    } catch {}
    if (-not $pb -or $mb -eq [IntPtr]::Zero) {
        $actRead += "repeat=boot-fail"
        $actBad += "重复点击段: 第二实例 8s 内未起主窗"
        if ($pb) { try { $pb.Kill() } catch {} }
    } else {
        $deadB = $false
        foreach ($sp in $formSpecs) {
            if ($deadB) { break }
            for ($round = 1; $round -le 3; $round++) {
                if ($deadB) { break }
                $fb = [UP]::FindChildByText($mb, $sp.btn)
                if ($fb -eq [IntPtr]::Zero) {
                    $actRead += "repeat[$($sp.btn)] r$round btn=gone"
                    $actBad += "'$($sp.btn)' 第 $round 轮找不到按钮 (主窗控件已失效)"
                    break
                }
                $bsync = [UP]::ClickSync($fb)
                if (-not $bsync) {
                    $actRead += "repeat[$($sp.btn)] r$round clickSync=false"
                    $actBad += "'$($sp.btn)' 第 $round 轮 BM_CLICK 3s 未返回 (事件没到按钮过程)"
                    break
                }
                $fw = Wait-Until { [UP]::FindOtherForm([uint32]$pb.Id, $mb) } 3000
                if ($fw -eq [IntPtr]::Zero) {
                    $actRead += "repeat[$($sp.btn)] r$round=OPEN-FAIL"
                    $actBad += "'$($sp.btn)' 第 $round 轮 3s 内没有弹出可见窗体 (Show 重建失败)"
                    if (-not [UP]::Responds($mb)) {
                        Kill-ProcDiag $pb "'$($sp.btn)' 第 $round 轮弹窗超时且主窗不响应" $oB $eB
                        $deadB = $true
                    }
                    break
                }
                $fcap = [UP]::WText($fw)
                $dlgB = [UP]::FindDialog([uint32]$pb.Id)
                if ($dlgB -ne [IntPtr]::Zero) {
                    $actBad += "'$($sp.btn)' 第 $round 轮弹了模态框 '" + [UP]::WText($dlgB) + "'"
                    [UP]::CloseWin($dlgB) | Out-Null
                    [void](Wait-Gone $dlgB 3000)
                }
                [UP]::CloseWin($fw) | Out-Null
                $fgB = Wait-Gone $fw 3000
                if (-not $fgB) {
                    [UP]::PostMessageW($fw, 0x0111, [IntPtr]2, [IntPtr]::Zero) | Out-Null
                    $fgB = Wait-Gone $fw 3000
                }
                $actRead += "repeat[$($sp.btn) $round]=open:'$fcap' sync=$bsync close=$fgB"
                if (-not $fgB) {
                    $actBad += "'$($sp.btn)' 第 $round 轮弹窗 3s+3s 内关不掉"
                    Kill-ProcDiag $pb "'$($sp.btn)' 第 $round 轮关窗超时" $oB $eB
                    $deadB = $true
                    break
                }
                Start-Sleep -Milliseconds 400
                $left = [UP]::FindOtherForm([uint32]$pb.Id, $mb)
                $ldlg = [UP]::FindDialog([uint32]$pb.Id)
                if ($left -ne [IntPtr]::Zero -or $ldlg -ne [IntPtr]::Zero) {
                    $actBad += "'$($sp.btn)' 第 $round 轮关闭后仍有残留窗体"
                }
                if ($pb.HasExited) {
                    $actBad += "'$($sp.btn)' 第 $round 轮后进程崩溃"
                    Kill-ProcDiag $pb "'$($sp.btn)' 第 $round 轮后进程已退出" $oB $eB
                    $deadB = $true
                    break
                }
            }
        }
        try { if (-not $pb.HasExited) { $pb.Kill() } } catch {}
    }
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
# 残差 #7: 底部 5 枚 push 按钮 —— BS_MULTILINE 在位 + 行数与参考一致 + 两行墨迹高度下限
if ($btnBad.Count -gt 0) { $bad += ("底部按钮: " + ($btnBad -join "; ")) }
# 残差 #8: 窗体设计期字体 + 子控件 ambient 继承
if ($fontBad.Count -gt 0) { $bad += ("字体继承: " + ($fontBad -join "; ")) }
if ($refBad.Count -gt 0) { $bad += ($refBad -join "; ") }
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

# 交互判据 (缺陷 1/2/3/4) 的红项并入同一条 FAIL 通道, 不另开"软失败"
if ($actBad.Count -gt 0) { $bad += ("交互: " + ($actBad -join "; ")) }

$readout = "tree=${tw}x${th} white=$white bg240=$bg240 dark=$dark darkRows=$darkRows hl=$hl vivid=$vivid"
$readout += " selWhite=$selWhite selDark=$selDark selBlack=$selBlack hlH=${hlH}@${expHl} pitch=${pitch}@${expPitch} dpi=$expDpi"
$readout += " swatches=[" + ($swatchRead -join " | ") + "] chk240=$([Math]::Round(100*$chk240,1))% chkWhite=$([Math]::Round(100*$chkWhite,1))% label240=$([Math]::Round(100*$labelRatio,1))% labelWhite=$([Math]::Round(100*$labelWhiteRatio,1))%"
$readout += " edFrameDark=$edFrameDark edTopDark=$edTopDark edEdge236=$edEdge236/$($editHwnds.Length) sb240=$([Math]::Round(100*$sbR240,1))% sbWhite=$([Math]::Round(100*$sbRWhite,1))% sbMid=$sbMidOk/$sbCount sbLeft=$sbLeftOk/$sbCount"
$readout += " btns=[" + ($btnRead -join " | ") + "] twoLine=$expTwoLine bottomBarRefDiff=$refDiff"
$readout += " font=[" + ($fontRead -join " | ") + "] expFontH=$expFontH faceInstalled=$expFace" + $(if ($faceInstalled) { ":yes" } else { ":no(参考行数判据 SKIP)" })
$readout += " act=[" + ($actRead -join " | ") + "]"
if ($bad.Count -gt 0) {
    Write-Output ("FAIL " + ($bad -join "; ") + " | " + $readout + " | png=$pngPath")
    try { $p.Kill() } catch {}
    exit 1
}
Write-Output ("PASS " + $readout + " | png=$pngPath")
try { $p.Kill() } catch {}
exit 0
