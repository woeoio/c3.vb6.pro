param(
    [Parameter(Mandatory=$true)][string]$Exe,
    [string]$OutDir = "",
    [int]$Sec = 12
)
# GUISample 代码编辑器 (Codejock SyntaxEdit) 右/底缘像素判据 (本机验收夹具;
# -Category guisample 分片调用, 与 tree_icons_px.ps1 并列).
#
# 抓的两个回归点 (2026-11-04 用户实测):
#   ① 编辑器右边缘/底边缘各有一条黑线 (实测 x=1659..1660 与 y=965..966 为 #000000);
#   ② 本该在右/底的灰色滚动条整块没出现 (实测 x=1600..1667 恒为编辑器底 #F6F6F6).
# 根因 (实测读数链):
#   · MDIClient 的客户区相对窗口外扩 (win 1403x843@259,125 vs client 1407x847@256,123),
#     Codejock 按该客户区摆 MDI 子窗 => 文档窗 1403x814@256,151 右/底各留 3px 缝,
#     露出 MDIClient 底色 => 黑线。修法: MDIClient 样式/扩展样式对齐 VB6
#     (style=0x56000001, ex=WS_EX_CLIENTEDGE) 后客户区 1402x842@259,125。
#   · `ScrollBarVertical.ZOrder 0` / `ScrollBarHorizontal.ZOrder 0` 落通用晚绑定后
#     静默不动作 => 两枚 Codejock ScrollBar 被 SyntaxEdit 整块盖住。
#
# 判据 (与真 VB6 (vs2015 / vc6 同) 读数对齐, 全部 0 容差):
#   A. MDIClient: client 原点 == 窗口原点, client 尺寸 == 窗口尺寸 - 1 (VB6 实测 1402x842)。
#   B. 文档窗 = MDIClient 客户区原点 + 顶部 29px 页签条, 尺寸 == MDIClient 窗口宽 x (高-29)。
#   C. 竖直滚动条 21x814 贴文档右缘 (left = doc.left+1381), 水平滚动条 1381x21 贴底缘
#      (top = doc.top+792); 两枚都在 Z 序里**高于**编辑器包装窗 (VB6 实测 top->bottom:
#      ScrollBarH, ScrollBarV, SyntaxEdit 包装窗)。
#   D. CopyFromScreen 真桌面像素: 右 22px / 底 22px 色带内**不许出现纯黑/近黑**
#      (max(R,G,B) < 0x30 即判黑); 改前这里各有 2px #000000, 改后 0。
#
# 退出码: 0=PASS, 2=SKIP(取不到主窗/桌面), 其它=FAIL。

Add-Type -AssemblyName System.Drawing
Add-Type @"
using System; using System.Text; using System.Runtime.InteropServices;
public class ESP {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder sb, int n);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool GetClientRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool ClientToScreen(IntPtr h, ref POINT p);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern IntPtr GetParent(IntPtr h);
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  public struct POINT { public int X, Y; }
  public struct RECT { public int L, T, R, B; }
  public static string Cls(IntPtr h) { var sb = new StringBuilder(256); GetClassName(h, sb, 256); return sb.ToString(); }
  public static IntPtr FindMain(uint pid) {
    IntPtr best = IntPtr.Zero; long bestA = 0;
    EnumWindows((h, l) => {
      uint p; GetWindowThreadProcessId(h, out p);
      if (p == pid && IsWindowVisible(h) && Cls(h).StartsWith("VB6_Form_")) {
        RECT r; GetWindowRect(h, out r); long a = (long)(r.R-r.L)*(r.B-r.T);
        if (a > bestA) { bestA = a; best = h; }
      }
      return true;
    }, IntPtr.Zero);
    return best;
  }
  public static IntPtr FindClass(IntPtr root, string cls) {
    IntPtr found = IntPtr.Zero;
    EnumChildWindows(root, (h, l) => { if (found == IntPtr.Zero && Cls(h) == cls) found = h; return true; }, IntPtr.Zero);
    return found;
  }
  public static IntPtr FindPrefix(IntPtr root, string pre) {
    IntPtr found = IntPtr.Zero;
    EnumChildWindows(root, (h, l) => { if (found == IntPtr.Zero && Cls(h).StartsWith(pre)) found = h; return true; }, IntPtr.Zero);
    return found;
  }
  // 直接子窗, 按 Z 序 (顶层在前) 返回 "class|hw|l|t|r|b" 每行一条
  public static string[] Kids(IntPtr parent) {
    var list = new System.Collections.Generic.List<string>();
    EnumChildWindows(parent, (h, l) => {
      if (GetParent(h) != parent) return true;
      RECT r; GetWindowRect(h, out r);
      list.Add(Cls(h) + "|" + h.ToString() + "|" + r.L + "|" + r.T + "|" + r.R + "|" + r.B);
      return true;
    }, IntPtr.Zero);
    return list.ToArray();
  }
  public static int[] Win(IntPtr h) { RECT r; GetWindowRect(h, out r); return new int[] { r.L, r.T, r.R, r.B }; }
  public static int[] Cli(IntPtr h) {
    RECT c; GetClientRect(h, out c); POINT p = new POINT(); p.X = 0; p.Y = 0; ClientToScreen(h, ref p);
    return new int[] { c.R, c.B, p.X, p.Y };
  }
}
"@
[ESP]::SetProcessDPIAware() | Out-Null

if (-not (Test-Path $Exe)) { Write-Output "FAIL: exe 不存在 $Exe"; exit 1 }
if ($OutDir -and -not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }
$exePath = (Resolve-Path $Exe).Path
$wd = Split-Path $exePath
$proc = Start-Process -FilePath $exePath -WorkingDirectory $wd -PassThru `
    -RedirectStandardOutput (Join-Path $env:TEMP ("esp_" + $PID + ".out")) `
    -RedirectStandardError (Join-Path $env:TEMP ("esp_" + $PID + ".err"))
try { $proc.WaitForInputIdle(15000) | Out-Null } catch {}

$main = [IntPtr]::Zero
$deadline = [DateTime]::Now.AddSeconds($Sec + 8)
while ([DateTime]::Now -lt $deadline) {
    if ($proc.HasExited) { break }
    $m = [ESP]::FindMain([uint32]$proc.Id)
    if ($m -ne [IntPtr]::Zero -and [ESP]::Win($m)[2] -gt 200) { $main = $m; break }
    Start-Sleep -Milliseconds 250
}
if ($main -eq [IntPtr]::Zero) { Write-Output "SKIP: 主窗未找到 (单实例占用或启动超时)"; try { $proc.Kill() } catch {}; exit 2 }

[ESP]::ShowWindow($main, 9) | Out-Null
[ESP]::SetForegroundWindow($main) | Out-Null
Start-Sleep -Milliseconds 1500

$issues = New-Object System.Collections.Generic.List[string]

# ---------- A. MDIClient 客户区 ----------
$mc = [ESP]::FindClass($main, "MDIClient")
if ($mc -eq [IntPtr]::Zero) { $issues.Add("MDIClient 未找到") }
else {
    $mw = [ESP]::Win($mc); $mc2 = [ESP]::Cli($mc)
    $mww = $mw[2] - $mw[0]; $mwh = $mw[3] - $mw[1]
    if ($mc2[2] -ne $mw[0] -or $mc2[3] -ne $mw[1]) {
        $issues.Add("MDIClient 客户区原点(" + $mc2[2] + "," + $mc2[3] + ") != 窗口原点(" + $mw[0] + "," + $mw[1] + ") (黑线根因: 客户区外扩)")
    }
    if ($mc2[0] -ne ($mww - 1) -or $mc2[1] -ne ($mwh - 1)) {
        $issues.Add("MDIClient 客户区尺寸(" + $mc2[0] + "x" + $mc2[1] + ") != 窗口尺寸-1(" + ($mww - 1) + "x" + ($mwh - 1) + ")")
    }
}

# ---------- B. 文档窗 ----------
$doc = [ESP]::FindPrefix($mc, "VB6_Form_frmDocument")
if ($doc -eq [IntPtr]::Zero) { $issues.Add("frmDocument 未找到") }
else {
    $dw = [ESP]::Win($doc)
    $expW = $mww; $expH = $mwh - 29
    if (($dw[2] - $dw[0]) -ne $expW) { $issues.Add("文档窗宽 " + ($dw[2]-$dw[0]) + " != MDIClient 窗宽 " + $expW) }
    if (($dw[3] - $dw[1]) -ne $expH) { $issues.Add("文档窗高 " + ($dw[3]-$dw[1]) + " != MDIClient 窗高-29 " + $expH) }
    if ($dw[0] -ne $mc2[2] -or $dw[1] -ne ($mc2[3] + 29)) { $issues.Add("文档窗原点(" + $dw[0] + "," + $dw[1] + ") != 客户区原点+29 (= " + $mc2[2] + "," + ($mc2[3] + 29) + ")") }
}

# ---------- C. 滚动条矩形 + Z 序 ----------
$hbRect = $null; $vbRect = $null; $edRect = $null
$order = @{}
$idx = 0
foreach ($k in [ESP]::Kids($doc)) {
    $p = $k.Split('|')
    if (-not $p[0].StartsWith("AfxOleControl")) { continue }
    $idx++
    $l = [int]$p[2]; $t = [int]$p[3]; $r = [int]$p[4]; $b = [int]$p[5]
    $w = $r - $l; $h = $b - $t
    if ($w -gt 100 -and $h -le 25) { if (-not $hbRect) { $hbRect = @($l,$t,$w,$h); $order["hbar"] = $idx } }
    elseif ($h -gt 100 -and $w -le 25) { if (-not $vbRect) { $vbRect = @($l,$t,$w,$h); $order["vbar"] = $idx } }
    elseif ($w -gt 300 -and $h -gt 300) { if (-not $edRect) { $edRect = @($l,$t,$w,$h); $order["edit"] = $idx } }
}
if (-not $hbRect) { $issues.Add("水平滚动条包装窗未找到") }
if (-not $vbRect) { $issues.Add("竖直滚动条包装窗未找到") }
if (-not $edRect) { $issues.Add("编辑器包装窗未找到") }
if ($vbRect) {
    if ($vbRect[2] -ne 21) { $issues.Add("竖直滚动条宽 " + $vbRect[2] + " != 21") }
    if ($vbRect[0] -ne ($dw[0] + 1381)) { $issues.Add("竖直滚动条 left " + $vbRect[0] + " != doc.left+1381 " + ($dw[0] + 1381)) }
    if ($vbRect[1] -ne $dw[1] -or $vbRect[3] -ne ($dw[3] - $dw[1])) { $issues.Add("竖直滚动条未贴文档上/下缘 (top=" + $vbRect[1] + " h=" + $vbRect[3] + ")") }
}
if ($hbRect) {
    if ($hbRect[3] -ne 21) { $issues.Add("水平滚动条高 " + $hbRect[3] + " != 21") }
    if ($hbRect[1] -ne ($dw[1] + 792)) { $issues.Add("水平滚动条 top " + $hbRect[1] + " != doc.top+792 " + ($dw[1] + 792)) }
    if ($hbRect[0] -ne $dw[0] -or $hbRect[2] -ne 1381) { $issues.Add("水平滚动条 left/宽 " + $hbRect[0] + "/" + $hbRect[2] + " != doc.left/1381 " + $dw[0] + "/1381") }
}
if ($order.ContainsKey("hbar") -and $order.ContainsKey("vbar") -and $order.ContainsKey("edit")) {
    # EnumChildWindows 顶层在前 => 序号小的在上面。VB6: hbar(0) < vbar(1) < edit(2)
    if ($order["hbar"] -ge $order["edit"]) { $issues.Add("水平滚动条在编辑器之下 (z: h=" + $order["hbar"] + " edit=" + $order["edit"] + ")") }
    if ($order["vbar"] -ge $order["edit"]) { $issues.Add("竖直滚动条在编辑器之下 (z: v=" + $order["vbar"] + " edit=" + $order["edit"] + ")") }
    if ($order["hbar"] -ge $order["vbar"]) { $issues.Add("水平滚动条不在竖直滚动条之上 (z: h=" + $order["hbar"] + " v=" + $order["vbar"] + ")") }
}

# ---------- D. 真桌面像素: 右/底 22px 带 + 窗**外** 4px 环内不许出现黑 ----------
# 改前黑线实测在 x=1659..1660 / y=965..966, 而当时文档窗右/底缘是 1659/965 ⇒ 黑线
# **落在窗矩形之外**的 3px 缝里。所以采样既要覆盖窗内 22px 色带 (滚动条), 也要向
# 外多取 4px (缝 + 拆分器), 否则改前的读数取不到那 2px #000000。
$blackCount = 0
$minLum = 255
$pad = 4
try {
    $dwW = [int]($dw[2] - $dw[0]); $dwH = [int]($dw[3] - $dw[1])
    $bmp = New-Object System.Drawing.Bitmap -ArgumentList ([int]($dwW + $pad)), ([int]($dwH + $pad))
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($dw[0], $dw[1], 0, 0, $bmp.Size)
    $g.Dispose()
    # 窗口内采样上界 (含窗缘外 pad px)
    for ($i = 1; $i -le 8; $i++) {
        $y = [int]($dwH * $i / 9)
        for ($x = $dwW - 22; $x -lt ($dwW + $pad); $x++) {
            if ($x -lt 0 -or $y -lt 0) { continue }
            $c = $bmp.GetPixel($x, $y)
            $mx = [Math]::Max($c.R, [Math]::Max($c.G, $c.B))
            if ($mx -lt $minLum) { $minLum = $mx }
            if ($mx -lt 0x30) { $blackCount++ }
        }
    }
    for ($i = 1; $i -le 8; $i++) {
        $x = [int]($dwW * $i / 9)
        for ($y = $dwH - 22; $y -lt ($dwH + $pad); $y++) {
            if ($y -lt 0 -or $x -lt 0) { continue }
            $c = $bmp.GetPixel($x, $y)
            $mx = [Math]::Max($c.R, [Math]::Max($c.G, $c.B))
            if ($mx -lt $minLum) { $minLum = $mx }
            if ($mx -lt 0x30) { $blackCount++ }
        }
    }
    if ($blackCount -gt 0) { $issues.Add("右/底 22px 带 + 窗外 4px 环内近黑像素 " + $blackCount + " (改前: 右 2px 与底 2px 纯黑 = 黑线)") }
    if ($OutDir) {
        $png = Join-Path $OutDir "editor_scroll_c3.png"
        $bmp.Save($png, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    $bmp.Dispose()
} catch {
    Write-Output ("SKIP: 桌面捕获失败 " + $_.Exception.Message)
    try { $proc.Kill() } catch {}
    exit 2
}

try { $proc.Kill() } catch {}

if ($issues.Count -gt 0) {
    Write-Output ("FAIL: " + ($issues -join "; "))
    exit 1
}
Write-Output ("PASS mdi=" + ($mc2[0]) + "x" + ($mc2[1]) + "@" + ($mc2[2]) + "," + ($mc2[3]) +
    " doc=" + ($dw[2]-$dw[0]) + "x" + ($dw[3]-$dw[1]) + "@" + $dw[0] + "," + $dw[1] +
    " vbar=" + ($vbRect[2]) + "x" + ($vbRect[3]) + "@" + ($vbRect[0]) + "," + ($vbRect[1]) +
    " hbar=" + ($hbRect[2]) + "x" + ($hbRect[3]) + "@" + ($hbRect[0]) + "," + ($hbRect[1]) +
    " z(h,v,edit)=" + $order["hbar"] + "," + $order["vbar"] + "," + $order["edit"] +
    " blackPx=" + $blackCount + " minLum=" + $minLum)
exit 0
