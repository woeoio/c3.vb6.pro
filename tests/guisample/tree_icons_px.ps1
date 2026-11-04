param(
    [Parameter(Mandatory=$true)][string]$Exe,
    [string]$OutDir = "",
    [int]$Sec = 10,
    [int]$IconPx = 16
)
# GUISample 树图标透明判据 (本机验收夹具; -Category guisample 分片调用).
#
# 判据 = 用户实测口径: **图标周围像素 = 树底色**, 不许出现纯黑 #000000 方底.
# 具体做法: 枚举 SysTreeView32 的每个可见条目, 用 TVM_GETITEMRECT 拿到该行的
# 图标格 (left = 图标左缘, 宽 = IconPx, 高 = 行高); 检查**四角 + 四周 1px 环**
# 这些"图标块边界"像素, 出现 (0,0,0) 即判失败.
#   - 为什么要四角: mask 没生效时 ImageList 存进去的是不透明黑块, 四角会变黑;
#     只检查"格内全部像素"会把图标自身的黑色描边误判成黑底.
#   - 为什么不用 PrintWindow: Codejock 子窗在 PrintWindow 里是空白, 只能取窗口 DC
#     或桌面像素; 这里按验收要求走 **CopyFromScreen** (真桌面像素), 取到空图则 SKIP.
#
# 读数口径 (2026-11-04 实测, C3 x86 / 96dpi):
#   改前(黑底): 图标条左侧 20x380 内黑像素 11 (4.45%)
#   改后: 19 个图标, 四角+环黑像素 0
#   VB6 vc6 参考 (.build/guishot/tree_icons_vb6vc6.png, 同 16px 格): 0 (仅行距 17 时
#   跨行 1px 环会碰到下一行图标, 那是采样口径问题, 不是黑底)

Add-Type -AssemblyName System.Drawing
Add-Type @"
using System; using System.Text; using System.Runtime.InteropServices;
public class TP {
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
  [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, int flags);
  [DllImport("user32.dll")] public static extern bool RedrawWindow(IntPtr h, IntPtr r, IntPtr rgn, int flags);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern IntPtr SendMessageW(IntPtr h, int m, IntPtr w, IntPtr l);
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  [DllImport("kernel32.dll")] public static extern IntPtr OpenProcess(int access, bool inherit, uint pid);
  [DllImport("kernel32.dll")] public static extern IntPtr VirtualAllocEx(IntPtr p, IntPtr addr, int size, int type, int protect);
  [DllImport("kernel32.dll")] public static extern bool ReadProcessMemory(IntPtr p, IntPtr addr, byte[] buf, int size, out int read);
  [DllImport("kernel32.dll")] public static extern bool WriteProcessMemory(IntPtr p, IntPtr addr, byte[] buf, int size, out int written);
  [DllImport("kernel32.dll")] public static extern bool VirtualFreeEx(IntPtr p, IntPtr addr, int size, int type);
  [DllImport("kernel32.dll")] public static extern bool CloseHandle(IntPtr h);
  public struct POINT { public int X, Y; }
  public struct RECT { public int L, T, R, B; }
  public static string Cls(IntPtr h) { var sb = new StringBuilder(256); GetClassName(h, sb, 256); return sb.ToString(); }
  // Enumeration lives in C# instead of PS scriptblock delegates: under Windows
  // PowerShell 5.1 the callback cannot see the caller's pid (compares as 0), which
  // produced a bogus "main window not found" SKIP; pwsh 7 passes the same script.
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
  public static IntPtr FindTree(IntPtr main) {
    IntPtr best = IntPtr.Zero; long bestA = 0;
    EnumChildWindows(main, (h, l) => {
      if (IsWindowVisible(h) && Cls(h) == "SysTreeView32") {
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
[TP]::SetProcessDPIAware() | Out-Null

$exePath = (Resolve-Path $Exe -ErrorAction SilentlyContinue)
# SKIP 口径 (2026-11-04 收紧): 只有**确认缺资产 / 位宽不符**才允许 SKIP;
# "窗口没起"一律 FAIL 并带诊断 —— 早先它被判 SKIP, 等于"没验"却看着像通过。
if (-not $exePath) { Write-Output "SKIP 缺产物 ($Exe 不存在)"; exit 2 }
$exePath = $exePath.Path
$wd = Split-Path $exePath
# 位宽: 32 位宿主装不下 64 位 exe (反之 WoW64 可以), 只能这一种情况按位宽 SKIP
$peOff = [BitConverter]::ToInt32([IO.File]::ReadAllBytes($exePath), 0x3C)
$machine = [BitConverter]::ToUInt16([IO.File]::ReadAllBytes($exePath), $peOff + 4)
if ($machine -eq 0x8664 -and [IntPtr]::Size -eq 4) { Write-Output "SKIP 位宽不符 (64 位产物, 32 位宿主)"; exit 2 }
# 缺资产: 无 Codejock 运行库时产物必然起不来, 那是环境缺件, 不是本判据要抓的黑底
if (-not @(Get-ChildItem -Path $wd -Filter 'Codejock.*.ocx' -ErrorAction SilentlyContinue)) {
    Write-Output "SKIP 缺资产 ($wd 下没有 Codejock.*.ocx)"; exit 2
}
if (-not $OutDir) { $OutDir = Join-Path $wd "tree_icons_px" }
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$pngPath = Join-Path $OutDir "tree_icons_c3_crt.png"

# 产物启动/退出诊断: 崩溃时 RTL 的 C3_CRASH_TRACE 会把 [C3_CRASH] 写进 stderr,
# 一并摘出来, 免得"启动即退"又变成一句没有证据的 SKIP。
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

# 必须强制走 CreateProcess：Windows PowerShell 5.1 的 Start-Process 默认 UseShellExecute=$true，
# 会经 ShellExecute → 附件管理器校验。历史根因：仓库根曾带 Low 完整性标签（(OI)(CI)(NW)），
# 其下产物全部继承 Low，ShellExecute 会弹出"打开文件 - 安全警告（无法验证发布者）"并阻塞到人工点"运行"。
# PS 5.1 没有 -UseShellExecute 参数; 带 -RedirectStandard* 时 Start-Process 自动走
# UseShellExecute=$false（CreateProcess），不经过 ShellExecute，也就不会被该对话框拦
# （不能同时给 -NoNewWindow: 5.1 的两个参数互斥，会报 ParameterArgumentValidationError）。
# 另: 输出必须落到**文件**而不是管道 —— 产物会把 rtfText 之类的 stderr 刷屏, 管道没人读会写满阻塞;
# 同一条 -RedirectStandard* 也把 -PassThru 的进程对象做实 (实测 5.1 下纯 -NoNewWindow 偶发拿不到 Id)。
$logOut = Join-Path $OutDir "sample.out"
$logErr = Join-Path $OutDir "sample.err"
$env:C3_CRASH_TRACE = '1'   # 真崩时 stderr 会写 [C3_CRASH]，由 Get-LaunchDiag 摘出
$p = Start-Process -FilePath $exePath -WorkingDirectory $wd -PassThru `
     -RedirectStandardOutput $logOut -RedirectStandardError $logErr
try { $p.WaitForInputIdle(15000) | Out-Null } catch {}
Start-Sleep -Seconds $Sec

# --- 找主窗 + 树 (枚举在 C# 侧: PS 5.1 的 scriptblock 委托回调取不到 pid) ---
$script:main = [TP]::FindMain([uint32]$p.Id)
if ($script:main -eq [IntPtr]::Zero) {
    Write-Output ("FAIL 主窗未起 (pid=" + $p.Id + " " + (Get-LaunchDiag $p $logOut $logErr) + ")")
    try { $p.Kill() } catch {}
    exit 1
}

[TP]::ShowWindow($script:main, 9) | Out-Null
[TP]::SetForegroundWindow($script:main) | Out-Null
[TP]::SetWindowPos($script:main, [IntPtr](-1), 0, 0, 0, 0, 0x0003) | Out-Null
Start-Sleep -Milliseconds 800

$script:tree = [TP]::FindTree($script:main)
if ($script:tree -eq [IntPtr]::Zero) {
    Write-Output ("FAIL 主窗可见但无可见 SysTreeView32 (pid=" + $p.Id + " " + (Get-LaunchDiag $p $logOut $logErr) + ")")
    try { $p.Kill() } catch {}
    exit 1
}

[TP]::RedrawWindow($script:tree, [IntPtr]::Zero, [IntPtr]::Zero, 0x0185) | Out-Null
Start-Sleep -Milliseconds 800

# --- 客户区真桌面捕获 (CopyFromScreen) ---
$cr = New-Object TP+RECT; [TP]::GetClientRect($script:tree, [ref]$cr) | Out-Null
$org = New-Object TP+POINT; $org.X = 0; $org.Y = 0
[TP]::ClientToScreen($script:tree, [ref]$org) | Out-Null
$cw = $cr.R; $ch = $cr.B
if ($cw -le 0 -or $ch -le 0) { Write-Output "SKIP 树客户区为空 ${cw}x${ch}"; try { $p.Kill() } catch {}; exit 2 }
$bmp = New-Object System.Drawing.Bitmap($cw, $ch)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen($org.X, $org.Y, 0, 0, (New-Object System.Drawing.Size($cw, $ch)))
$g.Dispose()
$bmp.Save($pngPath, [System.Drawing.Imaging.ImageFormat]::Png)

# 空图判据: 整幅只有一种颜色 ⇒ 桌面不可捕获 (不许把空白图当 PASS)
$first = $bmp.GetPixel(0, 0)
$blank = $true
for ($py = 0; $py -lt $ch -and $blank; $py += 4) {
    for ($px = 0; $px -lt $cw -and $blank; $px += 4) {
        $c = $bmp.GetPixel($px, $py)
        if ($c.R -ne $first.R -or $c.G -ne $first.G -or $c.B -ne $first.B) { $blank = $false }
    }
}
if ($blank) { Write-Output "SKIP 桌面捕获为空图 (单色 $($first.R),$($first.G),$($first.B) 于 ${cw}x${ch})"; $bmp.Dispose(); try { $p.Kill() } catch {}; exit 2 }

# --- 枚举树条目: TVM_GETNEXTITEM + TVM_GETITEMRECT (跨进程 RECT 用远端缓冲) ---
$TVM_GETNEXTITEM = 0x110A   # TV_FIRST + 10
$TVM_GETITEMRECT = 0x1104   # TV_FIRST + 4
$TVGN_ROOT = 0; $TVGN_NEXT = 1; $TVGN_CHILD = 4
$proc = [TP]::OpenProcess(0x0038, $false, $p.Id)   # PROCESS_VM_OPERATION|VM_READ|VM_WRITE
if ($proc -eq [IntPtr]::Zero) { Write-Output "SKIP OpenProcess 失败"; $bmp.Dispose(); try { $p.Kill() } catch {}; exit 2 }
$rbuf = [TP]::VirtualAllocEx($proc, [IntPtr]::Zero, 64, 0x1000, 0x04)

$rows = New-Object System.Collections.Generic.List[string]
$icons = 0; $badRows = 0; $ringBlack = 0; $cornerBlack = 0
$checkErr = 0

function Get-ItemRect([IntPtr]$item) {
    # 远端 RECT 的 left 字段先放 hItem (TVM_GETITEMRECT 的输入)
    $bytes = [BitConverter]::GetBytes([int64]$item)
    $written = 0
    [TP]::WriteProcessMemory($proc, $rbuf, $bytes, 8, [ref]$written) | Out-Null
    $ok = [TP]::SendMessageW($script:tree, $TVM_GETITEMRECT, [IntPtr]1, $rbuf)
    if ([int]$ok -eq 0) { if ($env:TREE_PX_DEBUG) { Write-Output "DBG RET item=$item ok=$ok tree=$($script:tree) buf=$rbuf" }; return $null }
    $buf = New-Object byte[] 16
    $read = 0
    [TP]::ReadProcessMemory($proc, $rbuf, $buf, 16, [ref]$read) | Out-Null
    $r = New-Object TP+RECT
    $r.L = [BitConverter]::ToInt32($buf, 0);  $r.T = [BitConverter]::ToInt32($buf, 4)
    $r.R = [BitConverter]::ToInt32($buf, 8);  $r.B = [BitConverter]::ToInt32($buf, 12)
    return $r
}

$stack = New-Object System.Collections.Generic.Stack[IntPtr]
$root = [TP]::SendMessageW($script:tree, $TVM_GETNEXTITEM, [IntPtr]$TVGN_ROOT, [IntPtr]0)
if ([int64]$root -ne 0) {
    $stack.Push($root)
    $s = $root
    while ($true) {
        $nx = [TP]::SendMessageW($script:tree, $TVM_GETNEXTITEM, [IntPtr]$TVGN_NEXT, $s)
        if ([int64]$nx -eq 0) { break }
        $stack.Push($nx); $s = $nx
    }
}
$seen = @{}
while ($stack.Count -gt 0) {
    $it = $stack.Pop()
    $key = [string][int64]$it
    if ($seen.ContainsKey($key)) { continue }
    $seen[$key] = 1
    $kid = [TP]::SendMessageW($script:tree, $TVM_GETNEXTITEM, [IntPtr]$TVGN_CHILD, $it)
    if ([int64]$kid -ne 0) {
        $stack.Push($kid)
        $s = $kid
        while ($true) {
            $nx = [TP]::SendMessageW($script:tree, $TVM_GETNEXTITEM, [IntPtr]$TVGN_NEXT, $s)
            if ([int64]$nx -eq 0) { break }
            $stack.Push($nx); $s = $nx
        }
    }
    $r = Get-ItemRect $it
    if ($null -eq $r) { if ($env:TREE_PX_DEBUG) { Write-Output "DBG item=$key rect=fail" }; continue }
    $rowH = $r.B - $r.T
    if ($rowH -le 0) { if ($env:TREE_PX_DEBUG) { Write-Output "DBG item=$key rowH=$rowH" }; continue }
    # 只判落在捕获区内的行
    if ($r.L -lt 0 -or $r.T -lt 0 -or ($r.L + $IconPx) -gt $cw -or ($r.T + $rowH) -gt $ch) {
        if ($env:TREE_PX_DEBUG) { Write-Output "DBG item=$key rect=$($r.L),$($r.T),$($r.R),$($r.B) oob cap=${cw}x${ch}" }
        continue
    }
    # 该行是否真有图标内容 (全底色则说明这一行没图, 跳过)
    $hasIcon = $false
    for ($py = $r.T; $py -lt $r.T + $rowH -and -not $hasIcon; $py++) {
        for ($px = $r.L; $px -lt $r.L + $IconPx; $px++) {
            $c = $bmp.GetPixel($px, $py)
            if ([Math]::Abs($c.R - $first.R) -gt 24 -or [Math]::Abs($c.G - $first.G) -gt 24 -or [Math]::Abs($c.B - $first.B) -gt 24) { $hasIcon = $true; break }
        }
    }
    if (-not $hasIcon) { continue }
    $icons++
    $bad = $false
    for ($py = $r.T - 1; $py -le $r.T + $rowH; $py++) {
        for ($px = $r.L - 1; $px -le $r.L + $IconPx; $px++) {
            if ($px -lt 0 -or $py -lt 0 -or $px -ge $cw -or $py -ge $ch) { continue }
            # 图标右侧紧邻节点文字, 首列笔画天生是黑的 (root 实测 x=39,4/13), 不是黑底证据
            if ($px -eq $r.L + $IconPx) { continue }
            $inside = ($px -ge $r.L -and $px -lt $r.L + $IconPx -and $py -ge $r.T -and $py -lt $r.T + $rowH)
            $isCorner = (($px -eq $r.L -or $px -eq $r.L + $IconPx - 1) -and ($py -eq $r.T -or $py -eq $r.T + $rowH - 1))
            if ($inside -and -not $isCorner) { continue }   # 格内非角像素属于图标本体, 不算黑底证据
            $c = $bmp.GetPixel($px, $py)
            if ($c.R -eq 0 -and $c.G -eq 0 -and $c.B -eq 0) {
                $bad = $true
                if ($isCorner) { $cornerBlack++ } else { $ringBlack++ }
                $rows.Add("item=$key rect=$($r.L),$($r.T),$($r.R),$($r.B) blackAt=$px,$py")
            }
        }
    }
    if ($bad) { $badRows++ }
}
[TP]::VirtualFreeEx($proc, $rbuf, 0, 0x8000) | Out-Null
[TP]::CloseHandle($proc) | Out-Null
$bmp.Dispose()
if (-not $p.HasExited) { try { $p.Kill() } catch {} }

if ($icons -eq 0) { Write-Output "SKIP 未枚举到带图标的条目 (树为空/全隐藏) tree=$($script:tree) client=${cw}x${ch} org=$($org.X),$($org.Y) items=$($seen.Count)"; exit 2 }
$status = if ($badRows -eq 0) { "PASS" } else { "FAIL" }
Write-Output "$status tree_icons_px icons=$icons badRows=$badRows cornerBlack=$cornerBlack ringBlack=$ringBlack png=$pngPath"
$rows | ForEach-Object { Write-Output ("  " + $_) }
if ($status -eq "PASS") { exit 0 } else { exit 1 }
