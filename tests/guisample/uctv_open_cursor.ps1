# uctv_open_cursor.ps1 - Tree Explorer「Open Cursors folder」用例（不许 SKIP）
#
# 行为期望（用户口径，与真 VB6 一致）：
#   点 Form2 的 Open Cursors folder → 展开 C:\Windows\ 到列表树，
#   并弹出 Shell 式路径框（这里是 MsgBox，标题 Proyecto1，静态文本 = 路径），
#   文本应为 C:\Windows\Cursors。
#
# 模态框会挂住脚本：本用例**必须**探测 #32770、截图存证、判定路径、WM_CLOSE 关掉，
# 不允许因"会阻塞"跳过。
#
# 抓屏口径（2026-11-07 用户要求，别再人工判断）：
#   抓图后抽样统计像素；若整幅同色/只有 <=3 种颜色，判定"被其它窗口压住、抓屏不可用"，
#   SetForegroundWindow + BringWindowToTop 后**重试一次**；仍不可用记 [occluded] 并继续。
#   （editor_scroll_px 那次红就是这条：桌面被资源管理器压住，重跑即 PASS。）
param(
  [string]$Exe = ".temp\ucb13\Proyecto1.exe",
  [string]$OutDir = ".temp\uctv_open_cursor"
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System; using System.Text; using System.Runtime.InteropServices;
public class TEW {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassNameW(IntPtr h, StringBuilder sb, int n);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowTextW(IntPtr h, StringBuilder sb, int n);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool PostMessageW(IntPtr h, uint m, IntPtr w, IntPtr l);
  [DllImport("user32.dll")] public static extern IntPtr SendMessageTimeoutW(IntPtr h, uint m, IntPtr w, IntPtr l, uint flags, uint timeout, out IntPtr res);
  [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint flags);
  [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr h);
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L,T,R,B; }
  public static string Cls(IntPtr h){ var b=new StringBuilder(256); GetClassNameW(h,b,256); return b.ToString(); }
  public static string Text(IntPtr h){ var b=new StringBuilder(512); GetWindowTextW(h,b,512); return b.ToString(); }
  public static string RealText(IntPtr h){
    IntPtr buf=Marshal.AllocHGlobal(4096);
    try { SendMessageTimeoutW(h,0x000D,(IntPtr)2048,buf,0x0002,1000,out _); return Marshal.PtrToStringUni(buf); }
    finally { Marshal.FreeHGlobal(buf); }
  }
  public static IntPtr[] TopWindows(uint pid){
    var list=new System.Collections.Generic.List<IntPtr>();
    EnumWindows((h,l)=>{ uint p; GetWindowThreadProcessId(h,out p); if(p==pid && IsWindowVisible(h)) list.Add(h); return true; }, IntPtr.Zero);
    return list.ToArray();
  }
  public static IntPtr[] Children(IntPtr p){
    var list=new System.Collections.Generic.List<IntPtr>();
    EnumChildWindows(p,(h,l)=>{ list.Add(h); return true; }, IntPtr.Zero);
    return list.ToArray();
  }
}
"@
# 抓一幅窗口位图；distinct<=3 视为"被压住/抓屏不可用"（PrintWindow 只交回一张底噪图）。
# 命中时把窗口顶到前台再抓一次，仍不可用就报 [occluded]（不静默当作通过）。
function Save-Win([IntPtr]$h,[string]$path){
  function Grab {
    $r = New-Object TEW+RECT
    [void][TEW]::GetWindowRect($h,[ref]$r)
    $w = $r.R-$r.L; $hh = $r.B-$r.T
    if($w -le 0 -or $hh -le 0){ return $null }
    $bmp = New-Object System.Drawing.Bitmap $w,$hh
    $g = [System.Drawing.Graphics]::FromImage($bmp); $dc = $g.GetHdc()
    [void][TEW]::PrintWindow($h,$dc,2)
    $g.ReleaseHdc($dc); $g.Dispose()
    return $bmp
  }
  function Distinct([System.Drawing.Bitmap]$b){
    $set = @{}
    for($y=0; $y -lt $b.Height; $y+=7){ for($x=0; $x -lt $b.Width; $x+=7){ $set[$b.GetPixel($x,$y).ToArgb()] = 1; if($set.Count -gt 3){ return $set.Count } } }
    return $set.Count
  }
  $bmp = Grab
  if($null -eq $bmp){ return $false }
  if((Distinct $bmp) -le 3){
    [void][TEW]::BringWindowToTop($h); [void][TEW]::SetForegroundWindow($h); Start-Sleep -Milliseconds 350
    $bmp.Dispose(); $bmp = Grab
    if($null -ne $bmp -and (Distinct $bmp) -le 3){
      Write-Output ("[occluded] 0x{0:X} 抓屏仍只有 <=3 种颜色, 记存证但不当作像素判据" -f $h.ToInt64())
    }
  }
  if($null -eq $bmp){ return $false }
  $bmp.Save($path,[System.Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()
  return $true
}
$exePath = (Resolve-Path $Exe).Path
if(-not (Test-Path $OutDir)){ New-Item -ItemType Directory -Path $OutDir | Out-Null }
$p = Start-Process -FilePath $exePath -WorkingDirectory (Split-Path $exePath) -PassThru
Write-Output ("[start] pid={0} exe={1}" -f $p.Id,$exePath)
function Find-ByCaption([uint32]$procId,[string]$cap){
  foreach($h in [TEW]::TopWindows($procId)){ if([TEW]::Text($h) -eq $cap){ return $h } }
  return [IntPtr]::Zero
}
function Find-Child([IntPtr]$parent,[string]$cap){
  foreach($h in [TEW]::Children($parent)){ if([TEW]::Text($h) -eq $cap){ return $h } }
  return [IntPtr]::Zero
}
$fail = 0
try {
  $main = [IntPtr]::Zero
  for($i=0;$i -lt 100;$i++){ Start-Sleep -Milliseconds 100; $main = Find-ByCaption $p.Id 'Form1'; if($main -ne [IntPtr]::Zero){ break } }
  if($main -eq [IntPtr]::Zero){ Write-Output '[fail] 主窗 Form1 未出现'; $fail = 2 }
  if($fail -eq 0){
    $btnTE = Find-Child $main 'Tree Explorer'
    if($btnTE -eq [IntPtr]::Zero){ Write-Output '[fail] 找不到 Tree Explorer'; $fail = 2 }
    else { [void][TEW]::PostMessageW($btnTE,0x00F5,[IntPtr]::Zero,[IntPtr]::Zero) }
  }
  $form2 = [IntPtr]::Zero
  if($fail -eq 0){
    for($i=0;$i -lt 60;$i++){ Start-Sleep -Milliseconds 100; $form2 = Find-ByCaption $p.Id 'Tree Explorer'; if($form2 -ne [IntPtr]::Zero){ break } }
    if($form2 -eq [IntPtr]::Zero){ Write-Output '[fail] Form2 (Tree Explorer) 未出现'; $fail = 2 }
  }
  if($fail -eq 0){
    Start-Sleep -Milliseconds 400
    [void](Save-Win $form2 (Join-Path $OutDir 'form2_before.png'))
    $btnOpen = Find-Child $form2 'Open Cursors folder'
    if($btnOpen -eq [IntPtr]::Zero){ Write-Output '[fail] 找不到 Open Cursors folder'; $fail = 2 }
    else { [void][TEW]::PostMessageW($btnOpen,0x00F5,[IntPtr]::Zero,[IntPtr]::Zero) }
  }
  if($fail -eq 0){
    $dialog = [IntPtr]::Zero; $crashed = $false
    for($i=0;$i -lt 80;$i++){
      Start-Sleep -Milliseconds 100
      if($p.HasExited){ $crashed = $true; break }
      foreach($h in [TEW]::TopWindows($p.Id)){ if([TEW]::Cls($h) -eq '#32770'){ $dialog = $h; break } }
      if($dialog -ne [IntPtr]::Zero){ break }
    }
    if($crashed){
      $p.WaitForExit(); Write-Output ("[crash] exit=0x{0:X8}" -f $p.ExitCode); $fail = 3
    } elseif($dialog -eq [IntPtr]::Zero){
      Write-Output '[fail] 未出现路径对话框 (80*100ms), 但进程存活'; $fail = 4
    } else {
      Write-Output ("[dialog] cls={0} title={1}" -f [TEW]::Cls($dialog),[TEW]::Text($dialog))
      [void](Save-Win $dialog (Join-Path $OutDir 'dialog.png'))
      $texts = @()
      foreach($h in [TEW]::Children($dialog)){ $texts += ("{0}:{1}" -f [TEW]::Cls($h),[TEW]::RealText($h)) }
      $joined = ($texts -join '|')
      if($joined -match 'C:\\Windows\\Cursors'){ Write-Output '[path] OK C:\Windows\Cursors' }
      else { Write-Output ("[path] MISMATCH {0}" -f $joined); $fail = 5 }
      [void][TEW]::PostMessageW($dialog,0x0010,[IntPtr]::Zero,[IntPtr]::Zero)
      $closed = $false
      for($i=0;$i -lt 30;$i++){ Start-Sleep -Milliseconds 100; if(-not [TEW]::IsWindow($dialog)){ $closed=$true; break } }
      Write-Output ("[close] closed={0} alive={1}" -f $closed,(-not $p.HasExited))
      if(-not $closed){ $fail = 6 }
      if(-not $p.HasExited){ [void](Save-Win $form2 (Join-Path $OutDir 'form2_after.png')) }
    }
  }
} finally { try{ $p.Kill() }catch{} }
if($fail -eq 0){ Write-Output '[PASS] Open Cursors folder' } else { Write-Output ("[FAIL] code={0}" -f $fail) }
exit $fail
