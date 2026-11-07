# uctv_coldclick.ps1 - 冷启动直点矩阵（问题 2 判据）
#
# 每个组合一个**独立进程**：冷启动 -> 只点目标按钮 -> 量该窗内容 -> 关窗。
# 不测顺序点击；要抓的是"第一次打开某窗时内容为空"这类共享状态依赖。
#
# 读数口径与 uctreeview_px.ps1 的 MetricsBuf 逐字一致 (textRows/vivid/photo/avatar)，
# 抓屏同样走 PrintWindow(2)。若整幅只有 <=3 种颜色 => 判"被其它窗口压住/抓屏不可用",
# 顶到前台重试一次，仍不可用记 [occluded] 并继续（不静默当通过）。
param(
  [Parameter(Mandatory=$true)][string]$Exe,
  [string]$OutDir = "",
  # 轮数：修后要求连续 3 轮可重复
  [int]$Rounds = 1,
  # 一个进程里按 $Order 依次点开（默认 0..4）；不给则每组合独立进程
  [switch]$SameProcess,
  [string]$Order = '0,1,2,3,4'
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System; using System.Text; using System.Runtime.InteropServices;
public class CK {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassNameW(IntPtr h, StringBuilder sb, int n);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowTextW(IntPtr h, StringBuilder sb, int n);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool GetClientRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool ClientToScreen(IntPtr h, ref POINT p);
  [DllImport("user32.dll")] public static extern bool PostMessageW(IntPtr h, uint m, IntPtr w, IntPtr l);
  [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint flags);
  [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr h);
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L,T,R,B; }
  [StructLayout(LayoutKind.Sequential)] public struct POINT { public int X,Y; }
  public static string Cls(IntPtr h){ var b=new StringBuilder(256); GetClassNameW(h,b,256); return b.ToString(); }
  public static string Text(IntPtr h){ var b=new StringBuilder(512); GetWindowTextW(h,b,512); return b.ToString(); }
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
  static bool IsGreenPx(int R,int G,int B){ return Math.Abs(R-34)<=20 && Math.Abs(G-170)<=25 && Math.Abs(B-85)<=25; }
  static bool IsGreenish(int R,int G,int B){ return G > R+20 && G > B+20; }
  public static string MetricsBuf(byte[] buf, int stride, int ox, int oy, int w, int hh) {
    int dark=0, vivid=0, white=0, photo=0, avatar=0, bg65=0;
    int textRows=0;
    bool inText=false;
    int quiet=0;
    for (int y = oy; y < oy + hh; y++) {
      int cd=0;
      for (int x = ox; x < ox + w; x++) {
        int bi = y * stride + x * 4;
        int B = buf[bi], G = buf[bi+1], R = buf[bi+2];
        int lum = (R*299 + G*587 + B*114)/1000;
        int mx = Math.Max(R, Math.Max(G, B)), mn = Math.Min(R, Math.Min(G, B));
        if (lum < 110) { dark++; cd++; }
        if (lum >= 200) white++;
        if (mx - mn >= 60) vivid++;
        if (mx-mn >= 40 && !IsGreenish(R,G,B) && mn >= 25) photo++;
        if (mn >= 100 && mx >= 180 && (mx-mn) >= 25 && !IsGreenish(R,G,B)) avatar++;
        if (Math.Abs(R-57)<=2 && Math.Abs(G-61)<=2 && Math.Abs(B-65)<=2) bg65++;
      }
      if (cd >= 3) { if (!inText) { inText = true; textRows++; quiet = 0; } else quiet = 0; }
      else if (inText) { quiet++; if (quiet >= 2) inText = false; }
    }
    return "textRows="+textRows+" dark="+dark+" vivid="+vivid+" white="+white+" photo="+photo+" avatar="+avatar+" bg65="+bg65;
  }
}
"@
if($OutDir -ne '' -and -not (Test-Path $OutDir)){ New-Item -ItemType Directory -Path $OutDir | Out-Null }
$exePath = (Resolve-Path $Exe).Path
$cases = @(
  @{ cap='Tree Explorer';       title='Tree Explorer' },
  @{ cap='Corporate Hierarchy'; title='Form3' },
  @{ cap='Drag and Drop';       title='Drag & Drop' },
  @{ cap='Menu / Tools';        title='Form5' },
  @{ cap='Contact  List';       title='Agenda tema oscuro' }
)
function Find-ByCaption([uint32]$procId,[string]$cap,[string]$exclude){
  foreach($h in [CK]::TopWindows($procId)){ $t=[CK]::Text($h); if($t -eq $cap -and $t -ne $exclude){ return $h } }
  return [IntPtr]::Zero
}
function Find-Child([IntPtr]$parent,[string]$cap){
  foreach($h in [CK]::Children($parent)){ if([CK]::Text($h) -eq $cap){ return $h } }
  return [IntPtr]::Zero
}
function Grab([IntPtr]$h){
  $r = New-Object CK+RECT; [void][CK]::GetWindowRect($h,[ref]$r)
  $w = $r.R-$r.L; $hh = $r.B-$r.T
  if($w -le 0 -or $hh -le 0){ return $null }
  $bmp = New-Object System.Drawing.Bitmap $w,$hh
  $g = [System.Drawing.Graphics]::FromImage($bmp); $dc = $g.GetHdc()
  [void][CK]::PrintWindow($h,$dc,2)
  $g.ReleaseHdc($dc); $g.Dispose(); return $bmp
}
function Distinct([System.Drawing.Bitmap]$b){
  $set = @{}
  for($y=0; $y -lt $b.Height; $y+=7){ for($x=0; $x -lt $b.Width; $x+=7){ $set[$b.GetPixel($x,$y).ToArgb()] = 1; if($set.Count -gt 3){ return $set.Count } } }
  return $set.Count
}
function Get-Metrics([System.Drawing.Bitmap]$b){
  $data = $b.LockBits((New-Object System.Drawing.Rectangle 0,0,$b.Width,$b.Height),
                      [System.Drawing.Imaging.ImageLockMode]::ReadOnly,
                      [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
  try {
    $bytes = New-Object byte[] ($data.Stride * $b.Height)
    [System.Runtime.InteropServices.Marshal]::Copy($data.Scan0,$bytes,0,$bytes.Length)
    return [CK]::MetricsBuf($bytes,$data.Stride,0,0,$b.Width,$b.Height)
  } finally { $b.UnlockBits($data) }
}
$fail = 0
# 点一次并量该窗；$CloseAfter=$true 时量完 WM_CLOSE。返回读数串，失败返回 $null。
function Test-One($proc,$cs,[string]$label,[bool]$CloseAfter){
  $main = [IntPtr]::Zero
  for($i=0;$i -lt 100;$i++){ Start-Sleep -Milliseconds 100; $main = Find-ByCaption $proc.Id 'Form1' ''; if($main -ne [IntPtr]::Zero){ break } }
  if($main -eq [IntPtr]::Zero){ Write-Host "$label launch=fail"; return $false }
  $btn = Find-Child $main $cs.cap
  if($btn -eq [IntPtr]::Zero){ Write-Host "$label button=fail"; return $false }
  [void][CK]::PostMessageW($btn,0x00F5,[IntPtr]::Zero,[IntPtr]::Zero)
  $win = [IntPtr]::Zero
  for($i=0;$i -lt 80;$i++){ Start-Sleep -Milliseconds 100; $win = Find-ByCaption $proc.Id $cs.title 'Form1'; if($win -ne [IntPtr]::Zero){ break } }
  if($win -eq [IntPtr]::Zero){ Write-Host "$label window=fail (title=$($cs.title))"; return $false }
  Start-Sleep -Milliseconds 700
  $bmp = Grab $win
  if($null -eq $bmp){ Write-Host "$label grab=fail"; return $false }
  if((Distinct $bmp) -le 3){
    [void][CK]::BringWindowToTop($win); [void][CK]::SetForegroundWindow($win); Start-Sleep -Milliseconds 350
    $bmp.Dispose(); $bmp = Grab $win
    if($null -ne $bmp -and (Distinct $bmp) -le 3){ Write-Host "$label [occluded]" }
  }
  if($null -eq $bmp){ Write-Host "$label grab=fail"; return $false }
  $st = Get-Metrics $bmp
  if($OutDir -ne ''){
    $safe = ("{0}_{1}" -f $label,$cs.title) -replace '[\\/:*?"<>|\[\] ]','_'
    $bmp.Save((Join-Path $OutDir ($safe + '.png')),[System.Drawing.Imaging.ImageFormat]::Png)
  }
  $bmp.Dispose()
  Write-Host ("{0} {1}" -f $label,$st)
  if($CloseAfter){ [void][CK]::PostMessageW($win,0x0010,[IntPtr]::Zero,[IntPtr]::Zero); Start-Sleep -Milliseconds 250 }
  return $true
}
if($SameProcess){
  $seq = $Order.Split(',') | ForEach-Object { [int]$_.Trim() }
  # 一个进程里按 $seq 依次点开（每个窗都留着），量的仍是"第一次打开"。
  for($round=1; $round -le $Rounds; $round++){
    $p = Start-Process -FilePath $exePath -WorkingDirectory (Split-Path $exePath) -PassThru
    try {
      foreach($idx in $seq){
        $ok = Test-One $p $cases[$idx] ("R$round S$idx [$($cases[$idx].cap)]") $false
        if(-not $ok){ $fail++ }
      }
    } finally { try{ if(-not $p.HasExited){ $p.Kill() } }catch{} }
  }
} else {
  for($round=1; $round -le $Rounds; $round++){
    foreach($cs in $cases){
      $p = Start-Process -FilePath $exePath -WorkingDirectory (Split-Path $exePath) -PassThru
      try { $ok = Test-One $p $cs ("R$round [$($cs.cap)]") $true; if(-not $ok){ $fail++ } }
      finally { try{ if(-not $p.HasExited){ $p.Kill() } }catch{} }
    }
  }
}
if($OutDir -ne ''){ Write-Output "[shots] $OutDir" }
if($fail -eq 0){ Write-Output '[DONE]' } else { Write-Output ("[FAIL] combos={0}" -f $fail) }
exit ([int]($fail -ne 0))
