param(
    [Parameter(Mandatory=$true)][string]$Exe,
    [string]$OutDir = "",
    [int]$Sec = 4
)
# 内置 TChart 的导出/剪贴板判据（tests\ctrlteechart\TcExportApp.vbp）。
#
# 夹具在 Form_Load 里做两件事：Export.SaveToJPEGFile "tc_export.jpg"、
# Export.CopyToClipboard。判据分别回头看这两样东西：
#   ① JPEG：SOI/EOI 完整、能解码、尺寸 401x257、柱色像素够多（不是白板）
#   ② 剪贴板：CF_BITMAP 存在且尺寸 401x257；像素从同一份剪贴板里的 CF_DIB 取样
# 只查“不报错”不够 —— 空实现也不报错，所以必须查产物内容。
#
# 两个写错过的地方（都在这里留痕）：
#   · CF_BITMAP = 2，不是 3（3 是 CF_METAFILEPICT）；第一版查 3 直接判“没复制”。
#   · 采样走 CF_DIB 而不是 GetDIBits(CF_BITMAP)：剪贴板位图是 DDB，
#     GetDIBits 在它身上返回 0 行（而且它不能同时选在 DC 里）。

Add-Type -AssemblyName System.Drawing
Add-Type @"
using System; using System.Runtime.InteropServices;
public class ExpPx {
  [DllImport("user32.dll")] public static extern bool OpenClipboard(IntPtr h);
  [DllImport("user32.dll")] public static extern bool CloseClipboard();
  [DllImport("user32.dll")] public static extern IntPtr GetClipboardData(uint fmt);
  [DllImport("user32.dll")] public static extern bool IsClipboardFormatAvailable(uint fmt);
  [DllImport("gdi32.dll")] public static extern int GetObject(IntPtr h, int cb, ref BITMAP bm);
  [DllImport("kernel32.dll")] public static extern IntPtr GlobalLock(IntPtr h);
  [DllImport("kernel32.dll")] public static extern bool GlobalUnlock(IntPtr h);
  [DllImport("kernel32.dll")] public static extern UIntPtr GlobalSize(IntPtr h);
  [StructLayout(LayoutKind.Sequential)] public struct BITMAP { public int bmType, bmWidth, bmHeight, bmWidthBytes; public ushort bmPlanes, bmBitsPixel; public IntPtr bmBits; }
}
"@

$exePath = (Resolve-Path $Exe -ErrorAction SilentlyContinue)
if (-not $exePath) { Write-Output "FAIL 缺产物 ($Exe 不存在)"; exit 1 }
$exePath = $exePath.Path
$wd = Split-Path $exePath
if (-not $OutDir) { $OutDir = Join-Path $wd "tcexp" }
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

$jpg = Join-Path $wd "tc_export.jpg"
if (Test-Path $jpg) { Remove-Item $jpg -Force }

$p = Start-Process -FilePath $exePath -WorkingDirectory $wd -PassThru `
     -RedirectStandardOutput (Join-Path $OutDir "tcexp.out") `
     -RedirectStandardError (Join-Path $OutDir "tcexp.err")
Start-Sleep -Seconds $Sec

$issues = New-Object System.Collections.Generic.List[string]
$read   = New-Object System.Collections.Generic.List[string]

# ---------- ① JPEG ----------
if (-not (Test-Path $jpg)) {
    $issues.Add("不存在 tc_export.jpg —— SaveToJPEGFile 没落盘")
} else {
    $bytes = [IO.File]::ReadAllBytes($jpg)
    $read.Add("jpg bytes=$($bytes.Length)")
    if ($bytes.Length -lt 1024) { $issues.Add("JPEG 只有 $($bytes.Length) 字节，太小") }
    if ($bytes[0] -ne 0xFF -or $bytes[1] -ne 0xD8) { $issues.Add("JPEG 缺 SOI (FFD8)") }
    if ($bytes[$bytes.Length - 2] -ne 0xFF -or $bytes[$bytes.Length - 1] -ne 0xD9) { $issues.Add("JPEG 缺 EOI (FFD9) —— 写入被截断") }
    try {
        $img = [System.Drawing.Image]::FromFile($jpg)
        $read.Add("jpg size=$($img.Width)x$($img.Height)")
        if ($img.Width -ne 401 -or $img.Height -ne 257) { $issues.Add("JPEG 尺寸 $($img.Width)x$($img.Height) != 401x257") }
        $bmp = New-Object System.Drawing.Bitmap($img)
        $img.Dispose()
        $hit = 0
        for ($y = 0; $y -lt $bmp.Height; $y += 2) {
            for ($x = 0; $x -lt 300; $x += 2) {
                $c = $bmp.GetPixel($x, $y)
                if ([Math]::Abs($c.R - 0) -le 24 -and [Math]::Abs($c.G - 128) -le 24 -and [Math]::Abs($c.B - 255) -le 24) { $hit++ }
            }
        }
        $read.Add("jpg barpx=$hit")
        # 有损压缩：按 1/2 抽样后仍应有 >500 个柱色点（整幅时约 1.2 万）
        if ($hit -lt 500) { $issues.Add("JPEG 里柱色像素只 $hit 个 —— 导出的是空白图") }
        $bmp.Dispose()
    } catch {
        $issues.Add("JPEG 解不开：" + $_.Exception.Message)
    }
}

# ---------- ② 剪贴板 ----------
if (-not [ExpPx]::IsClipboardFormatAvailable(2)) {
    $issues.Add("剪贴板没有 CF_BITMAP —— CopyToClipboard 没动作")
} elseif ([ExpPx]::OpenClipboard([IntPtr]::Zero)) {
    try {
        $h = [ExpPx]::GetClipboardData(2)
        if ($h -eq [IntPtr]::Zero) {
            $issues.Add("GetClipboardData(CF_BITMAP) 返回空")
        } else {
            $bm = New-Object ExpPx+BITMAP
            $ok = [ExpPx]::GetObject($h, [System.Runtime.InteropServices.Marshal]::SizeOf($bm), [ref]$bm)
            if ($ok -eq 0) {
                $issues.Add("GetObject(BITMAP) 失败")
            } else {
                # 像素取样走同一份剪贴板里的 CF_DIB
                $hdib = if ([ExpPx]::IsClipboardFormatAvailable(8)) { [ExpPx]::GetClipboardData(8) } else { [IntPtr]::Zero }
                $hit2 = -1
                if ($hdib -ne [IntPtr]::Zero) {
                    $ptr = [ExpPx]::GlobalLock($hdib)
                    if ($ptr -ne [IntPtr]::Zero) {
                        $hdr = New-Object byte[] 40
                        [System.Runtime.InteropServices.Marshal]::Copy($ptr, $hdr, 0, 40)
                        $dibW = [BitConverter]::ToInt32($hdr, 4)
                        $dibH = [BitConverter]::ToInt32($hdr, 8)
                        $dibBpp = [BitConverter]::ToUInt16($hdr, 14)
                        $bottomUp = $dibH -gt 0
                        $ah = [Math]::Abs($dibH)
                        $stride = [int](($dibW * $dibBpp / 8 + 3) -band -bnot 3)
                        $hit2 = 0
                        if ($dibBpp -ge 24 -and $ah -eq $bm.bmHeight -and $dibW -eq $bm.bmWidth) {
                            $row = New-Object byte[] $stride
                            for ($y = 40; $y -lt 220; $y += 4) {
                                $srcRow = if ($bottomUp) { $ah - 1 - $y } else { $y }
                                [System.Runtime.InteropServices.Marshal]::Copy([IntPtr]($ptr.ToInt64() + 40 + $srcRow * $stride), $row, 0, $stride)
                                for ($x = 30; $x -lt 300; $x += 4) {
                                    $o = $x * ($dibBpp / 8)
                                    $b = [int]$row[$o]; $g = [int]$row[$o + 1]; $r = [int]$row[$o + 2]
                                    if ([Math]::Abs($r - 0) -le 8 -and [Math]::Abs($g - 128) -le 8 -and [Math]::Abs($b - 255) -le 8) { $hit2++ }
                                }
                            }
                        }
                        [ExpPx]::GlobalUnlock($hdib) | Out-Null
                    }
                }
                $read.Add("clip=$($bm.bmWidth)x$($bm.bmHeight) barpx=$hit2")
                if ($bm.bmWidth -ne 401 -or $bm.bmHeight -ne 257) {
                    $issues.Add("剪贴板位图 $($bm.bmWidth)x$($bm.bmHeight) != 401x257")
                }
                if ($hit2 -lt 200) { $issues.Add("剪贴板位图里柱色像素只 $hit2 个（<200）—— 拷的是空图或 DIB 读不出来") }
            }
        }
    } finally {
        [ExpPx]::CloseClipboard() | Out-Null
    }
} else {
    $issues.Add("OpenClipboard 失败（被别的进程占用？）")
}

# 夹具日志里的 marker
$outFile = Join-Path $OutDir "tcexp.out"
if (-not $p.HasExited) { try { $p.Kill() } catch {} }
Start-Sleep -Milliseconds 300
if (Test-Path $outFile) {
    $txt = [Text.Encoding]::GetEncoding(936).GetString([IO.File]::ReadAllBytes($outFile))
    if ($txt -notmatch 'TC-EXP-DONE') { $issues.Add("夹具没走到 TC-EXP-DONE") }
}

if ($issues.Count -gt 0) {
    Write-Output ("FAIL tc_export_px " + ($read -join " | "))
    foreach ($it in $issues) { Write-Output ("  " + $it) }
    exit 1
}
Write-Output ("PASS tc_export_px " + ($read -join " | "))
exit 0
