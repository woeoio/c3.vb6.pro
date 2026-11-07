param(
    [string]$Fixture = "c:\Users\Administrator\Documents\c3.vb6.pro\tests\ucTreeView\Proyecto1.vbp",
    [string]$OutDir  = "c:\Users\Administrator\Documents\c3.vb6.pro\.temp\vb6ref\ucTreeView",
    [string]$VB6     = "C:\Program Files (x86)\VB6Mini\bin\VB6.EXE"
)
# uctv_vb6_baseline.ps1 — 用**真 VB6 /make** 产出 ucTreeView 基准 exe (五窗内容对照的判定基准).
#
# 必须在 pwsh (PowerShell 7) 下跑 (用户口径: 全流程 pwsh).
#
# 实测结论 (2026-11-07, pwsh 7.6 / VB6Mini):
#   A) 原字节副本 (Form1.frm / ucTreeViewPlus.ctl = cp1252, 无 BOM, CRLF, 西语字符)
#      → /make exit=1, Form1.log "行 416: 不能加载控件 ucTreeViewPlus1",
#         ucTreeViewPlus.log "行 31/50: 属性 SmoothScrollFactor 有无效值"
#   B) 同副本转 UTF-8 (无 BOM, 与用户机器上 D:\vb_yqt4qPac\ucTreeView 现状相同)
#      → 逐字相同的 exit=1 / 同样两条无效值 ⇒ **编码不是变量**, 两种都编不过.
#   根因是 `SmoothScrollFactor = 0,15` 的**逗号小数**: 该 .ctl 作者在逗号小数区域
#   (西语 Windows) 导出, 在 zh-CN (小数分隔符 ".") 的 VB6 里 `Single` 无法解析.
#   VB6 IDE 只当警告, 命令行 /make 当致命错误.
#
# 等价性依据 (不改语义, 只改**序列化**):
#   * ucScrollbar.ctl: `Public Property Let SmoothScrollFactor(ByVal SinValue As Single)`
#     ⇒ 类型 Single; 默认值 Read/WriteProperty 均为 0.3; Initialize 里赋 0.15;
#     `If factor > 1 Then factor = factor / 100` ⇒ 15 与 0.15 等价.
#   * 同一 .ctl 的 `FixDec(S) = Val(Replace(S, ",", "."))` 证明作者的运行时本来就把
#     逗号当小数点读 ⇒ `0,15` 的**本意就是 0.15**, 归一成 `0.15` 是同一个数
#     (IEEE-754 Single 位型一致), 不是"猜一个默认值".
#   * 更外一层: ucTreeViewPlus.ctl 的 UserControl_ReadProperties 先把内部两枚
#     ucScrollbar 的 SmoothScrollFactor 显式设为 1 (ucTreeViewPlus.ctl:2705/2710),
#     ⇒ 在本夹具里设计期值恒被覆盖, 对滚动行为无影响.
#
# 因此: 只 patch **副本**, 原夹具源码一字不动 (C3 侧输入不变, A/B 才干净);
# patch 走**原始字节**的 ASCII 替换 (不解码不重编码), 保证 cp1252 的西语字符
# (í/ú/ó/á/ñ) 逐字节保真.

$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSVersion.Major -lt 7) { throw "必须用 pwsh (PowerShell 7+), 当前 $($PSVersionTable.PSVersion)" }
if (-not (Test-Path $VB6)) { throw "缺 VB6: $VB6" }
if (-not (Test-Path $Fixture)) { throw "缺工程: $Fixture" }

Get-Process VB6 -ErrorAction SilentlyContinue | ForEach-Object { Write-Host "kill 残留 VB6 pid=$($_.Id)"; $_.Kill() }

if (Test-Path $OutDir) { Remove-Item -Recurse -Force $OutDir }
New-Item -ItemType Directory -Force (Split-Path $OutDir) | Out-Null
Copy-Item -Recurse (Split-Path $Fixture) $OutDir
$vbp = Join-Path $OutDir (Split-Path $Fixture -Leaf)

# 字节级替换: 0x3D 20 20 20 30 2C 31 35 ("=   0,15") -> "=   0.15"; 只动 ASCII 区,
# 非 ASCII 字节原样保留 (cp1252 的西语字符不会被转码).
$patched = @()
Get-ChildItem -Path (Join-Path $OutDir '*') -Include *.ctl,*.frm,*.cls -File | ForEach-Object {
    $b = [System.IO.File]::ReadAllBytes($_.FullName)
    $s = [System.Text.Encoding]::ASCII.GetString($b)   # ASCII 只做定位; 非 ASCII 字节保留在 $b
    $needle = '=   0,15'
    $n = 0
    $idx = $s.IndexOf($needle)
    if ($idx -ge 0) {
        $out = New-Object System.Collections.Generic.List[byte]
        $last = 0
        while ($idx -ge 0) {
            for ($i = $last; $i -lt $idx; $i++) { $out.Add($b[$i]) }
            foreach ($ch in [System.Text.Encoding]::ASCII.GetBytes('=   0.15')) { $out.Add($ch) }
            $last = $idx + $needle.Length
            $n++
            $idx = $s.IndexOf($needle, $last)
        }
        for ($i = $last; $i -lt $b.Length; $i++) { $out.Add($b[$i]) }
        [System.IO.File]::WriteAllBytes($_.FullName, $out.ToArray())
        $patched += "$($_.Name)x$n"
    }
}
Write-Host "patched(bytes, 0,15->0.15): $($patched -join ',')"

$log = Join-Path (Split-Path $OutDir) "vb6_make.log"
if (Test-Path $log) { Remove-Item -Force $log }
$p = Start-Process -FilePath $VB6 -WorkingDirectory $OutDir -ArgumentList @("/make", $vbp, "/out", $log) -Wait -PassThru
$exe = Join-Path $OutDir "Proyecto1.exe"
Write-Host "vb6 exit=$($p.ExitCode) exe=$(Test-Path $exe)"
if (-not (Test-Path $exe)) {
    Write-Host "--- /out 日志原文 ---"
    if (Test-Path $log) { [System.IO.File]::ReadAllText($log, [System.Text.Encoding]::GetEncoding(936)) }
    Get-ChildItem $OutDir -Filter *.log -File | Where-Object { $_.Name -ne (Split-Path $log -Leaf) } | ForEach-Object {
        Write-Host "--- $($_.Name) ---"
        [System.IO.File]::ReadAllText($_.FullName, [System.Text.Encoding]::GetEncoding(936))
    }
    exit 1
}
$fi = Get-Item $exe
Write-Host ("OK {0} size={1} mtime={2}" -f $exe, $fi.Length, $fi.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'))
exit 0
