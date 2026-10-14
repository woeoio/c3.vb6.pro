# check_rtl_embedded.ps1 - 账 §B99② 的结构哨兵 (第 43 道；读的是**产物**，不是人写的规矩)
#
# 为什么要这一道 (RTL 是嵌在 C3.exe 的资源里的，见 rc 顶部与 memory rtl-reembed-c3rtl-rc-landmine):
#   `src/driver/c3rtl.rc` 用 125 条 `RCDATA` 把整棵 RTL 打进 C3.exe；用户工程编译时 RTL 是从这份
#   **内嵌资源**解包出来交给 cl 的。于是有一条安静的失效路径：改了 `src/rtl/**` 但增量重编时
#   `rc.exe` 没重跑 ⇒ C3.exe 里嵌的还是旧 RTL，编译器前端是新的、运行时是旧的 ——
#   实测过两次：#225 是 id 与文件对调（同一种资源两份权威），57b1f64e 是改了 3 处 RTL 而没 touch
#   `c3rtl.rc`（CI 从零重编看不出来，本机却会拿旧 RTL 测出一堆假读数）。
#   以前这一格靠「人记得去 touch 那个 .rc」；这一道把它换成**比对产物本身**：
#   把 exe 里每个 RCDATA 的字节读出来，与 rc 点名的那份磁盘文件逐字节比哈希。
#
# 规则 (改坏了会红):
#   R1 rc 里每一条 `<id> RCDATA "path"` 都必须在 exe 里找得到同名 id，且**内容哈希相同**
#      —— 覆盖「内嵌的是旧 RTL」与「id 对调」两族。
#   R2 exe 里多出来的 RCDATA id（rc 没登记的）也红。
#   R3 两边数量都有下限 (>=100)：不许「两边都空」假绿。
#
# 用法: powershell -File scripts/check_rtl_embedded.ps1 [-Root <检出根>]   (PASS => exit 0)

param([string]$Root = "")
$ErrorActionPreference = 'Stop'
$root = if ($Root) { (Resolve-Path $Root).Path } else { Split-Path -Parent $PSScriptRoot }
$exe = Join-Path $root '.build\C3.exe'
$rc = Join-Path $root 'src\driver\c3rtl.rc'
foreach ($p in @($exe, $rc)) {
    if (-not (Test-Path -LiteralPath $p)) {
        Write-Host ("FAIL missing " + $p + '  (先跑 scripts/build.bat 产出 C3.exe)')
        exit 1
    }
}
$fails = @()

# ---------- 1) 读 rc：id -> 仓库内路径 ----------
$rcText = [System.IO.File]::ReadAllText($rc)
$rx = [regex]'(?m)^\s*(\d+)\s+RCDATA\s+"([^"]+)"'
$want = @{}
foreach ($m in $rx.Matches($rcText)) {
    $want[$m.Groups[1].Value] = $m.Groups[2].Value
}

# ---------- 2) 解析 PE 的 RT_RCDATA ----------
$b = [System.IO.File]::ReadAllBytes($exe)
function U16([int]$o) { [BitConverter]::ToUInt16($b, $o) }
function U32([int]$o) { [BitConverter]::ToUInt32($b, $o) }
function Sha256Hex([byte[]]$bytes) {
    $h = New-Object System.Security.Cryptography.SHA256Managed
    try { (($h.ComputeHash($bytes) | ForEach-Object { $_.ToString('x2') }) -join '') } finally { $h.Dispose() }
}
$peOff = U32 0x3C
if ((U16 $peOff) -ne 0x4550) { Write-Host 'FAIL not a PE file (no PE\0\0 signature)'; exit 1 }
$numSec = U16 ($peOff + 6)
$optSize = U16 ($peOff + 20)
$optOff = $peOff + 24
# 数据目录的起点随可选头格式变：PE32 = 0x10b -> +96，PE32+ = 0x20b -> +112
$dirOff = $optOff + 112
if ((U16 $optOff) -eq 0x10b) { $dirOff = $optOff + 96 }
$rsrcRva = U32 ($dirOff + 16)          # 目录项 2 = IMAGE_DIRECTORY_ENTRY_RESOURCE
if ($rsrcRva -eq 0) { Write-Host 'FAIL the exe has no resource directory (nothing embedded at all?)'; exit 1 }
$secOff = $optOff + $optSize
$sec = @()
for ($s = 0; $s -lt $numSec; $s++) {
    $e = $secOff + $s * 40
    $sec += [pscustomobject]@{ va = (U32 ($e + 12)); vs = (U32 ($e + 8)); raw = (U32 ($e + 20)) }
}
function RvaToOff([uint32]$rva) {
    foreach ($x in $sec) {
        if ($rva -ge $x.va -and $rva -lt ($x.va + $x.vs)) { return [int]($x.raw + ($rva - $x.va)) }
    }
    throw ("RVA 0x{0:x} is not inside any section" -f $rva)
}
# 目录项一律返回**对象**列表：PowerShell 的输出流会把「数组套数组」托平，
# 单个条目时会把「一条目的两个字段」变成「两个条目」—— 实测在语言层就栽过一次。
function DirEntries([int]$off) {
    $named = U16 ($off + 12); $ints = U16 ($off + 14)
    $r = @()
    for ($i = 0; $i -lt ($named + $ints); $i++) {
        $e = $off + 16 + $i * 8
        $r += [pscustomobject]@{ id = (U32 $e); off = (U32 ($e + 4)) }
    }
    return $r
}
$base = RvaToOff $rsrcRva
$rcEntry = $null
foreach ($t in (DirEntries $base)) {
    if ([int]([uint32]$t.id -band 0x7FFFFFFF) -eq 10) { $rcEntry = $t; break }   # RT_RCDATA = 10
}
if ($null -eq $rcEntry) { Write-Host 'FAIL no RT_RCDATA type in the exe'; exit 1 }
$nameBase = $base + [int]([uint32]$rcEntry.off -band 0x7FFFFFFF)
$have = @{}
foreach ($n in (DirEntries $nameBase)) {
    $id = [string]([uint32]$n.id -band 0x7FFFFFFF)     # 整数名 = 资源 id
    $langBase = $base + [int]([uint32]$n.off -band 0x7FFFFFFF)
    foreach ($l in (DirEntries $langBase)) {
        $de = $base + [int]([uint32]$l.off -band 0x7FFFFFFF)   # 语言层叶子直指 DATA 条目
        $rva = U32 $de
        $size = [int](U32 ($de + 4))
        if ($rva -eq 0 -or $size -le 0) {
            $fails += ('R2/R1 id ' + $id + ' has an empty data entry in the exe (rva=0x' +
                       ('{0:x}' -f $rva) + ' size=' + $size + ')')
            continue
        }
        $o = RvaToOff $rva
        $slice = New-Object byte[] $size
        [Array]::Copy($b, $o, $slice, 0, $size)
        $have[$id] = @{ size = $size; hex = (Sha256Hex $slice) }
    }
}

# ---------- 3) 两边对齐 ----------
if ($want.Count -lt 100) { $fails += ('R3 c3rtl.rc lists only ' + $want.Count + ' RCDATA entries (expect >=100)') }
if ($have.Count -lt 100) { $fails += ('R3 the exe carries only ' + $have.Count + ' RCDATA ids (expect >=100)') }
$checked = 0
foreach ($id in ($want.Keys | Sort-Object { [int]$_ })) {
    $rel = $want[$id] -replace '/', '\'
    $path = Join-Path (Join-Path $root 'src\driver') $rel
    if (-not (Test-Path -LiteralPath $path)) { $fails += ('R1 id ' + $id + ' names a file missing on disk: ' + $rel); continue }
    $disk = [System.IO.File]::ReadAllBytes($path)
    if (-not $have.ContainsKey($id)) { $fails += ('R1 id ' + $id + ' (' + $rel + ') is NOT embedded in C3.exe at all'); continue }
    if ($have[$id].hex -ne (Sha256Hex $disk)) {
        $fails += ('R1 id ' + $id + ' (' + $rel + ') DIFFERS: disk=' + $disk.Length + 'B ' +
                   ((Sha256Hex $disk).Substring(0, 12)) + '  vs embedded=' + $have[$id].size + 'B ' +
                   ($have[$id].hex).Substring(0, 12) +
                   '  —— C3.exe 里的 RTL 与磁盘不同步：重跑 scripts/build.bat（必要时 touch c3rtl.rc 让 rc.exe 重来）')
    }
    $checked++
}
$orphan = @($have.Keys | Where-Object { -not $want.ContainsKey($_) } | Sort-Object { [int]$_ })
if ($orphan.Count -gt 0) { $fails += ('R2 exe carries ' + $orphan.Count + ' RCDATA id(s) absent from c3rtl.rc: ' + (($orphan | Select-Object -First 8) -join ',')) }

if ($fails.Count -gt 0) {
    $fails | Select-Object -First 12 | ForEach-Object { Write-Host "FAIL $_" }
    exit 1
}
Write-Host ('OK check_rtl_embedded: R1..R3 (' + $checked + ' RTL resources byte-identical between disk and C3.exe,' +
             ' no orphan ids; exe sha=' + (Sha256Hex $b).Substring(0, 12) + ')')
exit 0
