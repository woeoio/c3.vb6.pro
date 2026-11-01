# check_form_icon_resources.ps1 - 账 #305 第328刀：读一台 EXE 里的图标资源并按期望逐帧比。
#
# 为什么单独一份尺：窗体图标这一族的**判据在产物里**，不在发码里 —— VB6 把 IconForm 那张窗体
# `.frm` 的 `Icon = "<名>.frx":NNNN` 记录的每一帧原样嵌成 RT_ICON(3) 30001..，再拼一枚
# RT_GROUP_ICON(14) id=1（语料里 151 份能找回自己 VB6 产物的 EXE 工程：137/138 逐帧同、
# 81/81 「窗体没写 Icon」的那一档落 VB6 默认图 = 3 帧 304/744/296）。所以「对不对」只能
# 打开刚链接出来的 exe 数资源，`--emit-c` 那条路一个字都看不见。
#
# 用法：
#   check_form_icon_resources.ps1 -Exe <path> -Expect "30001=744,30002=296" [-ExpectGroupSize 34]
# 退出码：0 = 逐帧同（且组号=1）；1 = 任一格不符或读不动。
param(
    [Parameter(Mandatory = $true)][string]$Exe,
    [Parameter(Mandatory = $true)][string]$Expect,
    [int]$ExpectGroupSize = -1,
    [switch]$ShowHash
)
$ErrorActionPreference = 'Continue'
$viol = @()
if (-not (Test-Path -LiteralPath $Exe)) {
    Write-Host ("FAIL icon-resources: exe missing: " + $Exe)
    exit 1
}
$data = [IO.File]::ReadAllBytes($Exe)
if ($data.Length -lt 4096) { $viol += "exe too small to carry a resource directory" }

function U16([byte[]]$b, [int]$o) { return [int][BitConverter]::ToUInt16($b, $o) }
# 一律交回 [long]：目录项的两个 DWORD 都可能带高位旗子，[uint32] 参与 -band 时 PS 会按左操作数
# 类型转换掩码，结果就不是「抹掉那一位」了（实测：同一份 exe，不加 [long] 读出 size=0）。
function U32([byte[]]$b, [int]$o) { return [long][BitConverter]::ToUInt32($b, $o) }

$peOff = U32 $data 0x3C
if ((U32 $data $peOff) -ne 0x00004550) { $viol += "not a PE image" }
$numSec = U16 $data ($peOff + 6)
$sizeOpt = U16 $data ($peOff + 20)
$magic = U16 $data ($peOff + 24)
$ddOff = $peOff + 24 + $(if ($magic -eq 0x10b) { 0x60 } else { 0x70 })
$rsrcRva = U32 $data ($ddOff + 16)
$rsrcSize = U32 $data ($ddOff + 20)
if ($rsrcRva -eq 0) { $viol += "no resource directory in this exe" }

$secOff = $peOff + 24 + $sizeOpt
$rsrcAddr = -1
for ($i = 0; $i -lt $numSec; $i++) {
    $o = $secOff + $i * 40
    $nameBytes = $data[$o..($o + 7)]
    $name = (($nameBytes | ForEach-Object { [char]$_ }) -join '').Trim([char]0)
    if ($name -eq '.rsrc') {
        $va = U32 $data ($o + 12); $raw = U32 $data ($o + 20)
        $rsrcAddr = $raw + ($rsrcRva - $va)
        break
    }
}
if ($rsrcAddr -lt 0) { $viol += ".rsrc section not found" }

# 资源目录三层：类型 → 名字 → 语言。两件事都会咬人：
#   · 条目第二个 DWORD 的高位 0x80000000 是「下面还有目录」的旗子（**第一个** DWORD 的高位才是
#     「名字是字符串」）—— 抹错那个就会跑到 0x8000xxxx 的地址上；
#   · 这些偏移都是**相对 .rsrc 段首**的，直接当文件下标读会读到文件头 ⇒ 每一层都要 +$rsrcAddr。
function Entries([int]$off) {
    $named = U16 $data ($off + 12)
    $idcnt = U16 $data ($off + 14)
    $list = @()
    for ($k = 0; $k -lt ($named + $idcnt); $k++) {
        $e = $off + 16 + $k * 8
        $list += , @((U32 $data $e), (U32 $data ($e + 4)))
    }
    # 前缀逗号是判据的一部分：只有 1 个条目的目录（RT_GROUP_ICON 常常就一个 id、语言层常常一个档）
    # 裸 return 会被 PowerShell 拆成「那个条目本身」，foreach 于是枚举它的两个数 ⇒ 偏移读成 0。
    return ,$list
}
function NameAt([int]$off) {
    $len = U16 $data $off
    $chars = New-Object char[] $len
    for ($c = 0; $c -lt $len; $c++) { $chars[$c] = [char](U16 $data ($off + 2 + $c * 2)) }
    return (-join $chars)
}
$SUB = [long]0x80000000
$icons = @{}
$groups = @{}
if ($viol.Count -eq 0) {
    foreach ($t in (Entries $rsrcAddr)) {
        $tn = $t[0]; $td = $t[1] -band (-bnot $SUB)
        $typeId = if (($tn -band $SUB) -ne 0) { NameAt ($rsrcAddr + ($tn -band 0x7FFFFFFF)) } else { $tn -band 0xFFFF }
        if ($typeId -ne 3 -and $typeId -ne 14) { continue }
        $sink = if ($typeId -eq 3) { $icons } else { $groups }
        foreach ($n in (Entries ($rsrcAddr + $td))) {
            $nn = $n[0]; $nd = $n[1] -band (-bnot $SUB)
            $resId = if (($nn -band $SUB) -ne 0) { NameAt ($rsrcAddr + ($nn -band 0x7FFFFFFF)) } else { $nn -band 0xFFFF }
            foreach ($l in (Entries ($rsrcAddr + $nd))) {
                $ld = $l[1] -band (-bnot $SUB)
                $drva = U32 $data ($rsrcAddr + $ld)
                $dsz = U32 $data ($rsrcAddr + $ld + 4)
                $doff = $rsrcAddr + ($drva - $rsrcRva)
                $buf = New-Object byte[] $dsz
                [Array]::Copy($data, $doff, $buf, 0, $dsz)
                $sink[[string]$resId] = $buf
            }
        }
    }
}

$got = @()
foreach ($k in ($icons.Keys | Sort-Object { [int]$_ })) { $got += ($k + '=' + $icons[$k].Length) }
$gotText = $got -join ','
$wantText = (($Expect -split ',') | ForEach-Object { $_.Trim() }) -join ','
Write-Host ("  icon ids/sizes: got [" + $gotText + "] want [" + $wantText + "]")
if ($gotText -ne $wantText) { $viol += "RT_ICON set differs" }

if ($groups.Count -ne 1) {
    $viol += ("RT_GROUP_ICON count = " + $groups.Count + " (want exactly 1)")
} else {
    $gk = @($groups.Keys)[0]
    if ($gk -ne '1') { $viol += ("RT_GROUP_ICON id = " + $gk + " (want 1: Explorer 按最小组号取图)") }
    $gb = $groups['1']
    if ($ExpectGroupSize -ge 0 -and $gb.Length -ne $ExpectGroupSize) {
        $viol += ("group bytes length = " + $gb.Length + " (want " + $ExpectGroupSize + ")")
    }
    $nInGrp = (U16 $gb 4)
    if ($nInGrp -ne $icons.Count) { $viol += ("group lists " + $nInGrp + " frames, RT_ICON has " + $icons.Count) }
    for ($f = 0; $f -lt $nInGrp; $f++) {
        $o = 6 + $f * 14
        $rid = U16 $gb ($o + 12)
        $frSize = U32 $gb ($o + 8)
        if (-not $icons.ContainsKey([string]$rid)) { $viol += ("group frame " + $f + " names id " + $rid + " with no RT_ICON"); continue }
        if ($icons[[string]$rid].Length -ne $frSize) {
            $viol += ("group frame " + $f + " (id " + $rid + ") lists " + $frSize + " bytes, RT_ICON carries " + $icons[[string]$rid].Length)
        }
    }
    if ($ShowHash) {
        $md5 = [System.Security.Cryptography.MD5]::Create()
        Write-Host ("  group md5 " + ((-join ($md5.ComputeHash($gb) | ForEach-Object { $_.ToString('x2') }))))
        foreach ($k in ($icons.Keys | Sort-Object { [int]$_ })) {
            Write-Host ("  icon " + $k + " md5 " + ((-join ($md5.ComputeHash($icons[$k]) | ForEach-Object { $_.ToString('x2') }))))
        }
    }
}

if ($viol.Count -gt 0) {
    Write-Host ("FAIL icon-resources: " + $Exe)
    $viol | ForEach-Object { Write-Host ("   " + $_) }
    exit 1
}
Write-Host ("OK icon-resources: " + $Exe)
exit 0
