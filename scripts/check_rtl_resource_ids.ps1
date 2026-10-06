# check_rtl_resource_ids.ps1 —— RTL 资源 id 的两份权威逐条对账（账 #225）
#
# 一条 RTL 文件要进产物，得在三处各写一次：
#   src/driver/c3rtl.rc            `<id> RCDATA "../rtl/.../<名字>"`   （资源编译器一侧的 id）
#   src/driver/rtl_embedded.hpp    `RTL_X = <id>,`                     （C++ 侧的 id 常量）
#   src/driver/rtl_embedded.cpp    `{ RTL_X, "<名字>" }`               （解包时写出的文件名）
# 三份里任意两份不齐，症状都不在编译期：`.rc` 与 `hpp` 对调时，解包会把 A 文件的内容
# 写成 B 文件名（实测 9a420157 那次：函数体被写成 vb6forms_draw.h）⇒ 每个 include 它的
# TU 都拿到一份定义 ⇒ 链接期一屏 LNK2005，措辞还指向错误的对象。所以这里逐条对账，
# 而不是只盯着新加的那一条。
#
# 判据面（任何一条不齐就红）：
#   R1 名字表（cpp）里每枚符号都在 hpp 有 id
#   R2 hpp 里每枚 id 都在 .rc 有条目
#   R3 .rc 那条目的文件名必须与名字表给该符号起的名字逐字符相同  ← 本账那条红
#   R4 .rc 里不许出现没被认领的 id；.rc / hpp / cpp 三处的条数必须相等
#   R5 .rc 里 id 不许重复

$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot
$rcPath  = Join-Path $Root 'src\driver\c3rtl.rc'
$hppPath = Join-Path $Root 'src\driver\rtl_embedded.hpp'
$cppPath = Join-Path $Root 'src\driver\rtl_embedded.cpp'

$viol = @()
foreach ($p in @($rcPath, $hppPath, $cppPath)) {
    if (-not (Test-Path -LiteralPath $p)) {
        Write-Host ("FAIL: 对账要读的文件不在: " + $p)
        exit 1
    }
}

$rcText  = [IO.File]::ReadAllText($rcPath)
$hppText = [IO.File]::ReadAllText($hppPath)
$cppText = [IO.File]::ReadAllText($cppPath)

# ---- .rc: id -> 文件名 ----
$rcNameOfId = @{}
$rcIdCount = 0
foreach ($m in [regex]::Matches($rcText, '(?m)^\s*(\d+)\s+RCDATA\s+"([^"]+)"')) {
    $rcIdCount++
    $id = [int]$m.Groups[1].Value
    $name = ($m.Groups[2].Value -replace '/', [char]92)
    $name = Split-Path -Leaf $name
    $key = [string]$id
    if ($rcNameOfId.ContainsKey($key)) {
        $viol += ("R5: c3rtl.rc 里 id {0} 出现两次（{1} 与 {2}）" -f $id, $rcNameOfId[$key], $name)
    } else {
        $rcNameOfId[$key] = $name
    }
}

# ---- hpp: 符号 -> id ----
$idOfSymbol = @{}
foreach ($m in [regex]::Matches($hppText, '(?m)\b(RTL_[A-Za-z0-9_]+)\s*=\s*(\d+)\s*,')) {
    $idOfSymbol[$m.Groups[1].Value] = [int]$m.Groups[2].Value
}

# ---- cpp: 符号 -> 解包后的文件名 ----
$nameOfSymbol = @{}
foreach ($m in [regex]::Matches($cppText, '(?m)\{\s*(RTL_[A-Za-z0-9_]+)\s*,\s*"([^"]+)"\s*\}')) {
    $nameOfSymbol[$m.Groups[1].Value] = $m.Groups[2].Value
}

# ---- R1 + R2 + R3 ----
foreach ($sym in ($nameOfSymbol.Keys | Sort-Object)) {
    $want = $nameOfSymbol[$sym]
    if (-not $idOfSymbol.ContainsKey($sym)) {
        $viol += ("R1: rtl_embedded.cpp 认领了 {0} -> {1}，但 hpp 里没有这枚 id 常量" -f $sym, $want)
        continue
    }
    $id = $idOfSymbol[$sym]
    $key = [string]$id
    if (-not $rcNameOfId.ContainsKey($key)) {
        $viol += ("R2: {0} 的 id {1} 在 c3rtl.rc 里没有 RCDATA 条目" -f $sym, $id)
        continue
    }
    $got = $rcNameOfId[$key]
    if ($got -ne $want) {
        $viol += ("R3: id {0} 两份权威对不上 —— c3rtl.rc 给的是 {1}，名字表要的是 {2}（符号 {3}）" -f $id, $got, $want, $sym)
    }
}

# ---- R4: 反向认领 + 条数 ----
$claimed = @($idOfSymbol.Values | ForEach-Object { [string]$_ })
$orphan = @($rcNameOfId.Keys | Where-Object { $claimed -notcontains $_ } | ForEach-Object { [int]$_ } | Sort-Object)
foreach ($o in $orphan) {
    $viol += ("R4: c3rtl.rc 的 id {0}（{1}）没被 hpp/cpp 认领 —— 加了 RTL 文件忘了登记" -f $o, $rcNameOfId[[string]$o])
}
$unbound = @($idOfSymbol.Keys | Where-Object { -not $nameOfSymbol.ContainsKey($_) } | Sort-Object)
foreach ($u in $unbound) {
    $viol += ("R4: hpp 有 {0} = {1}，但 rtl_embedded.cpp 的名字表里没登记文件名" -f $u, $idOfSymbol[$u])
}

Write-Host ("census: rc={0} hpp={1} cpp={2} orphanRcIds={3} unboundSymbols={4}" -f `
    $rcIdCount, $idOfSymbol.Count, $nameOfSymbol.Count, $orphan.Count, $unbound.Count)

if ($viol.Count -gt 0) {
    Write-Host ("FAIL: rtl-resource-id 对账不齐, {0} 条" -f $viol.Count)
    foreach ($v in $viol) { Write-Host ("  " + $v) }
    exit 1
}
Write-Host ("PASS: {0} 个 RTL 资源 id 三处逐条对上（头与体没有对调、无孤儿、无未登记）" -f $rcIdCount)
exit 0
