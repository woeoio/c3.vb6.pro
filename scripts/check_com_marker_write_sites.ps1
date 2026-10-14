# check_com_marker_write_sites.ps1 - 账 #260 的结构性哨兵 (只扫源码, 不起 cl)
#
# 为什么单开这一道: 「右值是 COM 成员读取时怎么解封」这个决定被抄成两份 (Fix 110i 与 P25b
# 各写一份 hint 表)，而第三处同类出口 —— With 块里的 COM 属性写 —— 一份都没问，于是
# `With w : .Count = d.Count` 把整枚读取丢掉、直接把 d 当值打包 (值永远错，编译与运行都不响)。
# 现在那张表住在一个函数里 (comMarkerValueForWrite)，写侧三处出口都问它。
#
# 规则 (改坏了会红, 不是装饰):
#   C1  comMarkerValueForWrite( 在 src/backend 里恰好 7 处 = 声明 + 定义 + **五个**写侧调用点
#       (SetProp 无索引写 / SetPropArg 带索引写 / 链式默认成员写 / 宿主伪对象与 Parent 写 /
#        With 块 COM 写 = 本账补上的那一处)；少一处 = 又有写侧出口没问它 (本账的形状)，
#       多一处 = 表开始被绕过/复制
#   C2  「packer → 解封类型」那张表只许有一份：六个 packer 档的 hint 赋值形在 src/backend 里
#       各恰好一次，且都在 comMarkerValueForWrite 体内 (改前它被抄了四份，其中一份漏 Boolean、
#       一份的默认档是 Variant —— 这两条差别就是本账的形状)
#   C3  With 那一支的调用必须在发 vb6_ComSetProp(... /* With COM SetProp */) 之前
#       (顺序错了就是把已清掉的标记当值用，等于没修)
#
# 用法:  pwsh -File scripts\check_com_marker_write_sites.ps1
# 退出码: 0 = 全绿; 1 = 红

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$bad = @()

$backend = Join-Path $root "src\backend"
if (-not (Test-Path -LiteralPath $backend)) {
    Write-Host ("FAIL C0 missing " + $backend) -ForegroundColor Red
    exit 1
}

function Get-BackendFiles {
    Get-ChildItem -LiteralPath $backend -Recurse -File |
        Where-Object { $_.Extension -in ".c", ".h", ".cpp", ".inc", ".hpp" }
}

$textOf = @{}
foreach ($f in Get-BackendFiles) { $textOf[$f.FullName] = [System.IO.File]::ReadAllText($f.FullName) }

# ---- C1: 五个站点 ----
$c1 = 0
foreach ($k in $textOf.Keys) { $c1 += @([regex]::Matches($textOf[$k], 'comMarkerValueForWrite\s*\(')).Count }
if ($c1 -ne 7) { $bad += ("C1 comMarkerValueForWrite call/decl/def sites = " + $c1 + " (need exactly 7: decl + def + 5 write-side consumers)") }

# ---- C2: hint 表只有一份 ----
foreach ($pair in @(@('vb6_ComPackDouble', 'Double'), @('vb6_ComPackInt', 'Long'),
                    @('vb6_ComPackBool', 'Long'), @('vb6_ComPackBSTR', 'BSTR'),
                    @('vb6_ComPackValue', 'Variant'), @('vb6_ComPackObject', 'Object'))) {
    $n = 0
    $pat = '"' + $pair[0] + '"\)\s+\w+ = "' + $pair[1] + '"'
    foreach ($k in $textOf.Keys) { $n += @([regex]::Matches($textOf[$k], $pat)).Count }
    if ($n -ne 1) { $bad += ("C2 the packer->unpack-type table is copied: hint for " + $pair[0] +
                             " assigned " + $n + " times (need exactly 1, inside comMarkerValueForWrite)") }
}

# ---- C3: With 那一支的顺序 ----
$propRel = "src\backend\detail\stmt\cgen_assign_prop_write.inc"
$propFile = Join-Path $root $propRel
if (-not (Test-Path -LiteralPath $propFile)) {
    Write-Host ("FAIL C3 missing " + $propFile) -ForegroundColor Red
    exit 1
}
$propText = [System.IO.File]::ReadAllText($propFile)
$iCall = $propText.IndexOf('comMarkerValueForWrite(packFn, valExpr);')
$iEmit = $propText.IndexOf('/* With COM SetProp */')
if ($iCall -lt 0) { $bad += "C3 the With COM-object write no longer asks comMarkerValueForWrite (that is the account 260 regression)" }
if ($iEmit -lt 0) { $bad += "C3 the With COM-object emit line is gone from the file the sentinel watches" }
if ($iCall -ge 0 -and $iEmit -ge 0 -and $iCall -gt $iEmit) {
    $bad += "C3 the marker is consumed AFTER the vb6_ComSetProp line is emitted -- the value packed there is still the bare object"
}

if ($bad.Count -ne 0) {
    foreach ($b in $bad) { Write-Host ("FAIL " + $b) -ForegroundColor Red }
    exit 1
}
Write-Host "PASS com_marker_write_sites (C1 7 sites = decl+def+5 consumers, C2 one hint table, C3 With branch consumes before emitting)"
