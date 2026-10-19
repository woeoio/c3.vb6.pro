# check_iface_view_authority.ps1 - 账 #305 §B221/#449（§B222）：「这一枚表达式交出来的是视图还是实例」
# 这一问只许有一位主人。三条钉：
#   ① 视图的成员调用发射形状 `->vt-><slot>(x, args)` 在全 src/backend 恰好 **1** 处，
#      且必须住在 module/cgen_iface_vtbl.cpp 的 ivrefVtableCall 里（com_bind 与「接收者是
#      视图值的调用」那一站都调它）。长出第 2 处 = 同一形状两位作者。
#   ② 问句 ifaceViewOfExpr 的**消费文件**恰好 **4** 份：Set 目标侧 / 实参侧 /
#      成员访问接收者 / 带实参调用的接收者。少一份 = 有一站回到各写一遍；多一份 = 判据被复制。
#   ③ 直接拿字面量 "vb6_ivref_" 拼类型/问档的文件恰好 **6** 份（今天在册的名单）。
#      第 7 份长出来就意味着「是不是接口」又有了第二位主人。
# 退出码：0 = 三条都过；1 = 任一条红。
param(
    [string]$Root = ""
)
$ErrorActionPreference = 'Continue'
if (-not $Root) { $Root = (Split-Path -Parent $PSScriptRoot) }
$backend = Join-Path $Root "src\backend"
if (-not (Test-Path $backend)) {
    Write-Host ("MISSING " + $backend)
    exit 2
}
$files = Get-ChildItem -Path $backend -Recurse -File |
    Where-Object { $_.Extension -match '^\.(cpp|inc|h|hpp)$' }

$fail = @()

# ① vtable 发射形状唯一作者
$vt = @($files | Select-String -Pattern '"->vt->"')
if ($vt.Count -ne 1) {
    $fail += ("vt-call sites = " + $vt.Count + " (want exactly 1)")
} elseif ($vt[0].Path -notmatch 'cgen_iface_vtbl\.cpp$') {
    $fail += ("vt-call formula moved into " + $vt[0].Path)
}

# ② ifaceViewOfExpr 的消费文件（去掉它自己的定义与声明）
$ivAll = @($files | Select-String -Pattern 'ifaceViewOfExpr')
$ivFiles = @($ivAll | ForEach-Object { Split-Path -Leaf $_.Path } | Sort-Object -Unique |
    Where-Object { $_ -ne 'cgen_util_classtype.cpp' -and $_ -ne 'cgen_helpers.inc' })
if ($ivFiles.Count -ne 4) {
    $fail += ("ifaceViewOfExpr consumer files = " + $ivFiles.Count +
              " (want exactly 4): " + ($ivFiles -join ', '))
}

# ③ 字面量 "vb6_ivref_" 的主人名单
$litFiles = @($files | Select-String -Pattern '"vb6_ivref_"' |
    ForEach-Object { Split-Path -Leaf $_.Path } | Sort-Object -Unique)
if ($litFiles.Count -ne 6) {
    $fail += ("vb6_ivref_ literal files = " + $litFiles.Count +
              " (want exactly 6): " + ($litFiles -join ', '))
}

if ($fail.Count -gt 0) {
    Write-Host "IFACE-VIEW-AUTHORITY FAIL"
    foreach ($f in $fail) { Write-Host ("  - " + $f) }
    exit 1
}
Write-Host "IFACE-VIEW-AUTHORITY OK (1 formula site / 4 consumers / 6 literal files)"
exit 0
