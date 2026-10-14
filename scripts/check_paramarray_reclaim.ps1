# check_paramarray_reclaim.ps1 - 账 #305 §B166 (第五十六刀): 一枚 ParamArray 往返的回收腿不许再没人管
#
# 为什么要这一道: 调用点每次调用都 vb6_PA_Create 现造一枚数组交进被调方, 而此前**全仓没有一声
# vb6_PA_Destroy**(那枚出口连 0 个调用者), For Each 的中间槽 _pa_vN 更是每轮裸覆盖 —— 夹具两台实测
# 每次调用漏 158/109 字节、中间槽每元素再漏 87/61 字节(存档 .build/k56_fix_base55_*.txt)。这一刀把两半
# 各接上一枚出口: 数组归**被调方统一出口**销毁(emitParamArrayReclaim), 中间槽归 §B164 那枚
# emitVariantSlotStore + 块尾收尾。这一道哨兵钉的就是「别改回没人收」。
#
# 五条规则:
#   P1  emitParamArrayReclaim: 定义恰好 1 处 / 声明恰好 1 处 / 调用点 >= $CallFloor (Sub+Function+Property 三处)
#   P1b 出口本体恰好 1 枚 emitLine 且发的就是 vb6_PA_Destroy(...) —— 空壳出口不算通过
#   P2  vb6_PA_Destroy 的发射只许住**定义出口的那份文件**, 恰好 1 处 (站点不许自己抄)
#   P3  vb6_PA_Create 的发射点在 src\backend 恰好 1 处 (出生地唯一, 配对才谈得上)
#   P4  For Each 的 ParamArray 那一站必须走 emitVariantSlotStore, 且裸形 _pa_v = ... 必须为 0
#   P5  块尾那一声收尾存在 (vb6_VariantClear(&_pa_v...))
param(
    [string]$Root = "",
    [int]$CallFloor = 3
)
$ErrorActionPreference = 'Continue'
if (-not $Root) { $Root = (Split-Path -Parent $PSScriptRoot) }
$backend = Join-Path $Root 'src\backend'
if (-not (Test-Path $backend)) { Write-Host ('FAIL 没有 ' + $backend); exit 2 }

$files = @(Get-ChildItem -Path $backend -Recurse -File -Include *.cpp, *.hpp, *.inc)
$bad = @()
$defFile = ''
$defBody = ''
$declCount = 0
$totalRefs = 0
$destroySites = @()
$createSites = @()
$barePA = 0

$DEF_RE    = '(?m)^\s*void\s+CCodeGen::emitParamArrayReclaim\s*\('
$DECL_RE   = '^\s*void\s+emitParamArrayReclaim\s*\('
$CALL_RE   = 'emitParamArrayReclaim\('
$DESTROY_RE = 'c_\.emitLine\("vb6_PA_Destroy\('
$CREATE_RE  = 'vb6_PA_Create\('
$STORE_RE   = 'emitVariantSlotStore\(\s*paTmp092q\s*,\s*paGet092q\s*,\s*true\s*\)'
$BARE_RE    = 'c_\.emitLine\(\s*paTmp092q\s*\+\s*" = "'
$TAIL_RE    = 'vb6_VariantClear\(&_pa_v'

foreach ($f in $files) {
    $t = [System.IO.File]::ReadAllText($f.FullName)
    foreach ($d in [regex]::Matches($t, $DESTROY_RE)) { $destroySites += $f.Name }
    foreach ($c in [regex]::Matches($t, $CREATE_RE)) { $createSites += $f.Name }
    foreach ($b in [regex]::Matches($t, $BARE_RE)) { $barePA++ }
    $refs = [regex]::Matches($t, $CALL_RE).Count
    if ($refs -eq 0) { continue }
    $totalRefs += $refs
    if ([regex]::IsMatch($t, $DECL_RE, 'Multiline')) { $declCount++ }
    $dm = [regex]::Matches($t, $DEF_RE)
    if ($dm.Count -gt 1) { $bad += ('P1 出口定义了 ' + $dm.Count + ' 遍 (' + $f.Name + ')') }
    if ($dm.Count -eq 1) {
        if ($defFile -ne '') { $bad += ('P1 出口定义落在两份文件: ' + $defFile + ' + ' + $f.Name) }
        $defFile = $f.FullName
        $i = $t.IndexOf($dm[0].Value)
        $j = $t.IndexOf([char]10 + '}', $i)
        $defBody = $t.Substring($i, [Math]::Max(0, $j - $i))
    }
}

if ($defFile -eq '') { $bad += 'P1 src\backend 里找不到 emitParamArrayReclaim 的定义 —— 那一问又没人管了' }
if ($declCount -ne 1) { $bad += ('P1 声明 found ' + $declCount + ' 处 (要恰好 1, 在 detail\util\cgen_helpers.inc)') }

$callSites = $totalRefs - $declCount - 1
if ($callSites -lt $CallFloor) {
    $bad += ('P1 调用点 ' + $callSites + ' < ' + $CallFloor + ' —— Sub/Function/Property 三个出口都要问这一枚')
}
if ($defBody -ne '') {
    $n = ([regex]::Matches($defBody, 'c_\.emitLine')).Count
    if ($n -ne 1) { $bad += ('P1b 出口本体有 ' + $n + ' 枚 emitLine (要恰好 1: 只发那一声销毁)') }
    if ($defBody.IndexOf('vb6_PA_Destroy(') -lt 0) { $bad += 'P1b 出口本体没发 vb6_PA_Destroy —— 空壳出口不算回收' }
}

$wantFile = Split-Path $defFile -Leaf
foreach ($g in ($destroySites | Group-Object)) {
    if ($g.Name -ne $wantFile -or $g.Count -ne 1) {
        $bad += ('P2 手抄销毁出现在 ' + $g.Name + ' x' + $g.Count + ' (只许住 ' + $wantFile + ' x1)')
    }
}
$gCreate = ($createSites | Measure-Object).Count
if ($gCreate -ne 1) { $bad += ('P3 vb6_PA_Create 的发射点有 ' + $gCreate + ' 处 (要恰好 1 —— 数组的出生地唯一)') }

$ctl = Join-Path $backend 'stmt\cgen_control.cpp'
if (Test-Path $ctl) {
    $tc = [System.IO.File]::ReadAllText($ctl)
    if ([regex]::Matches($tc, $STORE_RE).Count -ne 1) {
        $bad += 'P4 For Each 的 ParamArray 那一站没走 emitVariantSlotStore(paTmp092q, paGet092q, true)'
    }
    if ([regex]::Matches($tc, $TAIL_RE).Count -lt 1) {
        $bad += 'P5 块尾缺那一声收尾 (循环退出后 _pa_vN 手上还留着最后一份带权值)'
    }
}
if ($barePA -gt 0) { $bad += ('P4 站点裸发 _pa_v = ... ' + $barePA + ' 处 —— 中间槽每轮裸覆盖就是没人回收') }

if ($bad.Count -eq 0) {
    Write-Host ('PASS paramarray_reclaim: def 1 | calls ' + $callSites + ' | create 1 | destroy sites ' + `
        ($destroySites | Measure-Object).Count + ' | bare ' + $barePA + ' | violations 0') -ForegroundColor Green
    exit 0
}
foreach ($b in $bad) { Write-Host ('FAIL ' + $b) -ForegroundColor Red }
exit 1
